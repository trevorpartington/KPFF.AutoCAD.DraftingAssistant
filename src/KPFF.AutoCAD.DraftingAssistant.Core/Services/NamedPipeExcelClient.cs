using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using Microsoft.Extensions.Logging;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Excel reader client that communicates with external Excel reader process via named pipes
/// This eliminates EPPlus from AutoCAD process to prevent freezing issues
/// </summary>
public class NamedPipeExcelClient : IExcelReader
{
    private const string PIPE_NAME = "KPFF_AutoCAD_ExcelReader";
    private readonly IApplicationLogger _logger;
    private Process? _excelReaderProcess;
    private bool _isConnected;

    public NamedPipeExcelClient(IApplicationLogger logger)
    {
        _logger = logger;
    }

    public async Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
    {
        var request = new ExcelRequest
        {
            Operation = "readsheetindex",
            FilePath = filePath
        };

        var response = await SendRequestAsync(request);
        if (!response.Success || response.Data == null)
        {
            _logger.LogError($"Failed to read sheet index: {response.Error}");
            return new List<SheetInfo>();
        }

        try
        {
            var jsonElement = (JsonElement)response.Data;
            var sheets = JsonSerializer.Deserialize<List<SheetInfo>>(jsonElement.GetRawText());
            return sheets ?? new List<SheetInfo>();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error deserializing sheet index: {ex.Message}");
            return new List<SheetInfo>();
        }
    }

    public async Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config)
    {
        _logger.LogInformation($"Starting to read construction notes for series {series} from {filePath}");
        
        var request = new ExcelRequest
        {
            Operation = "readconstructionnotes",
            FilePath = filePath,
            Parameters = { { "series", series } }
        };

        _logger.LogInformation($"Sending request with operation: {request.Operation}");
        var response = await SendRequestAsync(request);
        
        _logger.LogInformation($"Response received - Success: {response.Success}, Data null: {response.Data == null}, Error: {response.Error}");
        
        if (!response.Success || response.Data == null)
        {
            _logger.LogError($"Failed to read construction notes for series {series}: {response.Error}");
            return new List<ConstructionNote>();
        }

        try
        {
            _logger.LogInformation("Deserializing construction notes from response data...");
            var jsonElement = (JsonElement)response.Data;
            var notes = JsonSerializer.Deserialize<List<ConstructionNote>>(jsonElement.GetRawText());
            _logger.LogInformation($"Successfully deserialized {notes?.Count ?? 0} construction notes");
            return notes ?? new List<ConstructionNote>();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error deserializing construction notes: {ex.Message}");
            _logger.LogError($"Exception type: {ex.GetType().Name}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");
            return new List<ConstructionNote>();
        }
    }

    public async Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config)
    {
        var request = new ExcelRequest
        {
            Operation = "readexcelnotes",
            FilePath = filePath
        };

        var response = await SendRequestAsync(request);
        if (!response.Success || response.Data == null)
        {
            _logger.LogError($"Failed to read excel notes: {response.Error}");
            return new List<SheetNoteMapping>();
        }

        try
        {
            var jsonElement = (JsonElement)response.Data;
            var mappings = JsonSerializer.Deserialize<List<SheetNoteMapping>>(jsonElement.GetRawText());
            return mappings ?? new List<SheetNoteMapping>();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error deserializing excel notes: {ex.Message}");
            return new List<SheetNoteMapping>();
        }
    }

    public async Task<bool> FileExistsAsync(string filePath)
    {
        return File.Exists(filePath);
    }

    public async Task<string[]> GetWorksheetNamesAsync(string filePath)
    {
        // For now, return empty array - can implement later if needed
        return Array.Empty<string>();
    }

    public async Task<string[]> GetTableNamesAsync(string filePath, string worksheetName)
    {
        // For now, return empty array - can implement later if needed  
        return Array.Empty<string>();
    }

    private async Task<ExcelResponse> SendRequestAsync(ExcelRequest request)
    {
        try
        {
            _logger.LogInformation($"Sending request: {request.Operation} for file: {request.FilePath}");
            await EnsureExcelReaderRunningAsync();

            using var client = new NamedPipeClientStream(".", PIPE_NAME, PipeDirection.InOut);
            
            _logger.LogInformation("Connecting to named pipe...");
            // Connect with timeout
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10)); // 10 second timeout
            await client.ConnectAsync(cts.Token);
            _logger.LogInformation("Connected to named pipe successfully");

            // Send request
            _logger.LogInformation("Sending request message...");
            await SendMessageAsync(client, request);
            _logger.LogInformation("Request message sent successfully");

            // Receive response with timeout
            _logger.LogInformation("Waiting for response...");
            using var responseCts = new CancellationTokenSource(TimeSpan.FromSeconds(30)); // 30 second timeout for response
            var response = await ReceiveMessageAsync(client, responseCts.Token);
            _logger.LogInformation($"Response received: Success={response.Success}, Error={response.Error}");
            return response;
        }
        catch (OperationCanceledException)
        {
            _logger.LogError("Operation timed out while communicating with Excel reader");
            return new ExcelResponse 
            { 
                Success = false, 
                Error = "Operation timed out" 
            };
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error communicating with Excel reader: {ex.Message}");
            _logger.LogError($"Exception type: {ex.GetType().Name}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");
            
            // Try to restart process once
            await RestartExcelReaderAsync();
            
            return new ExcelResponse 
            { 
                Success = false, 
                Error = $"Communication error: {ex.Message}" 
            };
        }
    }

    private async Task SendMessageAsync(NamedPipeClientStream pipe, ExcelRequest request)
    {
        var requestJson = JsonSerializer.Serialize(request);
        var requestBytes = Encoding.UTF8.GetBytes(requestJson);
        var lengthBytes = BitConverter.GetBytes(requestBytes.Length);

        _logger.LogInformation($"Sending message - Length: {requestBytes.Length} bytes, Content: '{requestJson}'");

        // Send length first, then content
        await pipe.WriteAsync(lengthBytes, 0, 4);
        await pipe.WriteAsync(requestBytes, 0, requestBytes.Length);
        await pipe.FlushAsync();
        
        _logger.LogInformation("Message sent and flushed");
    }

    private async Task<ExcelResponse> ReceiveMessageAsync(NamedPipeClientStream pipe, CancellationToken token)
    {
        // Read message length first (4 bytes)
        var lengthBuffer = new byte[4];
        var bytesRead = await pipe.ReadAsync(lengthBuffer, 0, 4, token);
        if (bytesRead != 4)
            throw new InvalidOperationException("Failed to read message length");

        var messageLength = BitConverter.ToInt32(lengthBuffer, 0);
        if (messageLength <= 0 || messageLength > 1024 * 1024) // Max 1MB
            throw new InvalidOperationException($"Invalid message length: {messageLength}");

        // Read message content
        var messageBuffer = new byte[messageLength];
        var totalBytesRead = 0;
        
        while (totalBytesRead < messageLength)
        {
            bytesRead = await pipe.ReadAsync(messageBuffer, totalBytesRead, messageLength - totalBytesRead, token);
            if (bytesRead == 0) break;
            totalBytesRead += bytesRead;
        }

        var responseJson = Encoding.UTF8.GetString(messageBuffer, 0, totalBytesRead);
        return JsonSerializer.Deserialize<ExcelResponse>(responseJson) ?? 
               new ExcelResponse { Success = false, Error = "Failed to deserialize response" };
    }

    private async Task EnsureExcelReaderRunningAsync()
    {
        if (_excelReaderProcess != null && !_excelReaderProcess.HasExited)
            return;

        // CRASH FIX: Clean up any existing Excel reader processes to prevent accumulation
        CleanupExistingExcelReaderProcesses();

        await StartExcelReaderAsync();
    }

    /// <summary>
    /// Cleans up any existing Excel reader processes to prevent accumulation
    /// </summary>
    private void CleanupExistingExcelReaderProcesses()
    {
        try
        {
            var existingProcesses = Process.GetProcessesByName("KPFF.AutoCAD.ExcelReader");
            if (existingProcesses.Length > 0)
            {
                _logger.LogInformation($"Found {existingProcesses.Length} existing Excel reader processes - cleaning up");
                foreach (var process in existingProcesses)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            process.Kill();
                            process.WaitForExit(3000); // Wait up to 3 seconds
                        }
                        process.Dispose();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Error cleaning up process {process.Id}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error during process cleanup: {ex.Message}");
        }
    }

    private async Task StartExcelReaderAsync()
    {
        try
        {
            // Try multiple possible locations for the Excel reader executable
            var excelReaderPath = FindExcelReaderExecutable();
            
            if (string.IsNullOrEmpty(excelReaderPath))
            {
                _logger.LogError("Excel reader executable not found in any expected location");
                throw new FileNotFoundException("Excel reader executable not found");
            }

            _logger.LogInformation($"Starting Excel reader: {excelReaderPath}");

            _excelReaderProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = excelReaderPath,
                    UseShellExecute = false,
                    CreateNoWindow = true, // Hide console window - not needed for production
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                }
            };

            // CRASH FIX: Capture console output for debugging
            _excelReaderProcess.OutputDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    _logger.LogInformation($"Excel Reader Console: {e.Data}");
                }
            };

            _excelReaderProcess.ErrorDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    _logger.LogError($"Excel Reader Error: {e.Data}");
                }
            };

            _excelReaderProcess.Start();
            
            // Start reading output streams
            _excelReaderProcess.BeginOutputReadLine();
            _excelReaderProcess.BeginErrorReadLine();
            
            // Wait a moment for the process to start
            await Task.Delay(2000);

            // CRASH FIX: Don't test connection during startup to avoid circular dependency
            // The connection will be tested when the first actual request is made
            _logger.LogInformation("Excel reader started successfully - connection will be tested on first use");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to start Excel reader: {ex.Message}");
            throw;
        }
    }

    private string? FindExcelReaderExecutable()
    {
        var possiblePaths = new List<string>();
        
        // 1. Try the current assembly directory (Core assembly location)
        var currentAssemblyDir = Path.GetDirectoryName(typeof(NamedPipeExcelClient).Assembly.Location);
        if (!string.IsNullOrEmpty(currentAssemblyDir))
        {
            possiblePaths.Add(Path.Combine(currentAssemblyDir, "KPFF.AutoCAD.ExcelReader.exe"));
        }
        
        // 2. Try the Plugin assembly directory (where the executable should be)
        var pluginAssembly = AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(a => a.GetName().Name?.Contains("Plugin") == true);
        if (pluginAssembly != null)
        {
            var pluginDir = Path.GetDirectoryName(pluginAssembly.Location);
            if (!string.IsNullOrEmpty(pluginDir))
            {
                possiblePaths.Add(Path.Combine(pluginDir, "KPFF.AutoCAD.ExcelReader.exe"));
            }
        }
        
        // 3. Try the AutoCAD plugin directory (common location)
        var acadPluginDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "Autodesk", "AutoCAD", "R25.0", "R25.0", "enu", "Support");
        possiblePaths.Add(Path.Combine(acadPluginDir, "KPFF.AutoCAD.ExcelReader.exe"));
        
        // 4. Try the current working directory
        possiblePaths.Add(Path.Combine(Directory.GetCurrentDirectory(), "KPFF.AutoCAD.ExcelReader.exe"));
        
        // 5. Try the executable directory
        var exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        if (!string.IsNullOrEmpty(exeDir))
        {
            possiblePaths.Add(Path.Combine(exeDir, "KPFF.AutoCAD.ExcelReader.exe"));
        }

        // Check each possible path
        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                _logger.LogInformation($"Found Excel reader at: {path}");
                return path;
            }
        }
        
        // Log all attempted paths for debugging
        _logger.LogWarning("Excel reader not found. Searched in:");
        foreach (var path in possiblePaths)
        {
            _logger.LogWarning($"  - {path}");
        }
        
        return null;
    }

    private async Task RestartExcelReaderAsync()
    {
        try
        {
            _excelReaderProcess?.Kill();
            _excelReaderProcess?.Dispose();
            _excelReaderProcess = null;
            
            await StartExcelReaderAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to restart Excel reader: {ex.Message}");
        }
    }

    private async Task TestConnectionAsync()
    {
        var pingRequest = new ExcelRequest { Operation = "ping" };
        var response = await SendRequestAsync(pingRequest);
        
        if (!response.Success)
            throw new InvalidOperationException($"Excel reader ping failed: {response.Error}");
    }

    /// <summary>
    /// Manually test the connection to the Excel reader process
    /// This can be called separately to verify communication
    /// </summary>
    public async Task<bool> TestConnectionManuallyAsync()
    {
        try
        {
            var pingRequest = new ExcelRequest { Operation = "ping" };
            var response = await SendRequestAsync(pingRequest);
            return response.Success;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Manual connection test failed: {ex.Message}");
            return false;
        }
    }

    public void Dispose()
    {
        try
        {
            if (_excelReaderProcess != null && !_excelReaderProcess.HasExited)
            {
                _logger.LogInformation("Terminating Excel reader process");
                _excelReaderProcess.Kill();
                _excelReaderProcess.WaitForExit(5000); // Wait up to 5 seconds
            }
            _excelReaderProcess?.Dispose();
            _excelReaderProcess = null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error disposing Excel reader: {ex.Message}");
        }
    }
}

// Request/Response models for communication
public class ExcelRequest
{
    [System.Text.Json.Serialization.JsonPropertyName("operation")]
    public string Operation { get; set; } = string.Empty;
    
    [System.Text.Json.Serialization.JsonPropertyName("filePath")]
    public string FilePath { get; set; } = string.Empty;
    
    [System.Text.Json.Serialization.JsonPropertyName("parameters")]
    public Dictionary<string, object> Parameters { get; set; } = new();
}

public class ExcelResponse
{
    [System.Text.Json.Serialization.JsonPropertyName("success")]
    public bool Success { get; set; }
    
    [System.Text.Json.Serialization.JsonPropertyName("data")]
    public object? Data { get; set; }
    
    [System.Text.Json.Serialization.JsonPropertyName("error")]
    public string? Error { get; set; }
}