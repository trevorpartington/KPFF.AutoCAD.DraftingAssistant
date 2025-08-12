using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using KPFF.AutoCAD.ExcelReader.Models;

namespace KPFF.AutoCAD.ExcelReader.Services;

public class NamedPipeServer
{
    private const string PIPE_NAME = "KPFF_AutoCAD_ExcelReader";
    private readonly ExcelProcessor _processor;
    private bool _isRunning;
    private CancellationTokenSource? _cancellationTokenSource;

    public NamedPipeServer(ExcelProcessor processor)
    {
        _processor = processor;
    }

    public async Task StartAsync()
    {
        _isRunning = true;
        _cancellationTokenSource = new CancellationTokenSource();
        
        Console.WriteLine($"Excel Reader starting on pipe: {PIPE_NAME}");
        
        // Run multiple pipe instances to handle concurrent requests
        var tasks = new List<Task>();
        for (int i = 0; i < 4; i++)
        {
            tasks.Add(HandleClientAsync(_cancellationTokenSource.Token));
        }
        
        await Task.WhenAll(tasks);
    }

    public void Stop()
    {
        _isRunning = false;
        _cancellationTokenSource?.Cancel();
        Console.WriteLine("Excel Reader stopping...");
    }

    private async Task HandleClientAsync(CancellationToken cancellationToken)
    {
        while (_isRunning && !cancellationToken.IsCancellationRequested)
        {
            try
            {
                using var pipeServer = new NamedPipeServerStream(
                    PIPE_NAME,
                    PipeDirection.InOut,
                    4, // max instances
                    PipeTransmissionMode.Message,
                    PipeOptions.Asynchronous);

                Console.WriteLine("Waiting for client connection...");
                await pipeServer.WaitForConnectionAsync(cancellationToken);
                Console.WriteLine("Client connected");

                // Read request
                var requestData = await ReadMessageAsync(pipeServer, cancellationToken);
                if (requestData == null) continue;

                ExcelRequest? request = null;
                try
                {
                    request = JsonSerializer.Deserialize<ExcelRequest>(requestData);
                }
                catch (JsonException ex)
                {
                    var errorResponse = new ExcelResponse 
                    { 
                        Success = false, 
                        Error = $"Invalid JSON request: {ex.Message}" 
                    };
                    await SendResponseAsync(pipeServer, errorResponse, cancellationToken);
                    continue;
                }

                if (request == null)
                {
                    var errorResponse = new ExcelResponse 
                    { 
                        Success = false, 
                        Error = "Null request received" 
                    };
                    await SendResponseAsync(pipeServer, errorResponse, cancellationToken);
                    continue;
                }

                Console.WriteLine($"Processing request: {request.Operation}");

                // Process request
                var response = await _processor.ProcessRequestAsync(request);

                // Send response
                await SendResponseAsync(pipeServer, response, cancellationToken);
                
                Console.WriteLine($"Request completed: {request.Operation}");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Operation cancelled");
                break;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error handling client: {ex.Message}");
                await Task.Delay(1000, cancellationToken); // Brief delay before retry
            }
        }
    }

    private async Task<string?> ReadMessageAsync(NamedPipeServerStream pipe, CancellationToken cancellationToken)
    {
        try
        {
            // Read message length first (4 bytes)
            var lengthBuffer = new byte[4];
            var bytesRead = await pipe.ReadAsync(lengthBuffer, 0, 4, cancellationToken);
            if (bytesRead != 4) 
            {
                Console.WriteLine($"Failed to read message length: got {bytesRead} bytes, expected 4");
                return null;
            }

            var messageLength = BitConverter.ToInt32(lengthBuffer, 0);
            Console.WriteLine($"Message length received: {messageLength} bytes");
            
            if (messageLength <= 0 || messageLength > 1024 * 1024) // Max 1MB
            {
                Console.WriteLine($"Invalid message length: {messageLength}");
                return null;
            }

            // Read message content
            var messageBuffer = new byte[messageLength];
            var totalBytesRead = 0;
            
            while (totalBytesRead < messageLength)
            {
                bytesRead = await pipe.ReadAsync(
                    messageBuffer, 
                    totalBytesRead, 
                    messageLength - totalBytesRead, 
                    cancellationToken);
                
                if (bytesRead == 0) break;
                totalBytesRead += bytesRead;
            }

            var messageContent = Encoding.UTF8.GetString(messageBuffer, 0, totalBytesRead);
            Console.WriteLine($"Message content received ({totalBytesRead}/{messageLength} bytes): '{messageContent}'");
            
            return messageContent;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading message: {ex.Message}");
            return null;
        }
    }

    private async Task SendResponseAsync(NamedPipeServerStream pipe, ExcelResponse response, CancellationToken cancellationToken)
    {
        try
        {
            var responseJson = JsonSerializer.Serialize(response);
            var responseBytes = Encoding.UTF8.GetBytes(responseJson);
            var lengthBytes = BitConverter.GetBytes(responseBytes.Length);

            // Send length first, then content
            await pipe.WriteAsync(lengthBytes, 0, 4, cancellationToken);
            await pipe.WriteAsync(responseBytes, 0, responseBytes.Length, cancellationToken);
            await pipe.FlushAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending response: {ex.Message}");
        }
    }
}