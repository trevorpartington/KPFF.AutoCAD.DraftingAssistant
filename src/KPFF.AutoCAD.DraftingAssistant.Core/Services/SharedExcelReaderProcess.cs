using System.Diagnostics;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Static class that manages a single shared Excel reader process across all service instances.
/// This ensures only one process runs while allowing safe concurrent access from multiple threads.
/// </summary>
public static class SharedExcelReaderProcess
{
    private static readonly object _lock = new object();
    private static Process? _excelReaderProcess;
    private static int _referenceCount = 0;
    private static bool _shutdownRegistered = false;

    /// <summary>
    /// Gets or starts the shared Excel reader process
    /// </summary>
    public static Process? GetOrStartProcess(string excelReaderPath, Action<string>? logAction = null)
    {
        lock (_lock)
        {
            // Register shutdown handler on first use
            if (!_shutdownRegistered)
            {
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
                _shutdownRegistered = true;
                logAction?.Invoke("Registered process exit handler for Excel reader cleanup");
            }

            // If process exists and is running, return it
            if (_excelReaderProcess != null && !_excelReaderProcess.HasExited)
            {
                _referenceCount++;
                logAction?.Invoke($"Returning existing Excel reader process (PID: {_excelReaderProcess.Id}, References: {_referenceCount})");
                return _excelReaderProcess;
            }

            // Clean up dead process reference
            if (_excelReaderProcess != null)
            {
                try
                {
                    _excelReaderProcess.Dispose();
                }
                catch { }
                _excelReaderProcess = null;
                _referenceCount = 0;
            }

            // Start new process
            try
            {
                logAction?.Invoke($"Starting new Excel reader process: {excelReaderPath}");

                _excelReaderProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = excelReaderPath,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    }
                };

                _excelReaderProcess.Start();
                _referenceCount = 1;

                logAction?.Invoke($"Excel reader process started successfully (PID: {_excelReaderProcess.Id})");
                return _excelReaderProcess;
            }
            catch (Exception ex)
            {
                logAction?.Invoke($"Failed to start Excel reader process: {ex.Message}");
                _excelReaderProcess = null;
                _referenceCount = 0;
                throw;
            }
        }
    }

    /// <summary>
    /// Increments the reference count for an already running process
    /// </summary>
    public static void IncrementReference(Action<string>? logAction = null)
    {
        lock (_lock)
        {
            if (_excelReaderProcess != null && !_excelReaderProcess.HasExited)
            {
                _referenceCount++;
                logAction?.Invoke($"Incremented Excel reader process reference (Total: {_referenceCount})");
            }
            else
            {
                logAction?.Invoke("Cannot increment reference - process not running");
            }
        }
    }

    /// <summary>
    /// Releases a reference to the shared process
    /// </summary>
    public static void ReleaseReference(Action<string>? logAction = null)
    {
        lock (_lock)
        {
            if (_referenceCount > 0)
            {
                _referenceCount--;
                logAction?.Invoke($"Released Excel reader process reference (Remaining: {_referenceCount})");

                // We keep the process running even with 0 references
                // It will be cleaned up on AppDomain exit
            }
        }
    }

    /// <summary>
    /// Forces termination of the Excel reader process
    /// </summary>
    public static void ForceTerminate(Action<string>? logAction = null)
    {
        lock (_lock)
        {
            if (_excelReaderProcess != null)
            {
                try
                {
                    if (!_excelReaderProcess.HasExited)
                    {
                        logAction?.Invoke($"Terminating Excel reader process (PID: {_excelReaderProcess.Id})");
                        _excelReaderProcess.Kill();
                        _excelReaderProcess.WaitForExit(5000);
                    }
                    _excelReaderProcess.Dispose();
                }
                catch (Exception ex)
                {
                    logAction?.Invoke($"Error terminating Excel reader process: {ex.Message}");
                }
                finally
                {
                    _excelReaderProcess = null;
                    _referenceCount = 0;
                }
            }
        }
    }

    /// <summary>
    /// Cleanup handler for AppDomain exit
    /// </summary>
    private static void OnProcessExit(object? sender, EventArgs e)
    {
        ForceTerminate(msg => 
        {
            // Try to log if possible, but don't throw during shutdown
            try
            {
                System.Diagnostics.Debug.WriteLine($"[Excel Reader Cleanup] {msg}");
            }
            catch { }
        });
    }

    /// <summary>
    /// Gets the current reference count (for debugging)
    /// </summary>
    public static int ReferenceCount
    {
        get
        {
            lock (_lock)
            {
                return _referenceCount;
            }
        }
    }

    /// <summary>
    /// Checks if the process is currently running
    /// </summary>
    public static bool IsRunning
    {
        get
        {
            lock (_lock)
            {
                return _excelReaderProcess != null && !_excelReaderProcess.HasExited;
            }
        }
    }
}