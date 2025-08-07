using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using System.Diagnostics;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Debug-based logger implementation for development
/// </summary>
public class DebugLogger : ILogger, IApplicationLogger
{
    private static readonly object _lock = new object();

    public void LogInformation(string message)
    {
        WriteLog("INFO", message);
    }

    public void LogWarning(string message)
    {
        WriteLog("WARN", message);
    }

    public void LogError(string message, System.Exception? exception = null)
    {
        WriteLog("ERROR", message);
        if (exception != null)
        {
            WriteLog("ERROR", $"Exception Details: {exception}");
        }
    }

    public void LogDebug(string message)
    {
        WriteLog("DEBUG", message);
    }

    public void LogCritical(string message, System.Exception? exception = null)
    {
        WriteLog("CRITICAL", message);
        if (exception != null)
        {
            WriteLog("CRITICAL", $"Exception Details: {exception}");
        }
    }

    private static void WriteLog(string level, string message)
    {
        lock (_lock)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var threadId = Environment.CurrentManagedThreadId;
                Debug.WriteLine($"[{level}] {timestamp} [Thread-{threadId}] - {message}");
            }
            catch
            {
                // Fallback if logging fails
                Debug.WriteLine($"[{level}] - {message}");
            }
        }
    }
}