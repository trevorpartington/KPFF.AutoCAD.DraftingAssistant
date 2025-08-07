using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using System.Diagnostics;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Debug-based logger implementation for development
/// </summary>
public class DebugLogger : ILogger
{
    public void LogInformation(string message)
    {
        Debug.WriteLine($"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }

    public void LogWarning(string message)
    {
        Debug.WriteLine($"[WARN] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }

    public void LogError(string message, System.Exception? exception = null)
    {
        Debug.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        if (exception != null)
        {
            Debug.WriteLine($"Exception: {exception}");
        }
    }

    public void LogDebug(string message)
    {
        Debug.WriteLine($"[DEBUG] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }
}