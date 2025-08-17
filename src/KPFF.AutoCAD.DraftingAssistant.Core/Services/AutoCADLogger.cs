using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using System.Diagnostics;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Logger implementation that writes to AutoCAD's command line
/// Falls back to Debug output if AutoCAD is not available
/// </summary>
public class AutoCADLogger : ILogger, IApplicationLogger
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
                var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                var formattedMessage = $"[{level}] {timestamp} - {message}";
                
                // Try to write to AutoCAD command line first
                if (TryWriteToAutoCAD(formattedMessage))
                {
                    return;
                }
                
                // Fallback to Debug output
                Debug.WriteLine($"[AutoCADLogger] {formattedMessage}");
            }
            catch (Exception ex)
            {
                // Fallback if logging fails
                Debug.WriteLine($"[AutoCADLogger] [{level}] - {message} (Logging Error: {ex.Message})");
            }
        }
    }

    private static bool TryWriteToAutoCAD(string message)
    {
        try
        {
            // Check if AutoCAD is available
            var docMgr = Application.DocumentManager;
            if (docMgr == null)
                return false;

            var doc = docMgr.MdiActiveDocument;
            if (doc == null)
                return false;

            var editor = doc.Editor;
            if (editor == null)
                return false;

            // Write to AutoCAD command line
            editor.WriteMessage($"\n{message}");
            return true;
        }
        catch
        {
            // If anything fails accessing AutoCAD, return false to use fallback
            return false;
        }
    }
}