namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Application-specific logger interface that extends standard logging capabilities
/// </summary>
public interface IApplicationLogger
{
    void LogInformation(string message);
    void LogWarning(string message);
    void LogError(string message, System.Exception? exception = null);
    void LogDebug(string message);
    void LogCritical(string message, System.Exception? exception = null);
}