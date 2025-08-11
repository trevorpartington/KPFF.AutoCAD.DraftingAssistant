namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Application-specific logger interface that extends standard logging capabilities
/// </summary>
public interface IApplicationLogger : ILogger
{
    void LogCritical(string message, System.Exception? exception = null);
}