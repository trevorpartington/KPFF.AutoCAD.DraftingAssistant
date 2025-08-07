namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for service resolution and dependency injection
/// </summary>
public interface IServiceProvider
{
    T GetService<T>() where T : class;
    T? GetOptionalService<T>() where T : class;
    bool IsServiceRegistered<T>() where T : class;
}