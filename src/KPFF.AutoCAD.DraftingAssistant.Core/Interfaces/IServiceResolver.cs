namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for service resolution and dependency injection
/// Renamed from IServiceProvider to avoid conflicts with System.IServiceProvider
/// </summary>
public interface IServiceResolver
{
    T GetService<T>() where T : class;
    T? GetOptionalService<T>() where T : class;
    bool IsServiceRegistered<T>() where T : class;
}