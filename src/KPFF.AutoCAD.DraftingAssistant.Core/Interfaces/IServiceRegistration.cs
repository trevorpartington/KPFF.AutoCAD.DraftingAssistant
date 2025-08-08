namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for service registration and dependency injection configuration
/// </summary>
public interface IServiceRegistration
{
    /// <summary>
    /// Register all application services
    /// </summary>
    void RegisterServices();
    
    /// <summary>
    /// Get a service instance by type
    /// </summary>
    T GetService<T>() where T : class;
    
    /// <summary>
    /// Get an optional service instance by type (returns null if not found)
    /// </summary>
    T? GetOptionalService<T>() where T : class;
    
    /// <summary>
    /// Check if a service is registered
    /// </summary>
    bool IsServiceRegistered<T>() where T : class;
}