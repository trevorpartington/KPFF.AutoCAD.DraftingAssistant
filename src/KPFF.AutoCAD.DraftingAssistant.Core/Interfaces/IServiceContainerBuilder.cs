namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for service registration and dependency injection configuration
/// Separated from service resolution to follow Single Responsibility Principle
/// </summary>
public interface IServiceContainerBuilder
{
    /// <summary>
    /// Register all application services
    /// </summary>
    void RegisterServices();
    
    /// <summary>
    /// Register a notification service implementation
    /// </summary>
    void RegisterNotificationService<T>() where T : class, INotificationService;
    
    /// <summary>
    /// Register a palette manager implementation
    /// </summary>
    void RegisterPaletteManager<T>() where T : class, IPaletteManager;
    
    /// <summary>
    /// Build the service container and return a resolver
    /// </summary>
    IServiceResolver Build();
}