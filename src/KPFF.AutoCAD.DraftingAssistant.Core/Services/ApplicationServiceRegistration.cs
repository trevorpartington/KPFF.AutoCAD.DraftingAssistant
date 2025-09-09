using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Centralized service registration for the entire application
/// Implements both builder and resolver to maintain backward compatibility
/// </summary>
public class ApplicationServiceRegistration : IServiceContainerBuilder, IServiceResolver
{
    private readonly DependencyInjectionServiceProvider _serviceProvider;
    private bool _servicesRegistered = false;

    public ApplicationServiceRegistration()
    {
        _serviceProvider = new DependencyInjectionServiceProvider();
    }

    public void RegisterServices()
    {
        if (_servicesRegistered) return;

        // Register core services
        RegisterCoreServices();
        
        // Build the service provider
        _serviceProvider.BuildServiceProvider();
        
        _servicesRegistered = true;
    }

    private void RegisterCoreServices()
    {
        // Logger services (register as singleton - DebugLogger implements both interfaces)
        _serviceProvider.RegisterSingleton<DebugLogger, DebugLogger>();
        _serviceProvider.RegisterSingleton<ILogger>(sp => (sp.GetService(typeof(DebugLogger)) as DebugLogger)!);
        _serviceProvider.RegisterSingleton<IApplicationLogger>(sp => (sp.GetService(typeof(DebugLogger)) as DebugLogger)!);

        // Configuration services
        _serviceProvider.RegisterTransient<IProjectConfigurationService, ProjectConfigurationService>();
        
        // Excel services (Phase 3 ClosedXML implementation)
        _serviceProvider.RegisterTransient<IExcelReader, ExcelReaderService>();

        // Construction Notes services
        _serviceProvider.RegisterTransient<IConstructionNotesService, ConstructionNotesService>();
        _serviceProvider.RegisterTransient<IConstructionNoteBlockManager, ConstructionNoteBlockManager>();
        _serviceProvider.RegisterTransient<IDrawingOperations, DrawingOperations>();
        
        // CRASH FIX: Removed CurrentDrawingBlockManager registration from DI container
        // It will be instantiated manually when needed to avoid premature AutoCAD access
        
        // Drawing availability service - plugin-specific implementation will be registered by Plugin layer
        // Core interface is defined here but implementation must be AutoCAD-aware
        
        // Notification services - these will be registered by specific layers (UI/Plugin)
        // Command handlers will be registered by Plugin layer
    }

    public void RegisterNotificationService<T>() where T : class, INotificationService
    {
        if (_servicesRegistered)
            throw new InvalidOperationException("Cannot register services after they have been built");
        
        _serviceProvider.RegisterSingleton<INotificationService, T>();
    }

    public void RegisterPaletteManager<T>() where T : class, IPaletteManager
    {
        if (_servicesRegistered)
            throw new InvalidOperationException("Cannot register services after they have been built");
        
        _serviceProvider.RegisterSingleton<IPaletteManager, T>();
    }

    public void RegisterDrawingAvailabilityService<T>() where T : class, IDrawingAvailabilityService
    {
        if (_servicesRegistered)
            throw new InvalidOperationException("Cannot register services after they have been built");
        
        _serviceProvider.RegisterSingleton<IDrawingAvailabilityService, T>();
    }

    public T GetService<T>() where T : class
    {
        if (!_servicesRegistered)
            throw new InvalidOperationException("Services must be registered before resolving");
        
        return _serviceProvider.GetService<T>();
    }

    public T? GetOptionalService<T>() where T : class
    {
        if (!_servicesRegistered)
            return null;
        
        return _serviceProvider.GetOptionalService<T>();
    }

    public bool IsServiceRegistered<T>() where T : class
    {
        if (!_servicesRegistered)
            return false;
        
        return _serviceProvider.IsServiceRegistered<T>();
    }
    
    /// <summary>
    /// Build the service container and return a resolver
    /// </summary>
    public IServiceResolver Build()
    {
        RegisterServices();
        return this; // Return self as we implement both interfaces
    }
}