using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Centralized service registration for the entire application
/// </summary>
public class ApplicationServiceRegistration : IServiceRegistration
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
        
        // Excel services
        _serviceProvider.RegisterTransient<IExcelReader, ExcelReaderService>();

        // Notification services - these will be registered by specific layers (UI/Plugin)
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
}