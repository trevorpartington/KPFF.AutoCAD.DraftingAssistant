using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Static access point for application services - used only at composition root
/// </summary>
public static class ApplicationServices
{
    private static IServiceRegistration? _serviceRegistration;
    private static readonly object _lock = new();

    /// <summary>
    /// Initialize the application services with a service registration implementation
    /// </summary>
    public static void Initialize(IServiceRegistration serviceRegistration)
    {
        lock (_lock)
        {
            if (_serviceRegistration != null)
                throw new InvalidOperationException("Application services have already been initialized");
            
            _serviceRegistration = serviceRegistration ?? throw new ArgumentNullException(nameof(serviceRegistration));
            _serviceRegistration.RegisterServices();
        }
    }

    /// <summary>
    /// Initialize with default service registration
    /// </summary>
    public static void Initialize()
    {
        Initialize(new ApplicationServiceRegistration());
    }

    /// <summary>
    /// Get a required service
    /// </summary>
    public static T GetService<T>() where T : class
    {
        ThrowIfNotInitialized();
        return _serviceRegistration!.GetService<T>();
    }

    /// <summary>
    /// Get an optional service
    /// </summary>
    public static T? GetOptionalService<T>() where T : class
    {
        return _serviceRegistration?.GetOptionalService<T>();
    }

    /// <summary>
    /// Check if a service is registered
    /// </summary>
    public static bool IsServiceRegistered<T>() where T : class
    {
        return _serviceRegistration?.IsServiceRegistered<T>() ?? false;
    }

    /// <summary>
    /// Check if application services are initialized
    /// </summary>
    public static bool IsInitialized => _serviceRegistration != null;

    /// <summary>
    /// Reset services - for testing purposes only
    /// </summary>
    public static void Reset()
    {
        lock (_lock)
        {
            _serviceRegistration = null;
        }
    }

    private static void ThrowIfNotInitialized()
    {
        if (_serviceRegistration == null)
            throw new InvalidOperationException("Application services have not been initialized. Call ApplicationServices.Initialize() first.");
    }
}