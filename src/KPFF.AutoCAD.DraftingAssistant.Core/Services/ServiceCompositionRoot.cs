using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Composition root for dependency injection - used only at application entry points
/// Replaces the static ApplicationServices anti-pattern with a proper composition root
/// </summary>
public sealed class ServiceCompositionRoot
{
    private readonly IServiceResolver _serviceResolver;
    private readonly object _lock = new();
    private bool _disposed = false;

    public ServiceCompositionRoot(IServiceResolver serviceResolver)
    {
        _serviceResolver = serviceResolver ?? throw new ArgumentNullException(nameof(serviceResolver));
    }

    /// <summary>
    /// Resolves a required service
    /// </summary>
    public T GetService<T>() where T : class
    {
        lock (_lock)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ServiceCompositionRoot));
                
            return _serviceResolver.GetService<T>();
        }
    }

    /// <summary>
    /// Resolves an optional service
    /// </summary>
    public T? GetOptionalService<T>() where T : class
    {
        lock (_lock)
        {
            if (_disposed)
                return null;
                
            return _serviceResolver.GetOptionalService<T>();
        }
    }

    /// <summary>
    /// Checks if a service is registered
    /// </summary>
    public bool IsServiceRegistered<T>() where T : class
    {
        lock (_lock)
        {
            if (_disposed)
                return false;
                
            return _serviceResolver.IsServiceRegistered<T>();
        }
    }

    /// <summary>
    /// Creates a command handler with resolved dependencies
    /// </summary>
    public T CreateCommandHandler<T>() where T : class, ICommandHandler
    {
        lock (_lock)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ServiceCompositionRoot));
                
            return _serviceResolver.GetService<T>();
        }
    }

    /// <summary>
    /// Disposes the composition root
    /// </summary>
    public void Dispose()
    {
        lock (_lock)
        {
            if (!_disposed)
            {
                if (_serviceResolver is IDisposable disposable)
                {
                    disposable.Dispose();
                }
                _disposed = true;
            }
        }
    }
}