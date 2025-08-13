using Microsoft.Extensions.DependencyInjection;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Dependency injection based service provider using Microsoft.Extensions.DependencyInjection
/// </summary>
public class DependencyInjectionServiceProvider : Interfaces.IServiceProvider, IDisposable
{
    private System.IServiceProvider _serviceProvider;
    private readonly ServiceCollection _services;
    private bool _isBuilt = false;
    private bool _disposed = false;

    public DependencyInjectionServiceProvider()
    {
        _services = new ServiceCollection();
        // Don't build the service provider yet - keep container open for registrations
        _serviceProvider = null!;
    }

    /// <summary>
    /// Registers a service with singleton lifetime
    /// </summary>
    public void RegisterSingleton<TInterface, TImplementation>()
        where TInterface : class
        where TImplementation : class, TInterface
    {
        ThrowIfBuilt();
        _services.AddSingleton<TInterface, TImplementation>();
    }

    /// <summary>
    /// Registers a service instance as singleton
    /// </summary>
    public void RegisterSingleton<T>(T instance) where T : class
    {
        ThrowIfBuilt();
        _services.AddSingleton(instance);
    }

    /// <summary>
    /// Registers a service with a factory function as singleton
    /// </summary>
    public void RegisterSingleton<T>(Func<System.IServiceProvider, T> factory) where T : class
    {
        ThrowIfBuilt();
        _services.AddSingleton(factory);
    }

    /// <summary>
    /// Registers a service with transient lifetime
    /// </summary>
    public void RegisterTransient<TInterface, TImplementation>()
        where TInterface : class
        where TImplementation : class, TInterface
    {
        ThrowIfBuilt();
        _services.AddTransient<TInterface, TImplementation>();
    }

    /// <summary>
    /// Registers a service with scoped lifetime
    /// </summary>
    public void RegisterScoped<TInterface, TImplementation>()
        where TInterface : class
        where TImplementation : class, TInterface
    {
        ThrowIfBuilt();
        _services.AddScoped<TInterface, TImplementation>();
    }

    /// <summary>
    /// Builds the service provider - must be called before resolving services
    /// </summary>
    public void BuildServiceProvider()
    {
        if (_isBuilt) return;
        
        var newServiceProvider = _services.BuildServiceProvider();
        
        // Replace the service provider
        if (_serviceProvider is IDisposable disposable)
        {
            disposable.Dispose();
        }
        
        _serviceProvider = newServiceProvider;
        _isBuilt = true;
    }

    /// <summary>
    /// Gets a required service
    /// </summary>
    public T GetService<T>() where T : class
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DependencyInjectionServiceProvider));
        if (!_isBuilt || _serviceProvider == null)
        {
            throw new InvalidOperationException("Service provider has not been built yet. Call BuildServiceProvider() first.");
        }
        
        var service = _serviceProvider.GetService(typeof(T)) as T;
        if (service == null)
        {
            throw new InvalidOperationException($"Service of type {typeof(T).Name} is not registered.");
        }
        return service;
    }

    /// <summary>
    /// Gets an optional service (returns null if not found)
    /// </summary>
    public T? GetOptionalService<T>() where T : class
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DependencyInjectionServiceProvider));
        if (!_isBuilt || _serviceProvider == null)
        {
            throw new InvalidOperationException("Service provider has not been built yet. Call BuildServiceProvider() first.");
        }
        
        return _serviceProvider.GetService(typeof(T)) as T;
    }

    /// <summary>
    /// Checks if a service is registered
    /// </summary>
    public bool IsServiceRegistered<T>() where T : class
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DependencyInjectionServiceProvider));
        if (!_isBuilt || _serviceProvider == null)
        {
            throw new InvalidOperationException("Service provider has not been built yet. Call BuildServiceProvider() first.");
        }
        
        return _serviceProvider.GetService(typeof(T)) != null;
    }

    /// <summary>
    /// Creates a new service scope for scoped services
    /// </summary>
    public IServiceScope CreateScope()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DependencyInjectionServiceProvider));
        if (!_isBuilt || _serviceProvider == null)
        {
            throw new InvalidOperationException("Service provider has not been built yet. Call BuildServiceProvider() first.");
        }
        
        return _serviceProvider.CreateScope();
    }

    private void ThrowIfBuilt()
    {
        if (_isBuilt)
        {
            throw new InvalidOperationException("Cannot register services after the service provider has been built.");
        }
    }


    public void Dispose()
    {
        if (!_disposed)
        {
            if (_serviceProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}