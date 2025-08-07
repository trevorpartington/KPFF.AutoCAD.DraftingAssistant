using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Simple dependency injection container for the application
/// </summary>
public class ServiceContainer
{
    private readonly Dictionary<Type, object> _services = new();
    private static ServiceContainer? _instance;
    
    public static ServiceContainer Instance => _instance ??= new ServiceContainer();

    /// <summary>
    /// Register a service instance
    /// </summary>
    public void Register<T>(T service) where T : class
    {
        _services[typeof(T)] = service;
    }

    /// <summary>
    /// Register a service with a factory function
    /// </summary>
    public void Register<T>(Func<T> factory) where T : class
    {
        _services[typeof(T)] = factory();
    }

    /// <summary>
    /// Resolve a service instance
    /// </summary>
    public T Resolve<T>() where T : class
    {
        if (_services.TryGetValue(typeof(T), out var service))
        {
            return (T)service;
        }
        
        throw new InvalidOperationException($"Service of type {typeof(T).Name} is not registered.");
    }

    /// <summary>
    /// Check if a service is registered
    /// </summary>
    public bool IsRegistered<T>() where T : class
    {
        return _services.ContainsKey(typeof(T));
    }

    /// <summary>
    /// Clear all registered services
    /// </summary>
    public void Clear()
    {
        _services.Clear();
    }
}