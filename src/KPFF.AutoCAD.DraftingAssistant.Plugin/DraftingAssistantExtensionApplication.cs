using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Services;
using IServiceProvider = KPFF.AutoCAD.DraftingAssistant.Core.Interfaces.IServiceProvider;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// Main extension application for the KPFF Drafting Assistant plugin
/// </summary>
public class DraftingAssistantExtensionApplication : IExtensionApplication
{
    private static DependencyInjectionServiceProvider? _serviceProvider;
    private ILogger? _logger;
    private PluginStartupManager? _startupManager;

    public void Initialize()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                SetupDependencyInjection();
                _logger = _serviceProvider!.GetService<ILogger>();
                var paletteManager = _serviceProvider.GetService<IPaletteManager>();

                _logger.LogInformation("KPFF Drafting Assistant Plugin initializing...");
                
                // Use the startup manager for robust initialization
                _startupManager = new PluginStartupManager(_logger, paletteManager);
                _startupManager.BeginInitialization();
                
                _logger.LogInformation("KPFF Drafting Assistant Plugin setup completed - initialization in progress");
            },
            logger: _logger ?? new DebugLogger(),
            context: "Plugin Initialization",
            showUserMessage: false
        );
    }

    public void Terminate()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger?.LogInformation("KPFF Drafting Assistant Plugin terminating...");
                
                // Cleanup startup manager
                _startupManager = null;
                
                // Cleanup services
                if (_serviceProvider?.IsServiceRegistered<IPaletteManager>() == true)
                {
                    var paletteManager = _serviceProvider.GetService<IPaletteManager>();
                    paletteManager.Cleanup();
                }
                
                // Dispose and cleanup the service provider
                _serviceProvider?.Dispose();
                _serviceProvider = null;
                
                _logger?.LogInformation("KPFF Drafting Assistant Plugin terminated successfully");
            },
            logger: _logger ?? new DebugLogger(),
            context: "Plugin Termination",
            showUserMessage: false
        );
    }

    /// <summary>
    /// Sets up the dependency injection container with all required services
    /// CRASH FIX: Use isolated services that never access drawing context during initialization
    /// </summary>
    private static void SetupDependencyInjection()
    {
        if (_serviceProvider != null)
        {
            return; // Already initialized
        }

        _serviceProvider = new DependencyInjectionServiceProvider();

        // Register core services - use isolated instances to avoid drawing context access
        var debugLogger = new DebugLogger();
        _serviceProvider.RegisterSingleton<ILogger>(debugLogger);
        _serviceProvider.RegisterSingleton<IApplicationLogger>(debugLogger);

        // Register notification service with logger dependency
        _serviceProvider.RegisterSingleton<INotificationService, AutoCadNotificationService>();
        
        // Register palette manager with dependencies
        _serviceProvider.RegisterSingleton<IPaletteManager, AutoCadPaletteManager>();

        // Register command handlers
        RegisterCommandHandlers(_serviceProvider);

        // Build the service provider
        _serviceProvider.BuildServiceProvider();
    }

    /// <summary>
    /// Registers all command handlers with the container
    /// </summary>
    private static void RegisterCommandHandlers(DependencyInjectionServiceProvider serviceProvider)
    {
        // Register command handlers as transient (new instance each time)
        serviceProvider.RegisterTransient<ShowPaletteCommandHandler, ShowPaletteCommandHandler>();
        serviceProvider.RegisterTransient<HidePaletteCommandHandler, HidePaletteCommandHandler>();
        serviceProvider.RegisterTransient<TogglePaletteCommandHandler, TogglePaletteCommandHandler>();
        serviceProvider.RegisterTransient<HelpCommandHandler, HelpCommandHandler>();
    }

    /// <summary>
    /// Gets the current service provider instance
    /// </summary>
    public static IServiceProvider? ServiceProvider => _serviceProvider;
}