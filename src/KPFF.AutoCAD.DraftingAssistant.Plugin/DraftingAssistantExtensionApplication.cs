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
    private static bool _servicesInitialized = false;
    private static readonly object _initializationLock = new object();
    private ILogger? _logger;

    public void Initialize()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                // Only setup basic dependency injection - no service initialization
                SetupBasicDependencyInjection();
                _logger = _serviceProvider!.GetService<ILogger>();
                
                _logger.LogInformation("KPFF Drafting Assistant Plugin loaded - services will initialize on first command use");
            },
            logger: _logger ?? new DebugLogger(),
            context: "Plugin Loading",
            showUserMessage: false
        );
    }

    public void Terminate()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger?.LogInformation("KPFF Drafting Assistant Plugin terminating...");
                
                // Cleanup services
                if (_serviceProvider?.IsServiceRegistered<IPaletteManager>() == true)
                {
                    var paletteManager = _serviceProvider.GetService<IPaletteManager>();
                    paletteManager.Cleanup();
                }
                
                // Dispose and cleanup the service provider
                _serviceProvider?.Dispose();
                _serviceProvider = null;
                
                // Force terminate the Excel reader process
                _logger?.LogInformation("Terminating Excel reader process...");
                Core.Services.SharedExcelReaderProcess.ForceTerminate(msg => _logger?.LogInformation(msg));
                
                _logger?.LogInformation("KPFF Drafting Assistant Plugin terminated successfully");
            },
            logger: _logger ?? new DebugLogger(),
            context: "Plugin Termination",
            showUserMessage: false
        );
    }

    /// <summary>
    /// Sets up only basic dependency injection for logging - no UI or service initialization
    /// </summary>
    private static void SetupBasicDependencyInjection()
    {
        if (_serviceProvider != null)
        {
            return; // Already initialized
        }

        _serviceProvider = new DependencyInjectionServiceProvider();

        // Register only basic logger - no other services yet
        var debugLogger = new DebugLogger();
        _serviceProvider.RegisterSingleton<ILogger>(debugLogger);
        _serviceProvider.RegisterSingleton<IApplicationLogger>(debugLogger);

        // Build the basic service provider
        _serviceProvider.BuildServiceProvider();
    }

    /// <summary>
    /// Ensures all services are initialized on first command use
    /// THREAD-SAFE: Uses lock to prevent concurrent initialization
    /// DOCUMENT-SAFE: Checks for valid document context before initialization
    /// </summary>
    public static bool EnsureServicesInitialized()
    {
        if (_servicesInitialized)
        {
            return true;
        }

        // Check document state first (per AutoCAD Developer's Guide)
        if (Application.DocumentManager == null || Application.DocumentManager.Count == 0)
        {
            // No documents open - cannot initialize
            return false;
        }

        if (Application.DocumentManager.MdiActiveDocument == null)
        {
            // No active document - cannot initialize
            return false;
        }

        lock (_initializationLock)
        {
            if (_servicesInitialized)
            {
                return true; // Double-check pattern
            }

            try
            {
                var logger = _serviceProvider?.GetService<ILogger>();
                logger?.LogInformation("Initializing KPFF Drafting Assistant services...");

                // Complete the service registration that was skipped during plugin load
                SetupFullDependencyInjection();
                
                // Services are now registered and ready for use - no additional initialization needed

                _servicesInitialized = true;
                logger?.LogInformation("KPFF Drafting Assistant services initialized successfully");
                return true;
            }
            catch (System.Exception ex)
            {
                var logger = _serviceProvider?.GetService<ILogger>();
                logger?.LogError($"Failed to initialize services: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// Sets up the complete dependency injection container with all required services
    /// Called lazily on first command use
    /// </summary>
    private static void SetupFullDependencyInjection()
    {
        if (_serviceProvider == null)
        {
            throw new InvalidOperationException("Basic service provider must be initialized first");
        }

        // Register Excel services as transient - process lifecycle managed by SharedExcelReaderProcess
        _serviceProvider.RegisterTransient<IExcelReader, ExcelReaderService>();

        // Register notification service with logger dependency
        _serviceProvider.RegisterSingleton<INotificationService, AutoCadNotificationService>();
        
        // Register palette manager with dependencies
        _serviceProvider.RegisterSingleton<IPaletteManager, AutoCadPaletteManager>();

        // Register command handlers
        RegisterCommandHandlers(_serviceProvider);

        // Rebuild the service provider with new registrations
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