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
        // Use a direct logger to avoid triggering service provider build
        var logger = new DebugLogger();
        
        ExceptionHandler.TryExecute(
            action: () =>
            {
                // Only setup basic dependency injection - no service initialization
                SetupBasicDependencyInjection();
                // Don't access services yet - this would trigger automatic build
                
                // Hook into Application.Idle to trigger ProjectWise fix automatically
                Application.Idle += OnApplicationIdleProjectWiseFix;
                
                logger.LogInformation("KPFF Drafting Assistant Plugin loaded - services will initialize on first command use");
            },
            logger: logger,
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
    /// One-time Application.Idle handler to automatically trigger ProjectWise fix
    /// Self-removes after first execution and includes robust error handling
    /// </summary>
    private static void OnApplicationIdleProjectWiseFix(object? sender, EventArgs e)
    {
        // Unregister immediately to ensure this only runs once
        Application.Idle -= OnApplicationIdleProjectWiseFix;
        
        try 
        {
            System.Diagnostics.Debug.WriteLine("ProjectWise Auto-Fix: Starting from Application.Idle event");
            
            // Trigger the ProjectWise fix asynchronously - don't wait for completion
            _ = ProjectWiseFix.TriggerProjectWiseInitialization();
        }
        catch (System.Exception ex)
        {
            // Catch any exceptions to prevent disrupting AutoCAD startup
            System.Diagnostics.Debug.WriteLine($"ProjectWise Auto-Fix: Failed safely during Idle event - {ex.Message}");
            // Continue silently - this is a best-effort fix
        }
    }

    /// <summary>
    /// Sets up only basic dependency injection for logging - no UI or service initialization
    /// Container is left "open" for additional service registration
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

        // DON'T build the service provider yet - keep it open for additional registrations
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
                // Use a direct logger since service provider hasn't been built yet
                var logger = new DebugLogger();
                logger.LogInformation("Initializing KPFF Drafting Assistant services...");

                // Complete the service registration that was skipped during plugin load
                SetupFullDependencyInjection();
                
                // Now we can safely get the registered logger from the built service provider
                var registeredLogger = _serviceProvider?.GetService<ILogger>();
                
                // Initialize the palette manager after services are registered
                InitializePaletteManager(registeredLogger);

                _servicesInitialized = true;
                registeredLogger?.LogInformation("KPFF Drafting Assistant services initialized successfully");
                return true;
            }
            catch (System.Exception ex)
            {
                // Use direct logger since service provider may not be built yet
                var logger = new DebugLogger();
                logger.LogError($"Failed to initialize services: {ex.Message}");
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
    /// Initializes the palette manager safely after service registration
    /// </summary>
    private static void InitializePaletteManager(ILogger? logger)
    {
        try
        {
            if (_serviceProvider == null)
            {
                logger?.LogError("Cannot initialize palette manager - service provider is null");
                return;
            }

            var paletteManager = _serviceProvider.GetService<IPaletteManager>();
            if (paletteManager == null)
            {
                logger?.LogError("Cannot initialize palette manager - service not registered");
                return;
            }

            if (!paletteManager.IsInitialized)
            {
                logger?.LogInformation("Initializing palette manager...");
                paletteManager.Initialize();
                logger?.LogInformation("Palette manager initialized successfully");
            }
            else
            {
                logger?.LogDebug("Palette manager already initialized");
            }
        }
        catch (System.Exception ex)
        {
            logger?.LogError($"Failed to initialize palette manager: {ex.Message}");
            // Don't throw - allow the plugin to continue working without palette
        }
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