using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Services;
using IServiceResolver = KPFF.AutoCAD.DraftingAssistant.Core.Interfaces.IServiceResolver;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// Main extension application for the KPFF Drafting Assistant plugin
/// </summary>
public class DraftingAssistantExtensionApplication : IExtensionApplication
{
    private static ServiceCompositionRoot? _compositionRoot;
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
                
                // Register DocumentContextManager immediately for closed drawing operations
                var documentContextManager = new DocumentContextManager(logger);
                DocumentContextRegistry.Register(documentContextManager);
                logger.LogDebug("DocumentContextManager registered for closed drawing operations");
                
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
                if (_compositionRoot?.IsServiceRegistered<IPaletteManager>() == true)
                {
                    var paletteManager = _compositionRoot.GetService<IPaletteManager>();
                    paletteManager.Cleanup();
                }
                
                // Dispose and cleanup the composition root
                _compositionRoot?.Dispose();
                _compositionRoot = null;
                
                // Clear DocumentContextManager tracking
                DocumentContextManager.ClearProtectionTracking();
                
                // Note: Excel reader process has been removed in Phase 1 refactoring
                
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
        if (_compositionRoot != null)
        {
            return; // Already initialized
        }

        var serviceProvider = new DependencyInjectionServiceProvider();

        // Register only basic logger - no other services yet
        var debugLogger = new DebugLogger();
        serviceProvider.RegisterSingleton<ILogger>(debugLogger);
        serviceProvider.RegisterSingleton<IApplicationLogger>(debugLogger);

        // DON'T build the service provider yet - keep it open for additional registrations
        // We'll create the composition root in SetupFullDependencyInjection
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
                
                // Now we can safely get the registered logger from the built composition root
                var registeredLogger = _compositionRoot?.GetService<ILogger>();
                
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
        if (_compositionRoot != null)
        {
            return; // Already initialized
        }

        // Create the service registration builder
        var serviceRegistration = new ApplicationServiceRegistration();
        
        // Register plugin-specific services
        serviceRegistration.RegisterNotificationService<AutoCadNotificationService>();
        serviceRegistration.RegisterPaletteManager<AutoCadPaletteManager>();
        serviceRegistration.RegisterDrawingAvailabilityService<DrawingAvailabilityService>();
        
        // Register command handlers
        RegisterCommandHandlers(serviceRegistration);
        
        // Build the composition root
        var serviceResolver = serviceRegistration.Build();
        _compositionRoot = new ServiceCompositionRoot(serviceResolver);
    }

    /// <summary>
    /// Initializes the palette manager safely after service registration
    /// </summary>
    private static void InitializePaletteManager(ILogger? logger)
    {
        try
        {
            if (_compositionRoot == null)
            {
                logger?.LogError("Cannot initialize palette manager - composition root is null");
                return;
            }

            var paletteManager = _compositionRoot.GetService<IPaletteManager>();
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
    /// Registers command handlers with the service registration
    /// </summary>
    private static void RegisterCommandHandlers(ApplicationServiceRegistration serviceRegistration)
    {
        // Since ApplicationServiceRegistration doesn't expose direct registration methods,
        // we'll need to resolve command handlers at the composition root level
        // This is acceptable since command handlers are leaf nodes in the dependency graph
    }

    /// <summary>
    /// Gets the current composition root for dependency resolution
    /// Used only at application entry points (commands)
    /// </summary>
    public static ServiceCompositionRoot? CompositionRoot => _compositionRoot;
}