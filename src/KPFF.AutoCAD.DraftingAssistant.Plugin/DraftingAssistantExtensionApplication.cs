using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// Main extension application for the KPFF Drafting Assistant plugin
/// </summary>
public class DraftingAssistantExtensionApplication : IExtensionApplication
{
    private ILogger? _logger;
    private PluginStartupManager? _startupManager;

    public void Initialize()
    {
        try
        {
            SetupDependencyInjection();
            _logger = ServiceContainer.Instance.Resolve<ILogger>();
            var paletteManager = ServiceContainer.Instance.Resolve<IPaletteManager>();

            _logger.LogInformation("KPFF Drafting Assistant Plugin initializing...");
            
            // Use the startup manager for robust initialization
            _startupManager = new PluginStartupManager(_logger, paletteManager);
            _startupManager.BeginInitialization();
            
            _logger.LogInformation("KPFF Drafting Assistant Plugin setup completed - initialization in progress");
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Critical error initializing plugin: {ex.Message}");
            // Don't rethrow - we want AutoCAD to continue working
        }
    }

    public void Terminate()
    {
        try
        {
            _logger?.LogInformation("KPFF Drafting Assistant Plugin terminating...");
            
            // Cleanup startup manager
            _startupManager = null;
            
            // Cleanup services
            if (ServiceContainer.Instance.IsRegistered<IPaletteManager>())
            {
                var paletteManager = ServiceContainer.Instance.Resolve<IPaletteManager>();
                paletteManager.Cleanup();
            }
            
            ServiceContainer.Instance.Clear();
            _logger?.LogInformation("KPFF Drafting Assistant Plugin terminated successfully");
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error terminating plugin: {ex.Message}");
            // Don't rethrow during shutdown
        }
    }

    /// <summary>
    /// Sets up the dependency injection container with all required services
    /// </summary>
    private static void SetupDependencyInjection()
    {
        var container = ServiceContainer.Instance;

        // Register core services
        container.Register<ILogger>(new DebugLogger());
        container.Register<INotificationService>(() => 
            new AutoCadNotificationService());
        
        // Register palette manager with dependencies
        container.Register<IPaletteManager>(() => 
            new AutoCadPaletteManager(
                container.Resolve<ILogger>(),
                container.Resolve<INotificationService>()));

        // Register command handlers
        RegisterCommandHandlers(container);
    }

    /// <summary>
    /// Registers all command handlers with the container
    /// </summary>
    private static void RegisterCommandHandlers(ServiceContainer container)
    {
        var logger = container.Resolve<ILogger>();
        var paletteManager = container.Resolve<IPaletteManager>();

        // Create and register command handlers directly
        var showCommand = new ShowPaletteCommandHandler(paletteManager, logger);
        var hideCommand = new HidePaletteCommandHandler(paletteManager, logger);
        var toggleCommand = new TogglePaletteCommandHandler(paletteManager, logger);
        
        container.Register<ShowPaletteCommandHandler>(showCommand);
        container.Register<HidePaletteCommandHandler>(hideCommand);
        container.Register<TogglePaletteCommandHandler>(toggleCommand);

        // Create help command with list of other commands
        var commands = new List<ICommandHandler> { showCommand, hideCommand, toggleCommand };
        var helpCommand = new HelpCommandHandler(logger, commands);
        container.Register<HelpCommandHandler>(helpCommand);
    }
}