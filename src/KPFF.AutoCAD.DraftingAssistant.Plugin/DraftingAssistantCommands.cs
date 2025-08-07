using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using IServiceProvider = KPFF.AutoCAD.DraftingAssistant.Core.Interfaces.IServiceProvider;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// AutoCAD command methods that delegate to command handlers
/// </summary>
public class DraftingAssistantCommands
{
    [CommandMethod(CommandNames.DraftingAssistant)]
    public void ShowDraftingAssistant()
    {
        ExecuteCommand<ShowPaletteCommandHandler>();
    }

    [CommandMethod(CommandNames.HideDraftingAssistant)]
    public void HideDraftingAssistant()
    {
        ExecuteCommand<HidePaletteCommandHandler>();
    }

    [CommandMethod(CommandNames.ToggleDraftingAssistant)]
    public void ToggleDraftingAssistant()
    {
        ExecuteCommand<TogglePaletteCommandHandler>();
    }

    [CommandMethod(CommandNames.KpffStart)]
    public void StartDraftingAssistant()
    {
        // KPFFSTART is an alias for the main command
        ExecuteCommand<ShowPaletteCommandHandler>();
    }

    [CommandMethod(CommandNames.KpffHelp)]
    public void ShowHelp()
    {
        ExecuteCommand<HelpCommandHandler>();
    }

    /// <summary>
    /// Generic method to execute command handlers safely with robust error handling
    /// </summary>
    private static void ExecuteCommand<T>() where T : class, ICommandHandler
    {
        var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
        var logger = serviceProvider?.GetOptionalService<ILogger>();
        
        ExceptionHandler.TryExecute(
            action: () =>
            {
                // Ensure plugin is initialized
                EnsurePluginInitialized(serviceProvider, logger);
                
                if (serviceProvider == null)
                {
                    throw new InvalidOperationException("Service provider not available - plugin may not be initialized");
                }
                
                var command = serviceProvider.GetService<T>();
                command.Execute();
            },
            logger: logger ?? new DebugLogger(),
            context: $"Command Execution: {typeof(T).Name}",
            showUserMessage: true
        );
    }

    /// <summary>
    /// Ensures the plugin is properly initialized before executing commands
    /// </summary>
    private static void EnsurePluginInitialized(IServiceProvider? serviceProvider, ILogger? logger)
    {
        try
        {
            // First validate AutoCAD document context
            ValidateDocumentContext();
            
            if (serviceProvider == null)
            {
                throw new InvalidOperationException("Service provider not available");
            }
            
            // Check if services are registered
            if (!serviceProvider.IsServiceRegistered<IPaletteManager>())
            {
                throw new InvalidOperationException("Plugin not properly initialized - services not registered");
            }
            
            // Force initialization if needed
            var paletteManager = serviceProvider.GetService<IPaletteManager>();
            if (!paletteManager.IsInitialized)
            {
                logger?.LogWarning("Plugin not initialized - attempting to initialize now");
                paletteManager.Initialize();
            }
        }
        catch (System.Exception ex)
        {
            ExceptionHandler.HandleException(ex, logger ?? new DebugLogger(), null, "Plugin Initialization Check", false, true);
        }
    }

    /// <summary>
    /// Validates that AutoCAD is in a proper state for command execution
    /// </summary>
    private static void ValidateDocumentContext()
    {
        try
        {
            var docManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (docManager == null)
            {
                throw new InvalidOperationException("AutoCAD DocumentManager is not available");
            }

            // Check if we have at least one document
            if (docManager.Count == 0)
            {
                throw new InvalidOperationException("No AutoCAD documents are open");
            }

            // Check if current document is accessible
            var currentDoc = docManager.MdiActiveDocument;
            if (currentDoc == null)
            {
                throw new InvalidOperationException("No active AutoCAD document");
            }

            // Validate document is in a usable state
            var _ = currentDoc.Name; // This will throw if document is not ready
        }
        catch (System.Exception ex)
        {
            ExceptionHandler.HandleException(ex, new DebugLogger(), null, "Document Context Validation", false, true);
        }
    }
}