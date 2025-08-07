using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

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
        ILogger? logger = null;
        try
        {
            // Ensure plugin is initialized
            EnsurePluginInitialized();
            
            var container = ServiceContainer.Instance;
            logger = container.IsRegistered<ILogger>() ? container.Resolve<ILogger>() : null;
            
            var command = container.Resolve<T>();
            command.Execute();
        }
        catch (System.Exception ex)
        {
            logger?.LogError($"Error executing command {typeof(T).Name}", ex);
            
            // Show user-friendly error message
            ShowUserError($"Command failed: {ex.Message}");
            
            // Don't rethrow - we don't want to crash AutoCAD
        }
    }

    /// <summary>
    /// Ensures the plugin is properly initialized before executing commands
    /// </summary>
    private static void EnsurePluginInitialized()
    {
        try
        {
            // First validate AutoCAD document context
            ValidateDocumentContext();
            
            var container = ServiceContainer.Instance;
            
            // Check if services are registered
            if (!container.IsRegistered<IPaletteManager>())
            {
                throw new InvalidOperationException("Plugin not properly initialized - services not registered");
            }
            
            // Force initialization if needed
            var paletteManager = container.Resolve<IPaletteManager>();
            if (!paletteManager.IsInitialized)
            {
                var logger = container.IsRegistered<ILogger>() ? container.Resolve<ILogger>() : null;
                logger?.LogWarning("Plugin not initialized - attempting to initialize now");
                
                // Try to force initialization
                paletteManager.Initialize();
            }
        }
        catch (System.Exception ex)
        {
            throw new InvalidOperationException("Failed to initialize plugin", ex);
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
            throw new InvalidOperationException("AutoCAD document context is not valid for command execution", ex);
        }
    }

    /// <summary>
    /// Shows user-friendly error messages
    /// </summary>
    private static void ShowUserError(string message)
    {
        try
        {
            if (ServiceContainer.Instance.IsRegistered<INotificationService>())
            {
                var notificationService = ServiceContainer.Instance.Resolve<INotificationService>();
                notificationService.ShowError("KPFF Drafting Assistant", message);
            }
            else
            {
                // Fallback to AutoCAD's built-in error reporting
                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog($"KPFF Drafting Assistant Error\n\n{message}");
            }
        }
        catch
        {
            // If all else fails, at least output to debug
            System.Diagnostics.Debug.WriteLine($"KPFF Drafting Assistant Error: {message}");
        }
    }
}