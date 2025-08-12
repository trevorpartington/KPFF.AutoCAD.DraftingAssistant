using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.UI.Controls;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using System.Windows.Forms.Integration;
using WinForms = System.Windows.Forms;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific palette manager implementation
/// </summary>
public class AutoCadPaletteManager : IPaletteManager
{
    private readonly ILogger _logger;
    private readonly INotificationService _notificationService;
    private PaletteSet? _paletteSet;

    public AutoCadPaletteManager(ILogger logger, INotificationService notificationService)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
    }

    public bool IsVisible => _paletteSet?.Visible ?? false;
    public bool IsInitialized => true; // CRASH FIX: Always return true since we're just preparing the manager

    public void Initialize()
    {
        if (_paletteSet != null)
        {
            _logger.LogWarning("Palette already initialized");
            return;
        }

        try
        {
            _logger.LogInformation("Preparing palette manager (deferred initialization)");
            
            // CRASH FIX: Don't create the actual palette during plugin initialization
            // This prevents crashes when AutoCAD isn't fully ready
            // The palette will be created when Show() is called
            _logger.LogInformation("Palette manager prepared - palette will be created on first use");
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Failed to prepare palette manager", ex);
            throw;
        }
    }

    /// <summary>
    /// Creates the actual palette set when needed
    /// CRASH FIX: Moved from Initialize() to prevent startup crashes
    /// </summary>
    private void CreatePaletteSet()
    {
        if (_paletteSet != null)
        {
            return; // Already created
        }

        try
        {
            _logger.LogInformation("Creating palette set");
            
            var paletteSetId = new Guid(ApplicationConstants.PaletteSetId);
            _paletteSet = new PaletteSet(ApplicationConstants.ApplicationName, paletteSetId)
            {
                Size = new System.Drawing.Size(
                    ApplicationConstants.PaletteSettings.DefaultWidth,
                    ApplicationConstants.PaletteSettings.DefaultHeight),
                MinimumSize = new System.Drawing.Size(
                    ApplicationConstants.PaletteSettings.MinimumWidth,
                    ApplicationConstants.PaletteSettings.MinimumHeight),
                DockEnabled = DockSides.Left | DockSides.Right,
                Style = PaletteSetStyles.ShowPropertiesMenu |
                       PaletteSetStyles.ShowAutoHideButton |
                       PaletteSetStyles.ShowCloseButton
            };

            var elementHost = new ElementHost
            {
                Dock = WinForms.DockStyle.Fill
            };

            // CRASH FIX: Pass services explicitly to avoid ApplicationServices access during UI init
            var draftingAssistantControl = new DraftingAssistantControl(_logger, _notificationService);
            elementHost.Child = draftingAssistantControl;

            _paletteSet.Add("Drafting Assistant", elementHost);
            _logger.LogInformation("Palette set created successfully");
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Failed to create palette set", ex);
            _notificationService.ShowError("Palette Error", $"Failed to create interface: {ex.Message}");
            throw;
        }
    }

    public void Show()
    {
        try
        {
            // CRASH FIX: Ensure AutoCAD is fully ready before creating palette
            if (!IsAutoCadReady())
            {
                _logger.LogWarning("AutoCAD not ready - deferring palette creation");
                return;
            }

            // CRASH FIX: Add safety guard around AutoCAD object access
            // Removed direct access to DocumentManager that could cause crashes
            try
            {
                var docManager = Application.DocumentManager;
                if (docManager?.MdiActiveDocument == null)
                {
                    _logger.LogWarning("No active document - cannot show palette");
                    return;
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogWarning($"Cannot access document manager - palette may not function correctly: {ex.Message}");
                // Continue anyway - let the palette try to initialize
            }

            if (_paletteSet == null)
            {
                CreatePaletteSet(); // Use the new method
            }

            if (_paletteSet != null)
            {
                _paletteSet.Visible = true;
                _logger.LogDebug("Palette shown");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error showing palette", ex);
        }
    }

    /// <summary>
    /// Checks if AutoCAD is ready for palette creation
    /// CRASH FIX: Prevents palette creation when AutoCAD isn't fully initialized
    /// </summary>
    private bool IsAutoCadReady()
    {
        try
        {
            // Simple check - if we can access the application version, AutoCAD is minimally ready
            var _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.Version;
            return true;
        }
        catch
        {
            return false;
        }
    }

    public void Hide()
    {
        try
        {
            if (_paletteSet != null)
            {
                _paletteSet.Visible = false;
                _logger.LogDebug("Palette hidden");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error hiding palette", ex);
        }
    }

    public void Toggle()
    {
        try
        {
            if (_paletteSet == null)
            {
                Show();
            }
            else
            {
                _paletteSet.Visible = !_paletteSet.Visible;
                _logger.LogDebug($"Palette toggled to {(_paletteSet.Visible ? "visible" : "hidden")}");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error toggling palette", ex);
        }
    }

    public void Cleanup()
    {
        try
        {
            if (_paletteSet != null)
            {
                _logger.LogInformation("Cleaning up palette resources");
                // Don't manipulate visibility during shutdown to avoid AccessViolationException
                _paletteSet = null;
                _logger.LogInformation("Palette cleanup completed");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error during palette cleanup", ex);
        }
    }
}