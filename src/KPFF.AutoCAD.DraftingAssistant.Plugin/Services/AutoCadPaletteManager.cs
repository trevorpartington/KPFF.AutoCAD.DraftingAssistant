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
    public bool IsInitialized => _paletteSet != null;

    public void Initialize()
    {
        if (_paletteSet != null)
        {
            _logger.LogWarning("Palette already initialized");
            return;
        }

        try
        {
            _logger.LogInformation("Initializing palette set");
            
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

            var draftingAssistantControl = new DraftingAssistantControl();
            elementHost.Child = draftingAssistantControl;

            _paletteSet.Add("Drafting Assistant", elementHost);
            _logger.LogInformation("Palette set initialized successfully");
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Failed to initialize palette", ex);
            _notificationService.ShowError("Initialization Error", $"Failed to create interface: {ex.Message}");
            throw;
        }
    }

    public void Show()
    {
        try
        {
            // Ensure we're in a valid document context
            if (Application.DocumentManager.MdiActiveDocument == null)
            {
                _logger.LogWarning("No active document - cannot show palette");
                return;
            }

            if (_paletteSet == null)
            {
                Initialize();
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