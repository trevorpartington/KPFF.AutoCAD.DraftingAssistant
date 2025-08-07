using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using KPFF.AutoCAD.DraftingAssistant.UI.Controls;
using System.Windows.Forms.Integration;
using WinForms = System.Windows.Forms;

[assembly: CommandClass(typeof(KPFF.AutoCAD.DraftingAssistant.Plugin.Commands))]
[assembly: ExtensionApplication(typeof(KPFF.AutoCAD.DraftingAssistant.Plugin.PluginApplication))]

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

public class PluginApplication : IExtensionApplication
{
    private static PaletteSet? _paletteSet;

    public void Initialize()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("KPFF Drafting Assistant Plugin initializing...");
            CreatePalette();
            System.Diagnostics.Debug.WriteLine("KPFF Drafting Assistant Plugin initialized successfully");
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error initializing plugin: {ex.Message}");
        }
    }

    public void Terminate()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("KPFF Drafting Assistant Plugin terminating...");
            CleanupPalette();
            System.Diagnostics.Debug.WriteLine("KPFF Drafting Assistant Plugin terminated successfully");
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error terminating plugin: {ex.Message}");
        }
    }

    private static void CreatePalette()
    {
        if (_paletteSet != null) return;
        try
        {
            // Use a GUID to ensure consistent identity for AutoCAD persistence
            var paletteSetId = new System.Guid("B8E8A3D4-7C5E-4E2F-8D9A-1F2E3B4C5A6D");
            _paletteSet = new PaletteSet("KPFF Drafting Assistant", paletteSetId)
            {
                Size = new System.Drawing.Size(350, 600),
                MinimumSize = new System.Drawing.Size(300, 400),
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
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error creating drafting assistant palette: {ex.Message}");
            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog($"Error creating interface: {ex.Message}");
        }
    }

    public static void ShowPalette()
    {
        if (_paletteSet == null)
        {
            CreatePalette();
        }
        if (_paletteSet != null)
            _paletteSet.Visible = true;
    }

    public static void HidePalette()
    {
        if (_paletteSet != null)
            _paletteSet.Visible = false;
    }

    public static void TogglePalette()
    {
        if (_paletteSet == null)
        {
            ShowPalette();
        }
        else
        {
            _paletteSet.Visible = !_paletteSet.Visible;
        }
    }

    /// <summary>
    /// Cleanup palette resources - called during plugin termination
    /// </summary>
    public static void CleanupPalette()
    {
        try
        {
            // Don't try to manipulate the palette during AutoCAD shutdown
            // Just clear our reference and let AutoCAD handle cleanup
            if (_paletteSet != null)
            {
                try
                {
                    // Don't set Visible = false during shutdown - causes AccessViolationException
                    // AutoCAD handles this cleanup automatically
                    System.Diagnostics.Debug.WriteLine("Palette reference cleared for cleanup");
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error during palette cleanup: {ex.Message}");
                }
                finally
                {
                    _paletteSet = null;
                }
            }
        }
        catch (System.Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in CleanupPalette: {ex.Message}");
        }
    }
}

public class Commands
{
    [CommandMethod("DRAFTINGASSISTANT")]
    public void ShowDraftingAssistant()
    {
        PluginApplication.ShowPalette();
    }

    [CommandMethod("HIDEDRAFTINGASSISTANT")]
    public void HideDraftingAssistant()
    {
        PluginApplication.HidePalette();
    }

    [CommandMethod("TOGGLEDRAFTINGASSISTANT")]
    public void ToggleDraftingAssistant()
    {
        PluginApplication.TogglePalette();
    }

    [CommandMethod("KPFFSTART")]
    public void StartDraftingAssistant()
    {
        PluginApplication.ShowPalette();
    }

    [CommandMethod("KPFFHELP")]
    public void ShowHelp()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        ed.WriteMessage("\n=== KPFF AutoCAD Drafting Assistant Help ===\n");
        ed.WriteMessage("Available Commands:\n");
        ed.WriteMessage("  DRAFTINGASSISTANT    - Show the drafting assistant palette\n");
        ed.WriteMessage("  HIDEDRAFTINGASSISTANT - Hide the drafting assistant palette\n");
        ed.WriteMessage("  TOGGLEDRAFTINGASSISTANT- Toggle palette visibility\n");
        ed.WriteMessage("  KPFFSTART             - Start the drafting assistant (alias)\n");
        ed.WriteMessage("  KPFFHELP              - Show this help message\n");
        ed.WriteMessage("============================================\n");
    }
}