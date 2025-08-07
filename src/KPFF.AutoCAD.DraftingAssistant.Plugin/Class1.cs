using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(KPFF.AutoCAD.DraftingAssistant.Plugin.Commands))]

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

public class Commands
{
    [CommandMethod("KPFFSTART")]
    public void StartDraftingAssistant()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        ed.WriteMessage("\nKPFF AutoCAD Drafting Assistant started!\n");
        ed.WriteMessage("Plugin is ready for use.\n");
        
        // TODO: Launch UI window or show main interface
    }

    [CommandMethod("KPFFHELP")]
    public void ShowHelp()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        ed.WriteMessage("\n=== KPFF AutoCAD Drafting Assistant Help ===\n");
        ed.WriteMessage("Available Commands:\n");
        ed.WriteMessage("KPFFSTART - Start the drafting assistant\n");
        ed.WriteMessage("KPFFHELP  - Show this help message\n");
        ed.WriteMessage("============================================\n");
    }
}
