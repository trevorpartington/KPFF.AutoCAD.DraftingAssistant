using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Threading.Tasks;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Test command that demonstrates ProjectWise compatibility
/// </summary>
public class ProjectWiseTestCommand
{
    [CommandMethod("PWTEST")]
    public async void ProjectWiseTest()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\nProjectWise Test Command Starting...\n");
            
            // Trigger ProjectWise initialization if not already done
            await ProjectWiseFix.TriggerProjectWiseInitialization();
            
            ed.WriteMessage("ProjectWise initialization triggered.\n");
            ed.WriteMessage("Plugin should now be safe to use with ProjectWise.\n");
            ed.WriteMessage("Try opening a drawing from ProjectWise to test.\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }
}