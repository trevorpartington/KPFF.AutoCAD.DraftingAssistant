using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using Microsoft.Extensions.DependencyInjection;
using System.Text;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Test command for Phase 1: Read-only block discovery
/// This command lists all construction note blocks in the current drawing
/// </summary>
public class TestPhase1Command
{
    [CommandMethod("TESTPHASE1")]
    public void TestBlockDiscovery()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== PHASE 1 TEST: Construction Note Block Discovery ===\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n\n");

            // Get the service provider from the extension application
            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not initialized\n");
                return;
            }

            // Get the logger
            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
            
            // CRASH FIX: Manually create block manager instead of DI to avoid premature instantiation
            var blockManager = new CurrentDrawingBlockManager(logger);

            // Option 1: Test specific layout
            ed.WriteMessage("Enter layout name to test (or press Enter for all layouts): ");
            PromptResult layoutResult = ed.GetString("\nLayout name: ");
            
            if (layoutResult.Status == PromptStatus.OK && !string.IsNullOrWhiteSpace(layoutResult.StringResult))
            {
                // Test specific layout
                string layoutName = layoutResult.StringResult;
                ed.WriteMessage($"\nSearching for construction note blocks in layout '{layoutName}'...\n");
                
                var blocks = blockManager.GetConstructionNoteBlocks(layoutName);
                
                if (blocks.Count == 0)
                {
                    ed.WriteMessage($"No construction note blocks found in layout '{layoutName}'\n");
                }
                else
                {
                    ed.WriteMessage($"Found {blocks.Count} construction note blocks:\n");
                    ed.WriteMessage("----------------------------------------\n");
                    
                    foreach (var block in blocks)
                    {
                        ed.WriteMessage($"Block: {block.BlockName}\n");
                        ed.WriteMessage($"  Number: {(block.Number > 0 ? block.Number.ToString() : "(empty)")}\n");
                        ed.WriteMessage($"  Note: {TruncateString(block.Note, 60)}\n");
                        ed.WriteMessage($"  Visible: {block.IsVisible}\n");
                        ed.WriteMessage("----------------------------------------\n");
                    }
                }
            }
            else
            {
                // Test all layouts
                ed.WriteMessage("\nSearching for construction note blocks in all layouts...\n");
                
                var allBlocks = blockManager.GetAllConstructionNoteBlocks();
                
                if (allBlocks.Count == 0)
                {
                    ed.WriteMessage("No construction note blocks found in any layout\n");
                }
                else
                {
                    ed.WriteMessage($"Found construction note blocks in {allBlocks.Count} layout(s):\n");
                    ed.WriteMessage("========================================\n");
                    
                    foreach (var layoutGroup in allBlocks)
                    {
                        ed.WriteMessage($"\nLayout: {layoutGroup.Key}\n");
                        ed.WriteMessage($"Found {layoutGroup.Value.Count} blocks:\n");
                        ed.WriteMessage("----------------------------------------\n");
                        
                        foreach (var block in layoutGroup.Value)
                        {
                            ed.WriteMessage($"  {block.BlockName}: ");
                            
                            if (block.Number > 0)
                            {
                                ed.WriteMessage($"#{block.Number} ");
                            }
                            
                            if (!string.IsNullOrEmpty(block.Note))
                            {
                                ed.WriteMessage($"- {TruncateString(block.Note, 40)}");
                            }
                            
                            if (!block.IsVisible)
                            {
                                ed.WriteMessage(" [HIDDEN]");
                            }
                            
                            ed.WriteMessage("\n");
                        }
                    }
                    
                    ed.WriteMessage("========================================\n");
                    
                    // Summary
                    int totalBlocks = allBlocks.Sum(kvp => kvp.Value.Count);
                    int visibleBlocks = allBlocks.Sum(kvp => kvp.Value.Count(b => b.IsVisible));
                    int hiddenBlocks = totalBlocks - visibleBlocks;
                    
                    ed.WriteMessage($"\nSUMMARY:\n");
                    ed.WriteMessage($"  Total Layouts: {allBlocks.Count}\n");
                    ed.WriteMessage($"  Total Blocks: {totalBlocks}\n");
                    ed.WriteMessage($"  Visible: {visibleBlocks}\n");
                    ed.WriteMessage($"  Hidden: {hiddenBlocks}\n");
                }
            }
            
            ed.WriteMessage("\n=== Phase 1 Test Complete ===\n");
            ed.WriteMessage("Check the log file for detailed debug information.\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
            ed.WriteMessage($"Stack Trace:\n{ex.StackTrace}\n");
        }
    }

    [CommandMethod("TESTPHASE1DETAIL")]
    public void TestBlockDiscoveryDetailed()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== DETAILED PHASE 1 TEST ===\n");
            
            // Prompt for a specific block name
            PromptResult blockResult = ed.GetString("\nEnter specific block name (e.g., 101-NT01): ");
            if (blockResult.Status != PromptStatus.OK)
                return;
                
            string targetBlock = blockResult.StringResult;
            ed.WriteMessage($"\nSearching for block: {targetBlock}\n");

            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
            
            // CRASH FIX: Manually create block manager instead of DI to avoid premature instantiation
            var blockManager = new CurrentDrawingBlockManager(logger);

            // CRASH FIX: Updated for new NT## pattern - need to prompt for layout name
            // since block names no longer contain layout info
            PromptResult layoutPrompt = ed.GetString("\nEnter layout name to search in (e.g., ABC-101): ");
            if (layoutPrompt.Status != PromptStatus.OK)
                return;
                
            string layoutName = layoutPrompt.StringResult;
            
            var blocks = blockManager.GetConstructionNoteBlocks(layoutName);
            var foundBlock = blocks.FirstOrDefault(b => b.BlockName.Equals(targetBlock, StringComparison.OrdinalIgnoreCase));
            
            if (foundBlock != null)
            {
                ed.WriteMessage($"\nBlock Found!\n");
                ed.WriteMessage($"====================\n");
                ed.WriteMessage($"Block Name: {foundBlock.BlockName}\n");
                ed.WriteMessage($"Number Attribute: '{(foundBlock.Number > 0 ? foundBlock.Number.ToString() : "(empty)")}'\n");
                ed.WriteMessage($"Note Attribute: '{foundBlock.Note ?? "(null)"}'\n");
                ed.WriteMessage($"Is Visible: {foundBlock.IsVisible}\n");
                ed.WriteMessage($"====================\n");
                
                // Additional details
                if (!string.IsNullOrEmpty(foundBlock.Note))
                {
                    ed.WriteMessage($"\nFull Note Text:\n{foundBlock.Note}\n");
                    ed.WriteMessage($"\nNote Length: {foundBlock.Note.Length} characters\n");
                }
            }
            else
            {
                ed.WriteMessage($"\nBlock '{targetBlock}' not found in layout '{layoutName}'\n");
                
                if (blocks.Count > 0)
                {
                    ed.WriteMessage($"\nAvailable blocks in layout '{layoutName}':\n");
                    foreach (var block in blocks)
                    {
                        ed.WriteMessage($"  - {block.BlockName}\n");
                    }
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
        }
    }

    private string TruncateString(string str, int maxLength)
    {
        if (string.IsNullOrEmpty(str))
            return "(empty)";
            
        if (str.Length <= maxLength)
            return str;
            
        return str.Substring(0, maxLength) + "...";
    }
}