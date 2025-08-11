using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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

    /// <summary>
    /// Phase 1 Test: Construction Note Block Discovery
    /// Lists all construction note blocks in the current drawing
    /// </summary>
    [CommandMethod("TESTPHASE1")]
    public void TestPhase1BlockDiscovery()
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
                    
                    // If no blocks found, show available layouts for debugging
                    ed.WriteMessage("\n--- Available Layouts ---\n");
                    try
                    {
                        Database db = doc.Database;
                        using (Transaction layoutTr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = layoutTr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            
                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                                {
                                    ed.WriteMessage($"  - '{entry.Key}'\n");
                                }
                            }
                            layoutTr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"Could not list available layouts: {ex.Message}\n");
                    }
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

    /// <summary>
    /// Phase 1 Detailed Test: Specific block inspection
    /// </summary>
    [CommandMethod("TESTPHASE1DETAIL")]
    public void TestPhase1BlockDiscoveryDetailed()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== DETAILED PHASE 1 TEST ===\n");
            
            // Prompt for a specific block name
            PromptResult blockResult = ed.GetString("\nEnter specific block name (e.g., NT01): ");
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

    /// <summary>
    /// Phase 2 Test: Single Construction Note Block Update
    /// Tests the ability to safely modify a single block's attributes and visibility
    /// </summary>
    [CommandMethod("TESTPHASE2")]
    public void TestPhase2BlockUpdate()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== PHASE 2 TEST: Single Block Update ===\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n\n");

            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not initialized\n");
                return;
            }

            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
            var blockManager = new CurrentDrawingBlockManager(logger);

            // Get user inputs
            PromptResult layoutResult = ed.GetString("\nEnter layout name (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
            string layoutName = layoutResult.StringResult;

            PromptResult blockResult = ed.GetString("\nEnter block name (e.g., NT01): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
            string blockName = blockResult.StringResult;

            PromptIntegerOptions numberOptions = new PromptIntegerOptions("\nEnter note number (1-999): ");
            numberOptions.AllowNegative = false;
            numberOptions.AllowZero = false;
            numberOptions.LowerLimit = 1;
            numberOptions.UpperLimit = 999;
            PromptIntegerResult numberResult = ed.GetInteger(numberOptions);
            if (numberResult.Status != PromptStatus.OK)
                return;
            int noteNumber = numberResult.Value;

            PromptResult noteResult = ed.GetString($"\nEnter note text for #{noteNumber}: ");
            if (noteResult.Status != PromptStatus.OK)
                return;
            string noteText = noteResult.StringResult;

            PromptKeywordOptions visibilityOptions = new PromptKeywordOptions("\nMake block visible? ");
            visibilityOptions.Keywords.Add("Yes");
            visibilityOptions.Keywords.Add("No");
            visibilityOptions.Keywords.Default = "Yes";
            PromptResult visibilityResult = ed.GetKeywords(visibilityOptions);
            if (visibilityResult.Status != PromptStatus.OK)
                return;
            bool makeVisible = visibilityResult.StringResult.Equals("Yes", StringComparison.OrdinalIgnoreCase);

            ed.WriteMessage($"\n--- Update Summary ---\n");
            ed.WriteMessage($"Layout: {layoutName}\n");
            ed.WriteMessage($"Block: {blockName}\n");
            ed.WriteMessage($"Number: {noteNumber}\n");
            ed.WriteMessage($"Note: {TruncateString(noteText, 50)}\n");
            ed.WriteMessage($"Visible: {(makeVisible ? "Yes" : "No")}\n");
            ed.WriteMessage($"--- Starting Update ---\n");

            // Perform the update
            bool success = blockManager.UpdateConstructionNoteBlock(layoutName, blockName, noteNumber, noteText, makeVisible);

            if (success)
            {
                ed.WriteMessage("\n✓ UPDATE SUCCESSFUL!\n");
                ed.WriteMessage("The block has been updated with the new data.\n");
                
                // Verify the update by reading it back
                ed.WriteMessage("\n--- Verification ---\n");
                var blocks = blockManager.GetConstructionNoteBlocks(layoutName);
                var updatedBlock = blocks.FirstOrDefault(b => b.BlockName.Equals(blockName, StringComparison.OrdinalIgnoreCase));
                
                if (updatedBlock != null)
                {
                    ed.WriteMessage($"Verified Block: {updatedBlock.BlockName}\n");
                    ed.WriteMessage($"  Number: {(updatedBlock.Number > 0 ? updatedBlock.Number.ToString() : "(empty)")}\n");
                    ed.WriteMessage($"  Note: {TruncateString(updatedBlock.Note, 60)}\n");
                    ed.WriteMessage($"  Visible: {updatedBlock.IsVisible}\n");
                }
                else
                {
                    ed.WriteMessage("WARNING: Could not verify the update\n");
                }
            }
            else
            {
                ed.WriteMessage("\n✗ UPDATE FAILED!\n");
                ed.WriteMessage("The block could not be updated. Check the debug output for details.\n");
            }

            ed.WriteMessage("\n=== Phase 2 Test Complete ===\n");
            ed.WriteMessage("Check the Visual Studio Debug Output for detailed information.\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
            ed.WriteMessage($"Stack Trace:\n{ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Phase 2 Reset Test: Resets a block back to empty state
    /// Useful for testing multiple iterations
    /// </summary>
    [CommandMethod("TESTPHASE2RESET")]
    public void TestPhase2BlockReset()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== PHASE 2 RESET: Clear Block Data ===\n");

            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not initialized\n");
                return;
            }

            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
            var blockManager = new CurrentDrawingBlockManager(logger);

            // Get user inputs
            PromptResult layoutResult = ed.GetString("\nEnter layout name (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
            string layoutName = layoutResult.StringResult;

            PromptResult blockResult = ed.GetString("\nEnter block name to reset (e.g., NT01): ");
            if (blockResult.Status != PromptStatus.OK)
                return;
            string blockName = blockResult.StringResult;

            ed.WriteMessage($"\nResetting block {blockName} in layout {layoutName}...\n");

            // Reset the block (empty values, hidden)
            bool success = blockManager.UpdateConstructionNoteBlock(layoutName, blockName, 0, "", false);

            if (success)
            {
                ed.WriteMessage("✓ RESET SUCCESSFUL! Block is now empty and hidden.\n");
            }
            else
            {
                ed.WriteMessage("✗ RESET FAILED!\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Diagnostic command to list all available layouts in the current drawing
    /// Helps debug layout access issues
    /// </summary>
    [CommandMethod("TESTLAYOUTS")]
    public void TestLayouts()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== LAYOUT DIAGNOSTIC ===\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n");
            ed.WriteMessage($"Drawing Path: {doc.Database.Filename}\n\n");

            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the layout dictionary
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    
                    ed.WriteMessage($"Total layouts found: {layoutDict.Count}\n");
                    ed.WriteMessage("==========================================\n");

                    int layoutCount = 0;
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        layoutCount++;
                        string layoutName = entry.Key;
                        ObjectId layoutId = entry.Value;
                        
                        try
                        {
                            Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                            string tabName = layout.LayoutName;
                            bool isModelSpace = layoutName.Equals("Model", StringComparison.OrdinalIgnoreCase);
                            
                            ed.WriteMessage($"{layoutCount}. Dictionary Key: \"{layoutName}\"\n");
                            ed.WriteMessage($"   Tab Name: \"{tabName}\"\n");
                            ed.WriteMessage($"   Type: {(isModelSpace ? "Model Space" : "Paper Space")}\n");
                            ed.WriteMessage($"   ObjectId: {layoutId}\n");
                            
                            // Check if this layout has construction note blocks
                            if (!isModelSpace)
                            {
                                var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
                                var logger = serviceProvider?.GetService<Core.Interfaces.ILogger>();
                                var blockManager = new CurrentDrawingBlockManager(logger);
                                
                                var blocks = blockManager.GetConstructionNoteBlocks(layoutName);
                                ed.WriteMessage($"   Construction Note Blocks: {blocks.Count}\n");
                                
                                if (blocks.Count > 0)
                                {
                                    foreach (var block in blocks)
                                    {
                                        ed.WriteMessage($"     - {block.BlockName} (Visible: {block.IsVisible})\n");
                                    }
                                }
                            }
                            
                            ed.WriteMessage("------------------------------------------\n");
                        }
                        catch (System.Exception layoutEx)
                        {
                            ed.WriteMessage($"{layoutCount}. Dictionary Key: \"{layoutName}\" [ERROR: {layoutEx.Message}]\n");
                            ed.WriteMessage("------------------------------------------\n");
                        }
                    }
                    
                    ed.WriteMessage("\n=== LAYOUT TEST RESULTS ===\n");
                    ed.WriteMessage($"Expected Layouts: ABC-101, ABC-102\n");
                    
                    // Test specific layout access
                    bool abc101Found = layoutDict.Contains("ABC-101");
                    bool abc102Found = layoutDict.Contains("ABC-102");
                    
                    ed.WriteMessage($"ABC-101 found: {abc101Found}\n");
                    ed.WriteMessage($"ABC-102 found: {abc102Found}\n");
                    
                    // Try case variations
                    ed.WriteMessage("\n--- Case Variation Tests ---\n");
                    string[] variations = { "abc-101", "ABC-101", "Abc-101", "ABC101", "abc101" };
                    foreach (string variation in variations)
                    {
                        bool found = layoutDict.Contains(variation);
                        ed.WriteMessage($"'{variation}': {found}\n");
                    }
                    
                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"Error accessing layout dictionary: {ex.Message}\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
            ed.WriteMessage($"Stack Trace:\n{ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Utility method to truncate strings for display
    /// </summary>
    private static string TruncateString(string str, int maxLength)
    {
        if (string.IsNullOrEmpty(str))
            return "(empty)";
            
        if (str.Length <= maxLength)
            return str;
            
        return str.Substring(0, maxLength) + "...";
    }
}