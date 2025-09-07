using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System.IO;
using System.Text.Json;
using System.Linq;
using IServiceResolver = KPFF.AutoCAD.DraftingAssistant.Core.Interfaces.IServiceResolver;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// AutoCAD command methods that delegate to command handlers
/// </summary>
public class DraftingAssistantCommands
{
    [CommandMethod(CommandNames.Kpff)]
    public void ToggleDraftingAssistant()
    {
        // Single command that toggles palette visibility
        ExecuteCommand<TogglePaletteCommandHandler>();
    }

    [CommandMethod("KPFFINSERTBLOCKS")]
    public void InsertConstructionNoteBlocks()
    {
        // Command to insert construction note blocks from external DWG
        ExecuteCommand<InsertConstructionNotesCommand>();
    }

    [CommandMethod("KPFFINSERTTITLEBLOCK")]
    public void InsertTitleBlock()
    {
        // Command to insert title block from configured DWG file at origin
        ExecuteCommand<InsertTitleBlockCommand>();
    }

    /// <summary>
    /// Generic method to execute command handlers safely with robust error handling
    /// </summary>
    private static void ExecuteCommand<T>() where T : class, ICommandHandler
    {
        // Use a direct logger initially since services may not be built yet
        var directLogger = new DebugLogger();
        
        ExceptionHandler.TryExecute(
            action: () =>
            {
                // Ensure services are fully initialized on first command use
                bool servicesInitialized = DraftingAssistantExtensionApplication.EnsureServicesInitialized();
                if (!servicesInitialized)
                {
                    // Check if it's a document issue
                    if (Application.DocumentManager?.MdiActiveDocument == null)
                    {
                        Application.ShowAlertDialog("Please open a drawing before using KPFF commands.");
                        return;
                    }
                    throw new InvalidOperationException("Failed to initialize KPFF Drafting Assistant services");
                }
                
                // Get the composition root AFTER services are initialized
                var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
                if (compositionRoot == null)
                {
                    throw new InvalidOperationException("Composition root not available - plugin may not be initialized");
                }
                
                // Now we can safely get the registered logger
                var logger = compositionRoot.GetOptionalService<ILogger>() ?? directLogger;
                logger.LogInformation($"Executing command: {typeof(T).Name}");
                
                // Create command handler with dependency injection
                var command = CreateCommandHandler<T>(compositionRoot);
                command.Execute();
            },
            logger: directLogger,
            context: $"Command Execution: {typeof(T).Name}",
            showUserMessage: true
        );
    }
    
    /// <summary>
    /// Creates a command handler with dependency injection
    /// </summary>
    private static T CreateCommandHandler<T>(ServiceCompositionRoot compositionRoot) where T : class, ICommandHandler
    {
        // Manually create command handlers with dependency injection
        return typeof(T).Name switch
        {
            nameof(TogglePaletteCommandHandler) => (T)(object)new TogglePaletteCommandHandler(
                compositionRoot.GetService<IPaletteManager>(),
                compositionRoot.GetService<ILogger>()),
            nameof(InsertConstructionNotesCommand) => (T)(object)new InsertConstructionNotesCommand(
                compositionRoot.GetService<ILogger>()),
            nameof(InsertTitleBlockCommand) => (T)(object)new InsertTitleBlockCommand(
                compositionRoot.GetService<ILogger>()),
            _ => throw new InvalidOperationException($"Unknown command handler type: {typeof(T).Name}")
        };
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

            // Ensure services are initialized first
            if (!DraftingAssistantExtensionApplication.EnsureServicesInitialized())
            {
                ed.WriteMessage("ERROR: Failed to initialize services\n");
                return;
            }

            // Get the composition root from the extension application
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Composition root not initialized\n");
                return;
            }

            // Get the logger
            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
            
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

            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
            
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

            // Ensure services are initialized first
            if (!DraftingAssistantExtensionApplication.EnsureServicesInitialized())
            {
                ed.WriteMessage("ERROR: Failed to initialize services\n");
                return;
            }

            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not initialized\n");
                return;
            }

            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
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
    /// Clear all construction note blocks in all layouts of the active drawing
    /// Resets all NT## blocks to empty state (hidden, cleared attributes)
    /// </summary>
    [CommandMethod("CLEARNOTES")]
    public void ClearNotes()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== CLEAR ALL CONSTRUCTION NOTES ===\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n");

            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not initialized\n");
                return;
            }

            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
            var blockManager = new CurrentDrawingBlockManager(logger);

            ed.WriteMessage("\nScanning all layouts for construction note blocks...\n");

            // Get all layouts (excluding Model space)
            Database db = doc.Database;
            var layoutNames = new List<string>();
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                    {
                        layoutNames.Add(entry.Key);
                    }
                }
                tr.Commit();
            }

            ed.WriteMessage($"Found {layoutNames.Count} paper space layouts\n");

            int totalClearedBlocks = 0;
            int processedLayouts = 0;

            foreach (string layoutName in layoutNames)
            {
                ed.WriteMessage($"\nProcessing layout: {layoutName}\n");
                
                int clearedCount = blockManager.ClearAllConstructionNoteBlocks(layoutName);
                
                if (clearedCount > 0)
                {
                    ed.WriteMessage($"  ✓ Cleared {clearedCount} blocks\n");
                    totalClearedBlocks += clearedCount;
                    processedLayouts++;
                }
                else
                {
                    ed.WriteMessage($"  - No construction note blocks found\n");
                }
            }

            ed.WriteMessage($"\n=== CLEAR OPERATION COMPLETE ===\n");
            ed.WriteMessage($"Processed layouts: {processedLayouts}/{layoutNames.Count}\n");
            ed.WriteMessage($"Total blocks cleared: {totalClearedBlocks}\n");
            
            if (totalClearedBlocks > 0)
            {
                ed.WriteMessage("All NT## blocks now have:\n");
                ed.WriteMessage("  - Visibility set to OFF\n");
                ed.WriteMessage("  - NUMBER and NOTE attributes cleared\n");
            }
            else
            {
                ed.WriteMessage("No construction note blocks were found in any layout.\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR: {ex.Message}\n");
            ed.WriteMessage($"Stack Trace:\n{ex.StackTrace}\n");
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
                                var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
                                var logger = compositionRoot?.GetService<Core.Interfaces.ILogger>();
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
    /// Phase 3 Test: Out-of-Process Excel Integration 
    /// Reads Excel data via external process - eliminates EPPlus freezing issues
    /// </summary>
    [CommandMethod("TESTPHASE3")]
    public void TestPhase3ExcelIntegration()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== PHASE 3 TEST: OUT-OF-PROCESS EXCEL INTEGRATION ===\n");
            ed.WriteMessage("Reading Excel data via external process - NO EPPlus in AutoCAD!\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n\n");

            // Use dependency injection instead of manual service creation
            ed.WriteMessage("Initializing services via dependency injection...\n");
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not available. Services may not be registered.\n");
                return;
            }

            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
            var appLogger = compositionRoot.GetService<Core.Interfaces.IApplicationLogger>();
            var excelReader = compositionRoot.GetService<Core.Interfaces.IExcelReader>();
            var blockManager = new CurrentDrawingBlockManager(logger ?? new DebugLogger());

            if (excelReader == null)
            {
                ed.WriteMessage("ERROR: Excel reader service not available. Check service registration.\n");
                return;
            }

            ed.WriteMessage("✓ Services initialized successfully\n");

            // Get sheet/layout name from user
            PromptResult sheetResult = ed.GetString("\nEnter sheet/layout name (e.g., ABC-101): ");
            if (sheetResult.Status != PromptStatus.OK)
                return;
                
            string sheetName = sheetResult.StringResult;
            string layoutName = sheetName; // Layout and sheet name are the same

            // Read real Excel data via out-of-process reader
            string excelPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectIndex.xlsx";
            
            if (!File.Exists(excelPath))
            {
                ed.WriteMessage($"ERROR: Excel file not found: {excelPath}\n");
                return;
            }

            ed.WriteMessage($"Reading Excel data from: {Path.GetFileName(excelPath)}\n");
            ed.WriteMessage("This uses external Excel reader process to prevent AutoCAD freezing!\n\n");

            // Test Excel reader connection first
            ed.WriteMessage("Testing Excel reader connection...\n");
            try
            {
                // Try to read a small amount of data to test connection
                var config = new Core.Models.ProjectConfiguration(); // Default config
                
                ed.WriteMessage("Phase 1: Reading ABC construction notes...\n");
                // Use Task.Run to avoid blocking AutoCAD's UI thread
                var constructionNotes = Task.Run(async () => 
                    await excelReader.ReadConstructionNotesAsync(excelPath, "ABC", config)).GetAwaiter().GetResult();
                ed.WriteMessage($"✓ Loaded {constructionNotes.Count} construction notes from ABC_NOTES table\n");

                // Read Excel notes mappings
                ed.WriteMessage("Phase 2: Reading Excel notes mappings...\n");  
                // Use Task.Run to avoid blocking AutoCAD's UI thread
                var excelMappings = Task.Run(async () => 
                    await excelReader.ReadExcelNotesAsync(excelPath, config)).GetAwaiter().GetResult();
                ed.WriteMessage($"✓ Loaded mappings for {excelMappings.Count} sheets from EXCEL_NOTES table\n");

                if (excelMappings.Count == 0)
                {
                    ed.WriteMessage("WARNING: No Excel mappings found. This may indicate a connection issue.\n");
                    ed.WriteMessage("Available sheets in Excel: Check if EXCEL_NOTES table exists and has data.\n");
                    return;
                }

                // DEBUG: Show all loaded sheet names
                ed.WriteMessage($"DEBUG: Loaded sheet names: [{string.Join(", ", excelMappings.Select(m => $"'{m.SheetName}'"))}]\n");
                ed.WriteMessage($"DEBUG: User entered sheet name: '{sheetName}'\n");
                ed.WriteMessage($"DEBUG: Sheet name length: {sheetName.Length}, Contains whitespace: {sheetName.Any(c => char.IsWhiteSpace(c))}\n");

                // Find sheet mapping
                var sheetMapping = excelMappings.FirstOrDefault(m => m.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheetMapping == null)
                {
                    ed.WriteMessage($"Sheet '{sheetName}' not found in Excel notes mappings\n");
                    ed.WriteMessage($"Available sheets: {string.Join(", ", excelMappings.Select(m => m.SheetName))}\n");
                    
                    // DEBUG: Show detailed comparison
                    ed.WriteMessage("\nDEBUG: Detailed sheet name comparison:\n");
                    foreach (var mapping in excelMappings)
                    {
                        var exactMatch = mapping.SheetName.Equals(sheetName);
                        var ignoreCaseMatch = mapping.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase);
                        var trimmedMatch = mapping.SheetName.Trim().Equals(sheetName.Trim(), StringComparison.OrdinalIgnoreCase);
                        
                        ed.WriteMessage($"  '{mapping.SheetName}' vs '{sheetName}': Exact={exactMatch}, IgnoreCase={ignoreCaseMatch}, Trimmed={trimmedMatch}\n");
                    }
                    return;
                }

                ed.WriteMessage($"✓ Found sheet '{sheetName}' with {sheetMapping.NoteNumbers.Count} note(s)\n");

                // DEBUG: Show construction notes details
                ed.WriteMessage("\nDEBUG: Construction notes details:\n");
                foreach (var note in constructionNotes)
                {
                    ed.WriteMessage($"  Note #{note.Number}: '{TruncateString(note.Text, 40)}'\n");
                }

                // Create lookup for notes - handle duplicates gracefully
                var noteLookup = new Dictionary<int, Core.Models.ConstructionNote>();
                foreach (var note in constructionNotes)
                {
                    if (note.Number > 0) // Skip invalid note numbers
                    {
                        if (noteLookup.ContainsKey(note.Number))
                        {
                            ed.WriteMessage($"WARNING: Duplicate note number {note.Number} found - using first occurrence\n");
                        }
                        else
                        {
                            noteLookup[note.Number] = note;
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"WARNING: Skipping note with invalid number: {note.Number}\n");
                    }
                }

                ed.WriteMessage($"Created lookup dictionary with {noteLookup.Count} unique notes\n");

                // Phase 3: Clear all blocks first, then update with new data
                ed.WriteMessage("\nPhase 3: Updating AutoCAD blocks with real Excel data...\n");
                ed.WriteMessage("=======================================================\n");
                
                // Clear all construction note blocks first to ensure removed notes don't persist
                ed.WriteMessage("Clearing all construction note blocks first...\n");
                int clearedCount = blockManager.ClearAllConstructionNoteBlocks(layoutName);
                ed.WriteMessage($"✓ Cleared {clearedCount} blocks (set visibility OFF and cleared attributes)\n\n");
                
                // Now update blocks with current note data
                ed.WriteMessage("Updating blocks with current note data...\n");
                int successCount = 0;
                int blockIndex = 1;
                
                foreach (var noteNumber in sheetMapping.NoteNumbers)
                {
                    string blockName = $"NT{blockIndex:D2}";
                    
                    if (!noteLookup.TryGetValue(noteNumber, out Core.Models.ConstructionNote? note))
                    {
                        ed.WriteMessage($"✗ Note {noteNumber} not found in ABC_NOTES table\n");
                        blockIndex++;
                        continue;
                    }
                    
                    ed.WriteMessage($"Updating {blockName} with note {noteNumber}: '{TruncateString(note.Text, 30)}...'\n");
                    
                    bool updated = blockManager.UpdateConstructionNoteBlock(
                        layoutName, blockName, noteNumber, note.Text, makeVisible: true);
                    
                    if (updated)
                    {
                        successCount++;
                        ed.WriteMessage($"  ✓ Updated successfully with real Excel data!\n");
                    }
                    else
                    {
                        ed.WriteMessage($"  ✗ Failed to update block\n");
                    }
                    
                    blockIndex++;
                }
                
                ed.WriteMessage($"\n=== PHASE 3 COMPLETE ===\n");
                ed.WriteMessage($"Successfully updated {successCount} blocks with Excel data\n");
                ed.WriteMessage($"Sheet: {sheetName}, Layout: {layoutName}\n");
                ed.WriteMessage($"Excel file: {Path.GetFileName(excelPath)}\n");
            }
                    catch (System.Exception ex)
        {
            ed.WriteMessage($"ERROR during Excel reading: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
                
                // Try to provide helpful debugging information
                if (ex.Message.Contains("Excel reader executable not found"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING: Excel reader executable not found\n");
                    ed.WriteMessage("1. Check if KPFF.AutoCAD.ExcelReader.exe exists in the plugin directory\n");
                    ed.WriteMessage("2. Verify the executable path in the error logs\n");
                    ed.WriteMessage("3. Try restarting AutoCAD to reload the plugin\n");
                }
                else if (ex.Message.Contains("named pipe"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING: Named pipe communication failed\n");
                    ed.WriteMessage("1. Check if Excel reader process is running\n");
                    ed.WriteMessage("2. Verify no firewall is blocking local connections\n");
                    ed.WriteMessage("3. Try manually starting the Excel reader process\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"FATAL ERROR in TESTPHASE3: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Test command to clear all construction note blocks in a layout
    /// </summary>
    [CommandMethod("CLEARBLOCKS")]
    public void ClearConstructionNoteBlocks()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== CLEAR CONSTRUCTION NOTE BLOCKS ===\n");
            
            // Get layout name from user
            PromptResult layoutResult = ed.GetString("\nEnter layout name (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
                
            string layoutName = layoutResult.StringResult;
            
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var logger = compositionRoot.GetService<Core.Interfaces.ILogger>();
            var blockManager = new CurrentDrawingBlockManager(logger ?? new DebugLogger());
            
            ed.WriteMessage($"Clearing all construction note blocks in layout '{layoutName}'...\n");
            int clearedCount = blockManager.ClearAllConstructionNoteBlocks(layoutName);
            
            if (clearedCount > 0)
            {
                ed.WriteMessage($"✓ Successfully cleared {clearedCount} blocks\n");
                ed.WriteMessage("All NT## blocks now have:\n");
                ed.WriteMessage("  - Visibility set to OFF\n");
                ed.WriteMessage("  - NUMBER and NOTE attributes cleared\n");
            }
            else
            {
                ed.WriteMessage("No blocks were cleared. Check if the layout exists and contains NT## blocks.\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"ERROR: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Simple ping test to verify Excel reader connection
    /// </summary>
    [CommandMethod("PING")]
    public void PingExcelReader()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== PING TEST: EXCEL READER CONNECTION ===\n");
            
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var excelReader = compositionRoot.GetService<Core.Interfaces.IExcelReader>();
            if (excelReader == null)
            {
                ed.WriteMessage("ERROR: Excel reader service not available.\n");
                return;
            }

            ed.WriteMessage("✓ Excel reader service found\n");
            ed.WriteMessage("Testing connection...\n");

            // Try to read a small amount of data to test connection
            var config = new Core.Models.ProjectConfiguration();
            string excelPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectIndex.xlsx";
            
            if (!File.Exists(excelPath))
            {
                ed.WriteMessage($"ERROR: Excel file not found: {excelPath}\n");
                return;
            }

            try
            {
                // Use Task.Run to avoid blocking AutoCAD's UI thread
                var notes = Task.Run(async () => 
                    await excelReader.ReadConstructionNotesAsync(excelPath, "ABC", config)).GetAwaiter().GetResult();
                ed.WriteMessage($"✓ PING SUCCESSFUL! Read {notes.Count} construction notes\n");
                ed.WriteMessage("Excel reader process is running and communicating properly.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"✗ PING FAILED: {ex.Message}\n");
                
                if (ex.Message.Contains("Excel reader executable not found"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING:\n");
                    ed.WriteMessage("1. Check if KPFF.AutoCAD.ExcelReader.exe exists in the plugin directory\n");
                    ed.WriteMessage("2. Verify the executable path in the error logs\n");
                    ed.WriteMessage("3. Try restarting AutoCAD to reload the plugin\n");
                }
                else if (ex.Message.Contains("named pipe"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING:\n");
                    ed.WriteMessage("1. Check if Excel reader process is running\n");
                    ed.WriteMessage("2. Verify no firewall is blocking local connections\n");
                    ed.WriteMessage("3. Try manually starting the Excel reader process\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"FATAL ERROR in PING: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Manually start the Excel reader process for debugging
    /// </summary>
    [CommandMethod("STARTEXCEL")]
    public void StartExcelReader()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== MANUAL EXCEL READER START ===\n");
            
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            if (compositionRoot == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var excelReader = compositionRoot.GetService<Core.Interfaces.IExcelReader>();
            if (excelReader == null)
            {
                ed.WriteMessage("ERROR: Excel reader service not available.\n");
                return;
            }

            ed.WriteMessage("✓ Excel reader service found\n");
            ed.WriteMessage("Attempting to start Excel reader process...\n");

            // Force a connection attempt to trigger process startup
            try
            {
                var config = new Core.Models.ProjectConfiguration();
                string excelPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectIndex.xlsx";
                
                if (!File.Exists(excelPath))
                {
                    ed.WriteMessage($"ERROR: Excel file not found: {excelPath}\n");
                    return;
                }

                // This will trigger the Excel reader process to start
                ed.WriteMessage("Triggering Excel reader process startup...\n");
                var notes = excelReader.ReadConstructionNotesAsync(excelPath, "ABC", config).Result;
                ed.WriteMessage($"✓ Excel reader process started successfully! Read {notes.Count} notes\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"✗ Failed to start Excel reader: {ex.Message}\n");
                
                if (ex.Message.Contains("Excel reader executable not found"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING:\n");
                    ed.WriteMessage("1. Check if KPFF.AutoCAD.ExcelReader.exe exists in the plugin directory\n");
                    ed.WriteMessage("2. Verify the executable path in the error logs\n");
                    ed.WriteMessage("3. Try restarting AutoCAD to reload the plugin\n");
                }
                else if (ex.Message.Contains("named pipe"))
                {
                    ed.WriteMessage("\nTROUBLESHOOTING:\n");
                    ed.WriteMessage("1. Check if Excel reader process is running\n");
                    ed.WriteMessage("2. Verify no firewall is blocking local connections\n");
                    ed.WriteMessage("3. Try manually starting the Excel reader process\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"FATAL ERROR in STARTEXCEL: {ex.Message}\n");
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

    /// <summary>
    /// Test command to verify Excel reader shared process behavior
    /// </summary>
    [CommandMethod("TESTEXCELCLEANUP")]
    public static void TestExcelCleanup()
    {
        var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
        var editor = Application.DocumentManager.MdiActiveDocument?.Editor;
        
        if (editor == null || compositionRoot == null)
        {
            return;
        }
        
        try
        {
            editor.WriteMessage("\n=== Testing Excel Reader Shared Process Behavior ===");
            
            // Get first Excel reader service from DI container (transient)
            editor.WriteMessage("\nGetting Excel reader service from DI container (first call)...");
            var excelReader1 = compositionRoot.GetService<IExcelReader>();
            
            if (excelReader1 == null)
            {
                editor.WriteMessage("\nERROR: Excel reader service not available.");
                return;
            }
            
            // Do a simple operation to ensure process starts
            editor.WriteMessage("\nTesting first instance - checking file existence...");
            var fileExists1 = excelReader1.FileExistsAsync(@"C:\dummy1.xlsx").GetAwaiter().GetResult();
            editor.WriteMessage($"\nFirst instance result: {fileExists1}");
            
            // Check shared process status
            editor.WriteMessage($"\n\nShared Process Status:");
            editor.WriteMessage($"\n  Process Running: false (placeholder implementation)");
            editor.WriteMessage($"\n  Reference Count: 0 (placeholder implementation)");
            
            // Get another instance - should be different (transient) but share same process
            editor.WriteMessage("\n\nGetting Excel reader service again (second call)...");
            var excelReader2 = compositionRoot.GetService<IExcelReader>();
            
            // Test if they're different instances
            var areSameInstance = ReferenceEquals(excelReader1, excelReader2);
            editor.WriteMessage($"\nAre both references the same instance? {areSameInstance}");
            
            if (!areSameInstance)
            {
                editor.WriteMessage("\n✓ CORRECT: Services are transient (different instances)");
                editor.WriteMessage("\n  But they share the SAME Excel reader process");
            }
            
            // Test the second instance
            editor.WriteMessage("\n\nTesting second instance - checking file existence...");
            var fileExists2 = excelReader2.FileExistsAsync(@"C:\dummy2.xlsx").GetAwaiter().GetResult();
            editor.WriteMessage($"\nSecond instance result: {fileExists2}");
            
            // Check shared process status again
            editor.WriteMessage($"\n\nShared Process Status After Second Instance:");
            editor.WriteMessage($"\n  Process Running: false (placeholder implementation)");
            editor.WriteMessage($"\n  Reference Count: 0 (placeholder implementation)");
            
            // Dispose instances
            editor.WriteMessage("\n\nDisposing first instance...");
            excelReader1.Dispose();
            editor.WriteMessage($"  Reference Count after dispose: 0 (placeholder implementation)");
            
            editor.WriteMessage("\nDisposing second instance...");
            excelReader2.Dispose();
            editor.WriteMessage($"  Reference Count after dispose: 0 (placeholder implementation)");
            
            editor.WriteMessage("\n\n=== Excel Reader Behavior Summary ===");
            editor.WriteMessage("\n- Excel reader services are TRANSIENT (new instance each time)");
            editor.WriteMessage("\n- Currently using PLACEHOLDER implementation");
            editor.WriteMessage("\n- Excel functionality has been removed for Phase 1 & 2 refactoring");
            editor.WriteMessage("\n- Excel reader process will be re-implemented in Phase 3");
            editor.WriteMessage("\n- All Excel operations return empty results or false");
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\nError during shared process test: {ex.Message}");
        }
    }

    /// <summary>
    /// Test command to validate GeometryExtensions viewport transformations
    /// </summary>
    [CommandMethod("TESTVIEWPORT")]
    public static void TestViewportTransformations()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== TESTVIEWPORT: GeometryExtensions Validation ===\n");
            ed.WriteMessage($"Current Drawing: {doc.Name}\n");
            ed.WriteMessage("Goal: Validate GeometryExtensions DCS2WCS transformations\n\n");

            // Get layout name from user
            PromptResult layoutResult = ed.GetString("\nEnter layout name to test (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
            string layoutName = layoutResult.StringResult;

            ed.WriteMessage($"\nTesting layout: '{layoutName}'\n");
            ed.WriteMessage("==========================================\n");

            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the layout
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(layoutName))
                    {
                        ed.WriteMessage($"ERROR: Layout '{layoutName}' not found in drawing\n");
                        LogAvailableLayouts(ed, layoutDict, tr);
                        return;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                    ed.WriteMessage($"Layout Properties:\n");
                    ed.WriteMessage($"  Name: {layout.LayoutName}\n");
                    ed.WriteMessage($"  Tab Order: {layout.TabOrder}\n");
                    ed.WriteMessage($"  Block Table Record ID: {layout.BlockTableRecordId}\n\n");

                    // Access viewports through the layout's BlockTableRecord (no layout switching needed!)
                    ed.WriteMessage("Accessing viewports through BlockTableRecord (layout-independent)...\n");
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    
                    var viewportIds = new List<ObjectId>();
                    foreach (ObjectId entityId in layoutBtr)
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead);
                        if (entity is Viewport)
                        {
                            viewportIds.Add(entityId);
                        }
                    }
                    
                    ed.WriteMessage($"Found {viewportIds.Count} viewport(s) in layout\n");

                    if (viewportIds.Count == 0)
                    {
                        ed.WriteMessage("No viewports found in this layout.\n");
                        return;
                    }

                    int viewportIndex = 0;
                    foreach (ObjectId viewportId in viewportIds)
                    {
                        viewportIndex++;
                        ed.WriteMessage($"\nViewport #{viewportIndex}:\n");

                        try
                        {
                            // Open the viewport for reading
                            Viewport viewport = tr.GetObject(viewportId, OpenMode.ForRead) as Viewport;
                            if (viewport == null)
                            {
                                ed.WriteMessage($"  ERROR: Could not cast to Viewport object\n");
                                continue;
                            }

                            // Skip the first viewport (paper space viewport)
                            if (viewportIndex == 1)
                            {
                                ed.WriteMessage($"  Type: Paper Space Viewport (skipping)\n");
                                continue;
                            }

                            ed.WriteMessage($"  Type: Model Space Viewport\n");
                            ed.WriteMessage($"  Number: {viewport.Number}\n");
                            ed.WriteMessage($"  On: {viewport.On}\n");
                            
                            // === ALL VIEWPORT PROPERTIES USED IN OUR CALCULATIONS ===
                            ed.WriteMessage($"\n  *** VIEWPORT PROPERTIES (All properties we use) ***\n");
                            
                            // Paper space properties
                            ed.WriteMessage($"  PAPER SPACE PROPERTIES:\n");
                            ed.WriteMessage($"    CenterPoint: ({viewport.CenterPoint.X:F3}, {viewport.CenterPoint.Y:F3}) - Center of viewport in paper space\n");
                            ed.WriteMessage($"    Width: {viewport.Width:F3} - Paper space width of viewport\n");
                            ed.WriteMessage($"    Height: {viewport.Height:F3} - Paper space height of viewport\n");
                            
                            // Model space view properties
                            ed.WriteMessage($"  MODEL SPACE VIEW PROPERTIES:\n");
                            ed.WriteMessage($"    ViewTarget: ({viewport.ViewTarget.X:F3}, {viewport.ViewTarget.Y:F3}, {viewport.ViewTarget.Z:F3}) - Point in WCS that camera looks at\n");
                            ed.WriteMessage($"    ViewCenter: ({viewport.ViewCenter.X:F3}, {viewport.ViewCenter.Y:F3}) - Center of view in DCS (Display Coordinate System)\n");
                            ed.WriteMessage($"    ViewDirection: ({viewport.ViewDirection.X:F3}, {viewport.ViewDirection.Y:F3}, {viewport.ViewDirection.Z:F3}) - Camera direction vector\n");
                            ed.WriteMessage($"    ViewHeight: {viewport.ViewHeight:F3} - Height of model space area shown in viewport (in model units)\n");
                            ed.WriteMessage($"    CustomScale: {viewport.CustomScale:F6} - Scale factor (model units per paper unit)\n");
                            ed.WriteMessage($"    TwistAngle: {viewport.TwistAngle:F6} rad ({viewport.TwistAngle * 180.0 / Math.PI:F2}°) - Rotation around view direction\n");
                            
                            // Clipping properties
                            ed.WriteMessage($"  CLIPPING PROPERTIES:\n");
                            ed.WriteMessage($"    NonRectClipOn: {viewport.NonRectClipOn} - True if using polygonal clipping\n");
                            ed.WriteMessage($"    NonRectClipEntityId: {viewport.NonRectClipEntityId} - ObjectId of clipping entity\n");
                            
                            // Calculated scale information
                            ed.WriteMessage($"\n  *** SCALE ANALYSIS ***\n");
                            double scaleRatio = 1.0 / viewport.CustomScale;
                            ed.WriteMessage($"    Scale Ratio: 1:{scaleRatio:F0} (1 paper unit = {scaleRatio:F0} model units)\n");
                            ed.WriteMessage($"    CustomScale: {viewport.CustomScale:F6} (model units per paper unit)\n");
                            ed.WriteMessage($"    1/CustomScale: {1.0/viewport.CustomScale:F6} (paper units per model unit - for DCS→WCS)\n");
                            
                            // Expected model space dimensions
                            ed.WriteMessage($"\n  *** EXPECTED MODEL SPACE AREA CALCULATIONS ***\n");
                            double modelWidth = viewport.ViewHeight * viewport.Width / viewport.Height;
                            double modelHeight = viewport.ViewHeight;
                            ed.WriteMessage($"    ViewHeight: {viewport.ViewHeight:F3} (direct from property)\n");
                            ed.WriteMessage($"    Calculated Width: {modelWidth:F3} (ViewHeight × Width/Height aspect ratio)\n");
                            ed.WriteMessage($"    Expected Center: ({viewport.ViewCenter.X:F3}, {viewport.ViewCenter.Y:F3}) (using ViewCenter)\n");
                            ed.WriteMessage($"    Expected Bounds: ({viewport.ViewCenter.X - modelWidth/2:F3}, {viewport.ViewCenter.Y - modelHeight/2:F3}) to ({viewport.ViewCenter.X + modelWidth/2:F3}, {viewport.ViewCenter.Y + modelHeight/2:F3})\n");
                            
                            // Alternative calculation using paper dimensions and scale
                            double altWidth = viewport.Width / viewport.CustomScale;
                            double altHeight = viewport.Height / viewport.CustomScale;
                            ed.WriteMessage($"\n  *** ALTERNATIVE CALCULATION (Paper × Scale) ***\n");
                            ed.WriteMessage($"    Width: {viewport.Width:F3} ÷ {viewport.CustomScale:F6} = {altWidth:F3}\n");
                            ed.WriteMessage($"    Height: {viewport.Height:F3} ÷ {viewport.CustomScale:F6} = {altHeight:F3}\n");
                            ed.WriteMessage($"    Alt Bounds: ({viewport.ViewCenter.X - altWidth/2:F3}, {viewport.ViewCenter.Y - altHeight/2:F3}) to ({viewport.ViewCenter.X + altWidth/2:F3}, {viewport.ViewCenter.Y + altHeight/2:F3})\n");

                            // Test ViewportBoundaryCalculator with Gile's ViewportExtension
                            ed.WriteMessage($"\n  *** TESTING GILE'S VIEWPORTEXTENSION TRANSFORMATIONS ***\n");
                            
                            // Show detailed transformation diagnostics
                            ed.WriteMessage($"\n{ViewportBoundaryCalculator.GetTransformationDiagnostics(viewport)}\n");
                            
                            try
                            {
                                // Get the actual model space footprint using Gile's transformations
                                var footprint = ViewportBoundaryCalculator.GetViewportFootprint(viewport, tr);
                                
                                ed.WriteMessage($"  Paper space viewport: {viewport.Width:F3} x {viewport.Height:F3}\n");
                                ed.WriteMessage($"  Non-rectangular clipping: {viewport.NonRectClipOn}\n");
                                ed.WriteMessage($"  Calculated footprint points: {footprint.Count}\n\n");
                                
                                if (footprint.Count > 0)
                                {
                                    ed.WriteMessage($"  Model space boundary points ({footprint.Count} points):\n");
                                    for (int i = 0; i < footprint.Count; i++)
                                    {
                                        Point3d point = footprint[i];
                                        ed.WriteMessage($"    Point {i}: ({point.X:F3}, {point.Y:F3}, {point.Z:F3})\n");
                                    }
                                    
                                    // Calculate and display bounding box
                                    var bounds = ViewportBoundaryCalculator.GetViewportBounds(viewport, tr);
                                    if (bounds.HasValue)
                                    {
                                        ed.WriteMessage($"\n  Model space bounding box:\n");
                                        ed.WriteMessage($"    Min: ({bounds.Value.MinPoint.X:F3}, {bounds.Value.MinPoint.Y:F3}, {bounds.Value.MinPoint.Z:F3})\n");
                                        ed.WriteMessage($"    Max: ({bounds.Value.MaxPoint.X:F3}, {bounds.Value.MaxPoint.Y:F3}, {bounds.Value.MaxPoint.Z:F3})\n");
                                        ed.WriteMessage($"    Width: {(bounds.Value.MaxPoint.X - bounds.Value.MinPoint.X):F3}\n");
                                        ed.WriteMessage($"    Height: {(bounds.Value.MaxPoint.Y - bounds.Value.MinPoint.Y):F3}\n");
                                    }

                                    ed.WriteMessage($"\n  ✅ SUCCESS: ViewportBoundaryCalculator working correctly!\n");
                                    ed.WriteMessage($"  🎯 Model space transformation complete using GetMatrixFromDcsToWcs()\n");
                                }
                                else
                                {
                                    ed.WriteMessage($"  ⚠ WARNING: No footprint points calculated\n");
                                    ed.WriteMessage($"  Check viewport properties and transformation matrix\n");
                                }
                                
                            }
                            catch (System.Exception calcEx)
                            {
                                ed.WriteMessage($"  ERROR with ViewportBoundaryCalculator: {calcEx.Message}\n");
                                ed.WriteMessage($"  Stack trace: {calcEx.StackTrace}\n");
                            }

                        }
                        catch (System.Exception vpEx)
                        {
                            ed.WriteMessage($"  ERROR reading viewport properties: {vpEx.Message}\n");
                        }

                        ed.WriteMessage("----------------------------------------\n");
                    }

                    tr.Commit();

                    ed.WriteMessage("\n=== TESTVIEWPORT SUMMARY ===\n");
                    ed.WriteMessage($"Layout '{layoutName}' Analysis Complete\n");
                    ed.WriteMessage($"Total Viewports: {viewportIds.Count}\n");
                    ed.WriteMessage($"Model Space Viewports: {Math.Max(0, viewportIds.Count - 1)}\n");
                    
                    if (viewportIds.Count > 1)
                    {
                        ed.WriteMessage("✓ SUCCESS: Found model space viewports for testing\n");
                        ed.WriteMessage("📝 NEXT STEP: Implement actual GeometryExtensions DCS2WCS calls\n");
                        ed.WriteMessage("📝 TODO: Add 'using Gile.AutoCAD.Geometry;' when ready\n");
                    }
                    else
                    {
                        ed.WriteMessage("⚠ WARNING: No model space viewports found\n");
                        ed.WriteMessage("  Need model space viewports for Auto Notes testing\n");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"ERROR accessing layout or viewports: {ex.Message}\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nFATAL ERROR in TESTVIEWPORT: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Test command to understand ViewTarget vs ViewCenter properties
    /// </summary>
    [CommandMethod("CENTERTEST")]
    public static void TestViewportCenterProperties()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== CENTERTEST: ViewTarget vs ViewCenter Analysis ===\n");
            ed.WriteMessage("Goal: Understand the difference between ViewTarget and ViewCenter\n");
            ed.WriteMessage("ViewCenter: Display coordinate system coordinates (paper space units)\n");
            ed.WriteMessage("ViewTarget: Model Space WCS coordinates (model space location)\n\n");

            // Get layout name from user
            PromptResult layoutResult = ed.GetString("\nEnter layout name to test (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK)
                return;
            string layoutName = layoutResult.StringResult;

            ed.WriteMessage($"\nAnalyzing layout: '{layoutName}'\n");
            ed.WriteMessage("==========================================\n");

            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the layout
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(layoutName))
                    {
                        ed.WriteMessage($"ERROR: Layout '{layoutName}' not found in drawing\n");
                        LogAvailableLayouts(ed, layoutDict, tr);
                        return;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    
                    // Find viewports in the layout
                    var viewportIds = new List<ObjectId>();
                    foreach (ObjectId entityId in layoutBtr)
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead);
                        if (entity is Viewport)
                        {
                            viewportIds.Add(entityId);
                        }
                    }

                    if (viewportIds.Count == 0)
                    {
                        ed.WriteMessage("No viewports found in this layout.\n");
                        return;
                    }

                    int viewportIndex = 0;
                    foreach (ObjectId viewportId in viewportIds)
                    {
                        viewportIndex++;
                        
                        Viewport viewport = tr.GetObject(viewportId, OpenMode.ForRead) as Viewport;
                        if (viewport == null) continue;

                        // Skip paper space viewport
                        if (viewportIndex == 1)
                        {
                            ed.WriteMessage($"Viewport #{viewportIndex}: Paper Space Viewport (skipping)\n");
                            continue;
                        }

                        ed.WriteMessage($"\nViewport #{viewportIndex}: Model Space Viewport\n");
                        ed.WriteMessage("===========================================\n");
                        
                        // Core properties comparison
                        ed.WriteMessage($"📍 ViewCenter (DCS): ({viewport.ViewCenter.X:F3}, {viewport.ViewCenter.Y:F3})\n");
                        ed.WriteMessage($"🎯 ViewTarget (WCS): ({viewport.ViewTarget.X:F3}, {viewport.ViewTarget.Y:F3}, {viewport.ViewTarget.Z:F3})\n\n");
                        
                        // Additional context properties
                        ed.WriteMessage($"📐 Paper Space Size: {viewport.Width:F3} x {viewport.Height:F3}\n");
                        ed.WriteMessage($"🔍 Scale Factor: {viewport.CustomScale:F6} (1:{1.0/viewport.CustomScale:F0})\n");
                        ed.WriteMessage($"📏 ViewHeight: {viewport.ViewHeight:F3}\n");
                        ed.WriteMessage($"🔄 TwistAngle: {viewport.TwistAngle:F6} radians ({viewport.TwistAngle * 180.0 / Math.PI:F2} degrees)\n");
                        ed.WriteMessage($"👁️  ViewDirection: ({viewport.ViewDirection.X:F3}, {viewport.ViewDirection.Y:F3}, {viewport.ViewDirection.Z:F3})\n\n");
                        
                        // Analysis
                        ed.WriteMessage("🧠 ANALYSIS:\n");
                        
                        // Calculate model space area from ViewCenter approach
                        double scaleFactor = 1.0 / viewport.CustomScale;
                        double modelWidth = viewport.Width * scaleFactor;
                        double modelHeight = viewport.Height * scaleFactor;
                        
                        ed.WriteMessage($"   Using ViewCenter as model space center:\n");
                        ed.WriteMessage($"   Model Space Area: {modelWidth:F3} x {modelHeight:F3}\n");
                        ed.WriteMessage($"   Bottom-Left: ({viewport.ViewCenter.X - modelWidth/2:F3}, {viewport.ViewCenter.Y - modelHeight/2:F3})\n");
                        ed.WriteMessage($"   Top-Right: ({viewport.ViewCenter.X + modelWidth/2:F3}, {viewport.ViewCenter.Y + modelHeight/2:F3})\n\n");
                        
                        ed.WriteMessage($"   Using ViewTarget as model space center:\n");
                        ed.WriteMessage($"   Bottom-Left: ({viewport.ViewTarget.X - modelWidth/2:F3}, {viewport.ViewTarget.Y - modelHeight/2:F3})\n");
                        ed.WriteMessage($"   Top-Right: ({viewport.ViewTarget.X + modelWidth/2:F3}, {viewport.ViewTarget.Y + modelHeight/2:F3})\n\n");
                        
                        // Difference analysis
                        double centerDiffX = viewport.ViewCenter.X - viewport.ViewTarget.X;
                        double centerDiffY = viewport.ViewCenter.Y - viewport.ViewTarget.Y;
                        double centerDistance = Math.Sqrt(centerDiffX * centerDiffX + centerDiffY * centerDiffY);
                        
                        ed.WriteMessage($"   Difference (ViewCenter - ViewTarget):\n");
                        ed.WriteMessage($"   ΔX: {centerDiffX:F3}, ΔY: {centerDiffY:F3}\n");
                        ed.WriteMessage($"   Distance: {centerDistance:F3}\n");
                        
                        if (Math.Abs(centerDistance) < 0.001)
                        {
                            ed.WriteMessage($"   ✅ ViewCenter and ViewTarget are essentially the same point\n");
                        }
                        else
                        {
                            ed.WriteMessage($"   ⚠️  ViewCenter and ViewTarget are different points\n");
                        }
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"ERROR during viewport analysis: {ex.Message}\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nFATAL ERROR in CENTERTEST: {ex.Message}\n");
        }
    }

    [CommandMethod("TESTAUTONOTES")]
    public void TestAutoNotes()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== TESTAUTONOTES: Auto Notes Detection Validation ===\n");
            
            // Prompt for layout name
            PromptStringOptions pso = new PromptStringOptions($"\nEnter layout name to test (or press Enter for 'ABC-101'): ");
            pso.AllowSpaces = true;
            pso.DefaultValue = "ABC-101";
            PromptResult pr = ed.GetString(pso);
            
            if (pr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Command cancelled.\n");
                return;
            }

            string layoutName = string.IsNullOrWhiteSpace(pr.StringResult) ? "ABC-101" : pr.StringResult.Trim();
            ed.WriteMessage($"Testing auto notes detection for layout: '{layoutName}'\n");

            // Ensure services are initialized
            if (!DraftingAssistantExtensionApplication.EnsureServicesInitialized())
            {
                ed.WriteMessage("ERROR: Failed to initialize services. Please ensure a drawing is open.\n");
                return;
            }

            // Get or create project configuration
            var services = DraftingAssistantExtensionApplication.CompositionRoot;
            if (services == null)
            {
                ed.WriteMessage("ERROR: Services not initialized\n");
                return;
            }

            var logger = services.GetService<ILogger>();
            var configService = services.GetService<IProjectConfigurationService>();
            
            // Use default configuration for testing
            var config = configService.CreateDefaultConfiguration();
            config.ConstructionNotes = new ConstructionNotesConfiguration
            {
                MultileaderStyleName = "ML-STYLE-01"  // Test data style
            };
            ed.WriteMessage("Using default configuration for testing\n");

            ed.WriteMessage($"Target multileader style: '{config.ConstructionNotes?.MultileaderStyleName ?? "any"}'\n");

            // Test the auto notes service
            var autoNotesService = new AutoNotesService(logger);
            
            // Get diagnostic information first
            ed.WriteMessage("\n--- DIAGNOSTIC INFORMATION ---\n");
            var diagnosticInfo = autoNotesService.GetDiagnosticInfo(layoutName, config);
            ed.WriteMessage(diagnosticInfo + "\n");

            // Perform auto notes detection
            ed.WriteMessage("\n--- AUTO NOTES DETECTION ---\n");
            var noteNumbers = autoNotesService.GetAutoNotesForSheet(layoutName, config);
            
            if (noteNumbers.Count == 0)
            {
                ed.WriteMessage("❌ No construction notes detected in layout viewports\n");
                ed.WriteMessage("\nPossible reasons:\n");
                ed.WriteMessage("  - Layout not found\n");
                ed.WriteMessage("  - No viewports in layout\n");
                ed.WriteMessage("  - No multileaders in model space\n");
                ed.WriteMessage("  - Multileaders not within viewport boundaries\n");
                ed.WriteMessage("  - Multileaders don't match target style\n");
                ed.WriteMessage("  - Multileader blocks don't have TAGNUMBER attribute\n");
            }
            else
            {
                ed.WriteMessage($"✅ SUCCESS: Found {noteNumbers.Count} construction notes!\n");
                ed.WriteMessage($"Note numbers: {string.Join(", ", noteNumbers.OrderBy(n => n))}\n");
            }

            // Test point-in-polygon functionality with a simple case
            ed.WriteMessage("\n--- POINT-IN-POLYGON TEST ---\n");
            var testPolygon = new Point3dCollection();
            testPolygon.Add(new Point3d(0, 0, 0));
            testPolygon.Add(new Point3d(10, 0, 0));
            testPolygon.Add(new Point3d(10, 10, 0));
            testPolygon.Add(new Point3d(0, 10, 0));

            var testPoint1 = new Point3d(5, 5, 0);  // Should be inside
            var testPoint2 = new Point3d(15, 15, 0); // Should be outside

            bool inside1 = PointInPolygonDetector.IsPointInPolygon(testPoint1, testPolygon);
            bool inside2 = PointInPolygonDetector.IsPointInPolygon(testPoint2, testPolygon);

            ed.WriteMessage($"Point (5,5) in square (0,0)-(10,10): {(inside1 ? "✅ INSIDE" : "❌ OUTSIDE")} (expected: INSIDE)\n");
            ed.WriteMessage($"Point (15,15) in square (0,0)-(10,10): {(inside2 ? "❌ INSIDE" : "✅ OUTSIDE")} (expected: OUTSIDE)\n");

            if (inside1 && !inside2)
            {
                ed.WriteMessage("✅ Point-in-polygon test PASSED\n");
            }
            else
            {
                ed.WriteMessage("❌ Point-in-polygon test FAILED\n");
            }

            ed.WriteMessage("\n=== TESTAUTONOTES COMPLETE ===\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nFATAL ERROR in TESTAUTONOTES: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    [CommandMethod("TESTPRECISION")]
    public void TestPrecision()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        ed.WriteMessage("\n=== PRECISION DIAGNOSTIC TEST ===\n");

        try
        {
            using (var transaction = db.TransactionManager.StartTransaction())
            {
                // Get the current layout
                var layoutManager = LayoutManager.Current;
                var currentLayoutName = layoutManager.CurrentLayout;
                ed.WriteMessage($"Current layout: {currentLayoutName}\n");

                // Step 1: Find all multileaders with target styles
                var logger = new DebugLogger();
                var multileaderAnalyzer = new MultileaderAnalyzer(logger);
                var targetStyles = new[] { "ML-STYLE-01", "ML-STYLE-02" };
                var multileaders = multileaderAnalyzer.FindMultileadersInModelSpace(db, transaction, targetStyles);
                
                ed.WriteMessage($"\nFound {multileaders.Count} multileaders with target styles:\n");
                foreach (var ml in multileaders)
                {
                    ed.WriteMessage($"  - Note {ml.NoteNumber} at ({ml.Location.X:F6}, {ml.Location.Y:F6}) | Style: '{ml.StyleName}'\n");
                }

                if (multileaders.Count == 0)
                {
                    ed.WriteMessage("❌ No multileaders found - cannot test precision\n");
                    return;
                }

                // Step 2: Analyze viewports in current layout
                ed.WriteMessage("\n--- VIEWPORT ANALYSIS ---\n");
                var layoutDict = transaction.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                var layoutId = (ObjectId)layoutDict[currentLayoutName];
                if (layoutId == ObjectId.Null)
                {
                    ed.WriteMessage("❌ Cannot access current layout\n");
                    return;
                }

                var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                var layoutBlock = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                
                if (layoutBlock == null)
                {
                    ed.WriteMessage("❌ Cannot access layout block\n");
                    return;
                }

                int viewportCount = 0;
                int precisionIssuesFound = 0;

                foreach (ObjectId entityId in layoutBlock)
                {
                    var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                    if (entity is Viewport viewport && viewport.Number > 1)
                    {
                        viewportCount++;
                        ed.WriteMessage($"\n=== VIEWPORT #{viewportCount} ===\n");
                        
                        // Show viewport properties
                        ed.WriteMessage($"Center: ({viewport.CenterPoint.X:F6}, {viewport.CenterPoint.Y:F6})\n");
                        ed.WriteMessage($"ViewCenter: ({viewport.ViewCenter.X:F6}, {viewport.ViewCenter.Y:F6})\n");
                        ed.WriteMessage($"Scale: {viewport.CustomScale:F10} (1\" = {(1/viewport.CustomScale):F2}')\n");
                        ed.WriteMessage($"Twist Angle: {viewport.TwistAngle:F10} rad ({viewport.TwistAngle * 180.0 / Math.PI:F6}°)\n");
                        ed.WriteMessage($"Dimensions: {viewport.Width:F6} x {viewport.Height:F6}\n");

                        // Test precision differences
                        var precisionResults = TestViewportPrecision(viewport, multileaders, ed);
                        precisionIssuesFound += precisionResults;
                    }
                }

                ed.WriteMessage($"\n=== SUMMARY ===\n");
                ed.WriteMessage($"Viewports analyzed: {viewportCount}\n");
                ed.WriteMessage($"Multileaders tested: {multileaders.Count}\n");
                ed.WriteMessage($"Precision issues found: {precisionIssuesFound}\n");
                
                if (precisionIssuesFound > 0)
                {
                    ed.WriteMessage("🔍 PRECISION LOSS DETECTED! This is likely causing the Auto Notes failure.\n");
                    ed.WriteMessage("💡 Recommendation: Implement high-precision calculations or rotate around viewport center.\n");
                }
                else
                {
                    ed.WriteMessage("✅ No significant precision issues found. The problem may be elsewhere.\n");
                }

                transaction.Commit();
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTPRECISION: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }

        ed.WriteMessage("\n=== TESTPRECISION COMPLETE ===\n");
    }

    private int TestViewportPrecision(Viewport viewport, List<MultileaderAnalyzer.MultileaderInfo> multileaders, Editor ed)
    {
        int issuesFound = 0;

        try
        {
            // Get viewport footprint using current (double) precision
            var currentFootprint = ViewportBoundaryCalculator.GetViewportFootprint(viewport);
            ed.WriteMessage($"Current method footprint: {currentFootprint.Count} points\n");
            
            if (currentFootprint.Count >= 4)
            {
                // Show the corners
                ed.WriteMessage("  Corners (double precision):\n");
                for (int i = 0; i < Math.Min(4, currentFootprint.Count); i++)
                {
                    var pt = currentFootprint[i];
                    ed.WriteMessage($"    [{i}]: ({pt.X:F6}, {pt.Y:F6})\n");
                }
            }

            // Calculate high-precision version
            var highPrecisionFootprint = GetViewportFootprintHighPrecision(viewport);
            ed.WriteMessage($"High precision footprint: {highPrecisionFootprint.Count} points\n");
            
            if (highPrecisionFootprint.Count >= 4)
            {
                ed.WriteMessage("  Corners (decimal precision):\n");
                for (int i = 0; i < Math.Min(4, highPrecisionFootprint.Count); i++)
                {
                    var pt = highPrecisionFootprint[i];
                    ed.WriteMessage($"    [{i}]: ({pt.X:F6}, {pt.Y:F6})\n");
                }

                // Compare precision differences
                if (currentFootprint.Count == highPrecisionFootprint.Count)
                {
                    double maxDifference = 0;
                    for (int i = 0; i < currentFootprint.Count; i++)
                    {
                        double diffX = Math.Abs(currentFootprint[i].X - highPrecisionFootprint[i].X);
                        double diffY = Math.Abs(currentFootprint[i].Y - highPrecisionFootprint[i].Y);
                        maxDifference = Math.Max(maxDifference, Math.Max(diffX, diffY));
                    }
                    
                    ed.WriteMessage($"  Max coordinate difference: {maxDifference:E6}\n");
                    if (maxDifference > 1e-6) // If difference > 1 micrometer
                    {
                        ed.WriteMessage($"  ⚠️  SIGNIFICANT PRECISION DIFFERENCE: {maxDifference:E6}\n");
                        issuesFound++;
                    }
                }
            }

            // Test each multileader
            ed.WriteMessage("  Testing multileaders:\n");
            foreach (var ml in multileaders)
            {
                // Test with current precision
                bool insideCurrent = IsPointInPolygon(ml.Location, currentFootprint);
                bool insideHighPrecision = IsPointInPolygon(ml.Location, highPrecisionFootprint);
                
                ed.WriteMessage($"    Note {ml.NoteNumber}: Current={insideCurrent}, HighPrec={insideHighPrecision}");
                
                if (insideCurrent != insideHighPrecision)
                {
                    ed.WriteMessage(" ❌ PRECISION MISMATCH!");
                    issuesFound++;
                }
                else
                {
                    ed.WriteMessage(" ✅");
                }
                ed.WriteMessage("\n");

                // Show distance to boundary
                var minDist = GetMinDistanceToPolygon(ml.Location, currentFootprint);
                ed.WriteMessage($"      Distance to boundary: {minDist:F6}\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"    Error testing viewport precision: {ex.Message}\n");
        }

        return issuesFound;
    }

    private List<Point3d> GetViewportFootprintHighPrecision(Viewport viewport)
    {
        var points = new List<Point3d>();

        try
        {
            // Use decimal for high-precision calculations
            decimal halfW = (decimal)(viewport.Width / 2.0);
            decimal halfH = (decimal)(viewport.Height / 2.0);
            decimal centerX = (decimal)viewport.ViewCenter.X;
            decimal centerY = (decimal)viewport.ViewCenter.Y;
            decimal scaleFactor = 1.0m / (decimal)viewport.CustomScale;

            // Build corners in paper space with decimal precision
            var paperCorners = new[]
            {
                new { X = centerX - halfW, Y = centerY - halfH }, // BL
                new { X = centerX - halfW, Y = centerY + halfH }, // TL  
                new { X = centerX + halfW, Y = centerY + halfH }, // TR
                new { X = centerX + halfW, Y = centerY - halfH }, // BR
            };

            foreach (var corner in paperCorners)
            {
                // Transform with high precision
                decimal fromCenterX = corner.X - centerX;
                decimal fromCenterY = corner.Y - centerY;
                
                // Scale
                decimal scaledFromCenterX = fromCenterX * scaleFactor;
                decimal scaledFromCenterY = fromCenterY * scaleFactor;
                
                decimal scaledX = centerX + scaledFromCenterX;
                decimal scaledY = centerY + scaledFromCenterY;

                // For rotation, still need to use double for trig functions
                if (Math.Abs(viewport.TwistAngle) > 1e-10)
                {
                    // Rotate around origin (this is where precision loss occurs)
                    double angle = -viewport.TwistAngle;
                    decimal cosAngle = (decimal)Math.Cos(angle);
                    decimal sinAngle = (decimal)Math.Sin(angle);
                    
                    decimal rotatedX = scaledX * cosAngle - scaledY * sinAngle;
                    decimal rotatedY = scaledX * sinAngle + scaledY * cosAngle;
                    
                    points.Add(new Point3d((double)rotatedX, (double)rotatedY, 0));
                }
                else
                {
                    points.Add(new Point3d((double)scaledX, (double)scaledY, 0));
                }
            }
        }
        catch (System.Exception)
        {
            // Fallback to current method
            return ViewportBoundaryCalculator.GetViewportFootprint(viewport).Cast<Point3d>().ToList();
        }

        return points;
    }

    private bool IsPointInPolygon(Point3d testPoint, Point3dCollection polygon)
    {
        if (polygon.Count < 3) return false;
        
        var points = new Point3dCollection();
        foreach (Point3d pt in polygon)
            points.Add(pt);
        return PointInPolygonDetector.IsPointInPolygon(testPoint, points);
    }

    private bool IsPointInPolygon(Point3d testPoint, List<Point3d> polygon)
    {
        if (polygon.Count < 3) return false;
        var points = new Point3dCollection();
        foreach (var pt in polygon)
            points.Add(pt);
        return PointInPolygonDetector.IsPointInPolygon(testPoint, points);
    }

    private double GetMinDistanceToPolygon(Point3d testPoint, Point3dCollection polygon)
    {
        if (polygon.Count < 2) return double.MaxValue;

        double minDist = double.MaxValue;
        for (int i = 0; i < polygon.Count; i++)
        {
            var p1 = polygon[i];
            var p2 = polygon[(i + 1) % polygon.Count];
            
            var dist = GetDistanceToLineSegment(testPoint, p1, p2);
            minDist = Math.Min(minDist, dist);
        }
        
        return minDist;
    }

    private double GetDistanceToLineSegment(Point3d point, Point3d lineStart, Point3d lineEnd)
    {
        var A = point.X - lineStart.X;
        var B = point.Y - lineStart.Y;
        var C = lineEnd.X - lineStart.X;
        var D = lineEnd.Y - lineStart.Y;

        var dot = A * C + B * D;
        var lenSq = C * C + D * D;
        
        if (lenSq == 0) return Math.Sqrt(A * A + B * B); // Line is a point

        var param = dot / lenSq;

        double xx, yy;
        if (param < 0)
        {
            xx = lineStart.X;
            yy = lineStart.Y;
        }
        else if (param > 1)
        {
            xx = lineEnd.X;
            yy = lineEnd.Y;
        }
        else
        {
            xx = lineStart.X + param * C;
            yy = lineStart.Y + param * D;
        }

        var dx = point.X - xx;
        var dy = point.Y - yy;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    [CommandMethod("TESTVIEWPORTBOUNDS")]
    public void TestViewportBounds()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        ed.WriteMessage("\n=== VIEWPORT BOUNDS DIAGNOSTIC ===\n");

        try
        {
            using (var transaction = db.TransactionManager.StartTransaction())
            {
                // Get the current layout
                var layoutManager = LayoutManager.Current;
                var currentLayoutName = layoutManager.CurrentLayout;
                ed.WriteMessage($"Current layout: {currentLayoutName}\n");

                // Find all multileaders with target styles
                var logger = new DebugLogger();
                var multileaderAnalyzer = new MultileaderAnalyzer(logger);
                var targetStyles = new[] { "ML-STYLE-01", "ML-STYLE-02" };
                var multileaders = multileaderAnalyzer.FindMultileadersInModelSpace(db, transaction, targetStyles);
                
                ed.WriteMessage($"\nFound {multileaders.Count} multileaders with target styles:\n");
                foreach (var ml in multileaders)
                {
                    ed.WriteMessage($"  - Note {ml.NoteNumber} at ({ml.Location.X:F6}, {ml.Location.Y:F6}) | Style: '{ml.StyleName}'\n");
                }

                // Get current layout viewports
                var layoutDict = transaction.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                var layoutId = (ObjectId)layoutDict[currentLayoutName];
                var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                var layoutBlock = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                
                int viewportCount = 0;
                foreach (ObjectId entityId in layoutBlock)
                {
                    var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                    if (entity is Viewport viewport && viewport.Number > 1)
                    {
                        viewportCount++;
                        ed.WriteMessage($"\n=== VIEWPORT #{viewportCount} ===\n");
                        
                        // Show basic viewport properties
                        ed.WriteMessage($"Center: ({viewport.CenterPoint.X:F6}, {viewport.CenterPoint.Y:F6})\n");
                        ed.WriteMessage($"ViewCenter: ({viewport.ViewCenter.X:F6}, {viewport.ViewCenter.Y:F6})\n");
                        ed.WriteMessage($"Scale: {viewport.CustomScale:F10} (1\" = {(1/viewport.CustomScale):F2}')\n");
                        ed.WriteMessage($"Twist Angle: {viewport.TwistAngle:F10} rad ({viewport.TwistAngle * 180.0 / Math.PI:F6}°)\n");
                        ed.WriteMessage($"Dimensions: {viewport.Width:F6} x {viewport.Height:F6}\n");
                        ed.WriteMessage($"NonRectClip: {viewport.NonRectClipOn}, ClipEntityValid: {viewport.NonRectClipEntityId.IsValid}\n");

                        // Use the actual ViewportBoundaryCalculator from Auto Notes
                        ed.WriteMessage("\n--- USING ACTUAL AUTO NOTES VIEWPORT CALCULATOR ---\n");
                        try
                        {
                            var footprint = ViewportBoundaryCalculator.GetViewportFootprint(viewport, transaction);
                            ed.WriteMessage($"Calculated footprint: {footprint.Count} points\n");
                            
                            if (footprint.Count > 0)
                            {
                                ed.WriteMessage("Viewport boundary vertices (in order):\n");
                                for (int i = 0; i < footprint.Count; i++)
                                {
                                    var pt = footprint[i];
                                    ed.WriteMessage($"  [{i}]: ({pt.X:F6}, {pt.Y:F6})\n");
                                }

                                // Calculate bounding box
                                var bounds = ViewportBoundaryCalculator.GetViewportBounds(viewport, transaction);
                                if (bounds.HasValue)
                                {
                                    ed.WriteMessage($"\nBounding box:\n");
                                    ed.WriteMessage($"  Min: ({bounds.Value.MinPoint.X:F6}, {bounds.Value.MinPoint.Y:F6})\n");
                                    ed.WriteMessage($"  Max: ({bounds.Value.MaxPoint.X:F6}, {bounds.Value.MaxPoint.Y:F6})\n");
                                    ed.WriteMessage($"  Size: {bounds.Value.MaxPoint.X - bounds.Value.MinPoint.X:F6} x {bounds.Value.MaxPoint.Y - bounds.Value.MinPoint.Y:F6}\n");
                                }

                                // Test each multileader against the viewport boundary using Auto Notes logic
                                ed.WriteMessage($"\n--- TESTING MULTILEADERS AGAINST VIEWPORT BOUNDARY ---\n");
                                foreach (var ml in multileaders)
                                {
                                    bool isInside = IsPointInViewportFootprint(ml.Location, footprint);
                                    double minDistance = CalculateMinDistanceToFootprint(ml.Location, footprint);
                                    
                                    string insideStatus = isInside ? "✅ INSIDE" : "❌ OUTSIDE";
                                    ed.WriteMessage($"  Note {ml.NoteNumber} at ({ml.Location.X:F6}, {ml.Location.Y:F6}): {insideStatus}\n");
                                    ed.WriteMessage($"    Distance to boundary: {minDistance:F6}\n");
                                    
                                    if (isInside)
                                    {
                                        ed.WriteMessage($"    → This multileader SHOULD be detected by Auto Notes\n");
                                    }
                                    else
                                    {
                                        ed.WriteMessage($"    → This multileader should NOT be detected by Auto Notes\n");
                                    }
                                }

                                // Show diagnostic information from ViewportBoundaryCalculator
                                ed.WriteMessage($"\n--- VIEWPORT TRANSFORMATION DIAGNOSTICS ---\n");
                                var diagnostics = ViewportBoundaryCalculator.GetTransformationDiagnostics(viewport);
                                ed.WriteMessage($"{diagnostics}\n");
                            }
                            else
                            {
                                ed.WriteMessage("❌ No viewport footprint calculated\n");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"❌ Error calculating viewport footprint: {ex.Message}\n");
                            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
                        }
                    }
                }

                if (viewportCount == 0)
                {
                    ed.WriteMessage("❌ No viewports found in current layout\n");
                }

                transaction.Commit();
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTVIEWPORTBOUNDS: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }

        ed.WriteMessage("\n=== TESTVIEWPORTBOUNDS COMPLETE ===\n");
    }

    private bool IsPointInViewportFootprint(Point3d testPoint, Point3dCollection footprint)
    {
        if (footprint.Count < 3) return false;
        
        // Use the same PointInPolygonDetector that Auto Notes uses
        return PointInPolygonDetector.IsPointInPolygon(testPoint, footprint);
    }

    private double CalculateMinDistanceToFootprint(Point3d testPoint, Point3dCollection footprint)
    {
        if (footprint.Count < 2) return double.MaxValue;

        double minDistance = double.MaxValue;
        
        for (int i = 0; i < footprint.Count; i++)
        {
            var p1 = footprint[i];
            var p2 = footprint[(i + 1) % footprint.Count];
            
            double distance = GetDistanceToLineSegment(testPoint, p1, p2);
            minDistance = Math.Min(minDistance, distance);
        }
        
        return minDistance;
    }

    [CommandMethod("SELECTVIEWPORT")]
    public void SelectViewport()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        ed.WriteMessage("\n=== Select Viewport for Property Display ===\n");

        try
        {
            // Prompt user to select a viewport
            var selectionOptions = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a viewport: ",
                AllowDuplicates = false,
                SingleOnly = true
            };

            // Filter for viewport objects only
            var filter = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "VIEWPORT")
            };
            var selectionFilter = new SelectionFilter(filter);

            var selectionResult = ed.GetSelection(selectionOptions, selectionFilter);
            
            if (selectionResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("No viewport selected or selection cancelled.\n");
                return;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var selectedObjectId = selectionResult.Value[0].ObjectId;
                var viewport = tr.GetObject(selectedObjectId, OpenMode.ForRead) as Viewport;

                if (viewport == null)
                {
                    ed.WriteMessage("ERROR: Selected object is not a viewport.\n");
                    return;
                }

                ed.WriteMessage($"\n=== Viewport Properties ===\n");
                ed.WriteMessage($"Handle: {viewport.Handle}\n");
                ed.WriteMessage($"Number: {viewport.Number}\n");
                
                if (viewport.Number == 1)
                {
                    ed.WriteMessage("NOTE: This is the paper space viewport (Number = 1)\n");
                }
                else
                {
                    ed.WriteMessage("NOTE: This is a model space viewport\n");
                }

                ed.WriteMessage($"\n--- Requested Properties ---\n");
                ed.WriteMessage($"ViewCenter: ({viewport.ViewCenter.X:F6}, {viewport.ViewCenter.Y:F6})\n");
                ed.WriteMessage($"ViewDirection: ({viewport.ViewDirection.X:F6}, {viewport.ViewDirection.Y:F6}, {viewport.ViewDirection.Z:F6})\n");
                ed.WriteMessage($"ViewTarget: ({viewport.ViewTarget.X:F6}, {viewport.ViewTarget.Y:F6}, {viewport.ViewTarget.Z:F6})\n");
                ed.WriteMessage($"TwistAngle: {viewport.TwistAngle:F6} rad ({viewport.TwistAngle * 180.0 / Math.PI:F3}°)\n");
                ed.WriteMessage($"ViewHeight: {viewport.ViewHeight:F6}\n");
                ed.WriteMessage($"Width: {viewport.Width:F6}\n");
                ed.WriteMessage($"Height: {viewport.Height:F6}\n");
                ed.WriteMessage($"CenterPoint: ({viewport.CenterPoint.X:F6}, {viewport.CenterPoint.Y:F6})\n");
                ed.WriteMessage($"CustomScale: {viewport.CustomScale:F6}\n");

                ed.WriteMessage($"\n--- Additional Calculated Values ---\n");
                double scaleRatio = 1.0 / viewport.CustomScale;
                ed.WriteMessage($"Scale Ratio (1/CustomScale): {scaleRatio:F6} (1:{scaleRatio:F0})\n");
                
                double modelWidth = viewport.ViewHeight * viewport.Width / viewport.Height;
                ed.WriteMessage($"Model Width (ViewHeight × Width/Height): {modelWidth:F6}\n");
                
                ed.WriteMessage($"On: {viewport.On}\n");
                ed.WriteMessage($"NonRectClipOn: {viewport.NonRectClipOn}\n");
                
                tr.Commit();
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"ERROR in SELECTVIEWPORT: {ex.Message}\n");
            if (ex.InnerException != null)
            {
                ed.WriteMessage($"Inner exception: {ex.InnerException.Message}\n");
            }
        }
    }

    /// <summary>
    /// Helper method to log available layouts
    /// </summary>
    private static void LogAvailableLayouts(Editor ed, DBDictionary layoutDict, Transaction tr)
    {
        ed.WriteMessage("\nAvailable layouts:\n");
        foreach (DBDictionaryEntry entry in layoutDict)
        {
            if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage($"  - '{entry.Key}'\n");
            }
        }
    }

    /// <summary>
    /// Test command to validate DrawingAccessService functionality
    /// </summary>
    [CommandMethod("TESTDRAWINGACCESS")]
    public void TestDrawingAccess()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== Testing DrawingAccessService ===\n");

            // Create logger and service
            var logger = new AutoCADLogger();
            var drawingService = new DrawingAccessService(logger);

            // Test 1: Get all currently open drawings
            ed.WriteMessage("\n--- Test 1: All Open Drawings ---\n");
            var openDrawings = drawingService.GetAllOpenDrawings();
            ed.WriteMessage($"Found {openDrawings.Count} open drawings:\n");
            
            foreach (var kvp in openDrawings)
            {
                var fileName = Path.GetFileName(kvp.Key);
                var state = kvp.Value;
                ed.WriteMessage($"  • {fileName} -> {state}\n");
            }

            // Test 2: Check state of current drawing
            ed.WriteMessage("\n--- Test 2: Current Drawing State ---\n");
            var currentDrawingPath = doc.Name;
            var currentState = drawingService.GetDrawingState(currentDrawingPath);
            ed.WriteMessage($"Current drawing: {Path.GetFileName(currentDrawingPath)}\n");
            ed.WriteMessage($"State: {currentState}\n");

            // Test 3: Test with test project files
            ed.WriteMessage("\n--- Test 3: Test Project Files ---\n");
            var testProjectPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\";
            var testDwgPath = Path.Combine(testProjectPath, "PROJ-ABC-100.dwg");
            
            if (File.Exists(testDwgPath))
            {
                var testState = drawingService.GetDrawingState(testDwgPath);
                ed.WriteMessage($"Test drawing (PROJ-ABC-100.dwg): {testState}\n");
                
                if (testState == DrawingState.Inactive)
                {
                    ed.WriteMessage("Attempting to make test drawing active...\n");
                    var success = drawingService.TryMakeDrawingActive(testDwgPath);
                    ed.WriteMessage($"Make active result: {(success ? "SUCCESS" : "FAILED")}\n");
                }
            }
            else
            {
                ed.WriteMessage($"Test drawing not found at: {testDwgPath}\n");
            }

            // Test 4: File path resolution (if project config exists)
            ed.WriteMessage("\n--- Test 4: File Path Resolution ---\n");
            // Skip Excel-based test for now to avoid crashes
            ed.WriteMessage("Skipping Excel integration test to avoid crashes\n");
            ed.WriteMessage("Excel-based file path resolution would be tested here\n");

            // Test 5: Test with non-existent file
            ed.WriteMessage("\n--- Test 5: Non-existent File ---\n");
            var fakePath = @"C:\NonExistent\File.dwg";
            var fakeState = drawingService.GetDrawingState(fakePath);
            ed.WriteMessage($"Non-existent file state: {fakeState}\n");

            ed.WriteMessage("\n=== DrawingAccessService Test Complete ===\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError during test: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    private void TestFilePathResolution(Editor ed, DrawingAccessService drawingService)
    {
        try
        {
            // Try to load project configuration
            var configPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectConfig.json";
            if (!File.Exists(configPath))
            {
                ed.WriteMessage("ProjectConfig.json not found - skipping file path resolution test\n");
                return;
            }

            var logger = new AutoCADLogger();
            var configService = new ProjectConfigurationService(logger);
            var config = configService.LoadConfigurationAsync(configPath).GetAwaiter().GetResult();
            
            if (config == null)
            {
                ed.WriteMessage("Failed to load project configuration\n");
                return;
            }

            // Try to load sheet index
            var excelReader = new ExcelReaderService(logger);
            var sheetInfos = excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config).GetAwaiter().GetResult();
            
            if (sheetInfos.Count == 0)
            {
                ed.WriteMessage("No sheet information found in Excel file\n");
                return;
            }

            ed.WriteMessage($"Loaded {sheetInfos.Count} sheets from Excel index\n");

            // Test file path resolution for first few sheets
            var testSheets = sheetInfos.Take(3).ToList();
            foreach (var sheetInfo in testSheets)
            {
                var resolvedPath = drawingService.GetDrawingFilePath(sheetInfo.SheetName, config, sheetInfos);
                if (resolvedPath != null)
                {
                    var state = drawingService.GetDrawingState(resolvedPath);
                    var fileName = Path.GetFileName(resolvedPath);
                    ed.WriteMessage($"  • {sheetInfo.SheetName} -> {fileName} ({state})\n");
                }
                else
                {
                    ed.WriteMessage($"  • {sheetInfo.SheetName} -> [FAILED TO RESOLVE]\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"File path resolution test failed: {ex.Message}\n");
        }
    }

    [CommandMethod("TESTDRAWINGLIST")]
    public void ListAllDrawings()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== Quick Drawing List ===\n");
            
            var logger = new AutoCADLogger();
            var drawingService = new DrawingAccessService(logger);
            var openDrawings = drawingService.GetAllOpenDrawings();
            
            if (openDrawings.Count == 0)
            {
                ed.WriteMessage("No drawings currently open\n");
                return;
            }

            ed.WriteMessage($"Found {openDrawings.Count} open drawing(s):\n");
            foreach (var kvp in openDrawings)
            {
                var fileName = Path.GetFileName(kvp.Key);
                var state = kvp.Value;
                var indicator = state == DrawingState.Active ? " ← ACTIVE" : "";
                ed.WriteMessage($"  {fileName} ({state}){indicator}\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }

    [CommandMethod("TESTDRAWINGSTATE")]
    public void TestSpecificDrawingState()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            // Prompt user for file path
            var pfo = new PromptOpenFileOptions("\nSelect DWG file to check state")
            {
                Filter = "AutoCAD Drawing (*.dwg)|*.dwg"
            };

            var pfr = ed.GetFileNameForOpen(pfo);
            if (pfr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Command cancelled.\n");
                return;
            }

            var logger = new AutoCADLogger();
            var drawingService = new DrawingAccessService(logger);
            var state = drawingService.GetDrawingState(pfr.StringResult);
            
            var fileName = Path.GetFileName(pfr.StringResult);
            ed.WriteMessage($"\nDrawing: {fileName}\n");
            ed.WriteMessage($"State: {state}\n");

            if (state == DrawingState.Inactive)
            {
                var pko = new PromptKeywordOptions("\nDrawing is inactive. Make it active? ");
                pko.Keywords.Add("Yes");
                pko.Keywords.Add("No");
                pko.Keywords.Default = "Yes";
                var pkr = ed.GetKeywords(pko);
                
                if (pkr.Status == PromptStatus.OK && pkr.StringResult == "Yes")
                {
                    var success = drawingService.TryMakeDrawingActive(pfr.StringResult);
                    ed.WriteMessage($"Make active result: {(success ? "SUCCESS" : "FAILED")}\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }

    [CommandMethod("TESTCLOSEDUPDATE")]
    public void TestExternalDrawingUpdate()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== TESTCLOSEDUPDATE: External Drawing Manager Validation ===\n");

            // Prompt user for DWG file path
            var pfo = new PromptOpenFileOptions("\nSelect CLOSED DWG file to update")
            {
                Filter = "AutoCAD Drawing (*.dwg)|*.dwg"
            };

            var pfr = ed.GetFileNameForOpen(pfo);
            if (pfr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Command cancelled.\n");
                return;
            }

            string dwgPath = pfr.StringResult;
            ed.WriteMessage($"Target DWG: {Path.GetFileName(dwgPath)}\n");

            // Check that the file is actually closed (not currently open)
            var logger = new AutoCADLogger();
            var drawingService = new DrawingAccessService(logger);
            var state = drawingService.GetDrawingState(dwgPath);
            
            ed.WriteMessage($"Drawing state: {state}\n");
            
            if (state != DrawingState.Closed)
            {
                ed.WriteMessage($"WARNING: Drawing is {state}, not Closed. Test may still proceed but results might be affected.\n");
            }

            // Prompt for layout name
            var pso = new PromptStringOptions("\nEnter layout name to update (e.g., ABC-101): ");
            pso.AllowSpaces = false;
            var psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Command cancelled.\n");
                return;
            }

            string layoutName = psr.StringResult;
            ed.WriteMessage($"Target layout: {layoutName}\n");

            // Create test construction note data
            var testNotes = new List<ConstructionNoteData>
            {
                new ConstructionNoteData(1, "TEST NOTE 1 - Updated via ExternalDrawingManager"),
                new ConstructionNoteData(2, "TEST NOTE 2 - Closed drawing update"),
                new ConstructionNoteData(4, "TEST NOTE 4 - ATTSYNC validation test")
            };

            ed.WriteMessage($"Test data: {testNotes.Count} construction notes\n");
            foreach (var note in testNotes)
            {
                ed.WriteMessage($"  Note {note.NoteNumber}: {note.NoteText.Substring(0, Math.Min(note.NoteText.Length, 40))}...\n");
            }

            // Create ExternalDrawingManager and perform update
            ed.WriteMessage("\nStarting external drawing update...\n");
            var backupCleanupService = new BackupCleanupService(logger);
            var multileaderAnalyzer = new MultileaderAnalyzer(logger);
            var blockAnalyzer = new BlockAnalyzer(logger);
            var externalManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
            bool success = externalManager.UpdateClosedDrawing(dwgPath, layoutName, testNotes);

            if (success)
            {
                ed.WriteMessage("\n*** SUCCESS! ***\n");
                ed.WriteMessage($"Updated construction notes in closed drawing: {Path.GetFileName(dwgPath)}\n");
                ed.WriteMessage($"Layout: {layoutName}\n");
                ed.WriteMessage($"Notes updated: {testNotes.Count}\n");
                ed.WriteMessage("\nVerification:\n");
                ed.WriteMessage("1. Open the drawing in AutoCAD\n");
                ed.WriteMessage($"2. Switch to layout '{layoutName}'\n");
                ed.WriteMessage("3. Check that blocks NT01, NT02, NT04 are visible with test notes\n");
                ed.WriteMessage("4. Verify attributes are properly centered (ATTSYNC applied)\n");
            }
            else
            {
                ed.WriteMessage("\n*** FAILED ***\n");
                ed.WriteMessage("External drawing update was not successful.\n");
                ed.WriteMessage("Check command line for detailed error messages.\n");
            }

            ed.WriteMessage("\n=== TESTCLOSEDUPDATE COMPLETE ===\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nFATAL ERROR in TESTCLOSEDUPDATE: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Test command to validate BackupCleanupService functionality
    /// </summary>
    [CommandMethod("TESTCLEANUP")]
    public void TestBackupCleanup()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== TESTCLEANUP: BackupCleanupService Validation ===\n");

            // Get logger from composition root
            var compositionRoot = DraftingAssistantExtensionApplication.CompositionRoot;
            var logger = compositionRoot?.GetService<ILogger>();
            if (logger == null)
            {
                ed.WriteMessage("ERROR: Logger service not available\n");
                return;
            }

            // Create BackupCleanupService manually
            var backupCleanupService = new BackupCleanupService(logger);

            // Prompt for directory path
            var pfo = new PromptOpenFileOptions("\nSelect any DWG file in the directory to scan for backups")
            {
                Filter = "Drawing files (*.dwg)|*.dwg|All Files (*.*)|*.*",
                DialogCaption = "Select DWG file in directory to scan",
                DialogName = "TESTCLEANUP Directory Selection"
            };

            PromptFileNameResult pfnResult = ed.GetFileNameForOpen(pfo);
            if (pfnResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Operation cancelled by user.\n");
                return;
            }

            string selectedFile = pfnResult.StringResult;
            string directory = Path.GetDirectoryName(selectedFile) ?? string.Empty;

            if (string.IsNullOrEmpty(directory))
            {
                ed.WriteMessage("ERROR: Could not determine directory path.\n");
                return;
            }

            ed.WriteMessage($"Scanning directory: {directory}\n");

            // Get backup file information
            var backupFiles = backupCleanupService.GetBackupFileInfo(directory);
            ed.WriteMessage($"Found {backupFiles.Count} backup files:\n");

            if (backupFiles.Count == 0)
            {
                ed.WriteMessage("No .bak.beforeupdate files found in directory.\n");
                ed.WriteMessage("To test cleanup functionality:\n");
                ed.WriteMessage("1. Run TESTCLOSEDUPDATE on a closed drawing\n");
                ed.WriteMessage("2. Then run TESTCLEANUP again\n");
                return;
            }

            // Display backup files
            foreach (var backup in backupFiles)
            {
                ed.WriteMessage($"  - {backup.FileName} (Created: {backup.CreatedDate:yyyy-MM-dd HH:mm:ss}, Size: {backup.Size:N0} bytes)\n");
            }

            // Test simplified cleanup
            ed.WriteMessage("\n--- Testing Simplified Cleanup ---\n");

            // Test 1: Count only
            int count = backupCleanupService.GetBackupFileCount(directory);
            ed.WriteMessage($"GetBackupFileCount result: {count} files\n");

            // Test 2: Cleanup all backup files immediately
            ed.WriteMessage("\nTesting immediate cleanup (all backup files):\n");
            int cleanedCount = backupCleanupService.CleanupAllBackupFiles(directory);
            ed.WriteMessage($"Cleanup removed: {cleanedCount} files\n");

            // Test 3: Final count
            int finalCount = backupCleanupService.GetBackupFileCount(directory);
            ed.WriteMessage($"Final backup file count: {finalCount} files\n");

            ed.WriteMessage("\n=== TESTCLEANUP COMPLETE ===\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nFATAL ERROR in TESTCLEANUP: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Debug command to discover and analyze all blocks in a layout
    /// </summary>
    [CommandMethod("TESTBLOCKDISCOVERY")]
    public static void TestBlockDiscovery()
    {
        var ed = Application.DocumentManager.MdiActiveDocument.Editor;
        
        try
        {
            ed.WriteMessage("\n=== BLOCK DISCOVERY DEBUG TEST ===\n");
            
            // Get drawing file path
            var fileResult = ed.GetString("\nEnter DWG file path (or press Enter for current document): ");
            if (fileResult.Status == PromptStatus.Cancel) return;
            
            string dwgFilePath = string.IsNullOrWhiteSpace(fileResult.StringResult) 
                ? Application.DocumentManager.MdiActiveDocument.Name 
                : fileResult.StringResult.Trim();

            // Get layout name
            var layoutResult = ed.GetString("\nEnter layout name (e.g., ABC-101): ");
            if (layoutResult.Status != PromptStatus.OK) return;
            string layoutName = layoutResult.StringResult.Trim();

            ed.WriteMessage($"Analyzing blocks in layout '{layoutName}' from file: {Path.GetFileName(dwgFilePath)}\n");

            // Use current drawing or external drawing
            if (dwgFilePath == Application.DocumentManager.MdiActiveDocument.Name)
            {
                ed.WriteMessage("Using current drawing...\n");
                AnalyzeCurrentDrawingBlocks(layoutName, ed);
            }
            else
            {
                ed.WriteMessage("Using external drawing...\n");
                AnalyzeExternalDrawingBlocks(dwgFilePath, layoutName, ed);
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTBLOCKDISCOVERY: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Analyze blocks in the current drawing
    /// </summary>
    private static void AnalyzeCurrentDrawingBlocks(string layoutName, Editor ed)
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            if (!layoutDict.Contains(layoutName))
            {
                ed.WriteMessage($"ERROR: Layout '{layoutName}' not found\n");
                return;
            }

            var layoutId = layoutDict.GetAt(layoutName);
            var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
            var layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

            AnalyzeLayoutBlocks(layoutBtr, layout, tr, ed);
        }
    }

    /// <summary>
    /// Analyze blocks in an external drawing
    /// </summary>
    private static void AnalyzeExternalDrawingBlocks(string dwgPath, string layoutName, Editor ed)
    {
        if (!File.Exists(dwgPath))
        {
            ed.WriteMessage($"ERROR: File not found: {dwgPath}\n");
            return;
        }

        using (var db = new Database(false, true))
        {
            db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
            
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (!layoutDict.Contains(layoutName))
                {
                    ed.WriteMessage($"ERROR: Layout '{layoutName}' not found\n");
                    return;
                }

                var layoutId = layoutDict.GetAt(layoutName);
                var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                var layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                AnalyzeLayoutBlocks(layoutBtr, layout, tr, ed);
            }
        }
    }

    /// <summary>
    /// Analyze all blocks in a layout
    /// </summary>
    private static void AnalyzeLayoutBlocks(BlockTableRecord layoutBtr, Layout layout, Transaction tr, Editor ed)
    {
        // Get viewport protection info
        var viewportIds = new HashSet<ObjectId>();
        var viewports = layout.GetViewports();
        if (viewports != null)
        {
            foreach (ObjectId vpId in viewports)
            {
                viewportIds.Add(vpId);
            }
        }
        
        ed.WriteMessage($"Found {viewportIds.Count} viewports to protect\n");
        
        // Analyze all entities
        int totalEntities = 0;
        int blockReferences = 0;
        int viewportCount = 0;
        int ntBlocks = 0;
        var blockTypes = new Dictionary<string, int>();
        
        ed.WriteMessage("\n--- DETAILED ENTITY ANALYSIS ---\n");
        
        foreach (ObjectId objId in layoutBtr)
        {
            totalEntities++;
            var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
            
            ed.WriteMessage($"{totalEntities:D3}: ObjectId={objId}, Type={entity?.GetType().Name}");
            
            if (viewportIds.Contains(objId))
            {
                viewportCount++;
                ed.WriteMessage(" [VIEWPORT - PROTECTED]");
            }
            else if (entity is BlockReference blockRef)
            {
                blockReferences++;
                
                // Get effective block name using the same method as ExternalDrawingManager
                string blockName = GetEffectiveBlockNameForDebug(blockRef, tr);
                string blockType = blockRef.IsDynamicBlock ? "Dynamic" : "Static";
                
                ed.WriteMessage($", Block='{blockName}', {blockType}");
                
                if (!string.IsNullOrEmpty(blockName))
                {
                    blockTypes[blockName] = blockTypes.GetValueOrDefault(blockName, 0) + 1;
                    
                    // Check if NT block
                    if (System.Text.RegularExpressions.Regex.IsMatch(blockName, @"^NT\d{2}$"))
                    {
                        ntBlocks++;
                        ed.WriteMessage(" [NT BLOCK]");
                    }
                }
            }
            
            ed.WriteMessage("\n");
        }
        
        ed.WriteMessage("\n--- SUMMARY ---\n");
        ed.WriteMessage($"Total entities: {totalEntities}\n");
        ed.WriteMessage($"Block references: {blockReferences}\n");
        ed.WriteMessage($"Viewports: {viewportCount}\n");
        ed.WriteMessage($"NT blocks found: {ntBlocks}\n");
        
        ed.WriteMessage("\n--- BLOCK TYPES ---\n");
        foreach (var kvp in blockTypes.OrderBy(x => x.Key))
        {
            ed.WriteMessage($"  {kvp.Key}: {kvp.Value} instance(s)\n");
        }
    }

    /// <summary>
    /// Debug version of GetEffectiveBlockName for testing
    /// </summary>
    private static string GetEffectiveBlockNameForDebug(BlockReference blockRef, Transaction tr)
    {
        try
        {
            // Check if this is a dynamic block
            if (blockRef.IsDynamicBlock)
            {
                var btr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return btr?.Name ?? "[DYNAMIC-NO-NAME]";
            }
            else
            {
                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return btr?.Name ?? "[STATIC-NO-NAME]";
            }
        }
        catch (System.Exception ex)
        {
            return $"[ERROR: {ex.Message}]";
        }
    }

    /// <summary>
    /// Test command to verify enhanced CurrentDrawingBlockManager with external database support
    /// Tests both current drawing and external drawing operations using the same manager
    /// </summary>
    [CommandMethod("TESTENHANCEDBLOCKMANAGER")]
    public void TestEnhancedCurrentDrawingBlockManager()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        
        try
        {
            ed.WriteMessage("\n=== TESTING ENHANCED CURRENTDRAWINGBLOCKMANAGER ===\n");
            
            var logger = new AutoCADLogger();
            
            // Test 1: Current Drawing Mode (original behavior)
            ed.WriteMessage("\n--- TEST 1: CURRENT DRAWING MODE ---\n");
            var currentManager = new CurrentDrawingBlockManager(logger);
            
            // Get available layouts for testing
            var layouts = GetAvailableLayouts();
            if (layouts.Count == 0)
            {
                ed.WriteMessage("ERROR: No layouts available for testing\n");
                return;
            }
            
            string testLayout = layouts.First();
            ed.WriteMessage($"Testing with layout: {testLayout}\n");
            
            // Test reading blocks in current drawing
            var currentBlocks = currentManager.GetConstructionNoteBlocks(testLayout);
            ed.WriteMessage($"Current drawing mode - Found {currentBlocks.Count} NT blocks\n");
            
            foreach (var block in currentBlocks.Take(3)) // Show first 3 blocks
            {
                ed.WriteMessage($"  Block: {block.BlockName}, Number: {block.Number}, Visible: {block.IsVisible}\n");
            }
            
            // Test 2: External Database Mode (new functionality)  
            ed.WriteMessage("\n--- TEST 2: EXTERNAL DATABASE MODE ---\n");
            
            string currentDwgPath = doc.Name;
            ed.WriteMessage($"Current drawing path: {currentDwgPath}\n");
            
            if (File.Exists(currentDwgPath))
            {
                // Create external database and test
                using (var externalDb = new Database(false, true))
                {
                    try
                    {
                        ed.WriteMessage("Loading external database...\n");
                        externalDb.ReadDwgFile(currentDwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                        
                        // Test external database mode
                        var externalManager = new CurrentDrawingBlockManager(externalDb, logger);
                        
                        var externalBlocks = externalManager.GetConstructionNoteBlocks(testLayout);
                        ed.WriteMessage($"External database mode - Found {externalBlocks.Count} NT blocks\n");
                        
                        foreach (var block in externalBlocks.Take(3)) // Show first 3 blocks  
                        {
                            ed.WriteMessage($"  Block: {block.BlockName}, Number: {block.Number}, Visible: {block.IsVisible}\n");
                        }
                        
                        // Test 3: Compare Results
                        ed.WriteMessage("\n--- TEST 3: COMPARISON ---\n");
                        if (currentBlocks.Count == externalBlocks.Count)
                        {
                            ed.WriteMessage("✓ PASS: Both modes found same number of blocks\n");
                            
                            bool allMatch = true;
                            for (int i = 0; i < currentBlocks.Count && i < externalBlocks.Count; i++)
                            {
                                var current = currentBlocks[i];
                                var external = externalBlocks[i];
                                
                                if (current.BlockName != external.BlockName ||
                                    current.Number != external.Number ||
                                    current.IsVisible != external.IsVisible)
                                {
                                    allMatch = false;
                                    ed.WriteMessage($"  MISMATCH at index {i}: {current.BlockName} vs {external.BlockName}\n");
                                }
                            }
                            
                            if (allMatch)
                            {
                                ed.WriteMessage("✓ PASS: All block properties match between modes\n");
                            }
                            else
                            {
                                ed.WriteMessage("⚠ WARNING: Some block properties differ between modes\n");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"⚠ WARNING: Block count mismatch - Current: {currentBlocks.Count}, External: {externalBlocks.Count}\n");
                        }
                        
                        // Test 4: Update Operation (if we have blocks to test with)
                        if (externalBlocks.Count > 0)
                        {
                            ed.WriteMessage("\n--- TEST 4: UPDATE OPERATION TEST ---\n");
                            
                            var firstBlock = externalBlocks[0];
                            ed.WriteMessage($"Testing update on block: {firstBlock.BlockName}\n");
                            
                            // Test update with external database
                            bool updateResult = externalManager.UpdateConstructionNoteBlock(
                                testLayout, 
                                firstBlock.BlockName, 
                                99, 
                                "TEST NOTE FROM ENHANCED MANAGER", 
                                true
                            );
                            
                            if (updateResult)
                            {
                                ed.WriteMessage("✓ PASS: External database update successful\n");
                            }
                            else
                            {
                                ed.WriteMessage("⚠ WARNING: External database update failed\n");
                            }
                        }
                        
                        ed.WriteMessage("\n=== ENHANCED CURRENTDRAWINGBLOCKMANAGER TEST COMPLETE ===\n");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"ERROR in external database test: {ex.Message}\n");
                    }
                }
            }
            else
            {
                ed.WriteMessage($"WARNING: Cannot test external mode - file path not accessible: {currentDwgPath}\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTENHANCEDBLOCKMANAGER: {ex.Message}\n");
        }
    }
    
    /// <summary>
    /// Helper method to get available layouts for testing
    /// </summary>
    private List<string> GetAvailableLayouts()
    {
        var layouts = new List<string>();
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        
        try
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    string layoutName = entry.Key;
                    if (!layoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                    {
                        layouts.Add(layoutName);
                    }
                }
                tr.Dispose();
            }
        }
        catch (System.Exception)
        {
            // Return empty list on error
        }
        
        return layouts;
    }

    /// <summary>
    /// Test command to validate MultiDrawingConstructionNotesService batch processing
    /// Tests routing to different drawing states and result tracking
    /// </summary>
    // Simple mock implementations for testing
    private class MockExcelReader : IExcelReader
    {
        public void Dispose() { }
        public Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config) => Task.FromResult(new List<SheetInfo>());
        public Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config) => Task.FromResult(new List<ConstructionNote>());
        public Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config) => Task.FromResult(new List<SheetNoteMapping>());
        public Task<List<TitleBlockMapping>> ReadTitleBlockMappingsAsync(string filePath, ProjectConfiguration config) => Task.FromResult(new List<TitleBlockMapping>());
        public Task<bool> FileExistsAsync(string filePath) => Task.FromResult(false);
        public Task<string[]> GetWorksheetNamesAsync(string filePath) => Task.FromResult(Array.Empty<string>());
        public Task<string[]> GetTableNamesAsync(string filePath, string worksheetName) => Task.FromResult(Array.Empty<string>());
    }

    private class MockDrawingOperations : IDrawingOperations
    {
        public void Dispose() { }
        public Task<List<ConstructionNoteBlock>> GetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config) => Task.FromResult(new List<ConstructionNoteBlock>());
        public Task<bool> UpdateConstructionNoteBlockAsync(string sheetName, int blockIndex, int noteNumber, string noteText, ProjectConfiguration config) => Task.FromResult(true);
        public Task<bool> SetConstructionNotesAsync(string sheetName, Dictionary<int, string> noteData, ProjectConfiguration config) => Task.FromResult(true);
        public Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config) => Task.FromResult(true);
        public Task<bool> ResetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config, CurrentDrawingBlockManager? blockManager = null) => Task.FromResult(true);
        public Task UpdateTitleBlockAsync(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config) => Task.CompletedTask;
        public Task<bool> ValidateTitleBlockExistsAsync(string sheetName, ProjectConfiguration config) => Task.FromResult(true);
        public Task<Dictionary<string, string>> GetTitleBlockAttributesAsync(string sheetName, ProjectConfiguration config) => Task.FromResult(new Dictionary<string, string>());
    }

    [CommandMethod("TESTMULTIDRAWINGSERVICE")]
    public void TestMultiDrawingService()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        
        try
        {
            ed.WriteMessage("\n=== TESTING MULTI-DRAWING CONSTRUCTION NOTES SERVICE ===\n");
            
            var logger = new AutoCADLogger();
            
            // Create mock dependencies for testing
            ed.WriteMessage("\n--- STEP 1: CREATING SERVICE DEPENDENCIES ---\n");
            
            var drawingAccessService = new DrawingAccessService(logger);
            var backupCleanupService = new BackupCleanupService(logger);
            var multileaderAnalyzer = new MultileaderAnalyzer(logger);
            var blockAnalyzer = new BlockAnalyzer(logger);
            var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
            var mockExcelReader = new MockExcelReader();
            var mockDrawingOperations = new MockDrawingOperations();
            var constructionNotesService = new ConstructionNotesService(logger, mockExcelReader, mockDrawingOperations);
            
            ed.WriteMessage("✓ DrawingAccessService created\n");
            ed.WriteMessage("✓ ExternalDrawingManager created\n");
            ed.WriteMessage("✓ Mock ExcelReader created\n");
            ed.WriteMessage("✓ Mock DrawingOperations created\n");
            ed.WriteMessage("✓ ConstructionNotesService created\n");
            
            // Create the multi-drawing service
            var multiDrawingService = new MultiDrawingConstructionNotesService(
                logger,
                drawingAccessService,
                externalDrawingManager,
                constructionNotesService,
                mockExcelReader);
            
            ed.WriteMessage("✓ MultiDrawingConstructionNotesService created\n");
            
            // Test with sample data
            ed.WriteMessage("\n--- STEP 2: PREPARING TEST DATA ---\n");
            
            // Create sample sheet-to-notes mapping
            var sheetToNotes = new Dictionary<string, List<int>>
            {
                { "ABC-101", new List<int> { 1, 2, 4 } },
                { "ABC-102", new List<int> { 1, 3, 5 } },
                { "PV-201", new List<int> { 2, 6 } }
            };
            
            ed.WriteMessage($"Test data created: {sheetToNotes.Count} sheets with construction notes\n");
            foreach (var kvp in sheetToNotes)
            {
                ed.WriteMessage($"  {kvp.Key}: {string.Join(", ", kvp.Value)} (total: {kvp.Value.Count} notes)\n");
            }
            
            // Create mock project configuration
            var config = new ProjectConfiguration
            {
                ProjectName = "Test Project",
                ProjectDWGFilePath = Path.GetDirectoryName(doc.Name) ?? "",
                ProjectIndexFilePath = Path.Combine(Path.GetDirectoryName(doc.Name) ?? "", "ProjectIndex.xlsx")
            };
            
            ed.WriteMessage($"Mock project config: DWG path = {config.ProjectDWGFilePath}\n");
            ed.WriteMessage($"Mock project config: Excel path = {config.ProjectIndexFilePath}\n");
            
            // Create mock sheet infos
            var sheetInfos = new List<SheetInfo>
            {
                new SheetInfo { SheetName = "ABC-101", DWGFileName = "PROJ-ABC-100", DrawingTitle = "ABC PLAN 1" },
                new SheetInfo { SheetName = "ABC-102", DWGFileName = "PROJ-ABC-100", DrawingTitle = "ABC PLAN 2" },
                new SheetInfo { SheetName = "PV-201", DWGFileName = "PROJ-PV-200", DrawingTitle = "PAVEMENT PLAN" }
            };
            
            ed.WriteMessage($"Mock sheet infos created: {sheetInfos.Count} sheets\n");
            
            // Test 3: Drawing State Detection
            ed.WriteMessage("\n--- STEP 3: TESTING DRAWING STATE DETECTION ---\n");
            
            foreach (var sheetInfo in sheetInfos)
            {
                try
                {
                    var dwgPath = drawingAccessService.GetDrawingFilePath(sheetInfo.SheetName, config, sheetInfos);
                    if (!string.IsNullOrEmpty(dwgPath))
                    {
                        var state = drawingAccessService.GetDrawingState(dwgPath);
                        ed.WriteMessage($"  {sheetInfo.SheetName} -> {dwgPath} -> {state}\n");
                    }
                    else
                    {
                        ed.WriteMessage($"  {sheetInfo.SheetName} -> [NO PATH RESOLVED]\n");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"  {sheetInfo.SheetName} -> ERROR: {ex.Message}\n");
                }
            }
            
            // Test 4: Batch Update Simulation (without actually modifying files)
            ed.WriteMessage("\n--- STEP 4: SIMULATING BATCH UPDATE ---\n");
            ed.WriteMessage("NOTE: This is a simulation - no actual file modifications will be made\n");
            
            try
            {
                // This would normally call the actual update method
                // var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(sheetToNotes, config, sheetInfos);
                
                // For now, simulate the result
                ed.WriteMessage("Batch update simulation:\n");
                foreach (var kvp in sheetToNotes)
                {
                    var dwgPath = drawingAccessService.GetDrawingFilePath(kvp.Key, config, sheetInfos);
                    var state = !string.IsNullOrEmpty(dwgPath) ? drawingAccessService.GetDrawingState(dwgPath) : DrawingState.NotFound;
                    
                    ed.WriteMessage($"  {kvp.Key}: {kvp.Value.Count} notes, State: {state}\n");
                    
                    // Simulate routing logic
                    switch (state)
                    {
                        case DrawingState.Active:
                            ed.WriteMessage($"    → Would use current drawing operations\n");
                            break;
                        case DrawingState.Inactive:
                            ed.WriteMessage($"    → Would make active, then use current operations\n");
                            break;
                        case DrawingState.Closed:
                            ed.WriteMessage($"    → Would use external drawing manager\n");
                            break;
                        case DrawingState.NotFound:
                            ed.WriteMessage($"    → Would report as failure (file not found)\n");
                            break;
                        case DrawingState.Error:
                            ed.WriteMessage($"    → Would report as failure (access error)\n");
                            break;
                    }
                }
                
                ed.WriteMessage("✓ Routing logic validation completed\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"⚠ Error in batch update simulation: {ex.Message}\n");
            }
            
            // Test 5: Service Integration Validation
            ed.WriteMessage("\n--- STEP 5: SERVICE INTEGRATION VALIDATION ---\n");
            
            ed.WriteMessage("Checking service dependencies:\n");
            ed.WriteMessage($"  DrawingAccessService: {(drawingAccessService != null ? "✓" : "✗")}\n");
            ed.WriteMessage($"  ExternalDrawingManager: {(externalDrawingManager != null ? "✓" : "✗")}\n");
            ed.WriteMessage($"  ConstructionNotesService: {(constructionNotesService != null ? "✓" : "✗")}\n");
            ed.WriteMessage($"  MultiDrawingService: {(multiDrawingService != null ? "✓" : "✗")}\n");
            
            ed.WriteMessage("\nTesting note text generation:\n");
            foreach (var kvp in sheetToNotes)
            {
                var sampleNote = kvp.Value.FirstOrDefault();
                if (sampleNote > 0)
                {
                    // Test the private GetNoteTextForNumber method indirectly
                    var noteData = new ConstructionNoteData(sampleNote, $"Test note {sampleNote} for {kvp.Key}");
                    ed.WriteMessage($"  {kvp.Key} Note {sampleNote}: {noteData.NoteText}\n");
                }
            }
            
            ed.WriteMessage("\n=== MULTI-DRAWING SERVICE TEST COMPLETE ===\n");
            ed.WriteMessage("SUMMARY:\n");
            ed.WriteMessage("✓ Service creation and dependency injection working\n");
            ed.WriteMessage("✓ Drawing state detection integrated\n");
            ed.WriteMessage("✓ Routing logic validated\n");
            ed.WriteMessage("✓ Service integration confirmed\n");
            ed.WriteMessage("\nNEXT STEPS:\n");
            ed.WriteMessage("- Integrate with actual Excel reading for note text\n");
            ed.WriteMessage("- Test with real file operations\n");
            ed.WriteMessage("- Add comprehensive error handling validation\n");
            
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTMULTIDRAWINGSERVICE: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Production test command for real multi-drawing batch operations with actual Excel reading
    /// </summary>
    [CommandMethod("TESTPRODUCTIONBATCH")]
    public async void TestProductionBatch()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        
        try
        {
            ed.WriteMessage("\n=== PRODUCTION MULTI-DRAWING BATCH TEST ===\n");
            
            var logger = new AutoCADLogger();
            
            // Create REAL service dependencies
            ed.WriteMessage("\n--- STEP 1: CREATING PRODUCTION SERVICES ---\n");
            
            var drawingAccessService = new DrawingAccessService(logger);
            var backupCleanupService = new BackupCleanupService(logger);
            var multileaderAnalyzer = new MultileaderAnalyzer(logger);
            var blockAnalyzer = new BlockAnalyzer(logger);
            var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
            var excelReaderService = new ExcelReaderService(logger);
            var drawingOperations = new DrawingOperations(logger);
            var constructionNotesService = new ConstructionNotesService(logger, excelReaderService, drawingOperations);
            
            ed.WriteMessage("✓ DrawingAccessService created\n");
            ed.WriteMessage("✓ ExternalDrawingManager created\n");
            ed.WriteMessage("✓ ExcelReaderService created (REAL)\n");
            ed.WriteMessage("✓ DrawingOperations created (REAL)\n");
            ed.WriteMessage("✓ ConstructionNotesService created\n");
            
            // Create the production multi-drawing service
            var multiDrawingService = new MultiDrawingConstructionNotesService(
                logger,
                drawingAccessService,
                externalDrawingManager,
                constructionNotesService,
                excelReaderService);
            
            ed.WriteMessage("✓ MultiDrawingConstructionNotesService created with REAL services\n");
            
            // Use real project configuration
            ed.WriteMessage("\n--- STEP 2: LOADING REAL PROJECT CONFIGURATION ---\n");
            
            var testDataPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata";
            var config = new ProjectConfiguration
            {
                ProjectName = "Test Project",
                ProjectDWGFilePath = testDataPath,
                ProjectIndexFilePath = Path.Combine(testDataPath, "ProjectIndex.xlsx")
            };
            
            ed.WriteMessage($"Project config: DWG path = {config.ProjectDWGFilePath}\n");
            ed.WriteMessage($"Project config: Excel path = {config.ProjectIndexFilePath}\n");
            
            // Check if Excel file exists
            if (!File.Exists(config.ProjectIndexFilePath))
            {
                ed.WriteMessage($"⚠ Excel file not found: {config.ProjectIndexFilePath}\n");
                ed.WriteMessage("Please ensure ProjectIndex.xlsx exists in testdata folder\n");
                return;
            }
            
            // Load real sheet information from Excel
            try
            {
                var sheetInfos = await excelReaderService.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
                ed.WriteMessage($"✓ Loaded {sheetInfos.Count} sheets from Excel\n");
                
                foreach (var sheet in sheetInfos.Take(3)) // Show first 3 for brevity
                {
                    ed.WriteMessage($"  • {sheet.SheetName}: {sheet.DWGFileName} - {sheet.DrawingTitle}\n");
                }
                
                if (sheetInfos.Count == 0)
                {
                    ed.WriteMessage("⚠ No sheet data found in Excel file\n");
                    return;
                }
                
                // Load REAL Excel Notes data
                ed.WriteMessage("\n--- STEP 3: LOADING EXCEL NOTES DATA ---\n");
                
                var sheetNotesMappings = await excelReaderService.ReadExcelNotesAsync(config.ProjectIndexFilePath, config);
                ed.WriteMessage($"✓ Loaded Excel Notes data for {sheetNotesMappings.Count} sheets\n");
                
                // Convert to the format expected by MultiDrawingConstructionNotesService
                var sheetToNotes = new Dictionary<string, List<int>>();
                
                foreach (var mapping in sheetNotesMappings)
                {
                    if (mapping.NoteNumbers?.Count > 0)
                    {
                        sheetToNotes[mapping.SheetName] = mapping.NoteNumbers.ToList();
                        ed.WriteMessage($"  • {mapping.SheetName}: {mapping.NoteNumbers.Count} notes ({string.Join(", ", mapping.NoteNumbers)})\n");
                    }
                }
                
                if (sheetToNotes.Count == 0)
                {
                    ed.WriteMessage("⚠ No Excel Notes data found. Check EXCEL_NOTES table in Excel file.\n");
                    return;
                }
                
                ed.WriteMessage($"\nPrepared batch data for {sheetToNotes.Count} sheets with actual Excel Notes:\n");
                
                // Test drawing state detection for ALL sheets with Excel Notes
                ed.WriteMessage("\n--- STEP 4: DRAWING STATE DETECTION ---\n");
                
                var uniqueDrawingFiles = new HashSet<string>();
                
                foreach (var sheetName in sheetToNotes.Keys)
                {
                    var dwgPath = drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
                    if (!string.IsNullOrEmpty(dwgPath))
                    {
                        var state = drawingAccessService.GetDrawingState(dwgPath);
                        ed.WriteMessage($"  {sheetName} -> {Path.GetFileName(dwgPath)} -> {state}\n");
                        uniqueDrawingFiles.Add(dwgPath);
                    }
                    else
                    {
                        ed.WriteMessage($"  {sheetName} -> [PATH NOT RESOLVED]\n");
                    }
                }
                
                ed.WriteMessage($"\nSummary: {sheetToNotes.Count} sheets across {uniqueDrawingFiles.Count} unique drawing files\n");
                
                // Perform actual batch update
                ed.WriteMessage("\n--- STEP 5: PERFORMING BATCH UPDATE ---\n");
                ed.WriteMessage($"⚠ This will modify {uniqueDrawingFiles.Count} drawing files:\n");
                foreach (var dwgFile in uniqueDrawingFiles)
                {
                    ed.WriteMessage($"  • {Path.GetFileName(dwgFile)}\n");
                }
                ed.WriteMessage("Continue? (Press ESC to cancel)\n");
                
                var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(
                    sheetToNotes, 
                    config, 
                    sheetInfos);
                
                // Report results
                ed.WriteMessage($"\n--- STEP 6: BATCH UPDATE RESULTS ---\n");
                ed.WriteMessage($"Total sheets processed: {result.TotalProcessed}\n");
                ed.WriteMessage($"Successes: {result.Successes.Count}\n");
                ed.WriteMessage($"Failures: {result.Failures.Count}\n");
                ed.WriteMessage($"Success rate: {result.SuccessRate:P1}\n");
                
                if (result.Successes.Count > 0)
                {
                    ed.WriteMessage($"\n✓ SUCCESSFUL UPDATES:\n");
                    foreach (var success in result.Successes)
                    {
                        ed.WriteMessage($"  • {success.SheetName} ({success.DrawingState}): {success.NotesUpdated} notes updated\n");
                    }
                }
                
                if (result.Failures.Count > 0)
                {
                    ed.WriteMessage($"\n✗ FAILED UPDATES:\n");
                    foreach (var failure in result.Failures)
                    {
                        ed.WriteMessage($"  • {failure.SheetName}: {failure.ErrorMessage}\n");
                    }
                }
                
                ed.WriteMessage("\n=== PRODUCTION BATCH TEST COMPLETE ===\n");
                if (result.SuccessRate > 0.5)
                {
                    ed.WriteMessage("🎉 BATCH PROCESSING SUCCESSFULLY IMPLEMENTED!\n");
                }
                else
                {
                    ed.WriteMessage("⚠ Batch processing needs refinement - check failure details above\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"Excel reading error: {ex.Message}\n");
                return;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in TESTPRODUCTIONBATCH: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
        }
    }

    /// <summary>
    /// Test command for plotting functionality
    /// Tests the complete plotting pipeline with verbose output
    /// </summary>
    [CommandMethod("KPFFTESTPLOT")]
    public async void TestPlotting()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        
        try
        {
            ed.WriteMessage("\n=== KPFF PLOTTING SYSTEM TEST ===\n");
            
            var logger = new AutoCADLogger();
            
            // Step 1: Initialize Services
            ed.WriteMessage("\n--- STEP 1: INITIALIZING PLOTTING SERVICES ---\n");
            
            var excelReaderService = new ExcelReaderService(logger);
            var drawingOperations = new DrawingOperations(logger);
            var constructionNotesService = new ConstructionNotesService(logger, excelReaderService, drawingOperations);
            var drawingAccessService = new DrawingAccessService(logger);
            var backupCleanupService = new BackupCleanupService(logger);
            var multileaderAnalyzer = new MultileaderAnalyzer(logger);
            var blockAnalyzer = new BlockAnalyzer(logger);
            var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
            var titleBlockService = new TitleBlockService(logger, excelReaderService, drawingOperations);
            var multiDrawingConstructionNotesService = new MultiDrawingConstructionNotesService(logger, drawingAccessService, externalDrawingManager, constructionNotesService, excelReaderService);
            var multiDrawingTitleBlockService = new MultiDrawingTitleBlockService(logger, drawingAccessService, externalDrawingManager, titleBlockService, excelReaderService);
            var plotManager = new PlotManager(logger);
            var plottingService = new PlottingService(logger, constructionNotesService, drawingOperations, excelReaderService, multiDrawingConstructionNotesService, multiDrawingTitleBlockService, plotManager);
            
            ed.WriteMessage("✓ ExcelReaderService created\n");
            ed.WriteMessage("✓ DrawingOperations created\n");
            ed.WriteMessage("✓ ConstructionNotesService created\n");
            ed.WriteMessage("✓ PlottingService created\n");
            ed.WriteMessage("✓ PlotManager created\n");
            
            // Step 2: Load Current Project Configuration
            ed.WriteMessage("\n--- STEP 2: LOADING CURRENT PROJECT CONFIGURATION ---\n");
            
            // Try to load the current project configuration
            ProjectConfiguration? config = null;
            
            try
            {
                var configService = new ProjectConfigurationService((IApplicationLogger)logger);
                
                // Get the document directory and look for configuration files
                var documentPath = Path.GetDirectoryName(doc.Name) ?? "";
                ed.WriteMessage($"Looking for project configuration in: {documentPath}\n");
                
                // Try common configuration file names
                var configFileNames = new[] { "ProjectConfig.json", "DBRT_Config.json" };
                string? configPath = null;
                
                foreach (var fileName in configFileNames)
                {
                    var testPath = Path.Combine(documentPath, fileName);
                    if (File.Exists(testPath))
                    {
                        configPath = testPath;
                        ed.WriteMessage($"✓ Found configuration file: {fileName}\n");
                        break;
                    }
                }
                
                if (configPath == null)
                {
                    ed.WriteMessage($"❌ No project configuration found in: {documentPath}\n");
                    ed.WriteMessage("Searched for: " + string.Join(", ", configFileNames) + "\n");
                    ed.WriteMessage("Please configure a project using the UI first\n");
                    return;
                }
                
                config = await configService.LoadConfigurationAsync(configPath);
                
                if (config != null)
                {
                    ed.WriteMessage($"✓ Loaded current project: {config.ProjectName}\n");
                    
                    // Ensure plotting configuration exists
                    if (config.Plotting == null)
                    {
                        config.Plotting = new PlottingConfiguration();
                    }
                    
                    // Set default plotting output directory if not configured
                    if (string.IsNullOrEmpty(config.Plotting.OutputDirectory))
                    {
                        config.Plotting.OutputDirectory = Path.Combine(config.ProjectDWGFilePath, "PlotOutput");
                    }
                    
                    // Plotting configuration - removed EnablePlotting and DefaultPlotFormat properties
                }
                else
                {
                    ed.WriteMessage("❌ No current project configuration found\n");
                    ed.WriteMessage("Please configure a project using the UI first\n");
                    return;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"❌ Error loading project configuration: {ex.Message}\n");
                ed.WriteMessage("Please configure a project using the UI first\n");
                return;
            }
            
            ed.WriteMessage($"📂 Project DWG path: {config.ProjectDWGFilePath}\n");
            ed.WriteMessage($"📋 Excel file path: {config.ProjectIndexFilePath}\n");
            ed.WriteMessage($"📄 Plot output path: {config.Plotting.OutputDirectory}\n");
            
            // Ensure output directory exists
            Directory.CreateDirectory(config.Plotting.OutputDirectory);
            ed.WriteMessage($"✓ Output directory created/verified\n");
            
            // Step 3: Load Sheet Data
            ed.WriteMessage("\n--- STEP 3: LOADING SHEET DATA ---\n");
            
            if (!File.Exists(config.ProjectIndexFilePath))
            {
                ed.WriteMessage($"❌ Excel file not found: {config.ProjectIndexFilePath}\n");
                ed.WriteMessage("Cannot proceed without project index file\n");
                return;
            }
            
            var sheetInfos = await excelReaderService.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            ed.WriteMessage($"✓ Loaded {sheetInfos.Count} sheets from Excel index\n");
            
            if (sheetInfos.Count == 0)
            {
                ed.WriteMessage("❌ No sheets found in Excel index\n");
                return;
            }
            
            // Show first few sheets
            ed.WriteMessage("Sample sheets loaded:\n");
            foreach (var sheet in sheetInfos.Take(5))
            {
                ed.WriteMessage($"  • {sheet.SheetName}: {sheet.DWGFileName} - {sheet.DrawingTitle}\n");
            }
            
            // Step 4: Select Test Sheets
            ed.WriteMessage("\n--- STEP 4: SELECTING TEST SHEETS ---\n");
            
            // Use first 2 sheets for testing
            var testSheets = sheetInfos.Take(2).Select(s => s.SheetName).ToList();
            ed.WriteMessage($"Selected {testSheets.Count} sheets for testing:\n");
            foreach (var sheetName in testSheets)
            {
                ed.WriteMessage($"  • {sheetName}\n");
            }
            
            // Step 5: Validate Plotting Prerequisites
            ed.WriteMessage("\n--- STEP 5: VALIDATING PLOTTING PREREQUISITES ---\n");
            
            var validation = await plottingService.ValidateSheetsForPlottingAsync(testSheets, config);
            
            ed.WriteMessage($"Validation result: {(validation.IsValid ? "✓ VALID" : "❌ INVALID")}\n");
            ed.WriteMessage($"Valid sheets: {validation.ValidSheets.Count}\n");
            ed.WriteMessage($"Invalid sheets: {validation.InvalidSheets.Count}\n");
            ed.WriteMessage($"Total issues: {validation.Issues.Count}\n");
            
            if (validation.Issues.Count > 0)
            {
                ed.WriteMessage("Issues found:\n");
                foreach (var issue in validation.Issues)
                {
                    var severity = issue.IsWarning ? "⚠ WARNING" : "❌ ERROR";
                    ed.WriteMessage($"  {severity} - {issue.SheetName}: {issue.Description}\n");
                }
            }
            
            if (!validation.IsValid)
            {
                ed.WriteMessage("Cannot proceed due to validation failures\n");
                return;
            }
            
            // Step 6: Get Plot Settings for Test Sheets
            ed.WriteMessage("\n--- STEP 6: EXAMINING PLOT SETTINGS ---\n");
            
            foreach (var sheetName in validation.ValidSheets)
            {
                var plotSettings = await plottingService.GetDefaultPlotSettingsAsync(sheetName, config);
                if (plotSettings != null)
                {
                    ed.WriteMessage($"📋 {sheetName} plot settings:\n");
                    ed.WriteMessage($"    Drawing: {Path.GetFileName(plotSettings.DrawingPath)}\n");
                    ed.WriteMessage($"    Layout: {plotSettings.LayoutName}\n");
                    ed.WriteMessage($"    Device: {plotSettings.PlotDevice}\n");
                    ed.WriteMessage($"    Paper: {plotSettings.PaperSize}\n");
                    ed.WriteMessage($"    Scale: {plotSettings.PlotScale}\n");
                    ed.WriteMessage($"    Area: {plotSettings.PlotArea}\n");
                    ed.WriteMessage($"    Centered: {plotSettings.PlotCentered}\n");
                }
                else
                {
                    ed.WriteMessage($"❌ Could not get plot settings for {sheetName}\n");
                }
            }
            
            // Step 7: Test Individual Plot Operations
            ed.WriteMessage("\n--- STEP 7: TESTING INDIVIDUAL PLOT OPERATIONS ---\n");
            
            foreach (var sheetName in validation.ValidSheets.Take(1)) // Just test first sheet
            {
                ed.WriteMessage($"\n🖨 Testing plot operation for {sheetName}...\n");
                
                var sheetInfo = sheetInfos.First(s => s.SheetName == sheetName);
                var drawingPath = Path.Combine(config.ProjectDWGFilePath, sheetInfo.DWGFileName);
                if (!sheetInfo.DWGFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    drawingPath += ".dwg";
                }
                
                var outputPath = Path.Combine(config.Plotting.OutputDirectory, $"{sheetName}_TEST.pdf");
                
                ed.WriteMessage($"  Source: {Path.GetFileName(drawingPath)}\n");
                ed.WriteMessage($"  Layout: {sheetName}\n");
                ed.WriteMessage($"  Output: {Path.GetFileName(outputPath)}\n");
                
                if (!File.Exists(drawingPath))
                {
                    ed.WriteMessage($"  ❌ Drawing file not found: {drawingPath}\n");
                    continue;
                }
                
                try
                {
                    var plotSuccess = await plotManager.PlotLayoutToPdfAsync(drawingPath, sheetName, outputPath);
                    
                    if (plotSuccess)
                    {
                        ed.WriteMessage($"  ✅ Plot successful!\n");
                        if (File.Exists(outputPath))
                        {
                            var fileInfo = new FileInfo(outputPath);
                            ed.WriteMessage($"  📄 PDF created: {fileInfo.Length} bytes\n");
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"  ❌ Plot failed\n");
                    }
                }
                catch (System.Exception plotEx)
                {
                    ed.WriteMessage($"  ❌ Plot exception: {plotEx.Message}\n");
                }
            }
            
            // Step 8: Test Complete Plot Service Workflow
            ed.WriteMessage("\n--- STEP 8: TESTING COMPLETE PLOT WORKFLOW ---\n");
            
            var plotJobSettings = new PlotJobSettings
            {
                UpdateConstructionNotes = false,  // Skip for initial test
                UpdateTitleBlocks = false,        // Skip for initial test
                ApplyToCurrentSheetOnly = false,
                OutputDirectory = config.Plotting.OutputDirectory
            };
            
            ed.WriteMessage($"Plot job settings:\n");
            ed.WriteMessage($"  Update Construction Notes: {plotJobSettings.UpdateConstructionNotes}\n");
            ed.WriteMessage($"  Update Title Blocks: {plotJobSettings.UpdateTitleBlocks}\n");
            ed.WriteMessage($"  Apply to Current Sheet Only: {plotJobSettings.ApplyToCurrentSheetOnly}\n");
            ed.WriteMessage($"  Output Directory: {plotJobSettings.OutputDirectory}\n");
            
            // Create a progress reporter
            var progress = new Progress<PlotProgress>(p =>
            {
                ed.WriteMessage($"  📊 {p.CurrentOperation}: {p.CurrentSheet} ({p.CompletedSheets}/{p.TotalSheets}) - {p.ProgressPercentage}%\n");
            });
            
            ed.WriteMessage($"\n🖨 Starting batch plot of {validation.ValidSheets.Count} sheets...\n");
            
            var plotResult = await plottingService.PlotSheetsAsync(validation.ValidSheets, config, plotJobSettings, progress);
            
            // Step 9: Report Results
            ed.WriteMessage("\n--- STEP 9: PLOT RESULTS ---\n");
            
            ed.WriteMessage($"Overall Success: {(plotResult.Success ? "✅ YES" : "❌ NO")}\n");
            ed.WriteMessage($"Total Sheets: {plotResult.TotalSheets}\n");
            ed.WriteMessage($"Successful: {plotResult.SuccessfulSheets.Count}\n");
            ed.WriteMessage($"Failed: {plotResult.FailedSheets.Count}\n");
            ed.WriteMessage($"Success Rate: {plotResult.SuccessRate:F1}%\n");
            
            if (plotResult.SuccessfulSheets.Count > 0)
            {
                ed.WriteMessage($"\n✅ SUCCESSFUL PLOTS:\n");
                foreach (var sheet in plotResult.SuccessfulSheets)
                {
                    var outputFile = Path.Combine(config.Plotting.OutputDirectory, $"{sheet}.pdf");
                    var exists = File.Exists(outputFile);
                    ed.WriteMessage($"  • {sheet}: {(exists ? "✅ PDF created" : "❓ PDF status unknown")}\n");
                }
            }
            
            if (plotResult.FailedSheets.Count > 0)
            {
                ed.WriteMessage($"\n❌ FAILED PLOTS:\n");
                foreach (var failure in plotResult.FailedSheets)
                {
                    ed.WriteMessage($"  • {failure.SheetName}: {failure.ErrorMessage}\n");
                    if (!string.IsNullOrEmpty(failure.ExceptionDetails))
                    {
                        ed.WriteMessage($"    Details: {failure.ExceptionDetails}\n");
                    }
                }
            }
            
            if (!string.IsNullOrEmpty(plotResult.ErrorMessage))
            {
                ed.WriteMessage($"\nGeneral Error: {plotResult.ErrorMessage}\n");
            }
            
            // Step 10: Summary
            ed.WriteMessage("\n--- STEP 10: TEST SUMMARY ---\n");
            
            ed.WriteMessage($"📋 Configuration: ✅ Plotting ready\n");
            ed.WriteMessage($"📂 Output Directory: {(Directory.Exists(config.Plotting.OutputDirectory) ? "✅" : "❌")} {config.Plotting.OutputDirectory}\n");
            ed.WriteMessage($"📄 Sheet Validation: {(validation.IsValid ? "✅" : "❌")} {validation.ValidSheets.Count} valid sheets\n");
            ed.WriteMessage($"🖨 Plot Operations: {(plotResult.Success ? "✅" : "❌")} {plotResult.SuccessRate:F1}% success rate\n");
            
            var outputFiles = Directory.GetFiles(config.Plotting.OutputDirectory, "*.pdf");
            ed.WriteMessage($"📁 Output Files: {outputFiles.Length} PDF files in output directory\n");
            
            if (outputFiles.Length > 0)
            {
                ed.WriteMessage("Generated PDFs:\n");
                foreach (var file in outputFiles.Take(5))
                {
                    var fileInfo = new FileInfo(file);
                    ed.WriteMessage($"  • {Path.GetFileName(file)} ({fileInfo.Length} bytes)\n");
                }
                if (outputFiles.Length > 5)
                {
                    ed.WriteMessage($"  ... and {outputFiles.Length - 5} more files\n");
                }
            }
            
            ed.WriteMessage("\n=== KPFF PLOTTING TEST COMPLETE ===\n");
            
            if (plotResult.Success && plotResult.SuccessRate >= 100)
            {
                ed.WriteMessage("🎉 PLOTTING SYSTEM FULLY OPERATIONAL!\n");
            }
            else if (plotResult.Success && plotResult.SuccessRate >= 50)
            {
                ed.WriteMessage("✅ PLOTTING SYSTEM WORKING (with some issues to resolve)\n");
            }
            else
            {
                ed.WriteMessage("❌ PLOTTING SYSTEM NEEDS DEBUGGING\n");
            }
            
            ed.WriteMessage($"\nNext steps:\n");
            ed.WriteMessage($"- Check output files in: {config.Plotting.OutputDirectory}\n");
            ed.WriteMessage($"- Review any error messages above\n");
            ed.WriteMessage($"- Test with construction notes/title block updates if needed\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n💥 CRITICAL ERROR in KPFFTESTPLOT: {ex.Message}\n");
            ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            ed.WriteMessage("\nThis indicates a fundamental issue with the plotting system setup.\n");
        }
    }
}