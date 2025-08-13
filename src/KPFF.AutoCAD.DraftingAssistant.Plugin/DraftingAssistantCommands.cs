using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Models;
using System.IO;
using System.Text.Json;
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

    [CommandMethod("KPFF")]
    public void MainDraftingAssistant()
    {
        // Main entry point command that triggers service initialization
        ExecuteCommand<ShowPaletteCommandHandler>();
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
            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not available. Services may not be registered.\n");
                return;
            }

            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
            var appLogger = serviceProvider.GetService<Core.Interfaces.IApplicationLogger>();
            var excelReader = serviceProvider.GetService<Core.Interfaces.IExcelReader>();
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
            
            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var logger = serviceProvider.GetService<Core.Interfaces.ILogger>();
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
            
            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var excelReader = serviceProvider.GetService<Core.Interfaces.IExcelReader>();
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
            
            var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
            if (serviceProvider == null)
            {
                ed.WriteMessage("ERROR: Service provider not available.\n");
                return;
            }

            var excelReader = serviceProvider.GetService<Core.Interfaces.IExcelReader>();
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
        var serviceProvider = DraftingAssistantExtensionApplication.ServiceProvider;
        var editor = Application.DocumentManager.MdiActiveDocument?.Editor;
        
        if (editor == null || serviceProvider == null)
        {
            return;
        }
        
        try
        {
            editor.WriteMessage("\n=== Testing Excel Reader Shared Process Behavior ===");
            
            // Get first Excel reader service from DI container (transient)
            editor.WriteMessage("\nGetting Excel reader service from DI container (first call)...");
            var excelReader1 = serviceProvider.GetService<IExcelReader>();
            
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
            editor.WriteMessage($"\n  Process Running: {SharedExcelReaderProcess.IsRunning}");
            editor.WriteMessage($"\n  Reference Count: {SharedExcelReaderProcess.ReferenceCount}");
            
            // Get another instance - should be different (transient) but share same process
            editor.WriteMessage("\n\nGetting Excel reader service again (second call)...");
            var excelReader2 = serviceProvider.GetService<IExcelReader>();
            
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
            editor.WriteMessage($"\n  Process Running: {SharedExcelReaderProcess.IsRunning}");
            editor.WriteMessage($"\n  Reference Count: {SharedExcelReaderProcess.ReferenceCount}");
            
            // Dispose instances
            editor.WriteMessage("\n\nDisposing first instance...");
            excelReader1.Dispose();
            editor.WriteMessage($"  Reference Count after dispose: {SharedExcelReaderProcess.ReferenceCount}");
            
            editor.WriteMessage("\nDisposing second instance...");
            excelReader2.Dispose();
            editor.WriteMessage($"  Reference Count after dispose: {SharedExcelReaderProcess.ReferenceCount}");
            
            editor.WriteMessage("\n\n=== Shared Process Behavior Summary ===");
            editor.WriteMessage("\n- Excel reader services are TRANSIENT (new instance each time)");
            editor.WriteMessage("\n- All instances share ONE Excel reader process");
            editor.WriteMessage("\n- Process stays alive even with 0 references");
            editor.WriteMessage("\n- Process will be terminated when AutoCAD shuts down");
            editor.WriteMessage("\n- Check Task Manager: only one KPFF.AutoCAD.ExcelReader.exe should be running");
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\nError during shared process test: {ex.Message}");
        }
    }
}