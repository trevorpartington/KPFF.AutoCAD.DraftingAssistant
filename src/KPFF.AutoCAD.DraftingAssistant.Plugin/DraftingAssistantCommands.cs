using Autodesk.AutoCAD.ApplicationServices;
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