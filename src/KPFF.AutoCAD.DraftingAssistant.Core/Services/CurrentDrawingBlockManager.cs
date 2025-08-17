using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.Text.RegularExpressions;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Phase 1: Read-only operations for construction note blocks in the current drawing
/// Uses simplified NT## naming pattern (NT01, NT02, etc.) that works across all layouts
/// This service safely reads block information without making any modifications
/// </summary>
public class CurrentDrawingBlockManager
{
    private readonly ILogger _logger;
    private readonly Regex _noteBlockPattern = new Regex(@"^NT\d{2}$", RegexOptions.Compiled);
    private bool _isInitialized = false;

    public CurrentDrawingBlockManager(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Ensures AutoCAD context is properly initialized before operations
    /// </summary>
    private bool EnsureInitialized()
    {
        if (_isInitialized)
        {
            return true;
        }

        try
        {
            _logger.LogDebug("Initializing CurrentDrawingBlockManager - verifying AutoCAD context");
            
            // Verify AutoCAD document access
            var doc = Application.DocumentManager?.MdiActiveDocument;
            if (doc == null)
            {
                _logger.LogError("No active AutoCAD document found during initialization");
                return false;
            }

            // Verify database access
            var db = doc.Database;
            if (db == null)
            {
                _logger.LogError("AutoCAD database not accessible during initialization");
                return false;
            }

            // Test basic transaction access
            using (var testTrans = db.TransactionManager.StartTransaction())
            {
                testTrans.Commit();
            }

            _isInitialized = true;
            _logger.LogDebug("CurrentDrawingBlockManager initialized successfully");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to initialize CurrentDrawingBlockManager: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets all construction note blocks for a specific layout in the current drawing
    /// </summary>
    /// <param name="layoutName">Name of the layout (e.g., "101")</param>
    /// <returns>List of construction note blocks found</returns>
    public List<ConstructionNoteBlock> GetConstructionNoteBlocks(string layoutName)
    {
        var blocks = new List<ConstructionNoteBlock>();
        _logger.LogInformation($"Starting search for construction note blocks in layout: {layoutName}");

        if (!EnsureInitialized())
        {
            _logger.LogError("Failed to initialize AutoCAD context for block operations");
            return blocks;
        }

        try
        {
            // CRASH FIX: Add safety guard around AutoCAD object access
            Document? doc = null;
            try
            {
                doc = Application.DocumentManager?.MdiActiveDocument;
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to access AutoCAD document manager: {ex.Message}", ex);
                return blocks;
            }

            if (doc == null)
            {
                _logger.LogWarning("No active document found");
                return blocks;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the layout
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(layoutName))
                    {
                        _logger.LogWarning($"Layout '{layoutName}' not found in drawing");
                        LogAvailableLayouts(layoutDict, tr);
                        return blocks;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                    // Get the block table record for this layout
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    _logger.LogDebug($"Scanning layout '{layoutName}' for construction note blocks");

                    // Iterate through all entities in the layout
                    foreach (ObjectId objId in layoutBtr)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        // Check if this is a block reference
                        if (entity is BlockReference blockRef)
                        {
                            string blockName = GetEffectiveBlockName(blockRef, tr);
                            _logger.LogDebug($"Found block: {blockName}");

                            // Check if this matches our construction note pattern
                            if (IsConstructionNoteBlock(blockName))
                            {
                                _logger.LogDebug($"Found construction note block: {blockName}");
                                
                                var noteBlock = ReadConstructionNoteBlock(blockRef, blockName, tr);
                                if (noteBlock != null)
                                {
                                    blocks.Add(noteBlock);
                                }
                            }
                        }
                    }

                    _logger.LogInformation($"Found {blocks.Count} NT construction note blocks in layout '{layoutName}'");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error reading blocks from layout: {ex.Message}", ex);
                }
                finally
                {
                    // No commit needed for read-only operation
                    tr.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get construction note blocks: {ex.Message}", ex);
        }

        return blocks;
    }

    /// <summary>
    /// Gets the effective name of a block, handling dynamic blocks
    /// </summary>
    private string GetEffectiveBlockName(BlockReference blockRef, Transaction tr)
    {
        try
        {
            if (blockRef.IsDynamicBlock)
            {
                // For dynamic blocks, get the dynamic block table record
                BlockTableRecord dynamicBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return dynamicBtr.Name;
            }
            else
            {
                // For regular blocks, get the block table record
                BlockTableRecord btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return btr.Name;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Could not get effective block name: {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// Checks if a block name matches the construction note pattern
    /// </summary>
    private bool IsConstructionNoteBlock(string blockName)
    {
        if (string.IsNullOrEmpty(blockName))
            return false;

        // Pattern: NT{00}
        // Example: NT01, NT02, NT03, etc.
        return _noteBlockPattern.IsMatch(blockName);
    }

    /// <summary>
    /// Reads all properties of a construction note block
    /// </summary>
    private ConstructionNoteBlock ReadConstructionNoteBlock(BlockReference blockRef, string blockName, Transaction tr)
    {
        try
        {
            var noteBlock = new ConstructionNoteBlock
            {
                BlockName = blockName,
                IsVisible = true // Default to visible
            };

            // Read attributes
            var attributes = ReadBlockAttributes(blockRef, tr);
            if (attributes.ContainsKey("NUMBER"))
            {
                if (int.TryParse(attributes["NUMBER"], out int noteNumber))
                {
                    noteBlock.Number = noteNumber;
                }
            }
            if (attributes.ContainsKey("NOTE"))
            {
                noteBlock.Note = attributes["NOTE"];
            }

            // Read dynamic block visibility
            if (blockRef.IsDynamicBlock)
            {
                noteBlock.IsVisible = ReadDynamicBlockVisibility(blockRef);
            }

            _logger.LogDebug($"Read NT block {blockName}: Number='{noteBlock.Number}', Note='{noteBlock.Note?.Substring(0, Math.Min(noteBlock.Note?.Length ?? 0, 30))}...', Visible={noteBlock.IsVisible}");

            return noteBlock;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to read construction note block {blockName}: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Reads all attributes from a block reference
    /// </summary>
    private Dictionary<string, string> ReadBlockAttributes(BlockReference blockRef, Transaction tr)
    {
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            // Get the attribute collection
            AttributeCollection attCol = blockRef.AttributeCollection;
            
            foreach (ObjectId attId in attCol)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef != null)
                {
                    string tag = attRef.Tag.ToUpper();
                    string value = attRef.TextString;
                    attributes[tag] = value;
                    _logger.LogDebug($"  Attribute: {tag} = '{value}'");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to read block attributes: {ex.Message}", ex);
        }

        return attributes;
    }

    /// <summary>
    /// Reads the visibility state of a dynamic block
    /// </summary>
    private bool ReadDynamicBlockVisibility(BlockReference blockRef)
    {
        try
        {
            // Get dynamic block properties
            DynamicBlockReferencePropertyCollection props = blockRef.DynamicBlockReferencePropertyCollection;
            
            foreach (DynamicBlockReferenceProperty prop in props)
            {
                if (prop.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase) ||
                    prop.PropertyName.Equals("Visibility1", StringComparison.OrdinalIgnoreCase))
                {
                    string visibilityState = prop.Value.ToString();
                    _logger.LogDebug($"  Visibility state: {visibilityState}");
                    
                    // Check if visibility is "ON" or similar
                    return visibilityState.Equals("ON", StringComparison.OrdinalIgnoreCase) ||
                           visibilityState.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
                           visibilityState.Equals("Hex", StringComparison.OrdinalIgnoreCase); // From LISP code
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Could not read dynamic block visibility: {ex.Message}");
        }

        // Default to visible if we can't determine
        return true;
    }

    /// <summary>
    /// Gets a summary of all construction note blocks across all layouts
    /// Each layout will contain instances of NT01, NT02, etc. with layout-specific attribute values
    /// </summary>
    public Dictionary<string, List<ConstructionNoteBlock>> GetAllConstructionNoteBlocks()
    {
        var allBlocks = new Dictionary<string, List<ConstructionNoteBlock>>();
        _logger.LogInformation("Starting search for construction note blocks in all layouts");

        try
        {
            // CRASH FIX: Add safety guard around AutoCAD object access
            Document? doc = null;
            try
            {
                doc = Application.DocumentManager?.MdiActiveDocument;
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to access AutoCAD document manager: {ex.Message}", ex);
                return allBlocks;
            }

            if (doc == null)
            {
                _logger.LogWarning("No active document found");
                return allBlocks;
            }

            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get all layouts
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        string layoutName = entry.Key;
                        
                        // Skip Model space
                        if (layoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                            continue;

                        _logger.LogDebug($"Checking layout: {layoutName}");
                        var blocks = GetConstructionNoteBlocks(layoutName);
                        
                        if (blocks.Count > 0)
                        {
                            allBlocks[layoutName] = blocks;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error scanning layouts: {ex.Message}", ex);
                }
                finally
                {
                    tr.Dispose();
                }
            }

            _logger.LogInformation($"Found NT construction note blocks in {allBlocks.Count} layouts");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get all construction note blocks: {ex.Message}", ex);
        }

        return allBlocks;
    }

    /// <summary>
    /// Clears all construction note blocks in a layout (sets visibility OFF and clears attributes)
    /// This should be called before updating blocks to ensure removed notes don't persist
    /// </summary>
    /// <param name="layoutName">Layout containing the blocks</param>
    /// <returns>Number of blocks cleared</returns>
    public int ClearAllConstructionNoteBlocks(string layoutName)
    {
        _logger.LogInformation($"Clearing all construction note blocks in layout {layoutName}");
        int clearedCount = 0;
        
        try
        {
            Document? doc = null;
            try
            {
                doc = Application.DocumentManager?.MdiActiveDocument;
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to access AutoCAD document manager: {ex.Message}", ex);
                return 0;
            }

            if (doc == null)
            {
                _logger.LogWarning("No active document found");
                return 0;
            }

            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the layout
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(layoutName))
                    {
                        _logger.LogWarning($"Layout '{layoutName}' not found in drawing");
                        return 0;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    // Find all NT## blocks
                    foreach (ObjectId objId in layoutBtr)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (entity is BlockReference blockRef)
                        {
                            string currentBlockName = GetEffectiveBlockName(blockRef, tr);
                            
                            // Check if this is an NT## block
                            if (System.Text.RegularExpressions.Regex.IsMatch(currentBlockName, @"^NT\d{2}$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                            {
                                // Open for write
                                BlockReference writeBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                                
                                // Clear attributes
                                AttributeCollection attCol = writeBlockRef.AttributeCollection;
                                foreach (ObjectId attId in attCol)
                                {
                                    AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                    if (attRef != null)
                                    {
                                        string tag = attRef.Tag.ToUpper();
                                        if (tag == "NUMBER" || tag == "NOTE")
                                        {
                                            attRef.TextString = "";
                                        }
                                    }
                                }
                                
                                // Set visibility to OFF
                                if (writeBlockRef.IsDynamicBlock)
                                {
                                    DynamicBlockReferencePropertyCollection props = writeBlockRef.DynamicBlockReferencePropertyCollection;
                                    foreach (DynamicBlockReferenceProperty prop in props)
                                    {
                                        if (prop.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase) ||
                                            prop.PropertyName.Equals("Visibility1", StringComparison.OrdinalIgnoreCase))
                                        {
                                            try
                                            {
                                                prop.Value = "OFF";
                                                clearedCount++;
                                                _logger.LogDebug($"Cleared block {currentBlockName}");
                                                break;
                                            }
                                            catch (Exception ex)
                                            {
                                                _logger.LogWarning($"Could not set visibility for {currentBlockName}: {ex.Message}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    tr.Commit();
                    _logger.LogInformation($"Successfully cleared {clearedCount} construction note blocks in layout {layoutName}");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error clearing blocks: {ex.Message}", ex);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to clear construction note blocks: {ex.Message}", ex);
        }

        return clearedCount;
    }

    /// <summary>
    /// Phase 2: Updates a single construction note block with new data
    /// Safe modification method with proper transaction handling
    /// </summary>
    /// <param name="layoutName">Layout containing the block</param>
    /// <param name="blockName">Block name (e.g., "NT01")</param>
    /// <param name="noteNumber">Note number to set</param>
    /// <param name="noteText">Note text to set</param>
    /// <param name="makeVisible">Whether to make the block visible</param>
    /// <returns>True if update was successful</returns>
    public bool UpdateConstructionNoteBlock(string layoutName, string blockName, int noteNumber, string noteText, bool makeVisible)
    {
        _logger.LogInformation($"=== UpdateConstructionNoteBlock ENTRY ====");
        _logger.LogInformation($"Parameters: layoutName='{layoutName}', blockName='{blockName}', noteNumber={noteNumber}, noteText='{noteText?.Substring(0, Math.Min(noteText?.Length ?? 0, 30))}...', makeVisible={makeVisible}");
        _logger.LogInformation($"Logger type: {_logger.GetType().Name}");
        
        if (!EnsureInitialized())
        {
            _logger.LogError("Failed to initialize AutoCAD context for block update operations");
            return false;
        }

        try
        {
            // CRASH FIX: Add safety guard around AutoCAD object access
            Document? doc = null;
            try
            {
                _logger.LogDebug("Accessing AutoCAD DocumentManager...");
                doc = Application.DocumentManager?.MdiActiveDocument;
                _logger.LogDebug($"Document obtained: {doc?.Name ?? "null"}");
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to access AutoCAD document manager: {ex.Message}", ex);
                return false;
            }

            if (doc == null)
            {
                _logger.LogError("CRITICAL: No active document found - cannot update blocks");
                return false;
            }

            _logger.LogDebug($"Active document: {doc.Name}");
            Database db = doc.Database;
            _logger.LogDebug($"Database filename: {db.Filename}");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    _logger.LogDebug("Starting transaction for block update...");
                    
                    // Get the layout dictionary
                    _logger.LogDebug("Getting layout dictionary...");
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    _logger.LogDebug($"Layout dictionary contains {layoutDict?.Count ?? 0} entries");
                    
                    // Check if our target layout exists
                    _logger.LogInformation($"Checking if layout '{layoutName}' exists in dictionary...");
                    bool layoutExists = layoutDict.Contains(layoutName);
                    _logger.LogInformation($"Layout '{layoutName}' exists: {layoutExists}");
                    
                    if (!layoutExists)
                    {
                        _logger.LogError($"LAYOUT NOT FOUND: Layout '{layoutName}' not found in drawing");
                        _logger.LogInformation("Listing all available layouts for debugging:");
                        LogAvailableLayouts(layoutDict, tr);
                        return false;
                    }

                    _logger.LogDebug($"Getting layout object for '{layoutName}'...");
                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    _logger.LogDebug($"Layout object obtained: {layout?.LayoutName}");

                    // Get the block table record for this layout
                    _logger.LogDebug("Getting layout block table record...");
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    
                    // Count entities in the layout (BlockTableRecord implements IEnumerable<ObjectId>)
                    int entityCount = 0;
                    if (layoutBtr != null)
                    {
                        foreach (ObjectId objId in layoutBtr)
                        {
                            entityCount++;
                        }
                    }
                    _logger.LogDebug($"Layout BTR contains {entityCount} entities");

                    // Find the target block
                    _logger.LogInformation($"Searching for block '{blockName}' in layout '{layoutName}'...");
                    BlockReference targetBlockRef = null;
                    ObjectId targetBlockId = ObjectId.Null;
                    int blocksFound = 0;
                    
                    foreach (ObjectId objId in layoutBtr)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (entity is BlockReference blockRef)
                        {
                            string currentBlockName = GetEffectiveBlockName(blockRef, tr);
                            _logger.LogDebug($"Found block reference: '{currentBlockName}'");
                            blocksFound++;
                            
                            if (currentBlockName.Equals(blockName, StringComparison.OrdinalIgnoreCase))
                            {
                                _logger.LogInformation($"MATCH FOUND: Block '{blockName}' located in layout");
                                targetBlockId = objId;
                                break;
                            }
                        }
                    }
                    
                    _logger.LogInformation($"Block search complete: {blocksFound} total block references found in layout");

                    if (targetBlockId.IsNull)
                    {
                        _logger.LogError($"BLOCK NOT FOUND: Block '{blockName}' not found in layout '{layoutName}'");
                        _logger.LogError($"This means the block either doesn't exist or has a different name than expected");
                        
                        // List some of the blocks we did find for debugging
                        _logger.LogInformation("First few blocks found in layout (for debugging):");
                        int debugCount = 0;
                        foreach (ObjectId objId in layoutBtr)
                        {
                            if (debugCount >= 10) break; // Limit debug output
                            Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                            if (entity is BlockReference blockRef)
                            {
                                string currentBlockName = GetEffectiveBlockName(blockRef, tr);
                                _logger.LogInformation($"  Debug block #{debugCount + 1}: '{currentBlockName}'");
                                debugCount++;
                            }
                        }
                        
                        return false;
                    }

                    // Now open the block for write
                    _logger.LogDebug($"Opening block for write access...");
                    targetBlockRef = tr.GetObject(targetBlockId, OpenMode.ForWrite) as BlockReference;
                    _logger.LogDebug($"Block opened for write: {targetBlockRef != null}");

                    _logger.LogInformation($"Found target block {blockName}, updating attributes and visibility");

                    // Update block attributes
                    _logger.LogDebug($"Updating block attributes...");
                    bool attributesUpdated = UpdateBlockAttributes(targetBlockRef, noteNumber, noteText, tr);
                    _logger.LogInformation($"Attributes updated: {attributesUpdated}");
                    
                    // Update visibility state if it's a dynamic block
                    _logger.LogDebug($"Checking if block is dynamic: {targetBlockRef.IsDynamicBlock}");
                    bool visibilityUpdated = true;
                    if (targetBlockRef.IsDynamicBlock)
                    {
                        _logger.LogDebug($"Updating dynamic block visibility...");
                        visibilityUpdated = UpdateDynamicBlockVisibility(targetBlockRef, makeVisible);
                        _logger.LogInformation($"Visibility updated: {visibilityUpdated}");
                    }
                    else
                    {
                        _logger.LogInformation("Block is not dynamic - skipping visibility update");
                    }

                    if (attributesUpdated && visibilityUpdated)
                    {
                        _logger.LogDebug("Committing transaction...");
                        tr.Commit();
                        _logger.LogInformation($"SUCCESS: Block {blockName} updated successfully!");
                        _logger.LogInformation($"  Number: {noteNumber}");
                        _logger.LogInformation($"  Note: '{noteText?.Substring(0, Math.Min(noteText?.Length ?? 0, 50))}...'");
                        _logger.LogInformation($"  Visible: {makeVisible}");
                        return true;
                    }
                    else
                    {
                        _logger.LogError($"TRANSACTION ROLLBACK: Block update failed");
                        _logger.LogError($"  Attributes updated: {attributesUpdated}");
                        _logger.LogError($"  Visibility updated: {visibilityUpdated}");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error updating block: {ex.Message}", ex);
                    return false;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update construction note block: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates the attributes of a construction note block
    /// </summary>
    private bool UpdateBlockAttributes(BlockReference blockRef, int noteNumber, string noteText, Transaction tr)
    {
        try
        {
            AttributeCollection attCol = blockRef.AttributeCollection;
            bool numberUpdated = false;
            bool noteUpdated = false;
            
            foreach (ObjectId attId in attCol)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null)
                {
                    string tag = attRef.Tag.ToUpper();
                    
                    if (tag == "NUMBER")
                    {
                        attRef.TextString = noteNumber > 0 ? noteNumber.ToString() : "";
                        numberUpdated = true;
                        _logger.LogDebug($"Updated NUMBER attribute: '{attRef.TextString}'");
                    }
                    else if (tag == "NOTE")
                    {
                        attRef.TextString = noteText ?? "";
                        noteUpdated = true;
                        _logger.LogDebug($"Updated NOTE attribute: '{attRef.TextString.Substring(0, Math.Min(attRef.TextString.Length, 50))}...'");
                    }
                }
            }

            if (!numberUpdated)
            {
                _logger.LogWarning("NUMBER attribute not found or not updated");
            }
            if (!noteUpdated)
            {
                _logger.LogWarning("NOTE attribute not found or not updated");
            }

            return numberUpdated && noteUpdated;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update block attributes: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates the visibility state of a dynamic block with retry logic
    /// </summary>
    private bool UpdateDynamicBlockVisibility(BlockReference blockRef, bool makeVisible)
    {
        const int maxRetries = 3;
        const int retryDelayMs = 50;

        for (int attempt = 0; attempt < maxRetries; attempt++)
        {
            try
            {
                _logger.LogDebug($"Updating dynamic block visibility (attempt {attempt + 1}/{maxRetries})");
                
                // Give AutoCAD time to initialize dynamic block properties on first attempt
                if (attempt > 0)
                {
                    System.Threading.Thread.Sleep(retryDelayMs);
                }

                DynamicBlockReferencePropertyCollection props;
                try
                {
                    props = blockRef.DynamicBlockReferencePropertyCollection;
                    if (props == null)
                    {
                        _logger.LogWarning($"Dynamic block properties not available (attempt {attempt + 1})");
                        continue;
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
                {
                    _logger.LogWarning($"AutoCAD exception accessing dynamic properties (attempt {attempt + 1}): {acadEx.Message}");
                    continue;
                }
                
                foreach (DynamicBlockReferenceProperty prop in props)
                {
                    if (prop.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase) ||
                        prop.PropertyName.Equals("Visibility1", StringComparison.OrdinalIgnoreCase))
                    {
                        string targetState = makeVisible ? "ON" : "OFF";
                        
                        // Check if this visibility state exists in the allowed values
                        object[] allowedValues;
                        try
                        {
                            allowedValues = prop.GetAllowedValues();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
                        {
                            _logger.LogWarning($"AutoCAD exception getting allowed values (attempt {attempt + 1}): {acadEx.Message}");
                            continue;
                        }

                        bool stateExists = false;
                        
                        foreach (object allowedValue in allowedValues)
                        {
                            string allowedState = allowedValue.ToString();
                            if (allowedState.Equals(targetState, StringComparison.OrdinalIgnoreCase))
                            {
                                stateExists = true;
                                break;
                            }
                            // Also check for alternative visibility states
                            if (makeVisible && (allowedState.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
                                              allowedState.Equals("Hex", StringComparison.OrdinalIgnoreCase)))
                            {
                                targetState = allowedState;
                                stateExists = true;
                                break;
                            }
                        }

                        if (stateExists)
                        {
                            try
                            {
                                prop.Value = targetState;
                                _logger.LogDebug($"Successfully set visibility to: {targetState}");
                                return true;
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
                            {
                                _logger.LogWarning($"AutoCAD exception setting visibility value (attempt {attempt + 1}): {acadEx.Message}");
                                continue;
                            }
                        }
                        else
                        {
                            _logger.LogWarning($"Visibility state '{targetState}' not found in allowed values: [{string.Join(", ", allowedValues)}]");
                            return false;
                        }
                    }
                }
                
                _logger.LogWarning("No visibility property found in dynamic block");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Exception updating dynamic block visibility (attempt {attempt + 1}): {ex.Message}");
                if (attempt == maxRetries - 1)
                {
                    _logger.LogError($"Failed to update dynamic block visibility after {maxRetries} attempts: {ex.Message}", ex);
                    return false;
                }
            }
        }

        _logger.LogError($"Failed to update dynamic block visibility after {maxRetries} attempts");
        return false;
    }

    /// <summary>
    /// Helper method to log available layouts for debugging
    /// </summary>
    private void LogAvailableLayouts(DBDictionary layoutDict, Transaction tr)
    {
        try
        {
            _logger.LogInformation("Available layouts in drawing:");
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                string layoutName = entry.Key;
                if (!layoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                        string tabName = layout?.LayoutName ?? "Unknown";
                        _logger.LogInformation($"  - Dictionary Key: '{layoutName}', Tab Name: '{tabName}'");
                    }
                    catch
                    {
                        _logger.LogInformation($"  - Dictionary Key: '{layoutName}' (could not read tab name)");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Could not enumerate available layouts: {ex.Message}", ex);
        }
    }
}