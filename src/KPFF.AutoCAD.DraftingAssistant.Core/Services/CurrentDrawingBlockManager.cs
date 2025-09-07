using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AcadException = Autodesk.AutoCAD.Runtime.Exception;
using SystemException = System.Exception;
using Autodesk.AutoCAD.Runtime;
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
    private Regex _noteBlockPattern = new Regex(@"^NT\d{2}$", RegexOptions.Compiled);
    private bool _isInitialized = false;
    private readonly Database? _externalDatabase;
    private readonly bool _useExternalDatabase;

    /// <summary>
    /// Constructor for current drawing operations (original behavior)
    /// </summary>
    public CurrentDrawingBlockManager(ILogger logger)
    {
        _logger = logger;
        _externalDatabase = null;
        _useExternalDatabase = false;
    }

    /// <summary>
    /// Constructor for current drawing operations with custom note block pattern
    /// </summary>
    public CurrentDrawingBlockManager(ILogger logger, string? noteBlockPattern = null)
    {
        _logger = logger;
        _externalDatabase = null;
        _useExternalDatabase = false;
        if (!string.IsNullOrEmpty(noteBlockPattern))
        {
            SetNoteBlockPattern(noteBlockPattern);
        }
    }

    /// <summary>
    /// Constructor for external database operations (new functionality)
    /// </summary>
    public CurrentDrawingBlockManager(Database externalDatabase, ILogger logger, string? noteBlockPattern = null)
    {
        _logger = logger;
        _externalDatabase = externalDatabase;
        _useExternalDatabase = true;
        _isInitialized = true; // External database is already initialized
        if (!string.IsNullOrEmpty(noteBlockPattern))
        {
            SetNoteBlockPattern(noteBlockPattern);
        }
    }

    /// <summary>
    /// Sets the note block pattern for identifying construction note blocks
    /// </summary>
    public void SetNoteBlockPattern(string pattern)
    {
        try
        {
            _noteBlockPattern = new Regex(pattern, RegexOptions.Compiled);
            _logger.LogDebug($"Note block pattern updated to: {pattern}");
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning($"Invalid note block pattern '{pattern}': {ex.Message}. Using default pattern.");
            _noteBlockPattern = new Regex(@"^NT\d{2}$", RegexOptions.Compiled);
        }
    }

    /// <summary>
    /// Gets the appropriate database for operations (current drawing or external)
    /// </summary>
    private (Database? database, Document? document) GetDatabaseAndDocument()
    {
        if (_useExternalDatabase)
        {
            return (_externalDatabase, null);
        }
        else
        {
            try
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager?.MdiActiveDocument;
                return (doc?.Database, doc);
            }
            catch (SystemException ex)
            {
                _logger.LogError($"Failed to access current drawing: {ex.Message}", ex);
                return (null, null);
            }
        }
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
            _logger.LogDebug($"Initializing CurrentDrawingBlockManager - {(_useExternalDatabase ? "external database" : "current drawing")} mode");
            
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible during initialization ({(_useExternalDatabase ? "external" : "current")})");
                return false;
            }

            // Test basic transaction access
            using (var testTrans = db.TransactionManager.StartTransaction())
            {
                testTrans.Commit();
            }

            _isInitialized = true;
            _logger.LogDebug($"CurrentDrawingBlockManager initialized successfully ({(_useExternalDatabase ? "external" : "current")})");
            return true;
        }
        catch (SystemException ex)
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
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for block operations ({(_useExternalDatabase ? "external" : "current")})");
                return blocks;
            }

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
                catch (SystemException ex)
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
        catch (SystemException ex)
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
        catch (SystemException ex)
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
                string numberText = attributes["NUMBER"];
                if (!string.IsNullOrEmpty(numberText) && int.TryParse(numberText, out int noteNumber))
                {
                    noteBlock.Number = noteNumber;
                }
                else
                {
                    noteBlock.Number = 0; // Keep 0 only for truly empty, not default
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
        catch (SystemException ex)
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
                using (AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference)
                {
                    if (attRef != null)
                    {
                        string tag = attRef.Tag.ToUpper();
                        string value = attRef.TextString;
                        attributes[tag] = value;
                        _logger.LogDebug($"  Attribute: {tag} = '{value}'");
                    }
                }
            }
        }
        catch (SystemException ex)
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
        catch (SystemException ex)
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
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for layout scanning ({(_useExternalDatabase ? "external" : "current")})");
                return allBlocks;
            }

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
                catch (SystemException ex)
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
        catch (SystemException ex)
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
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for clearing blocks ({(_useExternalDatabase ? "external" : "current")})");
                return 0;
            }

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
                                    using (AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference)
                                    {
                                        if (attRef != null)
                                        {
                                            string tag = attRef.Tag.ToUpper();
                                            if (tag == "NUMBER" || tag == "NOTE")
                                            {
                                                attRef.TextString = "";
                                            }
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
                                            catch (SystemException ex)
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
                catch (SystemException ex)
                {
                    _logger.LogError($"Error clearing blocks: {ex.Message}", ex);
                }
            }
        }
        catch (SystemException ex)
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
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for block update ({(_useExternalDatabase ? "external" : "current")})");
                return false;
            }

            _logger.LogDebug($"Database filename: {db.Filename ?? "external"}");

            // For current drawing, lock document. For external database, no locking needed
            IDisposable? documentLock = null;
            if (doc != null && !_useExternalDatabase)
            {
                documentLock = doc.LockDocument();
            }

            try
            {
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
                    catch (SystemException ex)
                    {
                        _logger.LogError($"Error updating block: {ex.Message}", ex);
                        return false;
                    }
                }
            }
            finally
            {
                // Dispose document lock if it was created
                documentLock?.Dispose();
            }
        }
        catch (SystemException ex)
        {
            _logger.LogError($"Failed to update construction note block: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates the attributes of a construction note block with proper alignment
    /// Uses the working archive approach with Justify and AdjustAlignment
    /// </summary>
    private bool UpdateBlockAttributes(BlockReference blockRef, int noteNumber, string noteText, Transaction tr)
    {
        try
        {
            AttributeCollection attCol = blockRef.AttributeCollection;
            bool numberUpdated = false;
            bool noteUpdated = false;
            bool wasModified = false;
            
            foreach (ObjectId attId in attCol)
            {
                using (AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference)
                {
                    if (attRef != null)
                {
                    string tag = attRef.Tag.ToUpper();
                    
                    if (tag == "NUMBER")
                    {
                        string newValue = noteNumber > 0 ? noteNumber.ToString() : "";
                        string currentValue = attRef.TextString ?? "";
                        if (currentValue != newValue)
                        {
                            attRef.Justify = AttachmentPoint.MiddleCenter;
                            attRef.TextString = newValue;
                            
                            // CRITICAL: For external databases, switch WorkingDatabase for alignment
                            if (_useExternalDatabase)
                            {
                                var originalWdb = HostApplicationServices.WorkingDatabase;
                                try
                                {
                                    HostApplicationServices.WorkingDatabase = attRef.Database;
                                    attRef.AdjustAlignment(attRef.Database);
                                }
                                finally
                                {
                                    HostApplicationServices.WorkingDatabase = originalWdb;
                                }
                            }
                            else
                            {
                                attRef.AdjustAlignment(attRef.Database);
                            }
                            
                            wasModified = true;
                            _logger.LogDebug($"Updated NUMBER attribute: '{newValue}' with alignment ({(_useExternalDatabase ? "external" : "current")})");
                        }
                        numberUpdated = true;
                    }
                    else if (tag == "NOTE")
                    {
                        string newValue = noteText ?? "";
                        string currentValue = attRef.TextString ?? "";
                        if (currentValue != newValue)
                        {
                            attRef.TextString = newValue;
                            
                            // CRITICAL: For external databases, switch WorkingDatabase for alignment
                            if (_useExternalDatabase)
                            {
                                var originalWdb = HostApplicationServices.WorkingDatabase;
                                try
                                {
                                    HostApplicationServices.WorkingDatabase = attRef.Database;
                                    attRef.AdjustAlignment(attRef.Database);
                                }
                                finally
                                {
                                    HostApplicationServices.WorkingDatabase = originalWdb;
                                }
                            }
                            else
                            {
                                attRef.AdjustAlignment(attRef.Database);
                            }
                            
                            wasModified = true;
                            _logger.LogDebug($"Updated NOTE attribute with alignment ({(_useExternalDatabase ? "external" : "current")})");
                        }
                        noteUpdated = true;
                    }
                }
                }
            }

            // Record graphics modification after attribute updates
            if (wasModified)
            {
                blockRef.RecordGraphicsModified(true);
                _logger.LogDebug("Applied RecordGraphicsModified to block reference");
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
        catch (SystemException ex)
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
            catch (SystemException ex)
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
        catch (SystemException ex)
        {
            _logger.LogError($"Could not enumerate available layouts: {ex.Message}", ex);
        }
    }

    #region Title Block Operations

    /// <summary>
    /// Gets title blocks matching the specified pattern in a layout
    /// </summary>
    /// <param name="layoutName">Layout name to search in</param>
    /// <param name="titleBlockPattern">Regex pattern to match title block names (e.g., "^PROJ-TITLE-BLK$")</param>
    /// <returns>List of title block information</returns>
    public List<TitleBlockInfo> GetTitleBlocks(string layoutName, string titleBlockPattern)
    {
        var titleBlocks = new List<TitleBlockInfo>();
        
        if (!EnsureInitialized())
        {
            _logger.LogError("Failed to initialize AutoCAD context for title block operations");
            return titleBlocks;
        }

        try
        {
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for title block operations ({(_useExternalDatabase ? "external" : "current")})");
                return titleBlocks;
            }

            var titleBlockRegex = new Regex(titleBlockPattern, RegexOptions.Compiled);
            
            // For current drawing, lock document. For external database, no locking needed
            IDisposable? documentLock = null;
            if (doc != null && !_useExternalDatabase)
            {
                documentLock = doc.LockDocument();
            }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the layout dictionary
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    
                    if (!layoutDict.Contains(layoutName))
                    {
                        _logger.LogWarning($"Layout '{layoutName}' not found in drawing");
                        return titleBlocks;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId objId in layoutBtr)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (entity is BlockReference blockRef)
                        {
                            string blockName = GetEffectiveBlockName(blockRef, tr);
                            
                            if (titleBlockRegex.IsMatch(blockName))
                            {
                                var attributes = new Dictionary<string, string>();
                                
                                // Read all attributes from the title block
                                AttributeCollection attCol = blockRef.AttributeCollection;
                                foreach (ObjectId attId in attCol)
                                {
                                    using (AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference)
                                    {
                                        if (attRef != null)
                                        {
                                            attributes[attRef.Tag] = attRef.TextString ?? "";
                                        }
                                    }
                                }

                                var titleBlockInfo = new TitleBlockInfo
                                {
                                    SheetName = layoutName,
                                    Title = blockName,
                                    Attributes = attributes,
                                    IsActive = true,
                                    LastUpdated = DateTime.Now
                                };

                                titleBlocks.Add(titleBlockInfo);
                                _logger.LogDebug($"Found title block '{blockName}' with {attributes.Count} attributes");
                            }
                        }
                    }
                    
                    tr.Commit();
                }
            }
            finally
            {
                documentLock?.Dispose();
            }
        }
        catch (SystemException ex)
        {
            _logger.LogError($"Failed to get title blocks for layout {layoutName}: {ex.Message}", ex);
        }

        _logger.LogInformation($"Found {titleBlocks.Count} title blocks in layout {layoutName}");
        return titleBlocks;
    }

    /// <summary>
    /// Updates a single title block attribute
    /// </summary>
    /// <param name="layoutName">Layout containing the title block</param>
    /// <param name="blockName">Title block name</param>
    /// <param name="attributeName">Attribute tag name to update</param>
    /// <param name="attributeValue">New attribute value</param>
    /// <returns>True if update was successful</returns>
    public bool UpdateTitleBlockAttribute(string layoutName, string blockName, string attributeName, string attributeValue)
    {
        if (!EnsureInitialized())
        {
            _logger.LogError("Failed to initialize AutoCAD context for title block attribute update");
            return false;
        }

        try
        {
            var (db, doc) = GetDatabaseAndDocument();
            if (db == null)
            {
                _logger.LogError($"Database not accessible for title block update ({(_useExternalDatabase ? "external" : "current")})");
                return false;
            }

            // For current drawing, lock document. For external database, no locking needed
            IDisposable? documentLock = null;
            if (doc != null && !_useExternalDatabase)
            {
                documentLock = doc.LockDocument();
            }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the layout dictionary
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    
                    if (!layoutDict.Contains(layoutName))
                    {
                        _logger.LogDebug($"Layout '{layoutName}' not found");
                        return false;
                    }

                    ObjectId layoutId = layoutDict.GetAt(layoutName);
                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    BlockTableRecord layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    // Find the target title block
                    BlockReference targetBlockRef = null;
                    foreach (ObjectId objId in layoutBtr)
                    {
                        Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        
                        if (entity is BlockReference blockRef)
                        {
                            string currentBlockName = GetEffectiveBlockName(blockRef, tr);
                            
                            if (currentBlockName.Equals(blockName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                                break;
                            }
                        }
                    }

                    if (targetBlockRef == null)
                    {
                        _logger.LogDebug($"Title block '{blockName}' not found in layout '{layoutName}'");
                        return false;
                    }

                    // Update the attribute
                    bool attributeUpdated = false;
                    AttributeCollection attCol = targetBlockRef.AttributeCollection;
                    
                    foreach (ObjectId attId in attCol)
                    {
                        using (AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference)
                        {
                            if (attRef != null && attRef.Tag.Equals(attributeName, StringComparison.OrdinalIgnoreCase))
                            {
                                string currentValue = attRef.TextString ?? "";
                                if (currentValue != attributeValue)
                                {
                                    attRef.TextString = attributeValue ?? "";
                                    
                                    // Handle alignment for external databases
                                    if (_useExternalDatabase)
                                    {
                                        var originalWdb = HostApplicationServices.WorkingDatabase;
                                        try
                                        {
                                            HostApplicationServices.WorkingDatabase = attRef.Database;
                                            attRef.AdjustAlignment(attRef.Database);
                                        }
                                        finally
                                        {
                                            HostApplicationServices.WorkingDatabase = originalWdb;
                                        }
                                    }
                                    else
                                    {
                                        attRef.AdjustAlignment(attRef.Database);
                                    }
                                    
                                    targetBlockRef.RecordGraphicsModified(true);
                                    attributeUpdated = true;
                                    _logger.LogDebug($"Updated attribute '{attributeName}' = '{attributeValue}'");
                                }
                                break;
                            }
                        }
                    }

                    if (attributeUpdated)
                    {
                        tr.Commit();
                        return true;
                    }
                    
                    return false;
                }
            }
            finally
            {
                documentLock?.Dispose();
            }
        }
        catch (SystemException ex)
        {
            _logger.LogError($"Failed to update title block attribute: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets all attributes from title blocks matching the pattern in a layout
    /// </summary>
    /// <param name="layoutName">Layout name to search in</param>
    /// <param name="titleBlockPattern">Regex pattern to match title block names</param>
    /// <returns>Dictionary of attribute names and values</returns>
    public Dictionary<string, string> GetTitleBlockAttributes(string layoutName, string titleBlockPattern)
    {
        var attributes = new Dictionary<string, string>();
        
        if (!EnsureInitialized())
        {
            _logger.LogError("Failed to initialize AutoCAD context for getting title block attributes");
            return attributes;
        }

        try
        {
            var titleBlocks = GetTitleBlocks(layoutName, titleBlockPattern);
            
            if (titleBlocks.Count > 0)
            {
                // Return attributes from the first title block found
                attributes = titleBlocks[0].Attributes;
                _logger.LogDebug($"Retrieved {attributes.Count} attributes from title block in layout {layoutName}");
            }
            else
            {
                _logger.LogWarning($"No title blocks found matching pattern '{titleBlockPattern}' in layout {layoutName}");
            }
        }
        catch (SystemException ex)
        {
            _logger.LogError($"Failed to get title block attributes for layout {layoutName}: {ex.Message}", ex);
        }

        return attributes;
    }

    /// <summary>
    /// Optimized batch update method that discovers and updates construction note blocks in a single operation.
    /// Eliminates transaction conflicts and memory leaks by doing everything in one transaction.
    /// </summary>
    /// <param name="layoutName">Name of the layout containing the blocks</param>
    /// <param name="blockUpdates">Dictionary of block name to update parameters</param>
    /// <returns>Tuple of (success, discovered block names) for validation</returns>
    public (bool success, List<string> discoveredBlocks) UpdateConstructionNoteBlocksBatchWithDiscovery(string layoutName, Dictionary<string, (int noteNumber, string noteText, bool makeVisible)> blockUpdates)
    {
        if (blockUpdates == null || blockUpdates.Count == 0)
        {
            _logger.LogWarning("No block updates provided for batch operation");
            return (true, new List<string>());
        }

        _logger.LogInformation($"=== BATCH UpdateConstructionNoteBlocks ENTRY ====");
        _logger.LogInformation($"Layout: {layoutName}");
        _logger.LogInformation($"Blocks to update: {blockUpdates.Count} ({string.Join(", ", blockUpdates.Keys)})");

        var document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var database = document.Database;

        _logger.LogDebug($"Database filename: {database.Filename}");

        int successCount = 0;
        int totalBlocks = blockUpdates.Count;

        // Lock document before starting transaction (like other methods in this file do)
        using (var documentLock = document.LockDocument())
        using (var transaction = database.TransactionManager.StartTransaction())
        {
            try
            {
                _logger.LogDebug("Starting single transaction for batch block update...");

                // STEP 1: Get layout and validate it exists
                var layoutDict = transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                
                if (!layoutDict.Contains(layoutName))
                {
                    _logger.LogError($"Layout '{layoutName}' not found in drawing");
                    return (false, new List<string>());
                }

                var layoutId = layoutDict.GetAt(layoutName);
                var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                var layoutBtr = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                _logger.LogDebug($"Layout BTR contains {layoutBtr.Cast<ObjectId>().Count()} entities");

                // STEP 2: Find ALL target blocks in ONE scan and cache their ObjectIds
                _logger.LogDebug($"Scanning layout once to find all target blocks...");
                var blockCache = new Dictionary<string, ObjectId>();
                int blocksScanned = 0;

                foreach (ObjectId objId in layoutBtr)
                {
                    blocksScanned++;
                    var entity = transaction.GetObject(objId, OpenMode.ForRead) as Entity;
                    
                    if (entity is BlockReference blockRef)
                    {
                        string currentBlockName = GetEffectiveBlockName(blockRef, transaction);
                        
                        // Check if this is one of our target blocks
                        if (blockUpdates.ContainsKey(currentBlockName))
                        {
                            blockCache[currentBlockName] = objId;
                            _logger.LogDebug($"Cached block '{currentBlockName}' with ObjectId {objId.Handle.Value}");
                        }
                    }
                }

                _logger.LogInformation($"Single scan completed: found {blockCache.Count} target blocks out of {blocksScanned} entities scanned");

                // STEP 3: Update all cached blocks in batch
                foreach (var (blockName, updateParams) in blockUpdates)
                {
                    if (!blockCache.TryGetValue(blockName, out var blockId))
                    {
                        _logger.LogWarning($"Block '{blockName}' not found in layout - skipping");
                        continue;
                    }

                    try
                    {
                        _logger.LogDebug($"Batch updating block '{blockName}' (note={updateParams.noteNumber}, visible={updateParams.makeVisible})...");
                        
                        try
                        {
                            // Open block directly for write - no upgrade needed
                            using (var blockRef = transaction.GetObject(blockId, OpenMode.ForWrite) as BlockReference)
                            {
                                if (blockRef == null)
                                {
                                    _logger.LogWarning($"Could not open block '{blockName}' for write");
                                    continue;
                                }

                                // Update attributes
                                bool attributesUpdated = UpdateBlockAttributesBatch(blockRef, updateParams.noteNumber, updateParams.noteText, transaction);
                                
                                // Update visibility
                                bool visibilityUpdated = UpdateBlockVisibilityBatch(blockRef, updateParams.makeVisible, transaction);

                                if (attributesUpdated && visibilityUpdated)
                                {
                                    successCount++;
                                    _logger.LogDebug($"✓ Successfully batch updated block '{blockName}'");
                                }
                                else
                                {
                                    _logger.LogWarning($"✗ Partial failure updating block '{blockName}' (attributes={attributesUpdated}, visibility={visibilityUpdated})");
                                }
                            }
                        }
                        catch (AcadException acEx) when (acEx.Message.Contains("eLockViolation") || acEx.Message.Contains("lock violation"))
                        {
                            _logger.LogError($"Lock violation opening block '{blockName}' for write - may be locked by another operation");
                            // Continue with other blocks instead of failing entire operation
                        }
                    }
                    catch (SystemException blockEx)
                    {
                        _logger.LogError($"Error updating block '{blockName}': {blockEx.Message}");
                    }
                }

                // STEP 4: Commit all changes at once
                _logger.LogDebug("Committing batch transaction...");
                transaction.Commit();
                
                _logger.LogInformation($"=== BATCH UPDATE COMPLETE ===");
                _logger.LogInformation($"Successfully updated {successCount} of {totalBlocks} blocks in layout '{layoutName}'");
                _logger.LogInformation($"Performance: {blocksScanned} entities scanned ONCE vs {totalBlocks} individual searches avoided");

                return (successCount == totalBlocks, blockCache.Keys.ToList());
            }
            catch (SystemException ex)
            {
                _logger.LogError($"Exception during batch block update: {ex.Message}", ex);
                transaction.Abort();
                return (false, new List<string>());
            }
        }
    }


    /// <summary>
    /// Updates block attributes in batch mode (helper method)
    /// </summary>
    private bool UpdateBlockAttributesBatch(BlockReference blockRef, int noteNumber, string noteText, Transaction transaction)
    {
        try
        {
            bool attributesUpdated = false;

            // Handle attribute references (for regular blocks)
            foreach (ObjectId attRefId in blockRef.AttributeCollection)
            {
                using (var attRef = transaction.GetObject(attRefId, OpenMode.ForWrite) as AttributeReference)
                {
                    if (attRef == null) continue;

                    string tag = attRef.Tag?.ToUpperInvariant() ?? "";
                    
                    if (tag == "NUMBER")
                    {
                        // Use empty string for note numbers <= 0 (empty/reset blocks)
                        attRef.TextString = noteNumber <= 0 ? "" : noteNumber.ToString();
                        attributesUpdated = true;
                    }
                    else if (tag == "NOTE")
                    {
                        attRef.TextString = noteText ?? "";
                        attributesUpdated = true;
                    }
                }
            }

            return attributesUpdated;
        }
        catch (SystemException ex)
        {
            _logger.LogDebug($"Error updating attributes in batch: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Updates block visibility in batch mode (helper method)
    /// </summary>
    private bool UpdateBlockVisibilityBatch(BlockReference blockRef, bool makeVisible, Transaction transaction)
    {
        try
        {
            if (blockRef.IsDynamicBlock)
            {
                var dynamicProps = blockRef.DynamicBlockReferencePropertyCollection;
                
                foreach (DynamicBlockReferenceProperty prop in dynamicProps)
                {
                    if (string.Equals(prop.PropertyName, "Visibility", StringComparison.OrdinalIgnoreCase))
                    {
                        string newVisibilityState = makeVisible ? "ON" : "OFF";
                        
                        // Check if the visibility state is available
                        if (prop.GetAllowedValues().Cast<object>().Any(val => 
                            string.Equals(val?.ToString(), newVisibilityState, StringComparison.OrdinalIgnoreCase)))
                        {
                            prop.Value = newVisibilityState;
                            _logger.LogDebug($"Set batch visibility to: {newVisibilityState}");
                            return true;
                        }
                        else
                        {
                            _logger.LogDebug($"Visibility state '{newVisibilityState}' not available for block");
                            return false;
                        }
                    }
                }
            }

            // For non-dynamic blocks or if no visibility property found, use layer visibility
            // (This could be extended based on your needs)
            return true;
        }
        catch (SystemException ex)
        {
            _logger.LogDebug($"Error updating visibility in batch: {ex.Message}");
            return false;
        }
    }

    #endregion
}