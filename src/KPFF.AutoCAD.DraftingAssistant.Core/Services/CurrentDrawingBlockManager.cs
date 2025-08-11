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

    public CurrentDrawingBlockManager(ILogger logger)
    {
        _logger = logger;
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
}