using Autodesk.AutoCAD.DatabaseServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.IO;
using System.Text.RegularExpressions;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Manages operations on closed drawings (not currently open in AutoCAD)
/// Uses Database.ReadDwgFile() to work with external drawings
/// Applies BlockTableRecordExtensions.SynchronizeAttributes() for proper attribute positioning
/// </summary>
public class ExternalDrawingManager
{
    private readonly ILogger _logger;
    private readonly Regex _noteBlockPattern = new Regex(@"^NT\d{2}$", RegexOptions.Compiled);

    public ExternalDrawingManager(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Updates construction note blocks in a closed drawing
    /// </summary>
    /// <param name="dwgPath">Full path to the DWG file</param>
    /// <param name="layoutName">Layout name containing the blocks</param>
    /// <param name="noteData">Construction note data to apply</param>
    /// <returns>True if successful, false otherwise</returns>
    public bool UpdateClosedDrawing(string dwgPath, string layoutName, List<ConstructionNoteData> noteData)
    {
        _logger.LogInformation($"=== ExternalDrawingManager.UpdateClosedDrawing ===");
        _logger.LogInformation($"DWG Path: {dwgPath}");
        _logger.LogInformation($"Layout: {layoutName}");
        _logger.LogInformation($"Note Count: {noteData.Count}");

        // Validate inputs
        if (string.IsNullOrEmpty(dwgPath) || !File.Exists(dwgPath))
        {
            _logger.LogError($"Drawing file not found: {dwgPath}");
            return false;
        }

        if (string.IsNullOrEmpty(layoutName))
        {
            _logger.LogError("Layout name cannot be empty");
            return false;
        }

        if (noteData == null)
        {
            noteData = new List<ConstructionNoteData>();
        }

        try
        {
            // Create external database and load the drawing
            _logger.LogDebug("Creating external database...");
            using (var db = new Database(false, true)) // buildDefaultDrawing=false, noDocument=true
            {
                _logger.LogDebug($"Reading DWG file: {Path.GetFileName(dwgPath)}");
                db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                _logger.LogDebug("DWG file loaded successfully");

                // Start transaction for all operations
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Update blocks in the external database
                        bool updateSuccess = UpdateBlocksInExternalDatabase(db, tr, layoutName, noteData);
                        
                        if (updateSuccess)
                        {
                            // Apply ATTSYNC to all NT## blocks for proper attribute positioning
                            ApplyAttributeSynchronization(db, tr);
                            
                            // Commit transaction
                            tr.Commit();
                            _logger.LogDebug("Transaction committed successfully");
                        }
                        else
                        {
                            _logger.LogError("Block update failed, rolling back transaction");
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during block updates: {ex.Message}", ex);
                        return false;
                    }
                }

                // Save the drawing back to disk
                _logger.LogDebug("Saving drawing to disk...");
                db.SaveAs(dwgPath, DwgVersion.Current);
                _logger.LogInformation($"Successfully updated closed drawing: {Path.GetFileName(dwgPath)}");
                
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update closed drawing: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates construction note blocks within an external database
    /// </summary>
    private bool UpdateBlocksInExternalDatabase(Database db, Transaction tr, string layoutName, List<ConstructionNoteData> noteData)
    {
        try
        {
            _logger.LogDebug($"Updating blocks in layout: {layoutName}");
            
            // Get layout dictionary
            var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            if (!layoutDict.Contains(layoutName))
            {
                _logger.LogError($"Layout '{layoutName}' not found in drawing");
                LogAvailableLayouts(layoutDict, tr);
                return false;
            }

            // Get layout and block table record
            var layoutId = layoutDict.GetAt(layoutName);
            var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
            var layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

            _logger.LogDebug($"Found layout '{layoutName}', scanning for NT## blocks...");

            // First pass: Clear all existing NT## blocks (reset them to empty state)
            int clearedCount = ClearConstructionNoteBlocks(layoutBtr, tr);
            _logger.LogDebug($"Cleared {clearedCount} existing NT## blocks");

            // Second pass: Update blocks with new note data
            int updatedCount = 0;
            foreach (var noteEntry in noteData)
            {
                string targetBlockName = $"NT{noteEntry.NoteNumber:D2}"; // Format as NT01, NT02, etc.
                
                if (UpdateSingleBlockInLayout(layoutBtr, tr, targetBlockName, noteEntry))
                {
                    updatedCount++;
                    _logger.LogDebug($"Updated block {targetBlockName} with note {noteEntry.NoteNumber}");
                }
                else
                {
                    _logger.LogWarning($"Failed to update block {targetBlockName}");
                }
            }

            _logger.LogInformation($"Updated {updatedCount} out of {noteData.Count} construction note blocks");
            return updatedCount > 0 || noteData.Count == 0; // Success if we updated blocks or there were no blocks to update
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error updating blocks in external database: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Clears all NT## blocks in a layout (sets visibility OFF and clears attributes)
    /// </summary>
    private int ClearConstructionNoteBlocks(BlockTableRecord layoutBtr, Transaction tr)
    {
        int clearedCount = 0;
        
        try
        {
            foreach (ObjectId objId in layoutBtr)
            {
                var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    
                    if (_noteBlockPattern.IsMatch(blockName))
                    {
                        // Open block for write
                        var writeBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                        
                        // Clear attributes
                        ClearBlockAttributes(writeBlockRef, tr);
                        
                        // Set visibility to OFF
                        if (writeBlockRef.IsDynamicBlock)
                        {
                            SetDynamicBlockVisibility(writeBlockRef, false);
                        }
                        
                        // Notify AutoCAD that block graphics have been modified
                        writeBlockRef.RecordGraphicsModified(true);
                        
                        clearedCount++;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error clearing construction note blocks: {ex.Message}", ex);
        }
        
        return clearedCount;
    }

    /// <summary>
    /// Updates a single construction note block in the layout
    /// </summary>
    private bool UpdateSingleBlockInLayout(BlockTableRecord layoutBtr, Transaction tr, string targetBlockName, ConstructionNoteData noteData)
    {
        try
        {
            // Find the target block
            foreach (ObjectId objId in layoutBtr)
            {
                var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    
                    if (blockName.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Open block for write
                        var writeBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                        
                        // Update attributes
                        UpdateBlockAttributes(writeBlockRef, tr, noteData.NoteNumber, noteData.NoteText);
                        
                        // Set visibility to ON
                        if (writeBlockRef.IsDynamicBlock)
                        {
                            SetDynamicBlockVisibility(writeBlockRef, true);
                        }
                        
                        // Notify AutoCAD that block graphics have been modified
                        writeBlockRef.RecordGraphicsModified(true);
                        
                        return true;
                    }
                }
            }
            
            _logger.LogWarning($"Block '{targetBlockName}' not found in layout");
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error updating single block {targetBlockName}: {ex.Message}", ex);
            return false;
        }
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
                var dynamicBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return dynamicBtr.Name;
            }
            else
            {
                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
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
    /// Clears all attributes in a construction note block with proper alignment
    /// </summary>
    private void ClearBlockAttributes(BlockReference blockRef, Transaction tr)
    {
        try
        {
            var attCol = blockRef.AttributeCollection;
            foreach (ObjectId attId in attCol)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null)
                {
                    string tag = attRef.Tag.ToUpper();
                    if (tag == "NUMBER" || tag == "NOTE")
                    {
                        if (!string.IsNullOrEmpty(attRef.TextString))
                        {
                            attRef.TextString = "";
                            attRef.AdjustAlignment(attRef.Database);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error clearing block attributes: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Updates attributes in a construction note block with proper alignment
    /// Uses BlockProcessor approach: set Justify and call AdjustAlignment for proper positioning
    /// </summary>
    private void UpdateBlockAttributes(BlockReference blockRef, Transaction tr, int noteNumber, string noteText)
    {
        try
        {
            var attCol = blockRef.AttributeCollection;
            foreach (ObjectId attId in attCol)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null)
                {
                    string tag = attRef.Tag.ToUpper();
                    
                    if (tag == "NUMBER")
                    {
                        string newValue = noteNumber > 0 ? noteNumber.ToString() : "";
                        if (attRef.TextString != newValue)
                        {
                            attRef.Justify = AttachmentPoint.MiddleCenter;
                            attRef.TextString = newValue;
                            attRef.AdjustAlignment(attRef.Database);
                        }
                    }
                    else if (tag == "NOTE")
                    {
                        string newValue = noteText ?? "";
                        if (attRef.TextString != newValue)
                        {
                            attRef.TextString = newValue;
                            attRef.AdjustAlignment(attRef.Database);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error updating block attributes: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Sets the visibility state of a dynamic block
    /// </summary>
    private void SetDynamicBlockVisibility(BlockReference blockRef, bool makeVisible)
    {
        try
        {
            var props = blockRef.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty prop in props)
            {
                if (prop.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase) ||
                    prop.PropertyName.Equals("Visibility1", StringComparison.OrdinalIgnoreCase))
                {
                    string targetState = makeVisible ? "ON" : "OFF";
                    
                    // Check if this visibility state exists in allowed values
                    var allowedValues = prop.GetAllowedValues();
                    foreach (object allowedValue in allowedValues)
                    {
                        string allowedState = allowedValue.ToString();
                        if (allowedState.Equals(targetState, StringComparison.OrdinalIgnoreCase))
                        {
                            prop.Value = targetState;
                            return;
                        }
                        // Also check for alternative visibility states
                        if (makeVisible && (allowedState.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
                                          allowedState.Equals("Hex", StringComparison.OrdinalIgnoreCase)))
                        {
                            prop.Value = allowedState;
                            return;
                        }
                    }
                    
                    _logger.LogWarning($"Visibility state '{targetState}' not found in allowed values: [{string.Join(", ", allowedValues)}]");
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error setting dynamic block visibility: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Applies BlockTableRecordExtensions.SynchronizeAttributes() to all NT## blocks
    /// This mimics AutoCAD's ATTSYNC command for proper attribute positioning
    /// </summary>
    private void ApplyAttributeSynchronization(Database db, Transaction tr)
    {
        try
        {
            _logger.LogDebug("Applying attribute synchronization (ATTSYNC) to NT## blocks...");
            
            var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            int syncCount = 0;
            
            foreach (ObjectId btrId in blockTable)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && _noteBlockPattern.IsMatch(btr.Name))
                {
                    // Apply SynchronizeAttributes from BlockTableRecordExtensions
                    btr.SynchronizeAttributes();
                    syncCount++;
                    _logger.LogDebug($"Applied ATTSYNC to block definition: {btr.Name}");
                }
            }
            
            _logger.LogInformation($"Applied attribute synchronization to {syncCount} NT## block definitions");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error applying attribute synchronization: {ex.Message}", ex);
        }
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
                        var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
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

/// <summary>
/// Data structure for passing construction note information
/// </summary>
public class ConstructionNoteData
{
    public int NoteNumber { get; set; }
    public string NoteText { get; set; } = string.Empty;
    
    public ConstructionNoteData(int noteNumber, string noteText)
    {
        NoteNumber = noteNumber;
        NoteText = noteText ?? string.Empty;
    }
}