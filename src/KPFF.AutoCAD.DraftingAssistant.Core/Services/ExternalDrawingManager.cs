using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System.IO;
using System.Linq;
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
    private readonly BackupCleanupService _backupCleanupService;
    private Regex _noteBlockPattern = new Regex(@"^NT\d{2}$", RegexOptions.Compiled);
    private Regex _titleBlockPattern = new Regex(@".*", RegexOptions.Compiled); // Will be set by configuration

    public ExternalDrawingManager(ILogger logger, BackupCleanupService backupCleanupService, string? noteBlockPattern = null)
    {
        _logger = logger;
        _backupCleanupService = backupCleanupService;
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
    /// Sets the title block pattern for identifying title blocks
    /// </summary>
    public void SetTitleBlockPattern(string pattern)
    {
        try
        {
            _titleBlockPattern = new Regex(pattern, RegexOptions.Compiled);
            _logger.LogDebug($"Title block pattern updated to: {pattern}");
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning($"Invalid title block pattern '{pattern}': {ex.Message}. Using default pattern.");
            _titleBlockPattern = new Regex(@".*", RegexOptions.Compiled); // Default pattern, should be overridden
        }
    }

    /// <summary>
    /// Updates construction note blocks in a closed drawing
    /// </summary>
    /// <param name="dwgPath">Full path to the DWG file</param>
    /// <param name="layoutName">Layout name containing the blocks</param>
    /// <param name="noteData">Construction note data to apply</param>
    /// <param name="projectDWGFilePath">Project DWG directory path for cleanup (optional)</param>
    /// <returns>True if successful, false otherwise</returns>
    public bool UpdateClosedDrawing(string dwgPath, string layoutName, List<ConstructionNoteData> noteData, string? projectDWGFilePath = null)
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
                _logger.LogDebug("DWG file loaded successfully (CloseInput called automatically by OpenForReadAndAllShare mode)");

                // Start transaction for all operations - NO WorkingDatabase switching
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Update blocks in the external database
                        bool updateSuccess = UpdateBlocksInExternalDatabase(db, tr, layoutName, noteData);
                        
                        if (updateSuccess)
                        {
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

                // CRITICAL: Save to temporary file first to prevent corruption
                string tempPath = dwgPath + ".tmp";
                _logger.LogDebug($"Saving to temporary file: {Path.GetFileName(tempPath)}");
                db.SaveAs(tempPath, DwgVersion.Current);
                _logger.LogDebug("Successfully saved to temporary file");
                
                // Replace original file with updated version
                if (File.Exists(dwgPath))
                {
                    string backupPath = dwgPath + ".bak.beforeupdate";
                    _logger.LogDebug($"Creating backup: {Path.GetFileName(backupPath)}");
                    File.Copy(dwgPath, backupPath, true); // Create backup
                    File.Delete(dwgPath); // Delete original
                }
                
                File.Move(tempPath, dwgPath); // Move temp to original location
                _logger.LogInformation($"Successfully updated closed drawing: {Path.GetFileName(dwgPath)}");
                
                // Always perform backup cleanup after successful update
                if (!string.IsNullOrEmpty(projectDWGFilePath))
                {
                    try
                    {
                        int cleanedCount = _backupCleanupService.CleanupAllBackupFiles(projectDWGFilePath);
                        if (cleanedCount > 0)
                        {
                            _logger.LogInformation($"Auto-cleaned up {cleanedCount} backup files from project directory");
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        _logger.LogWarning($"Failed to cleanup backup files: {cleanupEx.Message}");
                        // Don't fail the entire operation if cleanup fails
                    }
                }
                
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
    /// Updates title blocks in a closed drawing
    /// Similar to UpdateClosedDrawing but for title block attributes
    /// </summary>
    public bool UpdateTitleBlocksInClosedDrawing(string dwgPath, string layoutName, List<TitleBlockAttributeData> attributeData, string? projectDWGFilePath = null)
    {
        _logger.LogInformation($"=== ExternalDrawingManager.UpdateTitleBlocksInClosedDrawing ===");
        _logger.LogInformation($"DWG Path: {dwgPath}");
        _logger.LogInformation($"Layout: {layoutName}");
        _logger.LogInformation($"Attribute Count: {attributeData.Count}");

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

        if (attributeData == null)
        {
            attributeData = new List<TitleBlockAttributeData>();
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
                        // Update title blocks in the external database
                        bool updateSuccess = UpdateTitleBlocksInExternalDatabase(db, tr, layoutName, attributeData);
                        
                        if (updateSuccess)
                        {
                            // Commit transaction
                            tr.Commit();
                            _logger.LogDebug("Transaction committed successfully");
                        }
                        else
                        {
                            _logger.LogError("Title block update failed, rolling back transaction");
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during title block updates: {ex.Message}", ex);
                        return false;
                    }
                }

                // Save to temporary file first to prevent corruption
                string tempPath = dwgPath + ".tmp";
                _logger.LogDebug($"Saving to temporary file: {Path.GetFileName(tempPath)}");
                db.SaveAs(tempPath, DwgVersion.Current);
                _logger.LogDebug("Successfully saved to temporary file");
                
                // Replace original file with updated version
                if (File.Exists(dwgPath))
                {
                    string backupPath = dwgPath + ".bak.beforeupdate";
                    _logger.LogDebug($"Creating backup: {Path.GetFileName(backupPath)}");
                    File.Copy(dwgPath, backupPath, true); // Create backup
                    File.Delete(dwgPath); // Delete original
                }
                
                File.Move(tempPath, dwgPath); // Move temp to original location
                _logger.LogInformation($"Successfully updated title blocks in closed drawing: {Path.GetFileName(dwgPath)}");

                // Always perform backup cleanup after successful update
                if (!string.IsNullOrEmpty(projectDWGFilePath))
                {
                    try
                    {
                        int cleanedCount = _backupCleanupService.CleanupAllBackupFiles(projectDWGFilePath);
                        if (cleanedCount > 0)
                        {
                            _logger.LogInformation($"Auto-cleaned up {cleanedCount} backup files from project directory");
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        _logger.LogWarning($"Failed to cleanup backup files: {cleanupEx.Message}");
                        // Don't fail the entire operation if cleanup fails
                    }
                }
                
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update title blocks in closed drawing: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates construction note blocks within an external database using precise targeting
    /// </summary>
    private bool UpdateBlocksInExternalDatabase(Database db, Transaction tr, string layoutName, List<ConstructionNoteData> noteData)
    {
        try
        {
            _logger.LogDebug($"Updating blocks in layout: {layoutName} using precise targeting");
            
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

            _logger.LogDebug($"Found layout '{layoutName}' (BTR: {layoutBtr.Name})");

            // CRITICAL: Get all viewport ObjectIds to explicitly exclude them
            var viewportIds = new HashSet<ObjectId>();
            var viewports = layout.GetViewports();
            if (viewports != null)
            {
                foreach (ObjectId vpId in viewports)
                {
                    viewportIds.Add(vpId);
                }
                _logger.LogDebug($"Protected {viewportIds.Count} viewports from modification");
            }

            // HYBRID APPROACH: Iterate through layout but with explicit viewport protection
            var ntBlockRefs = new List<ObjectId>();
            int totalEntities = 0;
            int skippedViewports = 0;
            int skippedOtherBlocks = 0;
            int foundNTBlocks = 0;

            _logger.LogDebug("Starting layout entity iteration with viewport protection...");

            foreach (ObjectId objId in layoutBtr)
            {
                totalEntities++;
                var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                
                // CRITICAL: Skip viewports using explicit ObjectId check
                if (viewportIds.Contains(objId))
                {
                    skippedViewports++;
                    _logger.LogDebug($"Protected viewport: ObjectId {objId}, Type: {entity?.GetType().Name}");
                    continue;
                }
                
                // Additional entity type check for viewports (double protection)
                if (entity is Viewport)
                {
                    skippedViewports++;
                    _logger.LogDebug($"Protected viewport by type: ObjectId {objId}");
                    continue;
                }
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    _logger.LogDebug($"Found block: {blockName} (ObjectId: {objId}, Type: {(blockRef.IsDynamicBlock ? "Dynamic" : "Static")})");
                    
                    // Check if this is an NT block
                    if (!string.IsNullOrEmpty(blockName) && _noteBlockPattern.IsMatch(blockName))
                    {
                        ntBlockRefs.Add(objId);
                        foundNTBlocks++;
                        _logger.LogInformation($"Added NT block: {blockName} (ObjectId: {objId})");
                    }
                    else if (!string.IsNullOrEmpty(blockName))
                    {
                        skippedOtherBlocks++;
                        _logger.LogDebug($"Skipped non-NT block: {blockName}");
                    }
                    // If blockName is empty, it was filtered out by GetEffectiveBlockName
                }
                else
                {
                    _logger.LogDebug($"Non-block entity: {entity?.GetType().Name} (ObjectId: {objId})");
                }
            }

            _logger.LogInformation($"Layout scan complete: {totalEntities} total entities, {foundNTBlocks} NT blocks, {skippedViewports} viewports protected, {skippedOtherBlocks} other blocks skipped");

            // First pass: Clear all NT blocks
            int clearedCount = 0;
            foreach (var blockRefId in ntBlockRefs)
            {
                var blockRef = tr.GetObject(blockRefId, OpenMode.ForWrite) as BlockReference;
                ClearBlockAttributes(blockRef, tr);
                
                if (blockRef.IsDynamicBlock)
                {
                    SetDynamicBlockVisibility(blockRef, false);
                }
                clearedCount++;
            }
            _logger.LogDebug($"Cleared {clearedCount} NT blocks");

            // Second pass: Update specific blocks with note data using SEQUENTIAL assignment
            int updatedCount = 0;
            for (int i = 0; i < noteData.Count && i < ntBlockRefs.Count; i++)
            {
                var noteEntry = noteData[i];
                string targetBlockName = $"NT{(i + 1):D2}"; // Sequential assignment: first note → NT01, second note → NT02, etc.
                bool foundAndUpdated = false;

                foreach (var blockRefId in ntBlockRefs)
                {
                    var blockRef = tr.GetObject(blockRefId, OpenMode.ForWrite) as BlockReference;
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    
                    if (blockName.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogDebug($"Updating NT block: {blockName} with note {noteEntry.NoteNumber} (sequential assignment: position {i + 1})");
                        
                        // Update attributes
                        UpdateBlockAttributes(blockRef, tr, db, noteEntry.NoteNumber, noteEntry.NoteText);
                        
                        // Set visibility to ON
                        if (blockRef.IsDynamicBlock)
                        {
                            SetDynamicBlockVisibility(blockRef, true);
                        }
                        
                        updatedCount++;
                        foundAndUpdated = true;
                        break; // Only update the first match
                    }
                }

                if (!foundAndUpdated)
                {
                    _logger.LogWarning($"Block '{targetBlockName}' not found for note {noteEntry.NoteNumber} at position {i + 1}");
                }
            }


            _logger.LogInformation($"Updated {updatedCount} out of {noteData.Count} construction note blocks");
            return updatedCount > 0 || noteData.Count == 0;
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
                
                // CRITICAL: Skip viewports entirely to prevent corruption
                if (entity is Viewport)
                {
                    _logger.LogDebug("Skipping viewport entity");
                    continue;
                }
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    
                    // Only process if we got a valid NT block name back
                    if (!string.IsNullOrEmpty(blockName) && _noteBlockPattern.IsMatch(blockName))
                    {
                        _logger.LogDebug($"Clearing NT block: {blockName}");
                        
                        // NOW it's safe to open for write
                        var writeBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                        
                        // Clear attributes
                        ClearBlockAttributes(writeBlockRef, tr);
                        
                        // Set visibility to OFF
                        if (writeBlockRef.IsDynamicBlock)
                        {
                            SetDynamicBlockVisibility(writeBlockRef, false);
                        }
                        
                        clearedCount++;
                    }
                    else if (!string.IsNullOrEmpty(blockName))
                    {
                        _logger.LogDebug($"Skipping non-NT block: {blockName}");
                    }
                    // If blockName is empty, it was filtered out (anonymous/viewport-related)
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
    private bool UpdateSingleBlockInLayout(BlockTableRecord layoutBtr, Transaction tr, Database db, string targetBlockName, ConstructionNoteData noteData)
    {
        try
        {
            // Find the target block
            foreach (ObjectId objId in layoutBtr)
            {
                var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                
                // CRITICAL: Skip viewports to prevent corruption
                if (entity is Viewport)
                {
                    continue;
                }
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveBlockName(blockRef, tr);
                    
                    // Only process valid NT block names
                    if (!string.IsNullOrEmpty(blockName) && 
                        blockName.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogDebug($"Updating NT block: {blockName} with note {noteData.NoteNumber}");
                        
                        // NOW it's safe to open for write
                        var writeBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                        
                        // Update attributes
                        UpdateBlockAttributes(writeBlockRef, tr, db, noteData.NoteNumber, noteData.NoteText);
                        
                        // Set visibility to ON
                        if (writeBlockRef.IsDynamicBlock)
                        {
                            SetDynamicBlockVisibility(writeBlockRef, true);
                        }
                        
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
    /// Gets the effective name of a block, handling dynamic blocks and filtering out viewport-related anonymous blocks
    /// </summary>
    private string GetEffectiveBlockName(BlockReference blockRef, Transaction tr)
    {
        try
        {
            // Get the base block name first
            BlockTableRecord btr;
            if (blockRef.IsDynamicBlock)
            {
                btr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            else
            {
                btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            
            // CRITICAL: Skip anonymous/system blocks unless they're tied to NT blocks
            if (btr.Name.StartsWith("*"))
            {
                // For dynamic blocks with anonymous definitions, check if parent is NT block
                if (blockRef.IsDynamicBlock)
                {
                    var parentBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (!parentBtr.Name.StartsWith("NT"))
                    {
                        _logger.LogDebug($"Skipping anonymous block not tied to NT block: {btr.Name} (parent: {parentBtr.Name})");
                        return string.Empty; // Not our block, skip it
                    }
                    return parentBtr.Name;
                }
                
                // Anonymous non-dynamic block - skip it
                _logger.LogDebug($"Skipping anonymous non-dynamic block: {btr.Name}");
                return string.Empty;
            }
            
            // FINAL SAFETY CHECK: Ensure this is actually a construction note block
            if (btr.Name.StartsWith("NT") && btr.Name.Length == 4 && 
                btr.Name.Substring(2).All(char.IsDigit))
            {
                _logger.LogDebug($"Confirmed valid NT block: {btr.Name}");
                return btr.Name;
            }
            else if (!btr.Name.StartsWith("NT"))
            {
                _logger.LogDebug($"Skipping non-NT block: {btr.Name}");
                return string.Empty;
            }
            else
            {
                _logger.LogDebug($"Skipping invalid NT block format: {btr.Name}");
                return string.Empty;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Could not get effective block name: {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// Clears all attributes in a construction note block
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
                        attRef.TextString = "";
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
    /// Uses the working archive approach with Justify and AdjustAlignment
    /// </summary>
    private void UpdateBlockAttributes(BlockReference blockRef, Transaction tr, Database db, int noteNumber, string noteText)
    {
        try
        {
            var attCol = blockRef.AttributeCollection;
            bool wasModified = false;
            
            foreach (ObjectId attId in attCol)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
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
                            
                            // CRITICAL: Minimal WorkingDatabase switch ONLY for AdjustAlignment
                            var originalWdb = HostApplicationServices.WorkingDatabase;
                            try
                            {
                                _logger.LogDebug($"Switching WorkingDatabase from {originalWdb?.Filename ?? "null"} to {db.Filename ?? "external"} for NUMBER attribute alignment");
                                HostApplicationServices.WorkingDatabase = db;
                                attRef.AdjustAlignment(db);
                                _logger.LogDebug($"Successfully adjusted NUMBER attribute alignment");
                            }
                            finally
                            {
                                HostApplicationServices.WorkingDatabase = originalWdb;
                                _logger.LogDebug($"Restored WorkingDatabase to {originalWdb?.Filename ?? "null"}");
                            }
                            
                            wasModified = true;
                            _logger.LogDebug($"Updated NUMBER attribute: '{newValue}' with minimal-scope WorkingDatabase alignment");
                        }
                    }
                    else if (tag == "NOTE")
                    {
                        string newValue = noteText ?? "";
                        string currentValue = attRef.TextString ?? "";
                        if (currentValue != newValue)
                        {
                            attRef.TextString = newValue;
                            
                            // CRITICAL: Minimal WorkingDatabase switch ONLY for AdjustAlignment
                            var originalWdb = HostApplicationServices.WorkingDatabase;
                            try
                            {
                                _logger.LogDebug($"Switching WorkingDatabase from {originalWdb?.Filename ?? "null"} to {db.Filename ?? "external"} for NOTE attribute alignment");
                                HostApplicationServices.WorkingDatabase = db;
                                attRef.AdjustAlignment(db);
                                _logger.LogDebug($"Successfully adjusted NOTE attribute alignment");
                            }
                            finally
                            {
                                HostApplicationServices.WorkingDatabase = originalWdb;
                                _logger.LogDebug($"Restored WorkingDatabase to {originalWdb?.Filename ?? "null"}");
                            }
                            
                            wasModified = true;
                            _logger.LogDebug($"Updated NOTE attribute with minimal-scope WorkingDatabase alignment");
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
            // SAFETY: Double-check this is our NT block before changing visibility
            if (blockRef.IsDynamicBlock)
            {
                using (var tr = blockRef.Database.TransactionManager.StartTransaction())
                {
                    var dynamicBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (dynamicBtr != null && !dynamicBtr.Name.StartsWith("NT"))
                    {
                        _logger.LogWarning($"SAFETY: Skipping visibility change for non-NT dynamic block: {dynamicBtr.Name}");
                        tr.Dispose();
                        return;
                    }
                    tr.Dispose();
                }
            }
            
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
                            _logger.LogDebug($"Set visibility of NT block to '{targetState}'");
                            return;
                        }
                        // Also check for alternative visibility states
                        if (makeVisible && (allowedState.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
                                          allowedState.Equals("Hex", StringComparison.OrdinalIgnoreCase)))
                        {
                            prop.Value = allowedState;
                            _logger.LogDebug($"Set visibility of NT block to '{allowedState}'");
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
    /// Updates title blocks within an external database using precise targeting
    /// </summary>
    private bool UpdateTitleBlocksInExternalDatabase(Database db, Transaction tr, string layoutName, List<TitleBlockAttributeData> attributeData)
    {
        try
        {
            _logger.LogDebug($"Updating title blocks in layout: {layoutName}");
            
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

            _logger.LogDebug($"Found layout '{layoutName}' (BTR: {layoutBtr.Name})");

            // Get all viewport ObjectIds to explicitly exclude them
            var viewportIds = new HashSet<ObjectId>();
            var viewports = layout.GetViewports();
            if (viewports != null)
            {
                foreach (ObjectId vpId in viewports)
                {
                    viewportIds.Add(vpId);
                }
                _logger.LogDebug($"Protected {viewportIds.Count} viewports from modification");
            }

            // Find title block blocks
            ObjectId? titleBlockId = null;
            string foundBlockName = "";
            
            foreach (ObjectId objId in layoutBtr)
            {
                var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                
                // Skip viewports using explicit ObjectId check
                if (viewportIds.Contains(objId))
                {
                    continue;
                }
                
                // Additional entity type check for viewports (double protection)
                if (entity is Viewport)
                {
                    continue;
                }
                
                if (entity is BlockReference blockRef)
                {
                    string blockName = GetEffectiveTitleBlockName(blockRef, tr);
                    
                    // Check if this is a title block
                    if (!string.IsNullOrEmpty(blockName) && _titleBlockPattern.IsMatch(blockName))
                    {
                        titleBlockId = objId;
                        foundBlockName = blockName;
                        _logger.LogInformation($"Found title block: {blockName} (ObjectId: {objId})");
                        break; // Assume only one title block per layout
                    }
                }
            }

            if (!titleBlockId.HasValue)
            {
                _logger.LogError($"No title blocks matching pattern '{_titleBlockPattern}' found in layout {layoutName}");
                return false;
            }

            // Update the title block attributes
            var titleBlockRef = tr.GetObject(titleBlockId.Value, OpenMode.ForWrite) as BlockReference;
            bool success = UpdateTitleBlockAttributes(titleBlockRef, tr, db, attributeData);
            
            if (success)
            {
                _logger.LogInformation($"Successfully updated title block {foundBlockName} with {attributeData.Count} attributes");
            }

            return success;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error updating title blocks in external database: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets the effective name of a title block, handling dynamic blocks
    /// Similar to GetEffectiveBlockName but for title blocks
    /// </summary>
    private string GetEffectiveTitleBlockName(BlockReference blockRef, Transaction tr)
    {
        try
        {
            // Get the base block name first
            BlockTableRecord btr;
            if (blockRef.IsDynamicBlock)
            {
                btr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            else
            {
                btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            
            // Skip anonymous/system blocks unless they're title blocks
            if (btr.Name.StartsWith("*"))
            {
                // For dynamic blocks with anonymous definitions, check if parent is title block
                if (blockRef.IsDynamicBlock)
                {
                    var parentBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (!_titleBlockPattern.IsMatch(parentBtr.Name))
                    {
                        _logger.LogDebug($"Skipping anonymous block not tied to title block: {btr.Name} (parent: {parentBtr.Name})");
                        return string.Empty; // Not our block, skip it
                    }
                    return parentBtr.Name;
                }
                
                // Anonymous non-dynamic block - skip it
                _logger.LogDebug($"Skipping anonymous non-dynamic block: {btr.Name}");
                return string.Empty;
            }
            
            return btr.Name;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Could not get effective title block name: {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// Updates attributes in a title block with dynamic attribute mapping
    /// Similar to UpdateBlockAttributes but handles multiple attribute-value pairs
    /// </summary>
    private bool UpdateTitleBlockAttributes(BlockReference blockRef, Transaction tr, Database db, List<TitleBlockAttributeData> attributeData)
    {
        try
        {
            var attCol = blockRef.AttributeCollection;
            bool wasModified = false;
            
            foreach (var attrData in attributeData)
            {
                if (string.IsNullOrEmpty(attrData.AttributeName) || attrData.AttributeValue == null)
                {
                    continue;
                }

                // Generate attribute name variants for case-insensitive matching
                var candidateNames = GenerateAttributeNameVariants(attrData.AttributeName);
                bool foundMatch = false;

                foreach (ObjectId attId in attCol)
                {
                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                    if (attRef != null)
                    {
                        string tag = attRef.Tag.ToUpper();
                        
                        // Check if this attribute matches any of our candidate names
                        if (candidateNames.Any(candidate => candidate.Equals(tag, StringComparison.OrdinalIgnoreCase)))
                        {
                            string currentValue = attRef.TextString ?? "";
                            if (currentValue != attrData.AttributeValue)
                            {
                                // Store original position and justification to preserve placement
                                var originalJustify = attRef.Justify;
                                var originalPosition = attRef.Position;
                                var originalAlignmentPoint = attRef.AlignmentPoint;
                                
                                // Update the text
                                attRef.TextString = attrData.AttributeValue;
                                
                                // CRITICAL: Restore the justify property BEFORE adjusting alignment
                                attRef.Justify = originalJustify;
                                
                                // Restore position based on justification type
                                if (originalJustify == AttachmentPoint.BaseLeft || 
                                    originalJustify == AttachmentPoint.BaseAlign || 
                                    originalJustify == AttachmentPoint.BaseFit)
                                {
                                    // Left-justified attributes use Position
                                    attRef.Position = originalPosition;
                                }
                                else
                                {
                                    // All other justifications use AlignmentPoint
                                    attRef.AlignmentPoint = originalAlignmentPoint;
                                }
                                
                                // Apply proper attribute alignment with database context
                                var originalWdb = HostApplicationServices.WorkingDatabase;
                                try
                                {
                                    _logger.LogDebug($"Switching WorkingDatabase for {attrData.AttributeName} attribute alignment (justify: {originalJustify})");
                                    HostApplicationServices.WorkingDatabase = db;
                                    attRef.AdjustAlignment(db);
                                }
                                finally
                                {
                                    HostApplicationServices.WorkingDatabase = originalWdb;
                                }
                                
                                wasModified = true;
                                foundMatch = true;
                                _logger.LogDebug($"Updated title block attribute '{tag}' = '{attrData.AttributeValue}'");
                                break; // Found and updated this attribute, move to next
                            }
                            else
                            {
                                foundMatch = true;
                                _logger.LogDebug($"Title block attribute '{tag}' already has value '{attrData.AttributeValue}'");
                                break;
                            }
                        }
                    }
                }

                if (!foundMatch)
                {
                    _logger.LogWarning($"Title block attribute '{attrData.AttributeName}' not found. Tried variants: {string.Join(", ", candidateNames)}");
                }
            }

            // Record graphics modification after attribute updates
            if (wasModified)
            {
                blockRef.RecordGraphicsModified(true);
                _logger.LogDebug("Applied RecordGraphicsModified to title block reference");
                _logger.LogInformation($"Title block attributes updated successfully");
            }
            else
            {
                _logger.LogDebug($"No title block attribute changes needed");
            }
            
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error updating title block attributes: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Generates attribute name variants for flexible matching
    /// Similar to construction notes but for title block attributes
    /// </summary>
    private List<string> GenerateAttributeNameVariants(string originalName)
    {
        var variants = new List<string>();
        
        if (string.IsNullOrEmpty(originalName))
        {
            return variants;
        }
        
        // Original name
        variants.Add(originalName);
        
        // Uppercase
        variants.Add(originalName.ToUpper());
        
        // Replace spaces with underscores
        if (originalName.Contains(" "))
        {
            var underscoreVersion = originalName.Replace(" ", "_");
            variants.Add(underscoreVersion);
            variants.Add(underscoreVersion.ToUpper());
        }
        
        // Remove spaces completely
        if (originalName.Contains(" "))
        {
            var noSpaceVersion = originalName.Replace(" ", "");
            variants.Add(noSpaceVersion);
            variants.Add(noSpaceVersion.ToUpper());
        }
        
        // Remove underscores
        if (originalName.Contains("_"))
        {
            var noUnderscoreVersion = originalName.Replace("_", "");
            variants.Add(noUnderscoreVersion);
            variants.Add(noUnderscoreVersion.ToUpper());
        }
        
        return variants.Distinct().ToList();
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

    /// <summary>
    /// Gets Auto Notes for a specific sheet in a closed drawing by analyzing viewports and multileaders
    /// </summary>
    /// <param name="dwgFilePath">Path to the drawing file</param>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="multileaderStyleNames">List of multileader styles to filter by</param>
    /// <returns>List of note numbers found in the sheet's viewports</returns>
    public List<int> GetAutoNotesForClosedDrawing(string dwgFilePath, string sheetName, List<string> multileaderStyleNames)
    {
        return GetAutoNotesForClosedDrawing(dwgFilePath, sheetName, multileaderStyleNames, new List<NoteBlockConfiguration>());
    }

    /// <summary>
    /// Gets Auto Notes for a specific sheet in a closed drawing by analyzing viewports, multileaders, and blocks
    /// </summary>
    /// <param name="dwgFilePath">Path to the drawing file</param>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="multileaderStyleNames">List of multileader styles to filter by</param>
    /// <param name="blockConfigurations">List of block configurations to detect</param>
    /// <returns>List of note numbers found in the sheet's viewports</returns>
    public List<int> GetAutoNotesForClosedDrawing(string dwgFilePath, string sheetName, List<string> multileaderStyleNames, List<NoteBlockConfiguration> blockConfigurations)
    {
        var noteNumbers = new List<int>();

        try
        {
            _logger.LogInformation($"=== Getting Auto Notes for closed drawing ===");
            _logger.LogInformation($"DWG Path: {dwgFilePath}");
            _logger.LogInformation($"Sheet: {sheetName}");
            _logger.LogInformation($"Styles: {string.Join(", ", multileaderStyleNames ?? new List<string>())}");
            _logger.LogInformation($"Block Configs: {string.Join(", ", blockConfigurations?.Select(bc => $"{bc.BlockName}→{bc.AttributeName}") ?? new List<string>())}");

            _logger.LogDebug("Creating external database...");
            using (var db = new Database(false, true))
            {
                _logger.LogDebug($"Reading DWG file: {Path.GetFileName(dwgFilePath)}");
                db.ReadDwgFile(dwgFilePath, FileOpenMode.OpenForReadAndAllShare, true, null);
                _logger.LogDebug("DWG file loaded successfully (CloseInput called automatically by OpenForReadAndAllShare mode)");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Get the layout
                    var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (layoutDict == null || !layoutDict.Contains(sheetName))
                    {
                        _logger.LogWarning($"Layout '{sheetName}' not found in drawing");
                        return noteNumbers;
                    }

                    var layoutId = layoutDict.GetAt(sheetName);
                    var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    if (layout == null)
                    {
                        _logger.LogWarning($"Could not access layout '{sheetName}'");
                        return noteNumbers;
                    }

                    _logger.LogDebug($"Found layout '{sheetName}'");

                    // Get viewports using BlockTableRecord iteration (works for all layouts regardless of activation)
                    var layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    if (layoutBtr == null)
                    {
                        _logger.LogWarning($"Could not access BlockTableRecord for layout '{sheetName}'");
                        return noteNumbers;
                    }

                    var viewports = new List<ObjectId>();
                    foreach (ObjectId entityId in layoutBtr)
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead);
                        if (entity is Viewport)
                        {
                            viewports.Add(entityId);
                        }
                    }

                    if (viewports.Count == 0)
                    {
                        _logger.LogDebug($"No viewports found in layout '{sheetName}'");
                        return noteNumbers;
                    }

                    _logger.LogDebug($"Found {viewports.Count} viewports in layout");
                    
                    // Debug viewport details
                    int vpIndex = 0;
                    foreach (ObjectId vpId in viewports)
                    {
                        try
                        {
                            var vp = tr.GetObject(vpId, OpenMode.ForRead) as Viewport;
                            if (vp != null)
                            {
                                _logger.LogDebug($"Viewport {vpIndex}: ID={vpId}, Number={vp.Number}, Width={vp.Width}, Height={vp.Height}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogDebug($"Error accessing viewport {vpIndex}: {ex.Message}");
                        }
                        vpIndex++;
                    }

                    // Get all multileaders in model space using MultileaderAnalyzer
                    var multileaderAnalyzer = new MultileaderAnalyzer(_logger);
                    var allMultileaders = multileaderAnalyzer.FindMultileadersInModelSpace(db, tr, multileaderStyleNames);
                    _logger.LogDebug($"Found {allMultileaders.Count} multileaders in model space");
                    
                    // Get all configured blocks in model space using BlockAnalyzer
                    var blockAnalyzer = new BlockAnalyzer(_logger);
                    var allBlocks = blockAnalyzer.FindBlocksInModelSpace(db, tr, blockConfigurations);
                    _logger.LogDebug($"Found {allBlocks.Count} blocks in model space");
                    
                    if (allMultileaders.Count == 0 && allBlocks.Count == 0)
                    {
                        var styleNames = multileaderStyleNames?.Count > 0 ? string.Join(", ", multileaderStyleNames) : "none";
                        var blockNames = blockConfigurations?.Count > 0 ? 
                            string.Join(", ", blockConfigurations.Select(bc => $"{bc.BlockName}→{bc.AttributeName}")) : "none";
                        _logger.LogDebug($"No multileaders (styles: {styleNames}) or blocks (configs: {blockNames}) found in model space");
                        return noteNumbers;
                    }

                    // For each viewport, find multileaders and blocks within its bounds
                    int viewportIndex = 0;
                    foreach (ObjectId vpId in viewports)
                    {
                        try
                        {
                            var viewport = tr.GetObject(vpId, OpenMode.ForRead) as Viewport;
                            if (viewport == null) continue;

                            viewportIndex++;

                            // Skip the first viewport (paper space viewport) using index-based filtering
                            // This matches the approach used in TESTVIEWPORT and CENTERTEST commands
                            if (viewportIndex == 1) 
                            {
                                _logger.LogDebug($"Skipping paper space viewport (index: {viewportIndex}, Number: {viewport.Number})");
                                continue;
                            }

                            _logger.LogDebug($"Processing viewport {vpId} (index: {viewportIndex}, Number: {viewport.Number})");

                            // Get viewport bounds in model space
                            var bounds = ViewportBoundaryCalculator.GetViewportBounds(viewport, tr);
                            if (!bounds.HasValue)
                            {
                                _logger.LogWarning($"Could not calculate bounds for viewport {vpId} (Number: {viewport.Number})");
                                continue;
                            }

                            _logger.LogDebug($"Viewport bounds: {bounds.Value.MinPoint} to {bounds.Value.MaxPoint}");
                            
                            // Log multileader positions for comparison
                            _logger.LogDebug($"Checking {allMultileaders.Count} multileaders and {allBlocks.Count} blocks against viewport bounds:");
                            foreach (var ml in allMultileaders)
                            {
                                _logger.LogDebug($"  Multileader {ml.NoteNumber} at: {ml.Location}");
                            }
                            foreach (var bl in allBlocks)
                            {
                                _logger.LogDebug($"  Block {bl.NoteNumber} ({bl.BlockName}) at: {bl.Location}");
                            }

                            // Find multileaders within this viewport's bounds
                            foreach (var mleader in allMultileaders)
                            {
                                // Check if multileader location is within viewport bounds
                                if (IsPointWithinBounds(mleader.Location, bounds.Value))
                                {
                                    noteNumbers.Add(mleader.NoteNumber);
                                    _logger.LogDebug($"Found multileader note {mleader.NoteNumber} within viewport bounds");
                                }
                            }

                            // Find blocks within this viewport's bounds
                            foreach (var block in allBlocks)
                            {
                                // Check if block location is within viewport bounds
                                if (IsPointWithinBounds(block.Location, bounds.Value))
                                {
                                    noteNumbers.Add(block.NoteNumber);
                                    _logger.LogDebug($"Found block note {block.NoteNumber} (block: {block.BlockName}) within viewport bounds");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning($"Error processing viewport {vpId}: {ex.Message}");
                        }
                    }

                    tr.Commit();
                }
            }

            // Remove duplicates and sort
            noteNumbers = noteNumbers.Distinct().OrderBy(n => n).ToList();
            _logger.LogInformation($"Auto Notes detection found {noteNumbers.Count} unique notes: [{string.Join(", ", noteNumbers)}]");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting auto notes for closed drawing: {ex.Message}", ex);
        }

        return noteNumbers;
    }

    /// <summary>
    /// Checks if a point is within the given bounds
    /// </summary>
    private static bool IsPointWithinBounds(Autodesk.AutoCAD.Geometry.Point3d point, Autodesk.AutoCAD.DatabaseServices.Extents3d bounds)
    {
        return point.X >= bounds.MinPoint.X && point.X <= bounds.MaxPoint.X &&
               point.Y >= bounds.MinPoint.Y && point.Y <= bounds.MaxPoint.Y;
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