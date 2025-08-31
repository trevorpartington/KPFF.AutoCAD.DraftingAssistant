using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Drawing operations service that wraps CurrentDrawingBlockManager
/// Provides a cleaner interface and handles AutoCAD-specific complexity
/// </summary>
public class DrawingOperations : IDrawingOperations
{
    private readonly ILogger _logger;
    private bool _disposed = false;

    public DrawingOperations(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Creates a fresh CurrentDrawingBlockManager to ensure current drawing state
    /// </summary>
    private CurrentDrawingBlockManager GetBlockManager()
    {
        _logger.LogDebug("Creating fresh CurrentDrawingBlockManager");
        return new CurrentDrawingBlockManager(_logger);
    }

    public async Task<List<ConstructionNoteBlock>> GetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting construction note blocks for sheet {sheetName}");
            var blockManager = GetBlockManager();
            var blocks = blockManager.GetConstructionNoteBlocks(sheetName);
            
            _logger.LogInformation($"Found {blocks.Count} construction note blocks for sheet {sheetName}");
            return await Task.FromResult(blocks);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get construction note blocks for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(new List<ConstructionNoteBlock>());
        }
    }

    public async Task<bool> UpdateConstructionNoteBlockAsync(string sheetName, int blockIndex, int noteNumber, string noteText, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Updating construction note block {blockIndex} for sheet {sheetName} with note {noteNumber}");
            var blockManager = GetBlockManager();
            var blockName = $"NT{blockIndex:D2}";
            var success = blockManager.UpdateConstructionNoteBlock(sheetName, blockName, noteNumber, noteText, noteNumber > 0);
            
            if (success)
            {
                _logger.LogDebug($"Successfully updated block {blockIndex} for sheet {sheetName}");
            }
            else
            {
                _logger.LogWarning($"Failed to update block {blockIndex} for sheet {sheetName}");
            }
            
            return await Task.FromResult(success);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception updating construction note block {blockIndex} for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(false);
        }
    }

    public async Task<bool> UpdateConstructionNoteBlocksAsync(string sheetName, List<int> noteNumbers, List<ConstructionNote> notes, ProjectConfiguration config)
    {
        // Add immediate logging for debugging
        _logger.LogInformation($"=== ENTRY: UpdateConstructionNoteBlocksAsync for sheet {sheetName} ====");
        
        try
        {
                _logger.LogInformation($"=== STARTING UpdateConstructionNoteBlocksAsync ====");
                _logger.LogInformation($"Sheet: {sheetName}");
                _logger.LogInformation($"Note numbers: [{string.Join(", ", noteNumbers)}]");
                _logger.LogInformation($"Total notes available: {notes.Count}");
                _logger.LogInformation($"Max notes per sheet: {config.ConstructionNotes.MaxNotesPerSheet}");
                
                // Create block manager with explicit logging
                _logger.LogDebug($"Creating CurrentDrawingBlockManager with logger type: {_logger.GetType().Name}");
                var blockManager = GetBlockManager();
                _logger.LogDebug($"Block manager created successfully");
                
                
                // Discover which blocks exist in this layout
                _logger.LogInformation($"Discovering existing construction note blocks in layout '{sheetName}'...");
                List<ConstructionNoteBlock> availableBlocks;
                try
                {
                    availableBlocks = blockManager.GetConstructionNoteBlocks(sheetName);
                    _logger.LogInformation($"Found {availableBlocks.Count} construction note blocks in layout '{sheetName}'");
                    
                    if (availableBlocks.Count == 0)
                    {
                        _logger.LogError($"CRITICAL: No NT## blocks found in layout '{sheetName}'. Cannot update construction notes.");
                        return false;
                    }
                    
                    // Log which blocks are available
                    var blockNames = availableBlocks.Select(b => b.BlockName).OrderBy(n => n).ToList();
                    _logger.LogInformation($"Available blocks: [{string.Join(", ", blockNames)}]");
                }
                catch (Exception layoutEx)
                {
                    _logger.LogError($"CRITICAL: Cannot access layout '{sheetName}': {layoutEx.Message}", layoutEx);
                    return false;
                }
                
                // First reset all available blocks to invisible state
                _logger.LogInformation($"Resetting construction note blocks in layout '{sheetName}'...");
                var resetSuccess = ResetConstructionNoteBlocksAsync(sheetName, config, blockManager).Result;
                if (!resetSuccess)
                {
                    _logger.LogWarning($"Warning: Reset operation had issues for sheet {sheetName}, but continuing...");
                    // Don't fail the entire operation - continue with updates
                }
                else
                {
                    _logger.LogInformation($"Block reset completed successfully");
                }

                // Create a dictionary for quick note lookup
                var noteLookup = notes.ToDictionary(n => n.Number, n => n.Text);
                _logger.LogDebug($"Created note lookup dictionary with {noteLookup.Count} entries");
                
                // Check if we have enough blocks for all the notes
                int notesToPlace = noteNumbers.Count;
                int blocksAvailable = availableBlocks.Count;
                
                if (notesToPlace > blocksAvailable)
                {
                    _logger.LogWarning($"WARNING: Need to place {notesToPlace} notes but only have {blocksAvailable} blocks available!");
                    _logger.LogWarning($"Notes {string.Join(", ", noteNumbers.Skip(blocksAvailable))} will not be placed due to missing blocks.");
                }
                
                int successCount = 0;
                int notesPlaced = 0;
                var sortedBlocks = availableBlocks.OrderBy(b => b.BlockName).ToList(); // Ensure consistent order (NT01, NT02, etc.)
                
                // Update blocks with notes, using only the blocks that exist
                foreach (var noteNumber in noteNumbers)
                {
                    if (notesPlaced >= sortedBlocks.Count)
                    {
                        _logger.LogWarning($"Cannot place note #{noteNumber} - no more blocks available (have {sortedBlocks.Count} blocks, already placed {notesPlaced})");
                        break;
                    }
                    
                    var targetBlock = sortedBlocks[notesPlaced];
                    var blockName = targetBlock.BlockName;
                    
                    _logger.LogInformation($"Processing block {blockName} for note number {noteNumber}...");
                    
                    if (noteLookup.TryGetValue(noteNumber, out var noteText))
                    {
                        _logger.LogDebug($"Found note text for #{noteNumber}: '{noteText?.Substring(0, Math.Min(noteText?.Length ?? 0, 50))}...'");
                        
                        _logger.LogInformation($"Calling blockManager.UpdateConstructionNoteBlock('{sheetName}', '{blockName}', {noteNumber}, noteText, true)");
                        var success = blockManager.UpdateConstructionNoteBlock(sheetName, blockName, noteNumber, noteText, true);
                        
                        if (success)
                        {
                            successCount++;
                            _logger.LogInformation($"✓ Successfully updated block {blockName} with note {noteNumber}");
                        }
                        else
                        {
                            _logger.LogError($"✗ FAILED to update block {blockName} for note {noteNumber}");
                        }
                    }
                    else
                    {
                        _logger.LogWarning($"Note text not found for note number {noteNumber} in lookup dictionary");
                    }
                    
                    notesPlaced++;
                }
                
                _logger.LogInformation($"=== UpdateConstructionNoteBlocksAsync COMPLETE ====");
                _logger.LogInformation($"Successfully updated {successCount} of {notesToPlace} construction note blocks for sheet {sheetName}");
                _logger.LogInformation($"Notes placed: {Math.Min(notesToPlace, blocksAvailable)} of {notesToPlace} requested");
                _logger.LogInformation($"Blocks available: {blocksAvailable}, Blocks updated successfully: {successCount}");
                
                if (successCount == 0)
                {
                    _logger.LogError($"ERROR: No blocks were updated successfully for sheet {sheetName}");
                    return false;
                }
                
                if (notesToPlace > blocksAvailable)
                {
                    _logger.LogWarning($"PARTIAL SUCCESS: Placed {Math.Min(notesToPlace, blocksAvailable)} of {notesToPlace} notes due to limited blocks");
                }
                
                // Return true if we successfully updated at least some blocks
                return await Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogError($"EXCEPTION in UpdateConstructionNoteBlocksAsync for sheet {sheetName}: {ex.Message}", ex);
            _logger.LogError($"Exception type: {ex.GetType().Name}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");
            return await Task.FromResult(false);
        }
    }

    public async Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Validating construction note blocks exist for sheet {sheetName}");
            var blockManager = GetBlockManager();
            var blocks = blockManager.GetConstructionNoteBlocks(sheetName);
            
            var hasBlocks = blocks.Count >= config.ConstructionNotes.MaxNotesPerSheet;
            _logger.LogDebug($"Sheet {sheetName} has {blocks.Count} blocks (expected: {config.ConstructionNotes.MaxNotesPerSheet})");
            
            return await Task.FromResult(hasBlocks);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception validating construction note blocks for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(false);
        }
    }

    public async Task<bool> ResetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config, CurrentDrawingBlockManager? blockManager = null)
    {
        try
        {
            _logger.LogInformation($"=== RESETTING construction note blocks for sheet {sheetName} ====");
            blockManager ??= GetBlockManager();
            
            // First, discover which NT## blocks actually exist in this layout
            _logger.LogDebug($"Discovering existing construction note blocks in layout {sheetName}...");
            var existingBlocks = blockManager.GetConstructionNoteBlocks(sheetName);
            _logger.LogInformation($"Found {existingBlocks.Count} existing construction note blocks in layout {sheetName}");
            
            if (existingBlocks.Count == 0)
            {
                _logger.LogWarning($"No construction note blocks found in layout {sheetName} - nothing to reset");
                return await Task.FromResult(true); // Not an error - just means no blocks to work with
            }
            
            // Log which blocks were found
            var blockNames = existingBlocks.Select(b => b.BlockName).OrderBy(n => n).ToList();
            _logger.LogInformation($"Existing blocks: [{string.Join(", ", blockNames)}]");
            
            // Reset only the blocks that actually exist
            int successCount = 0;
            foreach (var block in existingBlocks)
            {
                _logger.LogDebug($"Resetting block {block.BlockName}...");
                var success = blockManager.UpdateConstructionNoteBlock(sheetName, block.BlockName, 0, "", false);
                if (success)
                {
                    successCount++;
                    _logger.LogDebug($"✓ Successfully reset {block.BlockName}");
                }
                else
                {
                    _logger.LogWarning($"✗ Failed to reset {block.BlockName}");
                }
            }
            
            _logger.LogInformation($"Reset {successCount} of {existingBlocks.Count} construction note blocks for sheet {sheetName}");
            
            // Return true if we reset at least some blocks, or if there were no blocks to reset
            return await Task.FromResult(successCount > 0 || existingBlocks.Count == 0);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception resetting construction note blocks for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(false);
        }
    }

    public async Task UpdateTitleBlockAsync(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"=== UPDATING title block for sheet {sheetName} ====");
            _logger.LogInformation($"Attributes to update: {string.Join(", ", mapping.AttributeValues.Keys)}");
            
            var blockManager = GetBlockManager();
            
            // Get the title block name from the file path (filename without extension)
            var titleBlockName = Path.GetFileNameWithoutExtension(config.TitleBlocks.TitleBlockFilePath);
            var titleBlockPattern = $"^{System.Text.RegularExpressions.Regex.Escape(titleBlockName)}$";
            
            // Find title blocks matching the actual name in the layout
            var titleBlocks = blockManager.GetTitleBlocks(sheetName, titleBlockPattern);
            
            if (titleBlocks.Count == 0)
            {
                _logger.LogWarning($"No title blocks found with name '{titleBlockName}' in layout {sheetName}");
                return;
            }
            
            if (titleBlocks.Count > 1)
            {
                _logger.LogWarning($"Found {titleBlocks.Count} title blocks in layout {sheetName}, expected 1. Using the first one.");
            }
            
            var titleBlock = titleBlocks[0];
            _logger.LogInformation($"Found title block: {titleBlock.Title} in layout {sheetName}");
            
            // Update title block attributes
            var successCount = 0;
            var totalAttributes = mapping.AttributeValues.Count;
            
            foreach (var (headerName, cellValue) in mapping.AttributeValues)
            {
                try
                {
                    // Convert header name to potential attribute names
                    // "Designed By" -> ["Designed_By", "DESIGNED_BY", "DesignedBy", "DESIGNEDBY"]
                    var attributeCandidates = GenerateAttributeNameVariants(headerName);
                    
                    var updated = false;
                    foreach (var candidateAttributeName in attributeCandidates)
                    {
                        if (blockManager.UpdateTitleBlockAttribute(sheetName, titleBlock.Title, candidateAttributeName, cellValue))
                        {
                            _logger.LogDebug($"✓ Updated attribute '{candidateAttributeName}' = '{cellValue}'");
                            updated = true;
                            successCount++;
                            break;
                        }
                    }
                    
                    if (!updated)
                    {
                        _logger.LogDebug($"⚠ No matching attribute found for header '{headerName}' (tried: {string.Join(", ", attributeCandidates)})");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Failed to update attribute for header '{headerName}': {ex.Message}");
                }
            }
            
            _logger.LogInformation($"Updated {successCount} of {totalAttributes} title block attributes for sheet {sheetName}");
            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception updating title block for sheet {sheetName}: {ex.Message}", ex);
            throw;
        }
    }

    public async Task<bool> ValidateTitleBlockExistsAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Validating title block exists for sheet {sheetName}");
            var blockManager = GetBlockManager();
            // Get the title block name from the file path (filename without extension)
            var titleBlockName = Path.GetFileNameWithoutExtension(config.TitleBlocks.TitleBlockFilePath);
            var titleBlockPattern = $"^{System.Text.RegularExpressions.Regex.Escape(titleBlockName)}$";
            var titleBlocks = blockManager.GetTitleBlocks(sheetName, titleBlockPattern);
            
            var hasBlock = titleBlocks.Count > 0;
            _logger.LogDebug($"Sheet {sheetName} has {titleBlocks.Count} title blocks");
            
            return await Task.FromResult(hasBlock);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception validating title block for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(false);
        }
    }

    public async Task<Dictionary<string, string>> GetTitleBlockAttributesAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting title block attributes for sheet {sheetName}");
            var blockManager = GetBlockManager();
            // Get the title block name from the file path (filename without extension)
            var titleBlockName = Path.GetFileNameWithoutExtension(config.TitleBlocks.TitleBlockFilePath);
            var titleBlockPattern = $"^{System.Text.RegularExpressions.Regex.Escape(titleBlockName)}$";
            var attributes = blockManager.GetTitleBlockAttributes(sheetName, titleBlockPattern);
            
            _logger.LogDebug($"Retrieved {attributes.Count} title block attributes for sheet {sheetName}");
            return await Task.FromResult(attributes);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception getting title block attributes for sheet {sheetName}: {ex.Message}", ex);
            return await Task.FromResult(new Dictionary<string, string>());
        }
    }

    /// <summary>
    /// Generates possible attribute name variants from a header name
    /// Examples: "Designed By" -> ["Designed_By", "DESIGNED_BY", "DesignedBy", "DESIGNEDBY"]
    /// </summary>
    private static List<string> GenerateAttributeNameVariants(string headerName)
    {
        var variants = new List<string>();
        
        if (string.IsNullOrWhiteSpace(headerName))
            return variants;
        
        // Replace spaces with underscores and create variants
        var underscoreVersion = headerName.Replace(" ", "_");
        var noSpaceVersion = headerName.Replace(" ", "");
        
        // Add different case combinations
        variants.Add(underscoreVersion); // Original case with underscores
        variants.Add(underscoreVersion.ToUpperInvariant()); // Upper case with underscores
        variants.Add(underscoreVersion.ToLowerInvariant()); // Lower case with underscores
        variants.Add(noSpaceVersion); // Original case no spaces
        variants.Add(noSpaceVersion.ToUpperInvariant()); // Upper case no spaces
        variants.Add(noSpaceVersion.ToLowerInvariant()); // Lower case no spaces
        
        // Remove duplicates while preserving order
        return variants.Distinct().ToList();
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // No resources to dispose - CurrentDrawingBlockManager is created fresh each time
                _logger?.LogDebug("DrawingOperations disposed");
            }
            _disposed = true;
        }
    }
}