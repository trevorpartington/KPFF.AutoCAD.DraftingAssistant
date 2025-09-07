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

    /// <summary>
    /// Unified clear-and-fill service for construction note blocks.
    /// Always clears all blocks first, then fills with provided note data.
    /// This ensures consistent updates regardless of what data changed.
    /// </summary>
    /// <param name="sheetName">The layout/sheet name to update</param>
    /// <param name="noteData">Dictionary of note number to note text mappings</param>
    /// <param name="config">Project configuration</param>
    /// <returns>True if successful, false otherwise</returns>
    public async Task<bool> SetConstructionNotesAsync(string sheetName, Dictionary<int, string> noteData, ProjectConfiguration config)
    {
        _logger.LogInformation($"=== ENTRY: SetConstructionNotesAsync for sheet {sheetName} ====");
        
        try
        {
            _logger.LogInformation($"=== STARTING SetConstructionNotesAsync (CLEAR-AND-FILL) ====");
            _logger.LogInformation($"Sheet: {sheetName}");
            _logger.LogInformation($"Note data: {noteData.Count} entries");
            if (noteData.Count > 0)
            {
                var noteInfo = string.Join(", ", noteData.Select(kvp => $"#{kvp.Key}=\"{kvp.Value.Substring(0, Math.Min(kvp.Value.Length, 30))}...\""));
                _logger.LogInformation($"Notes: {noteInfo}");
            }
            
            var blockManager = GetBlockManager();
            _logger.LogDebug($"Block manager created successfully");
            
            // STEP 1: Build clear-and-fill batch update dictionary for all 24 blocks
            _logger.LogInformation($"Building clear-and-fill update plan for all 24 blocks...");
            var batchUpdates = new Dictionary<string, (int noteNumber, string noteText, bool makeVisible)>();
            
            // Clear all blocks first (start with reset state)
            for (int i = 1; i <= 24; i++)
            {
                string blockName = $"NT{i:D2}";
                batchUpdates[blockName] = (-1, "", false); // Reset: invisible with empty data
            }
            
            // STEP 2: Fill blocks with provided note data
            if (noteData.Count > 0)
            {
                int blockIndex = 1;
                foreach (var kvp in noteData.OrderBy(x => x.Key)) // Sort by note number for consistent ordering
                {
                    if (blockIndex > 24)
                    {
                        _logger.LogWarning($"Cannot place note #{kvp.Key} - only 24 blocks available (NT01-NT24)");
                        break;
                    }
                    
                    string blockName = $"NT{blockIndex:D2}";
                    batchUpdates[blockName] = (kvp.Key, kvp.Value ?? "", true); // Visible with note data
                    
                    _logger.LogDebug($"✓ Planned update for {blockName} with note #{kvp.Key}: '{kvp.Value?.Substring(0, Math.Min(kvp.Value?.Length ?? 0, 30))}...'");
                    blockIndex++;
                }
                
                _logger.LogInformation($"Filled {Math.Min(noteData.Count, 24)} blocks with note data, reset remaining {Math.Max(0, 24 - noteData.Count)} blocks");
            }
            else
            {
                _logger.LogInformation($"No note data provided - all 24 blocks will be cleared");
            }
            
            // STEP 3: Execute single batch operation with discovery (clears all then fills)
            _logger.LogInformation($"Executing unified clear-and-fill batch update for {batchUpdates.Count} blocks...");
            var (batchSuccess, discoveredBlocks) = blockManager.UpdateConstructionNoteBlocksBatchWithDiscovery(sheetName, batchUpdates);
            
            _logger.LogInformation($"Batch update discovered {discoveredBlocks.Count} actual blocks: [{string.Join(", ", discoveredBlocks.OrderBy(x => x))}]");
            int actualBlocksAvailable = discoveredBlocks.Count;
            
            _logger.LogInformation($"=== SetConstructionNotesAsync COMPLETE ====");
            _logger.LogInformation($"Clear-and-fill operation result: {(batchSuccess ? "SUCCESS" : "FAILED")}");
            _logger.LogInformation($"Processed {noteData.Count} note entries across {actualBlocksAvailable} discovered blocks");
            _logger.LogInformation($"Blocks planned: {batchUpdates.Count}, Blocks discovered: {actualBlocksAvailable}");
            
            if (!batchSuccess)
            {
                _logger.LogError($"ERROR: Clear-and-fill batch operation failed for sheet {sheetName}");
            }
            
            return await Task.FromResult(batchSuccess);
        }
        catch (Exception ex)
        {
            _logger.LogError($"EXCEPTION in SetConstructionNotesAsync for sheet {sheetName}: {ex.Message}", ex);
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