using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Analyzes block references in AutoCAD drawings to extract construction note numbers.
/// Handles filtering by block name and extracting note numbers from specified attributes.
/// </summary>
public class BlockAnalyzer
{
    private readonly ILogger _logger;

    public BlockAnalyzer(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Represents a block reference with its note number and location.
    /// </summary>
    public class BlockInfo
    {
        public ObjectId BlockId { get; set; }
        public int NoteNumber { get; set; }
        public Point3d Location { get; set; }
        public string BlockName { get; set; } = string.Empty;
        public string AttributeName { get; set; } = string.Empty;
        public string AttributeValue { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"Note {NoteNumber} at {Location} (Block: {BlockName}, Attribute: {AttributeName}={AttributeValue})";
        }
    }

    /// <summary>
    /// Finds all configured block references in model space and extracts their note information.
    /// </summary>
    /// <param name="database">The drawing database to search</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="blockConfigurations">List of block configurations to search for</param>
    /// <returns>List of block references with note information</returns>
    public List<BlockInfo> FindBlocksInModelSpace(Database database, Transaction transaction, IEnumerable<NoteBlockConfiguration>? blockConfigurations)
    {
        var blocks = new List<BlockInfo>();

        try
        {
            var configList = blockConfigurations?.ToList();
            if (configList == null || configList.Count == 0)
            {
                _logger.LogDebug("No block configurations provided for search");
                return blocks;
            }

            var configDescriptions = string.Join(", ", configList.Select(bc => $"{bc.BlockName}→{bc.AttributeName}"));
            _logger.LogDebug($"Searching for blocks in model space (configurations: {configDescriptions})");

            // Get the model space block table record
            var blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (blockTable == null)
            {
                _logger.LogWarning("Could not access block table");
                return blocks;
            }

            var modelSpaceId = blockTable[BlockTableRecord.ModelSpace];
            var modelSpace = transaction.GetObject(modelSpaceId, OpenMode.ForRead) as BlockTableRecord;
            if (modelSpace == null)
            {
                _logger.LogWarning("Could not access model space");
                return blocks;
            }

            int totalBlocks = 0;
            int processedBlocks = 0;

            // Iterate through all entities in model space
            foreach (ObjectId entityId in modelSpace)
            {
                var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                
                if (entity is BlockReference blockRef)
                {
                    totalBlocks++;
                    
                    try
                    {
                        var blockInfo = AnalyzeBlockReference(blockRef, transaction, configList);
                        if (blockInfo != null)
                        {
                            blocks.Add(blockInfo);
                            processedBlocks++;
                            _logger.LogDebug($"Found block: {blockInfo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Failed to analyze block reference {entityId}: {ex.Message}");
                    }
                }
            }

            _logger.LogInformation($"Processed {processedBlocks} of {totalBlocks} block references in model space");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error finding blocks in model space: {ex.Message}", ex);
        }

        return blocks;
    }

    /// <summary>
    /// Analyzes a single block reference to extract note information.
    /// </summary>
    /// <param name="blockRef">The block reference to analyze</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="blockConfigurations">List of block configurations to match against</param>
    /// <returns>BlockInfo if valid note found, null otherwise</returns>
    private BlockInfo? AnalyzeBlockReference(BlockReference blockRef, Transaction transaction, List<NoteBlockConfiguration> blockConfigurations)
    {
        try
        {
            // Check if the block's layer is frozen
            if (IsLayerFrozen(blockRef.LayerId, transaction))
            {
                _logger.LogDebug($"Skipping block reference on frozen layer");
                return null;
            }

            // Get the effective block name (handles dynamic blocks)
            string blockName = GetEffectiveBlockName(blockRef, transaction);
            
            // Find matching configuration for this block name
            var matchingConfig = blockConfigurations.FirstOrDefault(config => 
                string.Equals(blockName, config.BlockName, StringComparison.OrdinalIgnoreCase));

            if (matchingConfig == null)
            {
                _logger.LogDebug($"Skipping block '{blockName}' - no matching configuration");
                return null;
            }

            // Extract note number from the specified attribute
            int noteNumber = ExtractNoteNumberFromAttribute(blockRef, transaction, matchingConfig.AttributeName);
            if (noteNumber <= 0)
            {
                _logger.LogDebug($"Block '{blockName}' does not contain valid note number in attribute '{matchingConfig.AttributeName}'");
                return null;
            }

            // Get block position
            var location = blockRef.Position;

            return new BlockInfo
            {
                BlockId = blockRef.ObjectId,
                NoteNumber = noteNumber,
                Location = location,
                BlockName = blockName,
                AttributeName = matchingConfig.AttributeName,
                AttributeValue = GetAttributeValue(blockRef, transaction, matchingConfig.AttributeName) ?? string.Empty
            };
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error analyzing block reference: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Gets the effective block name, handling dynamic blocks properly.
    /// </summary>
    /// <param name="blockRef">The block reference</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>The effective block name</returns>
    private string GetEffectiveBlockName(BlockReference blockRef, Transaction transaction)
    {
        try
        {
            // For dynamic blocks, get the original block name
            if (blockRef.IsDynamicBlock)
            {
                var dynamicBlockTableRecord = transaction.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return dynamicBlockTableRecord?.Name ?? blockRef.Name;
            }
            
            return blockRef.Name;
        }
        catch
        {
            return blockRef.Name;
        }
    }

    /// <summary>
    /// Extracts a note number from a specific attribute of a block reference.
    /// </summary>
    /// <param name="blockRef">The block reference</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="attributeName">Name of the attribute to search</param>
    /// <returns>The note number if found and valid, 0 otherwise</returns>
    private int ExtractNoteNumberFromAttribute(BlockReference blockRef, Transaction transaction, string attributeName)
    {
        try
        {
            string? attributeValue = GetAttributeValue(blockRef, transaction, attributeName);
            
            if (string.IsNullOrWhiteSpace(attributeValue))
            {
                return 0;
            }

            // Try to parse the attribute value as an integer
            if (int.TryParse(attributeValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int noteNumber))
            {
                return noteNumber > 0 ? noteNumber : 0;
            }

            _logger.LogDebug($"Could not parse note number from attribute '{attributeName}' value '{attributeValue}' in block '{blockRef.Name}'");
            return 0;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error extracting note number from attribute '{attributeName}' in block '{blockRef.Name}': {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Gets the value of a specific attribute from a block reference.
    /// </summary>
    /// <param name="blockRef">The block reference</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="attributeName">Name of the attribute to retrieve</param>
    /// <returns>The attribute value if found, null otherwise</returns>
    private string? GetAttributeValue(BlockReference blockRef, Transaction transaction, string attributeName)
    {
        try
        {
            // Check if block has attributes
            var attributeCollection = blockRef.AttributeCollection;
            if (attributeCollection == null || attributeCollection.Count == 0)
            {
                return null;
            }

            // Search through all attributes
            foreach (ObjectId attributeId in attributeCollection)
            {
                var attributeRef = transaction.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;
                if (attributeRef != null && 
                    string.Equals(attributeRef.Tag, attributeName, StringComparison.OrdinalIgnoreCase))
                {
                    return attributeRef.TextString;
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error getting attribute '{attributeName}' from block '{blockRef.Name}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Consolidates a list of block notes, removing duplicates and sorting.
    /// </summary>
    /// <param name="blockInfos">List of block information objects</param>
    /// <returns>Sorted list of unique note numbers</returns>
    public List<int> ConsolidateNoteNumbers(List<BlockInfo> blockInfos)
    {
        try
        {
            var uniqueNotes = blockInfos
                .Where(bi => bi.NoteNumber > 0)
                .Select(bi => bi.NoteNumber)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            _logger.LogDebug($"Consolidated {blockInfos.Count} block entries into {uniqueNotes.Count} unique note numbers");
            return uniqueNotes;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error consolidating note numbers: {ex.Message}", ex);
            return new List<int>();
        }
    }

    /// <summary>
    /// Filters blocks to those within the specified viewport boundary and not on frozen layers.
    /// </summary>
    /// <param name="allBlocks">All blocks to filter</param>
    /// <param name="viewportBoundary">The viewport boundary polygon</param>
    /// <param name="viewport">Viewport for checking viewport-specific frozen layers</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>List of blocks within the boundary and visible</returns>
    public List<BlockInfo> FilterBlocksInViewport(List<BlockInfo> allBlocks, Point3dCollection viewportBoundary, Viewport? viewport = null, Transaction? transaction = null)
    {
        var blocksInViewport = new List<BlockInfo>();
        int skippedBoundary = 0;
        int skippedFrozen = 0;

        try
        {
            foreach (var block in allBlocks)
            {
                // First check if within viewport boundary
                if (!Utilities.PointInPolygonDetector.IsPointInPolygon(block.Location, viewportBoundary))
                {
                    skippedBoundary++;
                    _logger.LogDebug($"Block {block.BlockName} (note {block.NoteNumber}) at {block.Location} is outside viewport boundary");
                    continue;
                }

                // Then check if layer is frozen in this viewport
                if (viewport != null && transaction != null)
                {
                    if (IsLayerFrozenInViewport(block.BlockId, viewport, transaction))
                    {
                        skippedFrozen++;
                        _logger.LogDebug($"Block {block.BlockName} (note {block.NoteNumber}) is on layer frozen in viewport");
                        continue;
                    }
                }

                blocksInViewport.Add(block);
                _logger.LogDebug($"Block {block.BlockName} (note {block.NoteNumber}) at {block.Location} is inside viewport boundary and visible");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error filtering blocks in viewport: {ex.Message}", ex);
        }

        _logger.LogInformation($"Filtered {blocksInViewport.Count} blocks from {allBlocks.Count} total (skipped {skippedBoundary} outside boundary, {skippedFrozen} on frozen layers)");
        return blocksInViewport;
    }

    /// <summary>
    /// Checks if a layer is frozen globally.
    /// </summary>
    /// <param name="layerId">The ObjectId of the layer to check</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>True if the layer is frozen, false otherwise</returns>
    private bool IsLayerFrozen(ObjectId layerId, Transaction transaction)
    {
        try
        {
            var layerRecord = transaction.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
            if (layerRecord != null)
            {
                bool isFrozen = layerRecord.IsFrozen;
                if (isFrozen)
                {
                    _logger.LogDebug($"Layer '{layerRecord.Name}' is globally frozen");
                }
                return isFrozen;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error checking if layer is frozen: {ex.Message}");
        }
        
        return false;
    }

    /// <summary>
    /// Checks if a layer is frozen in a specific viewport (includes both global and viewport-specific freezing).
    /// </summary>
    /// <param name="entityId">The ObjectId of the entity to check</param>
    /// <param name="viewport">The viewport to check</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>True if the layer is frozen in the viewport, false otherwise</returns>
    private bool IsLayerFrozenInViewport(ObjectId entityId, Viewport viewport, Transaction transaction)
    {
        try
        {
            var entity = transaction.GetObject(entityId, OpenMode.ForRead) as Entity;
            if (entity == null) return false;

            var layerId = entity.LayerId;
            
            // Check global freezing first
            var layerRecord = transaction.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
            if (layerRecord?.IsFrozen == true)
            {
                _logger.LogDebug($"Layer '{layerRecord.Name}' is globally frozen");
                return true;
            }
            
            // Check viewport-specific freezing (attempt for all databases)
            try
            {
                var frozenLayers = viewport.GetFrozenLayers();
                var isActiveDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Database == viewport.Database;
                _logger.LogDebug($"Successfully accessed viewport frozen layers (count: {frozenLayers.Count}, {(isActiveDb ? "active" : "side")} database)");
                
                if (frozenLayers.Contains(layerId))
                {
                    _logger.LogDebug($"Layer '{layerRecord?.Name}' is frozen in viewport {viewport.Number}");
                    return true;
                }
            }
            catch (Exception vpEx)
            {
                var isActiveDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Database == viewport.Database;
                _logger.LogDebug($"Unable to check viewport frozen layers ({(isActiveDb ? "active" : "side")} database): {vpEx.Message}");
                // Continue with global layer checking only
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error checking if layer is frozen in viewport: {ex.Message}");
            return false;
        }
    }
}