using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Analyzes multileaders in AutoCAD drawings to extract construction note numbers.
/// Handles filtering by style and extracting note numbers from block attributes.
/// </summary>
public class MultileaderAnalyzer
{
    private readonly ILogger _logger;

    public MultileaderAnalyzer(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Represents a multileader with its note number and location.
    /// </summary>
    public class MultileaderInfo
    {
        public ObjectId MultileaderId { get; set; }
        public int NoteNumber { get; set; }
        public Point3d Location { get; set; }
        public string StyleName { get; set; } = string.Empty;
        public string BlockName { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"Note {NoteNumber} at {Location} (Style: {StyleName}, Block: {BlockName})";
        }
    }

    /// <summary>
    /// Finds all multileaders in model space and extracts their note information.
    /// </summary>
    /// <param name="database">The drawing database to search</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="targetStyleName">Optional style name to filter by (case-insensitive)</param>
    /// <returns>List of multileaders with note information</returns>
    public List<MultileaderInfo> FindMultileadersInModelSpace(Database database, Transaction transaction, string? targetStyleName = null)
    {
        var multileaders = new List<MultileaderInfo>();

        try
        {
            _logger.LogDebug($"Searching for multileaders in model space (target style: {targetStyleName ?? "any"})");

            // Get the model space block table record
            var blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (blockTable == null)
            {
                _logger.LogWarning("Could not access block table");
                return multileaders;
            }

            var modelSpaceId = blockTable[BlockTableRecord.ModelSpace];
            var modelSpace = transaction.GetObject(modelSpaceId, OpenMode.ForRead) as BlockTableRecord;
            if (modelSpace == null)
            {
                _logger.LogWarning("Could not access model space");
                return multileaders;
            }

            int totalMLeaders = 0;
            int processedMLeaders = 0;

            // Iterate through all entities in model space
            foreach (ObjectId entityId in modelSpace)
            {
                var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                
                if (entity is MLeader mleader)
                {
                    totalMLeaders++;
                    
                    try
                    {
                        var mleaderInfo = AnalyzeMLeader(mleader, transaction, targetStyleName);
                        if (mleaderInfo != null)
                        {
                            multileaders.Add(mleaderInfo);
                            processedMLeaders++;
                            _logger.LogDebug($"Found multileader: {mleaderInfo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Failed to analyze multileader {entityId}: {ex.Message}");
                    }
                }
            }

            _logger.LogInformation($"Processed {processedMLeaders} of {totalMLeaders} multileaders in model space");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error finding multileaders in model space: {ex.Message}", ex);
        }

        return multileaders;
    }

    /// <summary>
    /// Analyzes a single multileader to extract note information.
    /// </summary>
    /// <param name="mleader">The multileader to analyze</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="targetStyleName">Optional style name to filter by</param>
    /// <returns>MultileaderInfo if valid note found, null otherwise</returns>
    private MultileaderInfo? AnalyzeMLeader(MLeader mleader, Transaction transaction, string? targetStyleName)
    {
        try
        {
            // Get style name for filtering
            string styleName = GetMLeaderStyleName(mleader, transaction);
            
            // Filter by style if specified
            if (!string.IsNullOrWhiteSpace(targetStyleName) && 
                !string.Equals(styleName, targetStyleName, StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogDebug($"Skipping multileader with style '{styleName}' (looking for '{targetStyleName}')");
                return null;
            }

            // Check if multileader has block content
            if (mleader.ContentType != ContentType.BlockContent)
            {
                _logger.LogDebug($"Skipping multileader with content type {mleader.ContentType} (need BlockContent)");
                return null;
            }

            // Extract note number from block attributes
            var noteNumber = ExtractNoteNumberFromBlock(mleader, transaction);
            if (!noteNumber.HasValue)
            {
                return null;
            }

            // Get the multileader location
            Point3d location = mleader.BlockPosition;

            return new MultileaderInfo
            {
                MultileaderId = mleader.ObjectId,
                NoteNumber = noteNumber.Value,
                Location = location,
                StyleName = styleName,
                BlockName = GetBlockName(mleader, transaction) ?? string.Empty
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error analyzing multileader: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets the style name of a multileader.
    /// </summary>
    /// <param name="mleader">The multileader</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>Style name or "Unknown" if not found</returns>
    private string GetMLeaderStyleName(MLeader mleader, Transaction transaction)
    {
        try
        {
            if (mleader.MLeaderStyle.IsValid)
            {
                var styleObj = transaction.GetObject(mleader.MLeaderStyle, OpenMode.ForRead);
                if (styleObj is MLeaderStyle style)
                {
                    return style.Name;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug($"Could not get MLeader style name: {ex.Message}");
        }

        return "Unknown";
    }

    /// <summary>
    /// Gets the block name referenced by a multileader.
    /// </summary>
    /// <param name="mleader">The multileader</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>Block name or null if not found</returns>
    private string? GetBlockName(MLeader mleader, Transaction transaction)
    {
        try
        {
            if (mleader.BlockContentId.IsValid)
            {
                var blockRecord = transaction.GetObject(mleader.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                return blockRecord?.Name;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug($"Could not get block name: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Extracts the note number from a multileader's block attributes.
    /// Looks specifically for the "TAGNUMBER" attribute.
    /// </summary>
    /// <param name="mleader">The multileader containing block content</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <returns>Integer note number if found and valid, null otherwise</returns>
    private int? ExtractNoteNumberFromBlock(MLeader mleader, Transaction transaction)
    {
        try
        {
            if (!mleader.BlockContentId.IsValid)
            {
                _logger.LogDebug("Multileader has invalid block content ID");
                return null;
            }

            // For multileaders, we need to access the attributes through the MLeader's GetBlockAttribute method
            // First, get the block definition to find the attribute definition
            var blockRecord = transaction.GetObject(mleader.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
            if (blockRecord == null)
            {
                _logger.LogDebug("Could not access block table record for multileader");
                return null;
            }

            // Look for attribute definition with tag "TAGNUMBER"
            ObjectId? tagNumberDefId = null;
            foreach (ObjectId entityId in blockRecord)
            {
                var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                if (entity is AttributeDefinition attDef && 
                    string.Equals(attDef.Tag, "TAGNUMBER", StringComparison.OrdinalIgnoreCase))
                {
                    tagNumberDefId = attDef.ObjectId;
                    break;
                }
            }

            if (!tagNumberDefId.HasValue)
            {
                _logger.LogDebug($"No TAGNUMBER attribute definition found in block '{blockRecord.Name}'");
                return null;
            }

            // Get the attribute value from the multileader
            var attributeRef = mleader.GetBlockAttribute(tagNumberDefId.Value);
            if (attributeRef == null)
            {
                _logger.LogDebug("Could not get TAGNUMBER attribute from multileader");
                return null;
            }

            string attributeValue = attributeRef.TextString?.Trim() ?? string.Empty;
            
            if (string.IsNullOrWhiteSpace(attributeValue))
            {
                _logger.LogDebug("TAGNUMBER attribute is empty or whitespace");
                return null;
            }

            // Try to parse as integer
            if (int.TryParse(attributeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int noteNumber))
            {
                if (noteNumber > 0) // Only accept positive note numbers
                {
                    _logger.LogDebug($"Extracted note number {noteNumber} from TAGNUMBER attribute");
                    return noteNumber;
                }
                else
                {
                    _logger.LogDebug($"TAGNUMBER attribute value {noteNumber} is not positive");
                }
            }
            else
            {
                _logger.LogDebug($"Could not parse TAGNUMBER attribute value '{attributeValue}' as integer");
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error extracting note number from multileader block: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Filters multileaders to only those within the specified viewport boundaries.
    /// </summary>
    /// <param name="multileaders">List of multileaders to filter</param>
    /// <param name="viewportBoundary">Viewport boundary polygon</param>
    /// <returns>Filtered list of multileaders within the viewport</returns>
    public List<MultileaderInfo> FilterMultileadersInViewport(List<MultileaderInfo> multileaders, Point3dCollection viewportBoundary)
    {
        if (viewportBoundary == null || viewportBoundary.Count < 3)
        {
            _logger.LogWarning("Invalid viewport boundary for filtering multileaders");
            return new List<MultileaderInfo>();
        }

        var filteredList = new List<MultileaderInfo>();

        foreach (var mleader in multileaders)
        {
            try
            {
                bool isInside = Utilities.PointInPolygonDetector.IsPointInPolygon(mleader.Location, viewportBoundary);
                
                if (isInside)
                {
                    filteredList.Add(mleader);
                    _logger.LogDebug($"Multileader {mleader} is inside viewport boundary");
                }
                else
                {
                    _logger.LogDebug($"Multileader {mleader} is outside viewport boundary");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error checking if multileader {mleader.MultileaderId} is in viewport: {ex.Message}");
            }
        }

        _logger.LogInformation($"Filtered {filteredList.Count} multileaders from {multileaders.Count} total within viewport boundary");
        return filteredList;
    }

    /// <summary>
    /// Consolidates multileaders by note number, removing duplicates.
    /// </summary>
    /// <param name="multileaders">List of multileaders to consolidate</param>
    /// <returns>List of unique note numbers</returns>
    public List<int> ConsolidateNoteNumbers(List<MultileaderInfo> multileaders)
    {
        var uniqueNotes = multileaders
            .Select(m => m.NoteNumber)
            .Distinct()
            .OrderBy(n => n)
            .ToList();

        if (uniqueNotes.Count != multileaders.Count)
        {
            _logger.LogInformation($"Consolidated {multileaders.Count} multileaders into {uniqueNotes.Count} unique notes: {string.Join(", ", uniqueNotes)}");
        }
        else
        {
            _logger.LogInformation($"Found {uniqueNotes.Count} unique notes: {string.Join(", ", uniqueNotes)}");
        }

        return uniqueNotes;
    }
}