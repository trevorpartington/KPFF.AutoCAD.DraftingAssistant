using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for performing title block operations across multiple drawings
/// Handles Active, Inactive, and Closed drawings seamlessly using appropriate managers
/// </summary>
public class MultiDrawingTitleBlockService
{
    private readonly ILogger _logger;
    private readonly DrawingAccessService _drawingAccessService;
    private readonly ExternalDrawingManager _externalDrawingManager;
    private readonly ITitleBlockService _titleBlockService;
    private readonly IExcelReader _excelReader;
    private readonly IDrawingAvailabilityService? _drawingAvailabilityService;

    public MultiDrawingTitleBlockService(
        ILogger logger,
        DrawingAccessService drawingAccessService,
        ExternalDrawingManager externalDrawingManager,
        ITitleBlockService titleBlockService,
        IExcelReader excelReader,
        IDrawingAvailabilityService? drawingAvailabilityService = null)
    {
        _logger = logger;
        _drawingAccessService = drawingAccessService;
        _externalDrawingManager = externalDrawingManager;
        _titleBlockService = titleBlockService;
        _excelReader = excelReader;
        _drawingAvailabilityService = drawingAvailabilityService;
    }

    /// <summary>
    /// Updates title blocks across multiple drawings regardless of their state
    /// </summary>
    /// <param name="selectedSheets">List of sheets to update title blocks for</param>
    /// <param name="config">Project configuration containing file paths and settings</param>
    /// <param name="sheetInfos">Sheet information containing DWG file mappings</param>
    /// <returns>Results summary with successes and failures</returns>
    public async Task<MultiDrawingUpdateResult> UpdateTitleBlocksAcrossDrawingsAsync(
        List<string> selectedSheets,
        ProjectConfiguration config,
        List<SheetInfo> sheetInfos)
    {
        var result = new MultiDrawingUpdateResult();
        
        _logger.LogInformation($"Starting multi-drawing title block update for {selectedSheets.Count} sheets");
        
        // Ensure drawing is available before starting title block operations
        if (_drawingAvailabilityService != null)
        {
            if (!_drawingAvailabilityService.EnsureDrawingAvailable(isPlottingOperation: false))
            {
                var error = "Failed to ensure drawing availability for title block update";
                _logger.LogError(error);
                result.Failures.Add(new DrawingUpdateFailure("ALL_SHEETS", "", error));
                return result;
            }
        }
        else
        {
            _logger.LogDebug("DrawingAvailabilityService not available for title block operation");
        }

        // First, get all title block mappings from Excel
        var allMappings = await _excelReader.ReadTitleBlockMappingsAsync(config.ProjectIndexFilePath, config);
        
        // Create mapping lookup, handling potential duplicates gracefully
        var mappingLookup = new Dictionary<string, TitleBlockMapping>(StringComparer.OrdinalIgnoreCase);
        foreach (var mapping in allMappings)
        {
            if (mappingLookup.ContainsKey(mapping.SheetName))
            {
                _logger.LogWarning($"Duplicate sheet name found in SHEET_INDEX: '{mapping.SheetName}'. Using the first occurrence.");
                continue;
            }
            mappingLookup[mapping.SheetName] = mapping;
        }

        foreach (var sheetName in selectedSheets)
        {
            try
            {
                _logger.LogDebug($"Processing title block update for sheet {sheetName}");

                // Find the mapping for this sheet
                if (!mappingLookup.TryGetValue(sheetName, out var mapping))
                {
                    _logger.LogWarning($"No title block mapping found for sheet {sheetName}");
                    result.Failures.Add(new DrawingUpdateFailure(sheetName, "", "No title block mapping found in SHEET_INDEX"));
                    continue;
                }

                // Find the sheet info to get drawing file information
                var sheetInfo = sheetInfos.FirstOrDefault(s => s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheetInfo == null)
                {
                    _logger.LogWarning($"No sheet info found for sheet {sheetName}");
                    result.Failures.Add(new DrawingUpdateFailure(sheetName, "", "Sheet info not found"));
                    continue;
                }

                // Get the drawing file path using the same logic as construction notes
                var fullDwgPath = _drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
                if (string.IsNullOrEmpty(fullDwgPath))
                {
                    _logger.LogWarning($"Could not resolve drawing file path for sheet {sheetName}");
                    result.Failures.Add(new DrawingUpdateFailure(sheetName, "", "Could not resolve drawing file path"));
                    continue;
                }

                // Determine drawing state and process accordingly
                var drawingState = _drawingAccessService.GetDrawingState(fullDwgPath);
                
                switch (drawingState)
                {
                    case DrawingState.Active:
                        await ProcessActiveTitleBlockUpdate(sheetName, mapping, config, result);
                        break;
                    
                    case DrawingState.Inactive:
                        await ProcessInactiveTitleBlockUpdate(sheetName, fullDwgPath, mapping, config, result);
                        break;
                    
                    case DrawingState.Closed:
                        await ProcessClosedTitleBlockUpdate(sheetName, fullDwgPath, mapping, config, result);
                        break;
                    
                    default:
                        _logger.LogError($"Unknown drawing state for sheet {sheetName}: {drawingState}");
                        result.Failures.Add(new DrawingUpdateFailure(sheetName, fullDwgPath, $"Unknown drawing state: {drawingState}"));
                        break;
                }
            }
            catch (Exception ex)
            {
                var errorMsg = $"Failed to process title block for sheet {sheetName}: {ex.Message}";
                _logger.LogError(errorMsg, ex);
                result.Failures.Add(new DrawingUpdateFailure(sheetName, "", errorMsg));
            }
        }

        _logger.LogInformation($"Multi-drawing title block update completed. Successful: {result.Successes.Count}, Failed: {result.Failures.Count}");
        return result;
    }

    private async Task ProcessActiveTitleBlockUpdate(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config, MultiDrawingUpdateResult result)
    {
        _logger.LogDebug($"Processing active title block update for sheet {sheetName}");
        
        try
        {
            await _titleBlockService.UpdateTitleBlockAsync(sheetName, mapping, config);
            result.Successes.Add(new DrawingUpdateSuccess(sheetName, "Active", DrawingState.Active, mapping.AttributeValues.Count));
            _logger.LogDebug($"Successfully updated title block for active sheet {sheetName}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Failed to update active title block: {ex.Message}";
            _logger.LogError(errorMsg, ex);
            result.Failures.Add(new DrawingUpdateFailure(sheetName, "Active", errorMsg));
        }
    }

    private async Task ProcessInactiveTitleBlockUpdate(string sheetName, string fullDwgPath, TitleBlockMapping mapping, ProjectConfiguration config, MultiDrawingUpdateResult result)
    {
        _logger.LogDebug($"Processing inactive title block update for sheet {sheetName} in file {fullDwgPath}");
        
        try
        {
            // Make the inactive drawing active temporarily
            var switched = _drawingAccessService.TryMakeDrawingActive(fullDwgPath);
            if (!switched)
            {
                throw new InvalidOperationException($"Could not make drawing active: {fullDwgPath}");
            }
            
            await _titleBlockService.UpdateTitleBlockAsync(sheetName, mapping, config);
            result.Successes.Add(new DrawingUpdateSuccess(sheetName, fullDwgPath, DrawingState.Inactive, mapping.AttributeValues.Count));
            _logger.LogDebug($"Successfully updated title block for inactive sheet {sheetName}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Failed to update inactive title block: {ex.Message}";
            _logger.LogError(errorMsg, ex);
            result.Failures.Add(new DrawingUpdateFailure(sheetName, fullDwgPath, errorMsg));
        }
    }

    private async Task ProcessClosedTitleBlockUpdate(string sheetName, string fullDwgPath, TitleBlockMapping mapping, ProjectConfiguration config, MultiDrawingUpdateResult result)
    {
        _logger.LogDebug($"Processing closed title block update for sheet {sheetName} in file {fullDwgPath}");
        
        try
        {
            if (!File.Exists(fullDwgPath))
            {
                var errorMsg = $"Drawing file not found: {fullDwgPath}";
                _logger.LogError(errorMsg);
                result.Failures.Add(new DrawingUpdateFailure(sheetName, fullDwgPath, errorMsg));
                return;
            }

            // Convert title block mapping to attribute data for external drawing manager
            var attributeData = new List<TitleBlockAttributeData>();
            foreach (var kvp in mapping.AttributeValues)
            {
                attributeData.Add(new TitleBlockAttributeData(kvp.Key, kvp.Value));
            }

            _logger.LogDebug($"Updating closed drawing with {attributeData.Count} title block attributes");

            // Set the title block pattern on the external drawing manager
            // Get the title block name from the file path (filename without extension)
            var titleBlockName = System.IO.Path.GetFileNameWithoutExtension(config.TitleBlocks.TitleBlockFilePath);
            var titleBlockPattern = $"^{System.Text.RegularExpressions.Regex.Escape(titleBlockName)}$";
            _externalDrawingManager.SetTitleBlockPattern(titleBlockPattern);

            // Use external drawing manager for closed drawings with project path for cleanup
            bool success = _externalDrawingManager.UpdateTitleBlocksInClosedDrawing(
                fullDwgPath, 
                sheetName, 
                attributeData, 
                config.ProjectDWGFilePath);

            if (success)
            {
                result.Successes.Add(new DrawingUpdateSuccess(sheetName, fullDwgPath, DrawingState.Closed, mapping.AttributeValues.Count));
                _logger.LogDebug($"Successfully updated title blocks for closed sheet {sheetName}");
            }
            else
            {
                var errorMsg = "External drawing manager failed to update title blocks";
                _logger.LogError(errorMsg);
                result.Failures.Add(new DrawingUpdateFailure(sheetName, fullDwgPath, errorMsg));
            }

            await Task.CompletedTask; // Remove async warning
        }
        catch (Exception ex)
        {
            var errorMsg = $"Failed to update closed title block: {ex.Message}";
            _logger.LogError(errorMsg, ex);
            result.Failures.Add(new DrawingUpdateFailure(sheetName, fullDwgPath, errorMsg));
        }
    }

    /// <summary>
    /// Gets title block mappings for the specified sheets
    /// </summary>
    public async Task<Dictionary<string, TitleBlockMapping>> GetTitleBlockMappingsAsync(List<string> sheetNames, ProjectConfiguration config)
    {
        try
        {
            var allMappings = await _excelReader.ReadTitleBlockMappingsAsync(config.ProjectIndexFilePath, config);
            var filteredMappings = new Dictionary<string, TitleBlockMapping>(StringComparer.OrdinalIgnoreCase);

            foreach (var mapping in allMappings)
            {
                if (sheetNames.Any(s => s.Equals(mapping.SheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    filteredMappings[mapping.SheetName] = mapping;
                }
            }

            _logger.LogInformation($"Retrieved title block mappings for {filteredMappings.Count} out of {sheetNames.Count} requested sheets");
            return filteredMappings;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get title block mappings: {ex.Message}", ex);
            return new Dictionary<string, TitleBlockMapping>();
        }
    }
}