using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Title block service that coordinates Excel reading and drawing operations for title blocks
/// Handles reading title block data from SHEET_INDEX table and updating title blocks
/// </summary>
public class TitleBlockService : ITitleBlockService
{
    private readonly ILogger _logger;
    private readonly IExcelReader _excelReader;
    private readonly IDrawingOperations _drawingOperations;

    public TitleBlockService(ILogger logger, IExcelReader excelReader, IDrawingOperations drawingOperations)
    {
        _logger = logger;
        _excelReader = excelReader;
        _drawingOperations = drawingOperations;
    }

    public async Task<TitleBlockMapping?> GetTitleBlockMappingForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting title block mapping for sheet {sheetName}");
            
            // Read title block mappings from the Excel file
            var mappings = await _excelReader.ReadTitleBlockMappingsAsync(config.ProjectIndexFilePath, config);
            
            // Find the mapping for the specified sheet
            var sheetMapping = mappings.FirstOrDefault(m => m.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            
            if (sheetMapping == null)
            {
                _logger.LogWarning($"No title block mapping found for sheet {sheetName}");
                return null;
            }
            
            _logger.LogInformation($"Found title block mapping for sheet {sheetName} with {sheetMapping.AttributeValues.Count} attributes");
            return sheetMapping;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get title block mapping for sheet {sheetName}: {ex.Message}", ex);
            return null;
        }
    }

    public async Task UpdateTitleBlockAsync(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Updating title block for sheet {sheetName} with {mapping.AttributeValues.Count} attributes");
            
            if (mapping.AttributeValues.Count == 0)
            {
                _logger.LogInformation($"No attributes to update for sheet {sheetName}");
                return;
            }

            // Update the title block using drawing operations
            await _drawingOperations.UpdateTitleBlockAsync(sheetName, mapping, config);
            
            _logger.LogInformation($"Successfully updated title block for sheet {sheetName}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update title block for sheet {sheetName}: {ex.Message}", ex);
            throw;
        }
    }

    public async Task<bool> ValidateTitleBlockExistsAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Validating title block exists for sheet {sheetName}");
            
            var exists = await _drawingOperations.ValidateTitleBlockExistsAsync(sheetName, config);
            
            if (!exists)
            {
                _logger.LogWarning($"Title block not found for sheet {sheetName}");
            }
            
            return exists;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to validate title block for sheet {sheetName}: {ex.Message}", ex);
            return false;
        }
    }

    public async Task<Dictionary<string, string>> GetTitleBlockAttributesAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting title block attributes for sheet {sheetName}");
            
            var attributes = await _drawingOperations.GetTitleBlockAttributesAsync(sheetName, config);
            
            _logger.LogDebug($"Retrieved {attributes.Count} attributes for title block in sheet {sheetName}");
            return attributes;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get title block attributes for sheet {sheetName}: {ex.Message}", ex);
            return new Dictionary<string, string>();
        }
    }
}