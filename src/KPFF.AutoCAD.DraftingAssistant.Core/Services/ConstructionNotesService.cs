using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Construction notes service that coordinates Excel reading and drawing operations
/// Supports both Excel Notes and Auto Notes modes
/// </summary>
public class ConstructionNotesService : IConstructionNotesService
{
    private readonly ILogger _logger;
    private readonly IExcelReader _excelReader;
    private readonly IDrawingOperations _drawingOperations;

    public ConstructionNotesService(ILogger logger, IExcelReader excelReader, IDrawingOperations drawingOperations)
    {
        _logger = logger;
        _excelReader = excelReader;
        _drawingOperations = drawingOperations;
    }

    public Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        _logger.LogDebug($"GetAutoNotesForSheetAsync called for {sheetName} - returning empty list (stub implementation)");
        return Task.FromResult(new List<int>());
    }

    public async Task<List<int>> GetExcelNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting Excel notes for sheet {sheetName}");
            
            // Read Excel notes mappings from the Excel file
            var mappings = await _excelReader.ReadExcelNotesAsync(config.ProjectIndexFilePath, config);
            
            // Find the mapping for the specified sheet
            var sheetMapping = mappings.FirstOrDefault(m => m.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            
            if (sheetMapping == null)
            {
                _logger.LogWarning($"No Excel notes mapping found for sheet {sheetName}");
                return new List<int>();
            }
            
            _logger.LogInformation($"Found {sheetMapping.NoteNumbers.Count} notes for sheet {sheetName} in Excel: {string.Join(", ", sheetMapping.NoteNumbers)}");
            return sheetMapping.NoteNumbers.ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get Excel notes for sheet {sheetName}: {ex.Message}", ex);
            return new List<int>();
        }
    }

    public async Task UpdateConstructionNoteBlocksAsync(string sheetName, List<int> noteNumbers, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Updating construction note blocks for sheet {sheetName} with {noteNumbers.Count} notes");
            
            if (noteNumbers.Count == 0)
            {
                _logger.LogInformation($"No notes to update for sheet {sheetName}, resetting all blocks");
                await _drawingOperations.ResetConstructionNoteBlocksAsync(sheetName, config);
                return;
            }
            
            // Extract series from sheet name to get the correct note text
            var configService = new ProjectConfigurationService((IApplicationLogger)_logger);
            var parts = configService.ExtractSeriesFromSheetName(sheetName, config.SheetNaming);
            
            if (parts.Length == 0)
            {
                _logger.LogError($"Could not extract series from sheet name {sheetName}");
                return;
            }
            
            var series = parts[0];
            _logger.LogDebug($"Extracted series '{series}' from sheet name {sheetName}");
            
            // Get construction notes for the series
            var notes = await GetNotesForSeriesAsync(series, config);
            if (notes.Count == 0)
            {
                _logger.LogError($"No construction notes found for series {series}");
                return;
            }
            
            // Update the blocks using DrawingOperations
            var success = await _drawingOperations.UpdateConstructionNoteBlocksAsync(sheetName, noteNumbers, notes, config);
            
            if (success)
            {
                _logger.LogInformation($"Successfully updated construction note blocks for sheet {sheetName}");
            }
            else
            {
                _logger.LogWarning($"Failed to update some construction note blocks for sheet {sheetName}");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Exception updating construction note blocks for sheet {sheetName}: {ex.Message}", ex);
        }
    }

    public async Task<List<ConstructionNote>> GetNotesForSeriesAsync(string series, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting construction notes for series {series}");
            
            // Use ExcelReader to get construction notes for the series
            var notes = await _excelReader.ReadConstructionNotesAsync(config.ProjectIndexFilePath, series, config);
            
            _logger.LogInformation($"Found {notes.Count} construction notes for series {series}");
            return notes;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get construction notes for series {series}: {ex.Message}", ex);
            return new List<ConstructionNote>();
        }
    }

    public async Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Validating note blocks exist for sheet {sheetName}");
            
            // Use DrawingOperations to validate blocks exist
            var blocksExist = await _drawingOperations.ValidateNoteBlocksExistAsync(sheetName, config);
            
            _logger.LogDebug($"Sheet {sheetName} blocks validation result: {blocksExist}");
            return blocksExist;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to validate note blocks for sheet {sheetName}: {ex.Message}", ex);
            return false;
        }
    }

    public async Task CreateNoteBlocksForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Creating note blocks for sheet {sheetName}");
            
            // For now, this is a placeholder - block creation would require more complex AutoCAD operations
            // In the current implementation, we assume blocks already exist and just need updating
            _logger.LogInformation($"Note block creation not yet implemented for sheet {sheetName}");
            
            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to create note blocks for sheet {sheetName}: {ex.Message}", ex);
        }
    }
}