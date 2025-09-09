using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System;

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
    private readonly AutoNotesService _autoNotesService;
    private readonly IDrawingAvailabilityService? _drawingAvailabilityService;

    public ConstructionNotesService(ILogger logger, IExcelReader excelReader, IDrawingOperations drawingOperations, IDrawingAvailabilityService? drawingAvailabilityService = null)
    {
        _logger = logger;
        _excelReader = excelReader;
        _drawingOperations = drawingOperations;
        _drawingAvailabilityService = drawingAvailabilityService;
        _autoNotesService = new AutoNotesService(logger);
    }

    public async Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        try
        {
            // Ensure drawing is available for auto notes detection
            if (_drawingAvailabilityService != null)
            {
                if (!_drawingAvailabilityService.EnsureDrawingAvailable(isPlottingOperation: false))
                {
                    _logger.LogError("Failed to ensure drawing availability for auto notes detection");
                    return new List<int>();
                }
            }
            else
            {
                _logger.LogDebug("DrawingAvailabilityService not available for auto notes detection");
            }

            _logger.LogDebug($"Getting auto notes for sheet {sheetName}");
            var noteNumbers = await _autoNotesService.GetAutoNotesForSheetAsync(sheetName, config);
            _logger.LogInformation($"Auto notes detection found {noteNumbers.Count} notes for sheet {sheetName}: {string.Join(", ", noteNumbers)}");
            return noteNumbers;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get auto notes for sheet {sheetName}: {ex.Message}", ex);
            return new List<int>();
        }
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
            // Ensure drawing is available for construction note block updates
            if (_drawingAvailabilityService != null)
            {
                if (!_drawingAvailabilityService.EnsureDrawingAvailable(isPlottingOperation: false))
                {
                    _logger.LogError("Failed to ensure drawing availability for construction note block updates");
                    return;
                }
            }
            else
            {
                _logger.LogDebug("DrawingAvailabilityService not available for construction note block updates");
            }

            _logger.LogDebug($"Updating construction note blocks for sheet {sheetName} with {noteNumbers.Count} notes");
            
            if (noteNumbers.Count == 0)
            {
                _logger.LogInformation($"No notes to update for sheet {sheetName}, using unified service with empty dictionary");
                var emptyNoteData = new Dictionary<int, string>();
                await _drawingOperations.SetConstructionNotesAsync(sheetName, emptyNoteData, config);
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
            
            // Build dictionary mapping note numbers to note text
            var noteData = new Dictionary<int, string>();
            foreach (var noteNumber in noteNumbers)
            {
                var note = notes.FirstOrDefault(n => n.Number == noteNumber);
                if (note != null)
                {
                    noteData[noteNumber] = note.Text ?? "";
                    _logger.LogDebug($"Mapped note {noteNumber} to text: {note.Text}");
                }
                else
                {
                    _logger.LogWarning($"No note text found for note number {noteNumber} in series {series}");
                    noteData[noteNumber] = $"Note {noteNumber}"; // Fallback text
                }
            }
            
            // Use the new unified service method
            var success = await _drawingOperations.SetConstructionNotesAsync(sheetName, noteData, config);
            
            if (success)
            {
                _logger.LogInformation($"Successfully updated construction note blocks for sheet {sheetName} using unified service");
            }
            else
            {
                _logger.LogWarning($"Failed to update construction note blocks for sheet {sheetName} using unified service");
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