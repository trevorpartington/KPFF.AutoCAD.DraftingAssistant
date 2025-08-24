using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for performing construction note operations across multiple drawings
/// Handles Active, Inactive, and Closed drawings seamlessly using appropriate managers
/// </summary>
public class MultiDrawingConstructionNotesService
{
    private readonly ILogger _logger;
    private readonly DrawingAccessService _drawingAccessService;
    private readonly ExternalDrawingManager _externalDrawingManager;
    private readonly IConstructionNotesService _constructionNotesService;
    private readonly IExcelReader _excelReader;

    public MultiDrawingConstructionNotesService(
        ILogger logger,
        DrawingAccessService drawingAccessService,
        ExternalDrawingManager externalDrawingManager,
        IConstructionNotesService constructionNotesService,
        IExcelReader excelReader)
    {
        _logger = logger;
        _drawingAccessService = drawingAccessService;
        _externalDrawingManager = externalDrawingManager;
        _constructionNotesService = constructionNotesService;
        _excelReader = excelReader;
    }

    /// <summary>
    /// Updates construction notes across multiple drawings regardless of their state
    /// </summary>
    /// <param name="sheetToNotes">Dictionary mapping sheet names to lists of note numbers</param>
    /// <param name="config">Project configuration containing file paths and settings</param>
    /// <param name="sheetInfos">Sheet information containing DWG file mappings</param>
    /// <returns>Results summary with successes and failures</returns>
    public async Task<MultiDrawingUpdateResult> UpdateConstructionNotesAcrossDrawingsAsync(
        Dictionary<string, List<int>> sheetToNotes,
        ProjectConfiguration config,
        List<SheetInfo> sheetInfos)
    {
        var result = new MultiDrawingUpdateResult();
        
        _logger.LogInformation($"Starting multi-drawing update for {sheetToNotes.Count} sheets");

        foreach (var (sheetName, noteNumbers) in sheetToNotes)
        {
            try
            {
                _logger.LogDebug($"Processing sheet: {sheetName} with {noteNumbers.Count} notes");
                
                // Get the drawing file path for this sheet
                var dwgPath = _drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
                if (string.IsNullOrEmpty(dwgPath))
                {
                    var error = $"Could not resolve DWG file path for sheet {sheetName}";
                    _logger.LogWarning(error);
                    result.Failures.Add(new DrawingUpdateFailure(sheetName, dwgPath, error));
                    continue;
                }

                // Determine the drawing state
                var drawingState = _drawingAccessService.GetDrawingState(dwgPath);
                _logger.LogDebug($"Sheet {sheetName} drawing state: {drawingState}");

                // Route to appropriate handler based on drawing state
                var success = await UpdateDrawingByState(
                    drawingState, 
                    sheetName, 
                    dwgPath, 
                    noteNumbers, 
                    config);

                if (success)
                {
                    result.Successes.Add(new DrawingUpdateSuccess(
                        sheetName, 
                        dwgPath, 
                        drawingState, 
                        noteNumbers.Count));
                    
                    _logger.LogInformation($"Successfully updated {sheetName} ({drawingState}) with {noteNumbers.Count} notes");
                }
                else
                {
                    var error = $"Update operation failed for unknown reason";
                    result.Failures.Add(new DrawingUpdateFailure(sheetName, dwgPath, error));
                    _logger.LogError($"Failed to update {sheetName}: {error}");
                }
            }
            catch (Exception ex)
            {
                var error = $"Exception during update: {ex.Message}";
                result.Failures.Add(new DrawingUpdateFailure(sheetName, "unknown", error));
                _logger.LogError($"Failed to update sheet {sheetName}: {error}", ex);
            }
        }

        _logger.LogInformation($"Multi-drawing update complete: {result.Successes.Count} successes, {result.Failures.Count} failures");
        return result;
    }

    /// <summary>
    /// Updates a single drawing based on its current state
    /// </summary>
    private async Task<bool> UpdateDrawingByState(
        DrawingState drawingState,
        string sheetName,
        string dwgPath,
        List<int> noteNumbers,
        ProjectConfiguration config)
    {
        switch (drawingState)
        {
            case DrawingState.Active:
                _logger.LogDebug($"Using current drawing operations for active sheet {sheetName}");
                return await UpdateActiveDrawing(sheetName, noteNumbers, config);

            case DrawingState.Inactive:
                _logger.LogDebug($"Making inactive drawing active for sheet {sheetName}");
                if (_drawingAccessService.TryMakeDrawingActive(dwgPath))
                {
                    return await UpdateActiveDrawing(sheetName, noteNumbers, config);
                }
                else
                {
                    _logger.LogWarning($"Could not make drawing active, falling back to external mode for {sheetName}");
                    return await UpdateClosedDrawingAsync(sheetName, dwgPath, noteNumbers, config);
                }

            case DrawingState.Closed:
                _logger.LogDebug($"Using external drawing operations for closed sheet {sheetName}");
                return await UpdateClosedDrawingAsync(sheetName, dwgPath, noteNumbers, config);

            case DrawingState.NotFound:
                _logger.LogError($"Drawing file not found for sheet {sheetName}: {dwgPath}");
                return false;

            case DrawingState.Error:
                _logger.LogError($"Error accessing drawing for sheet {sheetName}: {dwgPath}");
                return false;

            default:
                _logger.LogError($"Unknown drawing state {drawingState} for sheet {sheetName}");
                return false;
        }
    }

    /// <summary>
    /// Updates an active or recently activated drawing using the current drawing operations
    /// </summary>
    private async Task<bool> UpdateActiveDrawing(
        string sheetName,
        List<int> noteNumbers,
        ProjectConfiguration config)
    {
        try
        {
            // Use the existing construction notes service for active drawings
            await _constructionNotesService.UpdateConstructionNoteBlocksAsync(
                sheetName, 
                noteNumbers, 
                config);
            
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update active drawing for sheet {sheetName}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Updates a closed drawing using external drawing operations
    /// </summary>
    private async Task<bool> UpdateClosedDrawingAsync(
        string sheetName,
        string dwgPath,
        List<int> noteNumbers,
        ProjectConfiguration config)
    {
        try
        {
            // Convert note numbers to ConstructionNoteData objects with real note text
            var noteDataTasks = noteNumbers.Select(async noteNumber => 
                new ConstructionNoteData(
                    noteNumber,
                    await GetNoteTextForNumberAsync(noteNumber, sheetName, config)
                ));

            var noteData = await Task.WhenAll(noteDataTasks);

            // Use external drawing manager for closed drawings with project path for cleanup
            return _externalDrawingManager.UpdateClosedDrawing(dwgPath, sheetName, noteData.ToList(), config.ProjectDWGFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to update closed drawing for sheet {sheetName}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets the note text for a specific note number in a given sheet
    /// </summary>
    private async Task<string> GetNoteTextForNumberAsync(int noteNumber, string sheetName, ProjectConfiguration config)
    {
        try
        {
            // Extract series from sheet name (e.g., "ABC-101" -> "ABC")
            var series = ExtractSeriesFromSheetName(sheetName);
            if (string.IsNullOrEmpty(series))
            {
                _logger.LogWarning($"Could not extract series from sheet name: {sheetName}");
                return $"Note {noteNumber}";
            }

            // Read construction notes from Excel for the series
            var constructionNotes = await _excelReader.ReadConstructionNotesAsync(config.ProjectIndexFilePath, series, config);
            
            // Find the specific note
            var note = constructionNotes.FirstOrDefault(n => n.Number == noteNumber);
            if (note != null)
            {
                _logger.LogDebug($"Found note text for {sheetName} note {noteNumber}: {note.Text}");
                return note.Text;
            }
            else
            {
                _logger.LogWarning($"Note {noteNumber} not found in series {series} for sheet {sheetName}");
                return $"Note {noteNumber}";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting note text for number {noteNumber} in sheet {sheetName}: {ex.Message}", ex);
            return $"Note {noteNumber}";
        }
    }

    /// <summary>
    /// Extracts the series prefix from a sheet name (e.g., "ABC-101" -> "ABC")
    /// </summary>
    private string ExtractSeriesFromSheetName(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return string.Empty;

        var dashIndex = sheetName.IndexOf('-');
        return dashIndex > 0 ? sheetName.Substring(0, dashIndex) : string.Empty;
    }
}

/// <summary>
/// Result of a multi-drawing update operation
/// </summary>
public class MultiDrawingUpdateResult
{
    public List<DrawingUpdateSuccess> Successes { get; set; } = new List<DrawingUpdateSuccess>();
    public List<DrawingUpdateFailure> Failures { get; set; } = new List<DrawingUpdateFailure>();
    
    public int TotalProcessed => Successes.Count + Failures.Count;
    public bool HasFailures => Failures.Count > 0;
    public double SuccessRate => TotalProcessed == 0 ? 0 : (double)Successes.Count / TotalProcessed;
}

/// <summary>
/// Represents a successful drawing update
/// </summary>
public class DrawingUpdateSuccess
{
    public string SheetName { get; set; }
    public string DwgPath { get; set; }
    public DrawingState DrawingState { get; set; }
    public int NotesUpdated { get; set; }

    public DrawingUpdateSuccess(string sheetName, string dwgPath, DrawingState drawingState, int notesUpdated)
    {
        SheetName = sheetName;
        DwgPath = dwgPath;
        DrawingState = drawingState;
        NotesUpdated = notesUpdated;
    }
}

/// <summary>
/// Represents a failed drawing update
/// </summary>
public class DrawingUpdateFailure
{
    public string SheetName { get; set; }
    public string DwgPath { get; set; }
    public string ErrorMessage { get; set; }

    public DrawingUpdateFailure(string sheetName, string dwgPath, string errorMessage)
    {
        SheetName = sheetName;
        DwgPath = dwgPath;
        ErrorMessage = errorMessage;
    }
}