using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for coordinated batch operations across multiple drawings
/// Handles any combination of Construction Notes, Title Blocks, and PDF plotting
/// Supports Active, Inactive, and Closed drawing states seamlessly
/// </summary>
public class BatchOperationService : IBatchOperationService
{
    private readonly ILogger _logger;
    private readonly DrawingAccessService _drawingAccessService;
    private readonly MultiDrawingConstructionNotesService _multiDrawingConstructionNotesService;
    private readonly MultiDrawingTitleBlockService _multiDrawingTitleBlockService;
    private readonly IPlottingService _plottingService;
    private readonly IConstructionNotesService _constructionNotesService;
    private readonly IExcelReader _excelReader;

    public BatchOperationService(
        ILogger logger,
        DrawingAccessService drawingAccessService,
        MultiDrawingConstructionNotesService multiDrawingConstructionNotesService,
        MultiDrawingTitleBlockService multiDrawingTitleBlockService,
        IPlottingService plottingService,
        IConstructionNotesService constructionNotesService,
        IExcelReader excelReader)
    {
        _logger = logger;
        _drawingAccessService = drawingAccessService;
        _multiDrawingConstructionNotesService = multiDrawingConstructionNotesService;
        _multiDrawingTitleBlockService = multiDrawingTitleBlockService;
        _plottingService = plottingService;
        _constructionNotesService = constructionNotesService;
        _excelReader = excelReader;
    }

    public async Task<BatchOperationResult> ExecuteBatchOperationAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings,
        IProgress<BatchOperationProgress>? progressCallback = null)
    {
        var result = new BatchOperationResult();
        
        try
        {
            _logger.LogInformation($"Starting batch operation for {sheetNames.Count} sheets");
            
            // Filter sheets if ApplyToCurrentSheetOnly is enabled
            var filteredSheets = await FilterSheetsForCurrentSheetOnlyAsync(sheetNames, batchSettings);
            if (filteredSheets.Count == 0)
            {
                result.ErrorMessage = "No sheets to process after filtering";
                return result;
            }

            var totalSheets = filteredSheets.Count;
            var currentPhaseProgress = 0;

            // Phase 1: Validation (0-10%)
            progressCallback?.Report(CreateProgress("Validation", BatchOperationType.Validation, 
                "Validating sheets for batch operation", 5, 0, totalSheets));

            var validation = await ValidateSheetsForBatchOperationAsync(filteredSheets, config, batchSettings);
            if (!validation.IsValid && validation.ValidSheets.Count == 0)
            {
                result.ErrorMessage = $"Validation failed: {validation.Issues.Count} critical issues found";
                return result;
            }

            var validSheets = validation.ValidSheets;
            currentPhaseProgress = 10;

            // Get sheet info for drawing state resolution
            var sheetInfos = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);

            // Phase 2: Construction Notes Update (10-40%)
            if (batchSettings.UpdateConstructionNotes)
            {
                progressCallback?.Report(CreateProgress("Construction Notes", BatchOperationType.ConstructionNotes, 
                    "Updating construction notes", currentPhaseProgress, 0, totalSheets));

                var constructionNotesResult = await UpdateConstructionNotesInternalAsync(
                    validSheets, config, batchSettings, sheetInfos, progressCallback, currentPhaseProgress, 30);

                MergeConstructionNotesResults(result, constructionNotesResult);
                currentPhaseProgress = 40;
            }

            // Phase 3: Title Blocks Update (40-70%)
            if (batchSettings.UpdateTitleBlocks)
            {
                progressCallback?.Report(CreateProgress("Title Blocks", BatchOperationType.TitleBlocks, 
                    "Updating title blocks", currentPhaseProgress, 0, totalSheets));

                var titleBlockResult = await UpdateTitleBlocksInternalAsync(
                    validSheets, config, sheetInfos, progressCallback, currentPhaseProgress, 30);

                MergeTitleBlockResults(result, titleBlockResult);
                currentPhaseProgress = 70;
            }

            // Phase 4: Plotting (70-100%)
            if (batchSettings.PlotToPdf)
            {
                progressCallback?.Report(CreateProgress("Plotting", BatchOperationType.Plotting, 
                    "Plotting sheets to PDF", currentPhaseProgress, 0, totalSheets));

                var plotResult = await PlotSheetsInternalAsync(
                    validSheets, config, batchSettings, progressCallback, currentPhaseProgress);

                MergePlottingResults(result, plotResult);
            }

            // Final summary
            result.Success = result.SuccessfulSheets.Count > 0;
            result.Summary = GenerateSummary(result, batchSettings);
            
            _logger.LogInformation($"Batch operation completed: {result.SuccessfulSheets.Count}/{totalSheets} sheets successful");
            
            progressCallback?.Report(CreateProgress("Complete", BatchOperationType.Validation, 
                "Batch operation finished", 100, result.SuccessfulSheets.Count, totalSheets));
        }
        catch (Exception ex)
        {
            _logger.LogError($"Batch operation failed with exception: {ex.Message}", ex);
            result.Success = false;
            result.ErrorMessage = $"Batch operation failed: {ex.Message}";
        }

        return result;
    }

    public async Task<BatchOperationResult> UpdateConstructionNotesAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        bool isAutoNotesMode,
        bool applyToCurrentSheetOnly = false,
        IProgress<BatchOperationProgress>? progressCallback = null)
    {
        var batchSettings = new BatchOperationSettings
        {
            UpdateConstructionNotes = true,
            IsAutoNotesMode = isAutoNotesMode,
            ApplyToCurrentSheetOnly = applyToCurrentSheetOnly
        };

        return await ExecuteBatchOperationAsync(sheetNames, config, batchSettings, progressCallback);
    }

    public async Task<BatchOperationResult> UpdateTitleBlocksAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        bool applyToCurrentSheetOnly = false,
        IProgress<BatchOperationProgress>? progressCallback = null)
    {
        var batchSettings = new BatchOperationSettings
        {
            UpdateTitleBlocks = true,
            ApplyToCurrentSheetOnly = applyToCurrentSheetOnly
        };

        return await ExecuteBatchOperationAsync(sheetNames, config, batchSettings, progressCallback);
    }

    public async Task<BatchOperationValidationResult> ValidateSheetsForBatchOperationAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings)
    {
        var result = new BatchOperationValidationResult();
        
        try
        {
            _logger.LogDebug($"Validating {sheetNames.Count} sheets for batch operation");

            // Get sheet info from Excel
            var sheetInfos = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            var sheetInfoDict = sheetInfos.ToDictionary(s => s.SheetName, StringComparer.OrdinalIgnoreCase);

            foreach (var sheetName in sheetNames)
            {
                // Validate sheet exists in project index
                if (!sheetInfoDict.TryGetValue(sheetName, out var sheetInfo))
                {
                    result.Issues.Add(new BatchOperationValidationIssue
                    {
                        SheetName = sheetName,
                        AffectedOperation = BatchOperationType.Validation,
                        IssueType = BatchOperationValidationIssueType.MissingSheetInfo,
                        Description = $"Sheet '{sheetName}' not found in project index",
                        IsWarning = false
                    });
                    result.InvalidSheets.Add(sheetName);
                    continue;
                }

                // Validate drawing file exists
                var drawingPath = _drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
                if (string.IsNullOrEmpty(drawingPath) || !File.Exists(drawingPath))
                {
                    result.Issues.Add(new BatchOperationValidationIssue
                    {
                        SheetName = sheetName,
                        AffectedOperation = BatchOperationType.Validation,
                        IssueType = BatchOperationValidationIssueType.MissingDrawing,
                        Description = $"Drawing file not found: {drawingPath}",
                        IsWarning = false
                    });
                    result.InvalidSheets.Add(sheetName);
                    continue;
                }

                // Validate specific operation requirements
                var validForAllOperations = true;

                if (batchSettings.UpdateConstructionNotes)
                {
                    // Construction notes validation would go here if needed
                    result.ValidOperations.Add(BatchOperationType.ConstructionNotes);
                }

                if (batchSettings.UpdateTitleBlocks)
                {
                    // Title blocks validation would go here if needed
                    result.ValidOperations.Add(BatchOperationType.TitleBlocks);
                }

                if (batchSettings.PlotToPdf)
                {
                    // Validate plotting configuration
                    if (string.IsNullOrEmpty(config.Plotting?.OutputDirectory))
                    {
                        result.Issues.Add(new BatchOperationValidationIssue
                        {
                            SheetName = sheetName,
                            AffectedOperation = BatchOperationType.Plotting,
                            IssueType = BatchOperationValidationIssueType.PlottingConfig,
                            Description = "Plot output directory not configured",
                            IsWarning = false
                        });
                        validForAllOperations = false;
                    }
                    else
                    {
                        result.ValidOperations.Add(BatchOperationType.Plotting);
                    }
                }

                if (validForAllOperations)
                {
                    result.ValidSheets.Add(sheetName);
                }
                else
                {
                    result.InvalidSheets.Add(sheetName);
                }
            }

            result.IsValid = result.ValidSheets.Count > 0;
            _logger.LogDebug($"Validation complete: {result.ValidSheets.Count} valid, {result.InvalidSheets.Count} invalid sheets");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Sheet validation failed: {ex.Message}", ex);
            result.IsValid = false;
            result.Issues.Add(new BatchOperationValidationIssue
            {
                SheetName = "Validation",
                AffectedOperation = BatchOperationType.Validation,
                IssueType = BatchOperationValidationIssueType.Other,
                Description = $"Validation error: {ex.Message}",
                IsWarning = false
            });
        }

        return result;
    }

    #region Private Helper Methods

    private async Task<List<string>> FilterSheetsForCurrentSheetOnlyAsync(
        List<string> sheetNames, 
        BatchOperationSettings batchSettings)
    {
        if (!batchSettings.ApplyToCurrentSheetOnly)
        {
            return sheetNames;
        }

        var currentLayout = GetCurrentLayoutName();
        if (!string.IsNullOrEmpty(currentLayout) && 
            sheetNames.Contains(currentLayout, StringComparer.OrdinalIgnoreCase))
        {
            _logger.LogDebug($"ApplyToCurrentSheetOnly enabled: filtering to current sheet '{currentLayout}'");
            return new List<string> { currentLayout };
        }

        _logger.LogWarning($"Current layout '{currentLayout}' not found in selected sheets or could not be determined");
        return new List<string>(); // Return empty list if current sheet not found
    }

    private async Task<MultiDrawingUpdateResult> UpdateConstructionNotesInternalAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings,
        List<SheetInfo> sheetInfos,
        IProgress<BatchOperationProgress>? progressCallback,
        int baseProgress,
        int progressRange)
    {
        var sheetToNotes = new Dictionary<string, List<int>>();
        var processedSheets = 0;

        // Get note numbers for each sheet
        // For Auto Notes mode, we need to handle drawing states during collection
        if (batchSettings.IsAutoNotesMode)
        {
            processedSheets = await CollectAutoNotesAcrossDrawings(sheetNames, config, sheetInfos, sheetToNotes, progressCallback, baseProgress, progressRange / 2);
        }
        else
        {
            // Excel Notes mode doesn't require active drawings
            foreach (var sheetName in sheetNames)
            {
                try
                {
                    var noteNumbers = await _constructionNotesService.GetExcelNotesForSheetAsync(sheetName, config);
                    sheetToNotes[sheetName] = noteNumbers;
                    processedSheets++;

                    // Progress for note collection phase
                    var progress = baseProgress + (progressRange / 2) * processedSheets / sheetNames.Count;
                    progressCallback?.Report(CreateProgress("Construction Notes", BatchOperationType.ConstructionNotes, 
                        $"Collecting notes for {sheetName}", progress, processedSheets, sheetNames.Count));
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Failed to get notes for sheet {sheetName}: {ex.Message}");
                    sheetToNotes[sheetName] = new List<int>(); // Continue with empty notes
                }
            }
        }

        // Update construction notes using multi-drawing service
        return await _multiDrawingConstructionNotesService.UpdateConstructionNotesAcrossDrawingsAsync(
            sheetToNotes, config, sheetInfos);
    }

    /// <summary>
    /// Collects Auto Notes across multiple drawings, handling drawing states correctly.
    /// For inactive drawings, makes them active temporarily during Auto Notes detection.
    /// </summary>
    private async Task<int> CollectAutoNotesAcrossDrawings(
        List<string> sheetNames,
        ProjectConfiguration config,
        List<SheetInfo> sheetInfos,
        Dictionary<string, List<int>> sheetToNotes,
        IProgress<BatchOperationProgress>? progressCallback,
        int baseProgress,
        int progressRange)
    {
        var processedSheets = 0;
        
        try
        {
            // Remember the originally active drawing so we can restore it
            var originalActiveDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var originalActivePath = originalActiveDoc?.Name;
            
            _logger.LogDebug($"Starting Auto Notes collection for {sheetNames.Count} sheets. Original active drawing: {originalActivePath}");

            // Group sheets by their drawing file paths
            var sheetsByDrawing = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var sheetInfoDict = sheetInfos.ToDictionary(s => s.SheetName, StringComparer.OrdinalIgnoreCase);
            
            foreach (var sheetName in sheetNames)
            {
                var drawingPath = _drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
                if (!string.IsNullOrEmpty(drawingPath))
                {
                    if (!sheetsByDrawing.ContainsKey(drawingPath))
                    {
                        sheetsByDrawing[drawingPath] = new List<string>();
                    }
                    sheetsByDrawing[drawingPath].Add(sheetName);
                }
                else
                {
                    _logger.LogWarning($"Could not resolve drawing path for sheet {sheetName}");
                    sheetToNotes[sheetName] = new List<int>();
                    processedSheets++;
                }
            }

            // Process each drawing group
            foreach (var (drawingPath, sheetsInDrawing) in sheetsByDrawing)
            {
                try
                {
                    var drawingState = _drawingAccessService.GetDrawingState(drawingPath);
                    _logger.LogDebug($"Processing {sheetsInDrawing.Count} sheets in drawing '{drawingPath}' (state: {drawingState})");

                    switch (drawingState)
                    {
                        case DrawingState.Active:
                            // Drawing is already active, process sheets normally
                            foreach (var sheetName in sheetsInDrawing)
                            {
                                processedSheets = await CollectAutoNotesForActiveSheet(sheetName, config, sheetToNotes, progressCallback, baseProgress, progressRange, sheetNames.Count, processedSheets);
                            }
                            break;

                        case DrawingState.Inactive:
                            // Make drawing active temporarily
                            if (_drawingAccessService.TryMakeDrawingActive(drawingPath))
                            {
                                _logger.LogDebug($"Successfully made drawing active: {drawingPath}");
                                foreach (var sheetName in sheetsInDrawing)
                                {
                                    processedSheets = await CollectAutoNotesForActiveSheet(sheetName, config, sheetToNotes, progressCallback, baseProgress, progressRange, sheetNames.Count, processedSheets);
                                }
                            }
                            else
                            {
                                _logger.LogWarning($"Could not make drawing active: {drawingPath}. Skipping Auto Notes detection for sheets: {string.Join(", ", sheetsInDrawing)}");
                                foreach (var sheetName in sheetsInDrawing)
                                {
                                    sheetToNotes[sheetName] = new List<int>();
                                    processedSheets++;
                                }
                            }
                            break;

                        case DrawingState.Closed:
                            // Closed drawings cannot be processed for Auto Notes (requires viewports analysis)
                            _logger.LogWarning($"Cannot perform Auto Notes detection on closed drawing: {drawingPath}. Skipping sheets: {string.Join(", ", sheetsInDrawing)}");
                            foreach (var sheetName in sheetsInDrawing)
                            {
                                sheetToNotes[sheetName] = new List<int>();
                                processedSheets++;
                            }
                            break;

                        default:
                            _logger.LogError($"Cannot process drawing in state {drawingState}: {drawingPath}");
                            foreach (var sheetName in sheetsInDrawing)
                            {
                                sheetToNotes[sheetName] = new List<int>();
                                processedSheets++;
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error processing drawing {drawingPath}: {ex.Message}", ex);
                    foreach (var sheetName in sheetsInDrawing)
                    {
                        sheetToNotes[sheetName] = new List<int>();
                        processedSheets++;
                    }
                }
            }

            // Try to restore the original active drawing if it's different from current
            if (!string.IsNullOrEmpty(originalActivePath))
            {
                var currentActiveDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (currentActiveDoc?.Name != originalActivePath)
                {
                    var restored = _drawingAccessService.TryMakeDrawingActive(originalActivePath);
                    _logger.LogDebug($"Attempted to restore original active drawing '{originalActivePath}': {(restored ? "success" : "failed")}");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error during Auto Notes collection: {ex.Message}", ex);
            // Ensure all sheets get at least empty note lists
            foreach (var sheetName in sheetNames)
            {
                if (!sheetToNotes.ContainsKey(sheetName))
                {
                    sheetToNotes[sheetName] = new List<int>();
                    processedSheets++;
                }
            }
        }

        return processedSheets;
    }

    /// <summary>
    /// Collects Auto Notes for a single sheet in the currently active drawing.
    /// </summary>
    private async Task<int> CollectAutoNotesForActiveSheet(
        string sheetName,
        ProjectConfiguration config,
        Dictionary<string, List<int>> sheetToNotes,
        IProgress<BatchOperationProgress>? progressCallback,
        int baseProgress,
        int progressRange,
        int totalSheets,
        int processedSheets)
    {
        try
        {
            var noteNumbers = await _constructionNotesService.GetAutoNotesForSheetAsync(sheetName, config);
            sheetToNotes[sheetName] = noteNumbers;
            processedSheets++;

            // Report progress
            var progress = baseProgress + progressRange * processedSheets / totalSheets;
            progressCallback?.Report(CreateProgress("Construction Notes", BatchOperationType.ConstructionNotes,
                $"Collecting notes for {sheetName}", progress, processedSheets, totalSheets));
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Failed to get Auto Notes for sheet {sheetName}: {ex.Message}");
            sheetToNotes[sheetName] = new List<int>(); // Continue with empty notes
            processedSheets++;
        }

        return processedSheets;
    }

    private async Task<MultiDrawingUpdateResult> UpdateTitleBlocksInternalAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        List<SheetInfo> sheetInfos,
        IProgress<BatchOperationProgress>? progressCallback,
        int baseProgress,
        int progressRange)
    {
        return await _multiDrawingTitleBlockService.UpdateTitleBlocksAcrossDrawingsAsync(
            sheetNames, config, sheetInfos);
    }

    private async Task<PlotResult> PlotSheetsInternalAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings,
        IProgress<BatchOperationProgress>? progressCallback,
        int baseProgress)
    {
        var plotSettings = new PlotJobSettings
        {
            UpdateConstructionNotes = false, // Already done in previous phases
            UpdateTitleBlocks = false, // Already done in previous phases
            ApplyToCurrentSheetOnly = false, // Already filtered
            IsAutoNotesMode = batchSettings.IsAutoNotesMode,
            OutputDirectory = batchSettings.OutputDirectory
        };

        // Create a progress adapter to convert PlotProgress to BatchOperationProgress
        var progressAdapter = progressCallback != null ? 
            new Progress<PlotProgress>(plotProgress => 
            {
                var batchProgress = baseProgress + (30 * plotProgress.ProgressPercentage / 100);
                progressCallback.Report(CreateProgress("Plotting", BatchOperationType.Plotting,
                    $"Plotting {plotProgress.CurrentSheet}", batchProgress, 
                    plotProgress.CompletedSheets, plotProgress.TotalSheets));
            }) : null;

        return await _plottingService.PlotSheetsAsync(sheetNames, config, plotSettings, progressAdapter);
    }

    private void MergeConstructionNotesResults(BatchOperationResult batchResult, MultiDrawingUpdateResult constructionNotesResult)
    {
        foreach (var success in constructionNotesResult.Successes)
        {
            var existing = batchResult.SuccessfulSheets.FirstOrDefault(s => s.SheetName == success.SheetName);
            if (existing == null)
            {
                existing = new BatchOperationSheetResult
                {
                    SheetName = success.SheetName,
                    DrawingPath = success.DwgPath,
                    DrawingState = success.DrawingState
                };
                batchResult.SuccessfulSheets.Add(existing);
            }

            existing.CompletedOperations.Add(BatchOperationType.ConstructionNotes);
            existing.ConstructionNotesUpdated = success.NotesUpdated;
        }

        foreach (var failure in constructionNotesResult.Failures)
        {
            batchResult.FailedSheets.Add(new BatchOperationSheetError
            {
                SheetName = failure.SheetName,
                DrawingPath = failure.DwgPath,
                FailedOperation = BatchOperationType.ConstructionNotes,
                ErrorMessage = failure.ErrorMessage
            });
        }
    }

    private void MergeTitleBlockResults(BatchOperationResult batchResult, MultiDrawingUpdateResult titleBlockResult)
    {
        foreach (var success in titleBlockResult.Successes)
        {
            var existing = batchResult.SuccessfulSheets.FirstOrDefault(s => s.SheetName == success.SheetName);
            if (existing == null)
            {
                existing = new BatchOperationSheetResult
                {
                    SheetName = success.SheetName,
                    DrawingPath = success.DwgPath,
                    DrawingState = success.DrawingState
                };
                batchResult.SuccessfulSheets.Add(existing);
            }

            existing.CompletedOperations.Add(BatchOperationType.TitleBlocks);
            existing.TitleBlockAttributesUpdated = success.NotesUpdated; // Reusing NotesUpdated field for attributes
        }

        foreach (var failure in titleBlockResult.Failures)
        {
            batchResult.FailedSheets.Add(new BatchOperationSheetError
            {
                SheetName = failure.SheetName,
                DrawingPath = failure.DwgPath,
                FailedOperation = BatchOperationType.TitleBlocks,
                ErrorMessage = failure.ErrorMessage
            });
        }
    }

    private void MergePlottingResults(BatchOperationResult batchResult, PlotResult plotResult)
    {
        foreach (var successfulSheet in plotResult.SuccessfulSheets)
        {
            var existing = batchResult.SuccessfulSheets.FirstOrDefault(s => s.SheetName == successfulSheet);
            if (existing == null)
            {
                existing = new BatchOperationSheetResult
                {
                    SheetName = successfulSheet,
                    DrawingPath = "", // Plot result doesn't include drawing path
                    DrawingState = DrawingState.Active // Assume active for successful plots
                };
                batchResult.SuccessfulSheets.Add(existing);
            }

            existing.CompletedOperations.Add(BatchOperationType.Plotting);
            existing.PlottingSuccessful = true;
        }

        foreach (var failure in plotResult.FailedSheets)
        {
            batchResult.FailedSheets.Add(new BatchOperationSheetError
            {
                SheetName = failure.SheetName,
                DrawingPath = "",
                FailedOperation = BatchOperationType.Plotting,
                ErrorMessage = failure.ErrorMessage,
                ExceptionDetails = failure.ExceptionDetails
            });
        }
    }

    private BatchOperationSummary GenerateSummary(BatchOperationResult result, BatchOperationSettings batchSettings)
    {
        var summary = new BatchOperationSummary();

        foreach (var success in result.SuccessfulSheets)
        {
            if (success.CompletedOperations.Contains(BatchOperationType.ConstructionNotes))
            {
                summary.ConstructionNotesUpdated++;
                summary.TotalConstructionNotes += success.ConstructionNotesUpdated;
            }

            if (success.CompletedOperations.Contains(BatchOperationType.TitleBlocks))
            {
                summary.TitleBlocksUpdated++;
                summary.TotalTitleBlockAttributes += success.TitleBlockAttributesUpdated;
            }

            if (success.CompletedOperations.Contains(BatchOperationType.Plotting))
            {
                summary.SheetsPlotted++;
            }

            // Track drawing states
            if (!summary.DrawingStateBreakdown.ContainsKey(success.DrawingState))
            {
                summary.DrawingStateBreakdown[success.DrawingState] = 0;
            }
            summary.DrawingStateBreakdown[success.DrawingState]++;
        }

        return summary;
    }

    private BatchOperationProgress CreateProgress(
        string phase, 
        BatchOperationType operation, 
        string description, 
        int percentage, 
        int completed, 
        int total)
    {
        return new BatchOperationProgress
        {
            Phase = phase,
            CurrentOperation = operation,
            CurrentOperationDescription = description,
            ProgressPercentage = percentage,
            CompletedSheets = completed,
            TotalSheets = total,
            CurrentSheet = completed < total ? $"Sheet {completed + 1}" : "Complete"
        };
    }

    private string? GetCurrentLayoutName()
    {
        try
        {
            var layoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            return layoutManager?.CurrentLayout;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get current layout name: {ex.Message}", ex);
            return null;
        }
    }

    #endregion
}