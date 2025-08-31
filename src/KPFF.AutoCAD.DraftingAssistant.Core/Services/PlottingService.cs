using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for coordinating plotting operations with optional pre-plot updates
/// </summary>
public class PlottingService : IPlottingService
{
    private readonly ILogger _logger;
    private readonly IConstructionNotesService _constructionNotesService;
    private readonly IDrawingOperations _drawingOperations;
    private readonly IExcelReader _excelReader;
    private readonly IPlotManager? _plotManager;

    public PlottingService(
        ILogger logger,
        IConstructionNotesService constructionNotesService,
        IDrawingOperations drawingOperations,
        IExcelReader excelReader,
        IPlotManager? plotManager = null)
    {
        _logger = logger;
        _constructionNotesService = constructionNotesService;
        _drawingOperations = drawingOperations;
        _excelReader = excelReader;
        _plotManager = plotManager;
    }

    public async Task<PlotResult> PlotSheetsAsync(
        List<string> sheetNames, 
        ProjectConfiguration config,
        PlotJobSettings plotSettings,
        IProgress<PlotProgress>? progressCallback = null)
    {
        var result = new PlotResult();
        
        try
        {
            _logger.LogInformation($"Starting plot job for {sheetNames.Count} sheets");
            
            // Validate sheets before starting
            var validation = await ValidateSheetsForPlottingAsync(sheetNames, config);
            if (!validation.IsValid)
            {
                result.ErrorMessage = $"Pre-plot validation failed: {validation.Issues.Count} issues found";
                result.FailedSheets.AddRange(validation.Issues.Select(issue => new SheetPlotError
                {
                    SheetName = issue.SheetName,
                    ErrorMessage = issue.Description
                }));
                return result;
            }

            var totalSheets = validation.ValidSheets.Count;
            var currentSheetIndex = 0;

            foreach (var sheetName in validation.ValidSheets)
            {
                currentSheetIndex++;
                
                try
                {
                    // Report progress
                    progressCallback?.Report(new PlotProgress
                    {
                        CurrentSheet = sheetName,
                        CurrentOperation = "Processing sheet",
                        ProgressPercentage = (int)((double)currentSheetIndex / totalSheets * 100),
                        CompletedSheets = currentSheetIndex - 1,
                        TotalSheets = totalSheets
                    });

                    _logger.LogDebug($"Processing sheet {sheetName} ({currentSheetIndex}/{totalSheets})");

                    // Pre-plot updates
                    await PerformPrePlotUpdatesAsync(sheetName, config, plotSettings, progressCallback);

                    // Plot the sheet
                    var plotSuccess = await PlotSingleSheetAsync(sheetName, config, plotSettings, progressCallback);
                    
                    if (plotSuccess)
                    {
                        result.SuccessfulSheets.Add(sheetName);
                        _logger.LogDebug($"Successfully plotted sheet {sheetName}");
                    }
                    else
                    {
                        result.FailedSheets.Add(new SheetPlotError
                        {
                            SheetName = sheetName,
                            ErrorMessage = "Plot operation failed"
                        });
                        _logger.LogWarning($"Failed to plot sheet {sheetName}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error processing sheet {sheetName}: {ex.Message}", ex);
                    result.FailedSheets.Add(new SheetPlotError
                    {
                        SheetName = sheetName,
                        ErrorMessage = ex.Message,
                        ExceptionDetails = ex.ToString()
                    });
                }
            }

            result.Success = result.SuccessfulSheets.Count > 0;
            
            _logger.LogInformation($"Plot job completed: {result.SuccessfulSheets.Count}/{totalSheets} sheets successful");
            
            // Final progress report
            progressCallback?.Report(new PlotProgress
            {
                CurrentSheet = "Complete",
                CurrentOperation = "Plot job finished",
                ProgressPercentage = 100,
                CompletedSheets = result.SuccessfulSheets.Count,
                TotalSheets = totalSheets
            });
        }
        catch (Exception ex)
        {
            _logger.LogError($"Plot job failed with exception: {ex.Message}", ex);
            result.Success = false;
            result.ErrorMessage = $"Plot job failed: {ex.Message}";
        }

        return result;
    }

    public async Task<PlotValidationResult> ValidateSheetsForPlottingAsync(
        List<string> sheetNames,
        ProjectConfiguration config)
    {
        var result = new PlotValidationResult();
        
        try
        {
            _logger.LogDebug($"Validating {sheetNames.Count} sheets for plotting");

            // Check if plotting is enabled
            if (!config.Plotting.EnablePlotting)
            {
                result.Issues.Add(new PlotValidationIssue
                {
                    SheetName = "Project",
                    IssueType = PlotValidationIssueType.Other,
                    Description = "Plotting is disabled in project configuration",
                    IsWarning = false
                });
                result.IsValid = false;
                return result;
            }

            // Check output directory
            if (string.IsNullOrEmpty(config.Plotting.OutputDirectory))
            {
                result.Issues.Add(new PlotValidationIssue
                {
                    SheetName = "Project",
                    IssueType = PlotValidationIssueType.Other,
                    Description = "Plot output directory not configured",
                    IsWarning = false
                });
                result.IsValid = false;
                return result;
            }

            // Get sheet info from Excel
            var sheetInfoList = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            var sheetInfoDict = sheetInfoList.ToDictionary(s => s.SheetName, StringComparer.OrdinalIgnoreCase);

            foreach (var sheetName in sheetNames)
            {
                if (!sheetInfoDict.TryGetValue(sheetName, out var sheetInfo))
                {
                    result.Issues.Add(new PlotValidationIssue
                    {
                        SheetName = sheetName,
                        IssueType = PlotValidationIssueType.MissingLayout,
                        Description = $"Sheet '{sheetName}' not found in project index",
                        IsWarning = false
                    });
                    result.InvalidSheets.Add(sheetName);
                    continue;
                }

                // Build drawing file path
                var drawingPath = Path.Combine(config.ProjectDWGFilePath, sheetInfo.DWGFileName);
                if (!string.IsNullOrEmpty(sheetInfo.DWGFileName) && !sheetInfo.DWGFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    drawingPath += ".dwg";
                }

                // Check if drawing file exists
                if (!File.Exists(drawingPath))
                {
                    result.Issues.Add(new PlotValidationIssue
                    {
                        SheetName = sheetName,
                        IssueType = PlotValidationIssueType.MissingDrawing,
                        Description = $"Drawing file not found: {drawingPath}",
                        IsWarning = false
                    });
                    result.InvalidSheets.Add(sheetName);
                    continue;
                }

                result.ValidSheets.Add(sheetName);
            }

            result.IsValid = result.ValidSheets.Count > 0;
            _logger.LogDebug($"Validation complete: {result.ValidSheets.Count} valid, {result.InvalidSheets.Count} invalid sheets");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Sheet validation failed: {ex.Message}", ex);
            result.IsValid = false;
            result.Issues.Add(new PlotValidationIssue
            {
                SheetName = "Validation",
                IssueType = PlotValidationIssueType.Other,
                Description = $"Validation error: {ex.Message}",
                IsWarning = false
            });
        }

        return result;
    }

    public async Task<SheetPlotSettings?> GetDefaultPlotSettingsAsync(
        string sheetName,
        ProjectConfiguration config)
    {
        try
        {
            _logger.LogDebug($"Getting default plot settings for sheet {sheetName}");
            
            // Get sheet info from Excel
            var sheetInfoList = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            var sheetInfo = sheetInfoList.FirstOrDefault(s => s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            
            if (sheetInfo == null)
            {
                _logger.LogWarning($"Sheet {sheetName} not found in project index");
                return null;
            }

            // Build drawing file path
            var drawingPath = Path.Combine(config.ProjectDWGFilePath, sheetInfo.DWGFileName);
            if (!string.IsNullOrEmpty(sheetInfo.DWGFileName) && !sheetInfo.DWGFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                drawingPath += ".dwg";
            }

            return new SheetPlotSettings
            {
                SheetName = sheetName,
                LayoutName = sheetName, // Assume layout name matches sheet name
                DrawingPath = drawingPath,
                PlotDevice = "PDF", // Will be read from layout's actual settings
                PaperSize = "Auto", // Will be read from layout's actual settings
                PlotScale = "1:1", // Will be read from layout's actual settings
                PlotArea = "Layout", // Default to layout plotting
                PlotCentered = true
            };
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get plot settings for sheet {sheetName}: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Performs pre-plot updates (construction notes and title blocks) based on plot settings
    /// </summary>
    private async Task PerformPrePlotUpdatesAsync(
        string sheetName,
        ProjectConfiguration config,
        PlotJobSettings plotSettings,
        IProgress<PlotProgress>? progressCallback)
    {
        try
        {
            // Update construction notes if requested
            if (plotSettings.UpdateConstructionNotes)
            {
                progressCallback?.Report(new PlotProgress
                {
                    CurrentSheet = sheetName,
                    CurrentOperation = "Updating construction notes",
                    ProgressPercentage = 0,
                    CompletedSheets = 0,
                    TotalSheets = 1
                });

                _logger.LogDebug($"Updating construction notes for sheet {sheetName}");

                // Note: Construction notes mode (Auto vs Excel) is determined by the UI state
                // The ConstructionNotesService will need to be enhanced to accept a mode parameter
                // or check the UI state directly. For now, we'll use a placeholder approach.
                
                // Get notes using Auto Notes (this will need UI integration to determine mode)
                var noteNumbers = await _constructionNotesService.GetAutoNotesForSheetAsync(sheetName, config);
                await _constructionNotesService.UpdateConstructionNoteBlocksAsync(sheetName, noteNumbers, config);
            }

            // Update title blocks if requested  
            if (plotSettings.UpdateTitleBlocks)
            {
                progressCallback?.Report(new PlotProgress
                {
                    CurrentSheet = sheetName,
                    CurrentOperation = "Updating title blocks",
                    ProgressPercentage = 50,
                    CompletedSheets = 0,
                    TotalSheets = 1
                });

                _logger.LogDebug($"Updating title blocks for sheet {sheetName}");
                
                // Get title block mappings and update
                var titleBlockMappings = await _excelReader.ReadTitleBlockMappingsAsync(config.ProjectIndexFilePath, config);
                var sheetMapping = titleBlockMappings.FirstOrDefault(m => m.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                
                if (sheetMapping != null)
                {
                    await _drawingOperations.UpdateTitleBlockAsync(sheetName, sheetMapping, config);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Pre-plot updates failed for sheet {sheetName}: {ex.Message}");
            // Don't fail the entire plot job for pre-plot update failures
        }
    }

    /// <summary>
    /// Plots a single sheet using its default settings
    /// This method will delegate to the PlotManager in the Plugin project
    /// </summary>
    private async Task<bool> PlotSingleSheetAsync(
        string sheetName,
        ProjectConfiguration config,
        PlotJobSettings plotSettings,
        IProgress<PlotProgress>? progressCallback)
    {
        try
        {
            progressCallback?.Report(new PlotProgress
            {
                CurrentSheet = sheetName,
                CurrentOperation = "Plotting to PDF",
                ProgressPercentage = 75,
                CompletedSheets = 0,
                TotalSheets = 1
            });

            _logger.LogDebug($"Plotting sheet {sheetName}");

            // Get sheet info and build paths
            var sheetInfoList = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            var sheetInfo = sheetInfoList.FirstOrDefault(s => s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            
            if (sheetInfo == null)
            {
                _logger.LogError($"Sheet info not found for {sheetName}");
                return false;
            }

            var drawingPath = Path.Combine(config.ProjectDWGFilePath, sheetInfo.DWGFileName);
            if (!string.IsNullOrEmpty(sheetInfo.DWGFileName) && !sheetInfo.DWGFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                drawingPath += ".dwg";
            }

            var outputDirectory = plotSettings.OutputDirectory ?? config.Plotting.OutputDirectory;
            var outputPath = Path.Combine(outputDirectory, $"{sheetName}.pdf");

            // Ensure output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Use actual PlotManager if available, otherwise use placeholder
            if (_plotManager != null)
            {
                _logger.LogDebug($"Plotting {drawingPath} layout '{sheetName}' to {outputPath}");
                return await _plotManager.PlotLayoutToPdfAsync(drawingPath, sheetName, outputPath);
            }
            else
            {
                // Fallback to placeholder for testing without AutoCAD context
                _logger.LogDebug($"Would plot {drawingPath} layout '{sheetName}' to {outputPath}");
                await Task.Delay(100);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to plot sheet {sheetName}: {ex.Message}", ex);
            return false;
        }
    }
}