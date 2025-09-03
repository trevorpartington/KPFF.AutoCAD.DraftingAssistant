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
    private readonly MultiDrawingConstructionNotesService _multiDrawingConstructionNotesService;
    private readonly MultiDrawingTitleBlockService _multiDrawingTitleBlockService;

    public PlottingService(
        ILogger logger,
        IConstructionNotesService constructionNotesService,
        IDrawingOperations drawingOperations,
        IExcelReader excelReader,
        MultiDrawingConstructionNotesService multiDrawingConstructionNotesService,
        MultiDrawingTitleBlockService multiDrawingTitleBlockService,
        IPlotManager? plotManager = null)
    {
        _logger = logger;
        _constructionNotesService = constructionNotesService;
        _drawingOperations = drawingOperations;
        _excelReader = excelReader;
        _multiDrawingConstructionNotesService = multiDrawingConstructionNotesService;
        _multiDrawingTitleBlockService = multiDrawingTitleBlockService;
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
            // Filter sheets if ApplyToCurrentSheetOnly is enabled
            if (plotSettings.ApplyToCurrentSheetOnly)
            {
                var currentLayout = GetCurrentLayoutName();
                if (!string.IsNullOrEmpty(currentLayout))
                {
                    // Only include the current layout if it's in the selected sheets
                    if (sheetNames.Contains(currentLayout, StringComparer.OrdinalIgnoreCase))
                    {
                        sheetNames = new List<string> { currentLayout };
                        _logger.LogDebug($"ApplyToCurrentSheetOnly enabled: filtering to current sheet '{currentLayout}'");
                    }
                    else
                    {
                        _logger.LogWarning($"Current layout '{currentLayout}' not found in selected sheets, no sheets will be plotted");
                        result.ErrorMessage = $"Current layout '{currentLayout}' is not in the selected sheet list";
                        return result;
                    }
                }
                else
                {
                    _logger.LogError("Could not determine current layout for ApplyToCurrentSheetOnly option");
                    result.ErrorMessage = "Could not determine current layout";
                    return result;
                }
            }

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

            // Perform pre-plot updates first (if needed)
            if (plotSettings.UpdateConstructionNotes || plotSettings.UpdateTitleBlocks)
            {
                foreach (var sheetName in validation.ValidSheets)
                {
                    currentSheetIndex++;
                    
                    try
                    {
                        // Report progress for pre-plot updates
                        progressCallback?.Report(new PlotProgress
                        {
                            CurrentSheet = sheetName,
                            CurrentOperation = "Pre-plot updates",
                            ProgressPercentage = (int)((double)currentSheetIndex / totalSheets * 50),
                            CompletedSheets = 0,
                            TotalSheets = totalSheets
                        });

                        _logger.LogDebug($"Performing pre-plot updates for sheet {sheetName} ({currentSheetIndex}/{totalSheets})");
                        await PerformPrePlotUpdatesAsync(sheetName, config, plotSettings, progressCallback);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Pre-plot updates failed for sheet {sheetName}: {ex.Message}");
                        // Continue with plotting even if pre-plot updates fail
                    }
                }
                
                // Reset for plotting phase
                currentSheetIndex = 0;
            }

            // Use batch plotting with Publisher API if PlotManager supports it
            if (_plotManager != null && validation.ValidSheets.Count > 1)
            {
                try
                {
                    _logger.LogDebug($"Using batch plotting for {validation.ValidSheets.Count} sheets");
                    
                    progressCallback?.Report(new PlotProgress
                    {
                        CurrentSheet = "Batch plotting",
                        CurrentOperation = "Plotting to PDF",
                        ProgressPercentage = 75,
                        CompletedSheets = 0,
                        TotalSheets = totalSheets
                    });

                    // Build sheet collection for batch plotting
                    var sheetsToPlot = await BuildSheetCollectionAsync(validation.ValidSheets, config);
                    var outputDirectory = plotSettings.OutputDirectory ?? config.Plotting.OutputDirectory;
                    
                    var batchResult = await _plotManager.PublishSheetsToPdfAsync(sheetsToPlot, outputDirectory);
                    
                    if (batchResult)
                    {
                        // All sheets succeeded
                        result.SuccessfulSheets.AddRange(validation.ValidSheets);
                        _logger.LogInformation($"Batch plot successful for all {validation.ValidSheets.Count} sheets");
                    }
                    else
                    {
                        // Batch plot failed - fall back to individual plotting
                        _logger.LogWarning("Batch plotting failed, falling back to individual sheet plotting");
                        await PlotSheetsIndividuallyAsync(validation.ValidSheets, config, plotSettings, progressCallback, result);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Batch plotting failed: {ex.Message}", ex);
                    // Fall back to individual plotting
                    await PlotSheetsIndividuallyAsync(validation.ValidSheets, config, plotSettings, progressCallback, result);
                }
            }
            else
            {
                // Plot sheets individually (single sheet or no PlotManager)
                await PlotSheetsIndividuallyAsync(validation.ValidSheets, config, plotSettings, progressCallback, result);
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
    /// Uses multi-drawing services to support Active, Inactive, and Closed drawing states
    /// </summary>
    private async Task PerformPrePlotUpdatesAsync(
        string sheetName,
        ProjectConfiguration config,
        PlotJobSettings plotSettings,
        IProgress<PlotProgress>? progressCallback)
    {
        try
        {
            // Get sheet info list for drawing state resolution
            var sheetInfos = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);

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

                _logger.LogDebug($"Updating construction notes for sheet {sheetName} using {(plotSettings.IsAutoNotesMode ? "Auto Notes" : "Excel Notes")} mode");

                List<int> noteNumbers;
                if (plotSettings.IsAutoNotesMode)
                {
                    // Get notes using Auto Notes mode
                    noteNumbers = await _constructionNotesService.GetAutoNotesForSheetAsync(sheetName, config);
                }
                else
                {
                    // Get notes using Excel Notes mode
                    noteNumbers = await _constructionNotesService.GetExcelNotesForSheetAsync(sheetName, config);
                }

                // Use multi-drawing service to handle any drawing state
                var sheetToNotes = new Dictionary<string, List<int>> { { sheetName, noteNumbers } };
                var constructionNotesResult = await _multiDrawingConstructionNotesService.UpdateConstructionNotesAcrossDrawingsAsync(
                    sheetToNotes, 
                    config, 
                    sheetInfos);

                if (constructionNotesResult.HasFailures)
                {
                    var failure = constructionNotesResult.Failures.FirstOrDefault();
                    _logger.LogWarning($"Construction notes update failed for sheet {sheetName}: {failure?.ErrorMessage}");
                }
                else
                {
                    _logger.LogDebug($"Successfully updated construction notes for sheet {sheetName}");
                }
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
                
                // Use multi-drawing service to handle any drawing state
                var titleBlockResult = await _multiDrawingTitleBlockService.UpdateTitleBlocksAcrossDrawingsAsync(
                    new List<string> { sheetName }, 
                    config, 
                    sheetInfos);

                if (titleBlockResult.HasFailures)
                {
                    var failure = titleBlockResult.Failures.FirstOrDefault();
                    _logger.LogWarning($"Title block update failed for sheet {sheetName}: {failure?.ErrorMessage}");
                }
                else
                {
                    _logger.LogDebug($"Successfully updated title blocks for sheet {sheetName}");
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

    /// <summary>
    /// Builds a collection of SheetInfo objects for batch plotting
    /// </summary>
    private async Task<List<SheetInfo>> BuildSheetCollectionAsync(List<string> sheetNames, ProjectConfiguration config)
    {
        var sheets = new List<SheetInfo>();
        var sheetInfoList = await _excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);

        foreach (var sheetName in sheetNames)
        {
            var sheetInfo = sheetInfoList.FirstOrDefault(s => s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheetInfo == null)
            {
                _logger.LogWarning($"Sheet info not found for {sheetName}, skipping");
                continue;
            }

            // Update the DWGFileName to include the full path
            var drawingPath = Path.Combine(config.ProjectDWGFilePath, sheetInfo.DWGFileName);
            if (!string.IsNullOrEmpty(sheetInfo.DWGFileName) && !sheetInfo.DWGFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                drawingPath += ".dwg";
            }

            // Clone the sheet info and update the path
            var sheetWithPath = new SheetInfo
            {
                SheetName = sheetInfo.SheetName,
                DWGFileName = drawingPath,
                DrawingTitle = sheetInfo.DrawingTitle,
                ProjectNumber = sheetInfo.ProjectNumber,
                Scale = sheetInfo.Scale,
                SheetType = sheetInfo.SheetType,
                IssueDate = sheetInfo.IssueDate,
                DesignedBy = sheetInfo.DesignedBy,
                CheckedBy = sheetInfo.CheckedBy,
                DrawnBy = sheetInfo.DrawnBy,
                AdditionalProperties = sheetInfo.AdditionalProperties
            };
            
            sheets.Add(sheetWithPath);
        }

        return sheets;
    }

    /// <summary>
    /// Plots sheets individually (fallback method or for single sheet operations)
    /// </summary>
    private async Task PlotSheetsIndividuallyAsync(
        List<string> sheetNames, 
        ProjectConfiguration config, 
        PlotJobSettings plotSettings, 
        IProgress<PlotProgress>? progressCallback, 
        PlotResult result)
    {
        var totalSheets = sheetNames.Count;
        var currentSheetIndex = 0;

        foreach (var sheetName in sheetNames)
        {
            currentSheetIndex++;
            
            try
            {
                // Report progress
                progressCallback?.Report(new PlotProgress
                {
                    CurrentSheet = sheetName,
                    CurrentOperation = "Plotting to PDF",
                    ProgressPercentage = (int)((double)currentSheetIndex / totalSheets * 100),
                    CompletedSheets = currentSheetIndex - 1,
                    TotalSheets = totalSheets
                });

                _logger.LogDebug($"Plotting sheet {sheetName} ({currentSheetIndex}/{totalSheets})");

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
                _logger.LogError($"Error plotting sheet {sheetName}: {ex.Message}", ex);
                result.FailedSheets.Add(new SheetPlotError
                {
                    SheetName = sheetName,
                    ErrorMessage = ex.Message,
                    ExceptionDetails = ex.ToString()
                });
            }
        }
    }

    /// <summary>
    /// Gets the current active layout name
    /// </summary>
    private string? GetCurrentLayoutName()
    {
        try
        {
            // Try to get the current layout from AutoCAD
            var layoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            return layoutManager?.CurrentLayout;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get current layout name: {ex.Message}", ex);
            return null;
        }
    }
}