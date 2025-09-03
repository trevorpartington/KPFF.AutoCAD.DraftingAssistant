using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for plotting operations
/// Coordinates plotting with optional construction notes and title block updates
/// </summary>
public interface IPlottingService
{
    /// <summary>
    /// Plots the specified sheets with optional pre-plot updates
    /// </summary>
    /// <param name="sheetNames">List of sheet names to plot</param>
    /// <param name="config">Project configuration</param>
    /// <param name="plotSettings">Plot job settings including update options</param>
    /// <param name="progressCallback">Optional callback for progress updates</param>
    /// <returns>Plot result with success/failure information</returns>
    Task<PlotResult> PlotSheetsAsync(
        List<string> sheetNames, 
        ProjectConfiguration config,
        PlotJobSettings plotSettings,
        IProgress<PlotProgress>? progressCallback = null);

    /// <summary>
    /// Validates that sheets can be plotted and returns any issues
    /// </summary>
    /// <param name="sheetNames">List of sheet names to validate</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Validation result with any issues found</returns>
    Task<PlotValidationResult> ValidateSheetsForPlottingAsync(
        List<string> sheetNames,
        ProjectConfiguration config);

    /// <summary>
    /// Gets the default plot settings for a sheet based on its layout
    /// </summary>
    /// <param name="sheetName">Sheet name to get settings for</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Default plot settings for the sheet</returns>
    Task<SheetPlotSettings?> GetDefaultPlotSettingsAsync(
        string sheetName,
        ProjectConfiguration config);
}