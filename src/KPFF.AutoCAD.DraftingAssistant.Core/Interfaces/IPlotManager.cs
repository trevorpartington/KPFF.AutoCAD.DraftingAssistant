using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for AutoCAD-specific plotting operations
/// </summary>
public interface IPlotManager
{
    /// <summary>
    /// Plots a layout from a drawing file to PDF
    /// </summary>
    /// <param name="drawingPath">Full path to the drawing file</param>
    /// <param name="layoutName">Name of the layout to plot</param>
    /// <param name="outputPath">Full path for the output PDF file</param>
    /// <returns>True if plot succeeded, false otherwise</returns>
    Task<bool> PlotLayoutToPdfAsync(string drawingPath, string layoutName, string outputPath);

    /// <summary>
    /// Gets plot settings information for a layout without plotting
    /// </summary>
    /// <param name="drawingPath">Path to the drawing file</param>
    /// <param name="layoutName">Name of the layout</param>
    /// <returns>Plot settings info or null if failed</returns>
    Task<SheetPlotSettings?> GetLayoutPlotSettingsAsync(string drawingPath, string layoutName);
}