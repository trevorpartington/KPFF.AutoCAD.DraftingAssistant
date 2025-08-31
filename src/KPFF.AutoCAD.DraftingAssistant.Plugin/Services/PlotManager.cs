using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific plotting manager that handles the actual plot operations
/// </summary>
public class PlotManager : IPlotManager
{
    private readonly ILogger _logger;

    public PlotManager(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Plots a layout from a drawing file to PDF
    /// </summary>
    /// <param name="drawingPath">Full path to the drawing file</param>
    /// <param name="layoutName">Name of the layout to plot</param>
    /// <param name="outputPath">Full path for the output PDF file</param>
    /// <returns>True if plot succeeded, false otherwise</returns>
    public async Task<bool> PlotLayoutToPdfAsync(string drawingPath, string layoutName, string outputPath)
    {
        // AutoCAD API calls must be on the main thread, so we'll use Task.FromResult instead of Task.Run
        await Task.Delay(1); // Make this truly async but execute synchronously
        
        {
            try
            {
                _logger.LogDebug($"Starting plot operation for layout '{layoutName}' from '{drawingPath}' to '{outputPath}'");
                
                // Check if drawing file exists
                if (!File.Exists(drawingPath))
                {
                    _logger.LogError($"Drawing file not found: {drawingPath}");
                    return false;
                }
                _logger.LogDebug($"Drawing file exists: {drawingPath}");

                // Check if a plot is already in progress
                if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                {
                    _logger.LogError("Another plot operation is already in progress");
                    return false;
                }

                Database? targetDatabase = null;
                Document? targetDocument = null;
                bool wasDrawingAlreadyOpen = false;

                try
                {
                    _logger.LogDebug("Checking if drawing is already open...");
                    // Check if drawing is already open
                    targetDocument = GetOpenDocument(drawingPath);
                    if (targetDocument != null)
                    {
                        _logger.LogDebug($"Drawing {drawingPath} is already open");
                        wasDrawingAlreadyOpen = true;
                        targetDatabase = targetDocument.Database;
                    }
                    else
                    {
                        // Open the drawing
                        _logger.LogDebug($"Opening drawing {drawingPath} from disk...");
                        targetDatabase = new Database(false, true);
                        _logger.LogDebug("Created new database instance");
                        targetDatabase.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndAllShare, true, "");
                        _logger.LogDebug("Successfully read DWG file");
                    }

                    _logger.LogDebug("Starting transaction...");
                    using (var transaction = targetDatabase.TransactionManager.StartTransaction())
                    {
                        _logger.LogDebug("Transaction started successfully");
                        // Get the layout manager and find the specified layout
                        var layoutManager = LayoutManager.Current;
                        _logger.LogDebug("Getting layout dictionary...");
                        var layoutDict = transaction.GetObject(targetDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        _logger.LogDebug($"Layout dictionary retrieved: {layoutDict != null}");
                        
                        if (layoutDict == null)
                        {
                            _logger.LogError($"Layout dictionary is null in drawing {drawingPath}");
                            return false;
                        }
                        
                        if (!layoutDict.Contains(layoutName))
                        {
                            _logger.LogError($"Layout '{layoutName}' not found in drawing {drawingPath}");
                            
                            // Log available layouts for debugging
                            _logger.LogDebug("Available layouts:");
                            foreach (var entry in layoutDict)
                            {
                                _logger.LogDebug($"  - {entry.Key}");
                            }
                            return false;
                        }

                        var layoutId = layoutDict.GetAt(layoutName);
                        var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                        
                        if (layout == null)
                        {
                            _logger.LogError($"Could not open layout '{layoutName}' from drawing {drawingPath}");
                            return false;
                        }

                        // Create plot info
                        using (var plotInfo = new PlotInfo())
                        {
                            plotInfo.Layout = layout.ObjectId;

                            // Create plot settings based on the layout
                            using (var plotSettings = new PlotSettings(layout.ModelType))
                            {
                                plotSettings.CopyFrom(layout);

                                // Validate plot settings
                                var plotSettingsValidator = PlotSettingsValidator.Current;

                                // Set plot device AND media in the correct order (device + media together first)
                                try
                                {
                                    // Try common PDF device names with the layout's existing media
                                    var pdfDevices = new[] { "DWG To PDF.pc3", "PDF.pc3", "AutoCAD PDF (General Documentation).pc3" };
                                    var mediaName = layout.CanonicalMediaName;
                                    bool deviceSet = false;

                                    foreach (var device in pdfDevices)
                                    {
                                        try
                                        {
                                            // Set BOTH device and media in one call - this is the correct AutoCAD API pattern
                                            plotSettingsValidator.SetPlotConfigurationName(plotSettings, device, mediaName);
                                            deviceSet = true;
                                            _logger.LogDebug($"Using plot device: {device} with media: {mediaName ?? "default"}");
                                            break;
                                        }
                                        catch
                                        {
                                            // Try next device
                                            continue;
                                        }
                                    }

                                    if (!deviceSet)
                                    {
                                        _logger.LogWarning("No PDF plot device found, using layout's default device");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    _logger.LogWarning($"Could not set PDF plot device: {ex.Message}");
                                }

                                // AFTER setting device+media, now set plot type
                                try
                                {
                                    plotSettingsValidator.SetPlotType(plotSettings, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                                    _logger.LogDebug("Plot type set to Layout");
                                }
                                catch (System.Exception ex)
                                {
                                    _logger.LogWarning($"Could not set plot type: {ex.Message}");
                                }

                                // AFTER plot type, now set centering - this should work now
                                try
                                {
                                    plotSettingsValidator.SetPlotCentered(plotSettings, true);
                                    _logger.LogDebug("Plot centering enabled");
                                }
                                catch (System.Exception ex)
                                {
                                    _logger.LogWarning($"Could not set plot centering: {ex.Message}");
                                    // Continue without centering
                                }

                                // Override plot settings in plot info
                                plotInfo.OverrideSettings = plotSettings;

                                // Validate plot info
                                using (var plotInfoValidator = new PlotInfoValidator())
                                {
                                    plotInfoValidator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                    plotInfoValidator.Validate(plotInfo);

                                    // Create plot engine and execute plot
                                    _logger.LogDebug("Creating plot engine...");
                                    using (var plotEngine = PlotFactory.CreatePublishEngine())
                                    {
                                        try
                                        {
                                            // Set background plotting to false for better control
                                            Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                                            _logger.LogDebug("Set BACKGROUNDPLOT to 0");

                                            // Begin plotting
                                            _logger.LogDebug("Beginning plot...");
                                            plotEngine.BeginPlot(null, null);

                                            // Define plot output
                                            _logger.LogDebug($"Beginning document for output: {outputPath}");
                                            plotEngine.BeginDocument(plotInfo, 
                                                Path.GetFileNameWithoutExtension(drawingPath), 
                                                null, 1, true, outputPath);

                                            // Plot the page
                                            _logger.LogDebug("Plotting page...");
                                            using (var plotPageInfo = new PlotPageInfo())
                                            {
                                                plotEngine.BeginPage(plotPageInfo, plotInfo, true, null);
                                                plotEngine.BeginGenerateGraphics(null);
                                                plotEngine.EndGenerateGraphics(null);
                                                plotEngine.EndPage(null);
                                            }

                                            // End plotting
                                            _logger.LogDebug("Ending plot...");
                                            plotEngine.EndDocument(null);
                                            plotEngine.EndPlot(null);

                                            _logger.LogInformation($"Successfully plotted layout '{layoutName}' to '{outputPath}'");
                                            // Don't return here - let the transaction commit first
                                        }
                                        catch (System.Exception plotEx)
                                        {
                                            _logger.LogError($"Plot engine error: {plotEx.Message}", plotEx);
                                            throw;
                                        }
                                    }
                                }
                            }
                        }

                        transaction.Commit();
                        _logger.LogDebug("Transaction committed successfully");
                        // Plot succeeded and transaction committed
                        return true;
                    }
                }
                finally
                {
                    // Clean up - only dispose database if we opened it
                    if (!wasDrawingAlreadyOpen && targetDatabase != null)
                    {
                        targetDatabase.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Plot operation failed for layout '{layoutName}' in '{drawingPath}': {ex.Message}", ex);
                return false;
            }

            // This should not be reached in normal flow
            return false;
        }
    }

    /// <summary>
    /// Gets plot settings information for a layout without plotting
    /// </summary>
    /// <param name="drawingPath">Path to the drawing file</param>
    /// <param name="layoutName">Name of the layout</param>
    /// <returns>Plot settings info or null if failed</returns>
    public async Task<SheetPlotSettings?> GetLayoutPlotSettingsAsync(string drawingPath, string layoutName)
    {
        await Task.Delay(1); // Make this truly async but execute synchronously
        
        {
            try
            {
                _logger.LogDebug($"Reading plot settings for layout '{layoutName}' from '{drawingPath}'");

                Database? targetDatabase = null;
                Document? targetDocument = null;
                bool wasDrawingAlreadyOpen = false;

                try
                {
                    // Check if drawing is already open
                    targetDocument = GetOpenDocument(drawingPath);
                    if (targetDocument != null)
                    {
                        wasDrawingAlreadyOpen = true;
                        targetDatabase = targetDocument.Database;
                    }
                    else
                    {
                        // Open the drawing
                        targetDatabase = new Database(false, true);
                        targetDatabase.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndAllShare, true, "");
                    }

                    using (var transaction = targetDatabase.TransactionManager.StartTransaction())
                    {
                        // Get the layout
                        var layoutDict = transaction.GetObject(targetDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        
                        if (layoutDict == null || !layoutDict.Contains(layoutName))
                        {
                            _logger.LogWarning($"Layout '{layoutName}' not found in drawing {drawingPath}");
                            return null;
                        }

                        var layoutId = layoutDict.GetAt(layoutName);
                        var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                        
                        if (layout == null)
                        {
                            return null;
                        }

                        return new SheetPlotSettings
                        {
                            SheetName = layoutName,
                            LayoutName = layoutName,
                            DrawingPath = drawingPath,
                            PlotDevice = layout.PlotConfigurationName ?? "Default",
                            PaperSize = layout.CanonicalMediaName ?? "Unknown",
                            PlotScale = GetPlotScaleString(layout),
                            PlotArea = GetPlotAreaString(layout.PlotType),
                            PlotCentered = layout.PlotCentered
                        };
                    }
                }
                finally
                {
                    // Clean up - only dispose database if we opened it
                    if (!wasDrawingAlreadyOpen && targetDatabase != null)
                    {
                        targetDatabase.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to read plot settings for layout '{layoutName}' in '{drawingPath}': {ex.Message}", ex);
                return null;
            }
        }
    }

    /// <summary>
    /// Gets an open document by file path, if it exists
    /// </summary>
    private static Document? GetOpenDocument(string filePath)
    {
        var documents = Application.DocumentManager;
        foreach (Document doc in documents)
        {
            if (string.Equals(doc.Name, filePath, StringComparison.OrdinalIgnoreCase))
            {
                return doc;
            }
        }
        return null;
    }

    /// <summary>
    /// Converts plot scale to a readable string
    /// </summary>
    private static string GetPlotScaleString(Layout layout)
    {
        try
        {
            if (layout.UseStandardScale)
            {
                return layout.StdScaleType.ToString();
            }
            else
            {
                return $"{layout.CustomPrintScale.Numerator}:{layout.CustomPrintScale.Denominator}";
            }
        }
        catch
        {
            return "Unknown";
        }
    }

    /// <summary>
    /// Converts plot type to a readable string
    /// </summary>
    private static string GetPlotAreaString(Autodesk.AutoCAD.DatabaseServices.PlotType plotType)
    {
        return plotType switch
        {
            Autodesk.AutoCAD.DatabaseServices.PlotType.Layout => "Layout",
            Autodesk.AutoCAD.DatabaseServices.PlotType.Extents => "Extents",
            Autodesk.AutoCAD.DatabaseServices.PlotType.Limits => "Limits", 
            Autodesk.AutoCAD.DatabaseServices.PlotType.View => "View",
            Autodesk.AutoCAD.DatabaseServices.PlotType.Window => "Window",
            _ => "Unknown"
        };
    }
}