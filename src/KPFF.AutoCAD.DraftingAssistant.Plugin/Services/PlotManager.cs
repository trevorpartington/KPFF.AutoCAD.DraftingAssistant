using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Publishing;
using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific plotting manager that handles the actual plot operations using the Publisher API
/// </summary>
public class PlotManager : IPlotManager
{
    private readonly ILogger _logger;

    public PlotManager(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Plots multiple sheets to PDF using the Publisher API (recommended for batch operations)
    /// </summary>
    /// <param name="sheets">Collection of sheets to plot</param>
    /// <param name="outputDirectory">Directory for output PDF files</param>
    /// <returns>True if all plots succeeded, false otherwise</returns>
    public async Task<bool> PublishSheetsToPdfAsync(IEnumerable<SheetInfo> sheets, string outputDirectory)
    {
        await Task.Delay(1); // Make this truly async but execute synchronously

        try
        {
            var sheetList = sheets.ToList();
            if (!sheetList.Any())
            {
                _logger.LogWarning("No sheets provided for publishing");
                return false;
            }

            _logger.LogDebug($"Publishing {sheetList.Count} sheets to {outputDirectory}");

            // Ensure output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Build DSD entries for each sheet
            var dsdEntries = new List<DsdEntry>();
            foreach (var sheet in sheetList)
            {
                // Build the full path to the drawing file
                var drawingPath = Path.Combine(Path.GetDirectoryName(sheet.DWGFileName) ?? "", Path.GetFileName(sheet.DWGFileName));
                if (!Path.IsPathFullyQualified(drawingPath))
                {
                    // If DWGFileName is not a full path, we'll need the project path
                    // For now, assume it's a relative path from the current directory
                    drawingPath = Path.GetFullPath(sheet.DWGFileName);
                }
                
                if (!File.Exists(drawingPath))
                {
                    _logger.LogError($"Drawing file not found: {drawingPath}");
                    continue;
                }

                var entry = new DsdEntry
                {
                    DwgName = drawingPath,
                    Layout = sheet.SheetName,
                    Title = sheet.SheetName, // Use sheet name for unique PDF filenames
                    Nps = string.Empty, // Use layout's saved page setup
                    NpsSourceDwg = string.Empty
                };
                
                dsdEntries.Add(entry);
                _logger.LogDebug($"Added DSD entry: {drawingPath} - {sheet.SheetName}");
            }

            if (!dsdEntries.Any())
            {
                _logger.LogError("No valid sheets found for publishing");
                return false;
            }

            // Create DSD data for publishing
            using (var dsdData = new DsdData())
            {
                dsdData.ProjectPath = outputDirectory;
                dsdData.SheetType = SheetType.SinglePdf; // Individual PDFs per sheet
                dsdData.NoOfCopies = 1;
                dsdData.LogFilePath = Path.Combine(outputDirectory, "PlotLog.txt");
                
                // Set DSD entries
                using (var dsdCollection = new DsdEntryCollection())
                {
                    foreach (var entry in dsdEntries)
                    {
                        dsdCollection.Add(entry);
                    }
                    dsdData.SetDsdEntryCollection(dsdCollection);
                }

                // Set publishing options
                dsdData.SetUnrecognizedData("PromptForName", "FALSE");
                dsdData.SetUnrecognizedData("IncludeLayer", "TRUE");
                dsdData.SetUnrecognizedData("ShowPlotProgress", "FALSE");

                // Set background plotting to false for better control
                object oldBackgroundPlot = Application.GetSystemVariable("BACKGROUNDPLOT");
                
                try
                {
                    Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                    _logger.LogDebug("Set BACKGROUNDPLOT to 0");

                    // Get PDF plot configuration
                    PlotConfig plotConfig = PlotConfigManager.SetCurrentConfig("DWG to PDF.pc3");
                    _logger.LogDebug("Using plot config: DWG to PDF.pc3");

                    // Execute the publish operation
                    _logger.LogDebug("Starting Publisher.PublishExecute...");
                    Application.Publisher.PublishExecute(dsdData, plotConfig);
                    
                    _logger.LogInformation($"Successfully published {dsdEntries.Count} sheets to {outputDirectory}");
                    return true;
                }
                catch (System.Exception publishEx)
                {
                    _logger.LogError($"Publisher execution failed: {publishEx.Message}", publishEx);
                    return false;
                }
                finally
                {
                    // Restore original BACKGROUNDPLOT setting
                    Application.SetSystemVariable("BACKGROUNDPLOT", oldBackgroundPlot);
                    _logger.LogDebug("Restored original BACKGROUNDPLOT setting");
                }
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Publish operation failed: {ex.Message}", ex);
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
                        _logger.LogError($"Layout '{layoutName}' not found in drawing {drawingPath}");
                        return null;
                    }

                    var layoutId = layoutDict.GetAt(layoutName);
                    var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    
                    if (layout == null)
                    {
                        _logger.LogError($"Could not open layout '{layoutName}' from drawing {drawingPath}");
                        return null;
                    }

                    // Extract plot settings information
                    var settings = new SheetPlotSettings
                    {
                        DrawingPath = drawingPath,
                        LayoutName = layoutName,
                        PlotDevice = layout.PlotConfigurationName ?? "None",
                        PaperSize = layout.CanonicalMediaName ?? "Auto",
                        PlotScale = GetPlotScaleString(layout),
                        PlotArea = GetPlotAreaString(layout.PlotType),
                        PlotCentered = layout.PlotCentered
                    };

                    transaction.Commit();
                    return settings;
                }
            }
            finally
            {
                // If we opened a drawing, dispose it
                if (!wasDrawingAlreadyOpen && targetDatabase != null)
                {
                    targetDatabase.Dispose();
                }
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Failed to read plot settings for '{layoutName}' in '{drawingPath}': {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Helper method to find an already open document
    /// </summary>
    private static Document? GetOpenDocument(string drawingPath)
    {
        string normalizedPath = Path.GetFullPath(drawingPath);
        
        foreach (Document doc in Application.DocumentManager)
        {
            if (string.Equals(Path.GetFullPath(doc.Name), normalizedPath, StringComparison.OrdinalIgnoreCase))
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