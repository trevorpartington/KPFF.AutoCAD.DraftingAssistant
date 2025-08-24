using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Orchestrates automatic detection of construction notes from drawing viewports.
/// Analyzes viewports to find multileaders and extracts note numbers.
/// </summary>
public class AutoNotesService
{
    private readonly ILogger _logger;
    private readonly MultileaderAnalyzer _multileaderAnalyzer;

    public AutoNotesService(ILogger logger)
    {
        _logger = logger;
        _multileaderAnalyzer = new MultileaderAnalyzer(logger);
    }

    /// <summary>
    /// Gets construction note numbers automatically detected from a sheet's viewports.
    /// </summary>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="config">Project configuration containing multileader style settings</param>
    /// <returns>List of unique note numbers found in the sheet's viewports</returns>
    public async Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        return await Task.Run(() => GetAutoNotesForSheet(sheetName, config));
    }

    /// <summary>
    /// Synchronous implementation of auto notes detection.
    /// </summary>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="config">Project configuration containing multileader style settings</param>
    /// <returns>List of unique note numbers found in the sheet's viewports</returns>
    public List<int> GetAutoNotesForSheet(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"Starting auto notes detection for sheet '{sheetName}'");

            // Get current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (var transaction = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Find the layout
                    var layout = FindLayout(db, transaction, sheetName);
                    if (layout == null)
                    {
                        _logger.LogWarning($"Layout '{sheetName}' not found");
                        return new List<int>();
                    }

                    // Get all multileaders in model space
                    string? targetStyle = config.ConstructionNotes?.MultileaderStyleName;
                    var allMultileaders = _multileaderAnalyzer.FindMultileadersInModelSpace(db, transaction, targetStyle);
                    
                    if (allMultileaders.Count == 0)
                    {
                        _logger.LogInformation($"No multileaders found in model space for style '{targetStyle ?? "any"}'");
                        return new List<int>();
                    }

                    // Analyze viewports in the layout
                    var notesFromViewports = AnalyzeViewportsInLayout(layout, transaction, allMultileaders);

                    // Consolidate and return unique note numbers
                    var uniqueNotes = _multileaderAnalyzer.ConsolidateNoteNumbers(notesFromViewports);
                    
                    transaction.Commit();
                    
                    _logger.LogInformation($"Auto notes detection completed for sheet '{sheetName}': found {uniqueNotes.Count} notes");
                    return uniqueNotes;
                }
                catch (Exception ex)
                {
                    transaction.Abort();
                    _logger.LogError($"Error during auto notes detection for sheet '{sheetName}': {ex.Message}", ex);
                    return new List<int>();
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get auto notes for sheet '{sheetName}': {ex.Message}", ex);
            return new List<int>();
        }
    }

    /// <summary>
    /// Finds a layout by name in the database.
    /// </summary>
    /// <param name="database">The drawing database</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="layoutName">Name of the layout to find</param>
    /// <returns>Layout object if found, null otherwise</returns>
    private Layout? FindLayout(Database database, Transaction transaction, string layoutName)
    {
        try
        {
            var layoutDict = transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            if (layoutDict == null)
            {
                _logger.LogWarning("Could not access layout dictionary");
                return null;
            }

            if (layoutDict.Contains(layoutName))
            {
                var layoutId = layoutDict.GetAt(layoutName);
                var layout = transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
                _logger.LogDebug($"Found layout '{layoutName}'");
                return layout;
            }
            else
            {
                _logger.LogWarning($"Layout '{layoutName}' not found in drawing");
                return null;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error finding layout '{layoutName}': {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Analyzes all viewports in a layout to find multileaders within their boundaries.
    /// </summary>
    /// <param name="layout">The layout to analyze</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="allMultileaders">All multileaders found in model space</param>
    /// <returns>List of multileaders found within viewport boundaries</returns>
    private List<MultileaderAnalyzer.MultileaderInfo> AnalyzeViewportsInLayout(Layout layout, Transaction transaction, List<MultileaderAnalyzer.MultileaderInfo> allMultileaders)
    {
        var notesFromViewports = new List<MultileaderAnalyzer.MultileaderInfo>();

        try
        {
            // Get the layout's block table record
            var layoutBlock = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
            if (layoutBlock == null)
            {
                _logger.LogWarning($"Could not access block table record for layout '{layout.LayoutName}'");
                return notesFromViewports;
            }

            int viewportCount = 0;
            int processedViewports = 0;
            int viewportIndex = 0;

            // Find all viewports in the layout
            foreach (ObjectId entityId in layoutBlock)
            {
                var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                
                if (entity is Viewport viewport)
                {
                    viewportIndex++;
                    
                    // Skip the first viewport (paper space viewport)
                    if (viewportIndex == 1)
                    {
                        _logger.LogDebug("Skipping paper space viewport (index 1)");
                        continue;
                    }
                    
                    viewportCount++;
                    
                    try
                    {
                        var notesInViewport = AnalyzeSingleViewport(viewport, transaction, allMultileaders);
                        notesFromViewports.AddRange(notesInViewport);
                        processedViewports++;
                        
                        // Use viewportIndex instead of viewport.Number for logging (Number may be 0 for inactive layouts)
                        _logger.LogDebug($"Viewport #{viewportIndex} (Number={viewport.Number}): found {notesInViewport.Count} multileaders");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Error analyzing viewport #{viewportIndex} (Number={viewport.Number}): {ex.Message}");
                    }
                }
            }

            _logger.LogInformation($"Analyzed {processedViewports} of {viewportCount} viewports in layout '{layout.LayoutName}'");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error analyzing viewports in layout '{layout.LayoutName}': {ex.Message}", ex);
        }

        return notesFromViewports;
    }

    /// <summary>
    /// Analyzes a single viewport to find multileaders within its boundaries.
    /// </summary>
    /// <param name="viewport">The viewport to analyze</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="allMultileaders">All multileaders found in model space</param>
    /// <returns>List of multileaders found within this viewport's boundary</returns>
    private List<MultileaderAnalyzer.MultileaderInfo> AnalyzeSingleViewport(Viewport viewport, Transaction transaction, List<MultileaderAnalyzer.MultileaderInfo> allMultileaders)
    {
        try
        {
            // Get viewport boundary in model space coordinates
            var viewportBoundary = ViewportBoundaryCalculator.GetViewportFootprint(viewport, transaction);
            
            if (viewportBoundary.Count == 0)
            {
                _logger.LogWarning($"Could not calculate boundary for viewport {viewport.Number}");
                return new List<MultileaderAnalyzer.MultileaderInfo>();
            }

            _logger.LogDebug($"Viewport {viewport.Number} boundary: {viewportBoundary.Count} points");
            
            // Log boundary points for debugging
            for (int i = 0; i < Math.Min(viewportBoundary.Count, 4); i++)
            {
                _logger.LogDebug($"  Point {i}: {viewportBoundary[i]}");
            }

            // Filter multileaders to those within this viewport
            var multileadersInViewport = _multileaderAnalyzer.FilterMultileadersInViewport(allMultileaders, viewportBoundary);
            
            _logger.LogDebug($"Viewport {viewport.Number}: found {multileadersInViewport.Count} multileaders inside boundary");
            
            return multileadersInViewport;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error analyzing viewport {viewport.Number}: {ex.Message}", ex);
            return new List<MultileaderAnalyzer.MultileaderInfo>();
        }
    }

    /// <summary>
    /// Provides diagnostic information about multileaders and viewports for debugging.
    /// </summary>
    /// <param name="sheetName">Name of the sheet to diagnose</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Diagnostic information string</returns>
    public string GetDiagnosticInfo(string sheetName, ProjectConfiguration config)
    {
        try
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (var transaction = db.TransactionManager.StartTransaction())
            {
                var layout = FindLayout(db, transaction, sheetName);
                if (layout == null)
                {
                    return $"Layout '{sheetName}' not found";
                }

                // Get multileader information
                string? targetStyle = config.ConstructionNotes?.MultileaderStyleName;
                var allMultileaders = _multileaderAnalyzer.FindMultileadersInModelSpace(db, transaction, targetStyle);

                // Get viewport information with detailed boundary data
                var layoutBlock = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                int viewportCount = 0;
                int viewportIndex = 0;
                var viewportDetails = new List<string>();
                
                if (layoutBlock != null)
                {
                    foreach (ObjectId entityId in layoutBlock)
                    {
                        var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                        if (entity is Viewport viewport)
                        {
                            viewportIndex++;
                            
                            // Skip the first viewport (paper space viewport)
                            if (viewportIndex == 1)
                            {
                                continue;
                            }
                            
                            viewportCount++;
                            
                            try
                            {
                                // Get viewport boundary in model space coordinates
                                var viewportBoundary = ViewportBoundaryCalculator.GetViewportFootprint(viewport, transaction);
                                
                                var boundaryType = viewport.NonRectClipOn ? "Polygonal" : "Rectangular";
                                var scaleRatio = viewport.CustomScale > 0 ? $"1:{(1.0 / viewport.CustomScale):F0}" : "Unknown";
                                
                                // Use viewportIndex for display (more reliable than viewport.Number for inactive layouts)
                                var viewportInfo = new List<string>
                                {
                                    $"  Viewport #{viewportIndex} Boundary (Model Space):",
                                    $"    AutoCAD Number: {viewport.Number} (may be 0 for inactive layouts)",
                                    $"    Type: {boundaryType}",
                                    $"    Scale: {scaleRatio}",
                                    $"    Paper Size: {viewport.Width:F1} × {viewport.Height:F1}",
                                    $"    View Center: ({viewport.ViewCenter.X:F3}, {viewport.ViewCenter.Y:F3})"
                                };
                                
                                if (viewportBoundary.Count > 0)
                                {
                                    viewportInfo.Add($"    Boundary Points ({viewportBoundary.Count}):");
                                    for (int i = 0; i < viewportBoundary.Count; i++)
                                    {
                                        var pt = viewportBoundary[i];
                                        viewportInfo.Add($"      Point {i}: ({pt.X:F3}, {pt.Y:F3})");
                                    }
                                }
                                else
                                {
                                    viewportInfo.Add("    ❌ Could not calculate boundary points");
                                }
                                
                                viewportDetails.Add(string.Join("\n", viewportInfo));
                            }
                            catch (Exception ex)
                            {
                                viewportDetails.Add($"  Viewport #{viewportIndex}: ❌ Error calculating boundary - {ex.Message}");
                            }
                        }
                    }
                }

                transaction.Commit();

                var result = new List<string>
                {
                    $"Auto Notes Diagnostic for '{sheetName}':",
                    $"  Target Style: {targetStyle ?? "any"}",
                    $"  Multileaders Found: {allMultileaders.Count}",
                    $"  Viewports Found: {viewportCount}",
                    ""
                };
                
                if (viewportDetails.Count > 0)
                {
                    result.AddRange(viewportDetails);
                    result.Add("");
                }
                
                if (allMultileaders.Count > 0)
                {
                    result.Add("  Multileader Details:");
                    result.AddRange(allMultileaders.Select(m => $"    {m}"));
                }
                else
                {
                    result.Add("  No multileaders found in model space");
                }

                return string.Join("\n", result);
            }
        }
        catch (Exception ex)
        {
            return $"Error getting diagnostic info: {ex.Message}";
        }
    }
}