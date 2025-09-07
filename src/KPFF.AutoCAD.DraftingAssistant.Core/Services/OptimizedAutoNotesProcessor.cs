using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System.Diagnostics;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Optimized Auto Notes processor that eliminates duplicate model space scanning through intelligent caching.
/// Consolidates logic for active, inactive, and closed drawings with comprehensive performance monitoring.
/// </summary>
public class OptimizedAutoNotesProcessor
{
    private readonly ILogger _logger;
    private readonly ModelSpaceCacheService _modelSpaceCache;
    private readonly ViewportBoundaryCacheService _viewportBoundaryCache;
    private readonly MultileaderAnalyzer _multileaderAnalyzer;
    private readonly BlockAnalyzer _blockAnalyzer;

    public OptimizedAutoNotesProcessor(ILogger logger)
    {
        _logger = logger;
        _modelSpaceCache = new ModelSpaceCacheService(logger);
        _viewportBoundaryCache = new ViewportBoundaryCacheService(logger);
        _multileaderAnalyzer = new MultileaderAnalyzer(logger);
        _blockAnalyzer = new BlockAnalyzer(logger);
    }

    /// <summary>
    /// Performance metrics for auto notes processing
    /// </summary>
    public class PerformanceMetrics
    {
        public string DrawingPath { get; set; } = string.Empty;
        public string DrawingState { get; set; } = string.Empty;
        public int SheetsProcessed { get; set; }
        public TimeSpan TotalProcessingTime { get; set; }
        public TimeSpan ModelSpaceScanTime { get; set; }
        public TimeSpan ViewportAnalysisTime { get; set; }
        public bool ModelSpaceCacheHit { get; set; }
        public int ViewportCacheHits { get; set; }
        public int ViewportCacheMisses { get; set; }
        public int TotalMultileaders { get; set; }
        public int TotalBlocks { get; set; }
        public int TotalViewports { get; set; }
        public int TotalNotesFound { get; set; }
        public Dictionary<string, List<int>> SheetResults { get; set; } = new();

        public double ProcessingEfficiency => TotalProcessingTime.TotalMilliseconds > 0 
            ? (double)TotalNotesFound / TotalProcessingTime.TotalMilliseconds * 1000 
            : 0;

        public override string ToString()
        {
            var cacheEfficiency = ViewportCacheHits + ViewportCacheMisses > 0 
                ? $"{(double)ViewportCacheHits / (ViewportCacheHits + ViewportCacheMisses) * 100:F1}%" 
                : "N/A";
                
            return $"Performance[{Path.GetFileName(DrawingPath)}: {SheetsProcessed} sheets, " +
                   $"{TotalProcessingTime.TotalMilliseconds:F0}ms total, " +
                   $"ModelSpace: {(ModelSpaceCacheHit ? "HIT" : "MISS")}, " +
                   $"Viewport: {cacheEfficiency}, " +
                   $"Efficiency: {ProcessingEfficiency:F1} notes/sec]";
        }
    }

    /// <summary>
    /// Gets auto notes for multiple sheets from the same drawing with optimal performance
    /// </summary>
    /// <param name="drawingPath">Full path to the drawing file</param>
    /// <param name="sheetNames">List of sheet names to process</param>
    /// <param name="multileaderStyleNames">Multileader styles to filter</param>
    /// <param name="blockConfigurations">Block configurations to detect</param>
    /// <returns>Performance metrics with results for each sheet</returns>
    public PerformanceMetrics GetAutoNotesForMultipleSheets(
        string drawingPath,
        List<string> sheetNames,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        var overallStopwatch = Stopwatch.StartNew();
        var metrics = new PerformanceMetrics
        {
            DrawingPath = drawingPath,
            SheetsProcessed = sheetNames.Count
        };

        try
        {
            _logger.LogInformation($"=== OPTIMIZED AUTO NOTES PROCESSING ===");
            _logger.LogInformation($"Drawing: {Path.GetFileName(drawingPath)}");
            _logger.LogInformation($"Sheets: {sheetNames.Count} ({string.Join(", ", sheetNames)})");
            _logger.LogInformation($"Multileader Styles: {(multileaderStyleNames?.Count > 0 ? string.Join(", ", multileaderStyleNames) : "any")}");
            _logger.LogInformation($"Block Configs: {(blockConfigurations?.Count > 0 ? string.Join(", ", blockConfigurations.Select(bc => $"{bc.BlockName}→{bc.AttributeName}")) : "none")}");

            // Determine drawing state and get appropriate database/transaction
            var drawingState = DetermineDrawingState(drawingPath);
            metrics.DrawingState = drawingState.ToString();

            switch (drawingState)
            {
                case DrawingState.Active:
                    ProcessActiveDrawing(metrics, sheetNames, multileaderStyleNames, blockConfigurations);
                    break;
                
                case DrawingState.Inactive:
                    ProcessInactiveDrawing(drawingPath, metrics, sheetNames, multileaderStyleNames, blockConfigurations);
                    break;
                
                case DrawingState.Closed:
                    ProcessClosedDrawing(drawingPath, metrics, sheetNames, multileaderStyleNames, blockConfigurations);
                    break;
                
                default:
                    throw new InvalidOperationException($"Unable to process drawing in state: {drawingState}");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error in optimized auto notes processing: {ex.Message}", ex);
            throw;
        }
        finally
        {
            overallStopwatch.Stop();
            metrics.TotalProcessingTime = overallStopwatch.Elapsed;
            
            _logger.LogInformation($"=== PROCESSING COMPLETE ===");
            _logger.LogInformation(metrics.ToString());
            LogCacheStatistics();
        }

        return metrics;
    }

    /// <summary>
    /// Gets auto notes for a single sheet with optimal performance
    /// </summary>
    public List<int> GetAutoNotesForSheet(
        string drawingPath,
        string sheetName,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        var metrics = GetAutoNotesForMultipleSheets(
            drawingPath, 
            new List<string> { sheetName }, 
            multileaderStyleNames, 
            blockConfigurations);

        return metrics.SheetResults.GetValueOrDefault(sheetName, new List<int>());
    }

    /// <summary>
    /// Processes an active drawing (currently active document)
    /// </summary>
    private void ProcessActiveDrawing(
        PerformanceMetrics metrics,
        List<string> sheetNames,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        using (var transaction = db.TransactionManager.StartTransaction())
        {
            try
            {
                ProcessDrawingWithTransaction(
                    metrics,
                    db,
                    transaction,
                    sheetNames,
                    multileaderStyleNames,
                    blockConfigurations);
                
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Abort();
                _logger.LogError($"Transaction aborted for active drawing: {ex.Message}", ex);
                throw;
            }
        }
    }

    /// <summary>
    /// Processes an inactive drawing (open but not active)
    /// </summary>
    private void ProcessInactiveDrawing(
        string drawingPath,
        PerformanceMetrics metrics,
        List<string> sheetNames,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        // Find the document for this drawing path
        Document? targetDoc = null;
        
        foreach (Document doc in Application.DocumentManager)
        {
            if (string.Equals(doc.Name, drawingPath, StringComparison.OrdinalIgnoreCase))
            {
                targetDoc = doc;
                break;
            }
        }

        if (targetDoc == null)
        {
            throw new InvalidOperationException($"Cannot find open document for path: {drawingPath}");
        }

        using (targetDoc.LockDocument())
        {
            var db = targetDoc.Database;
            
            using (var transaction = db.TransactionManager.StartTransaction())
            {
                try
                {
                    ProcessDrawingWithTransaction(
                        metrics,
                        db,
                        transaction,
                        sheetNames,
                        multileaderStyleNames,
                        blockConfigurations);
                    
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Abort();
                    _logger.LogError($"Transaction aborted for inactive drawing: {ex.Message}", ex);
                    throw;
                }
            }
        }
    }

    /// <summary>
    /// Processes a closed drawing using external database
    /// </summary>
    private void ProcessClosedDrawing(
        string drawingPath,
        PerformanceMetrics metrics,
        List<string> sheetNames,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        using (var db = new Database(false, true)) // buildDefaultDrawing=false, noDocument=true
        {
            _logger.LogDebug($"Reading closed DWG file: {Path.GetFileName(drawingPath)}");
            db.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndAllShare, true, null);
            
            using (var transaction = db.TransactionManager.StartTransaction())
            {
                try
                {
                    ProcessDrawingWithTransaction(
                        metrics,
                        db,
                        transaction,
                        sheetNames,
                        multileaderStyleNames,
                        blockConfigurations);
                    
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Abort();
                    _logger.LogError($"Transaction aborted for closed drawing: {ex.Message}", ex);
                    throw;
                }
            }
        }
    }

    /// <summary>
    /// Core processing logic that works with any database/transaction combination
    /// </summary>
    private void ProcessDrawingWithTransaction(
        PerformanceMetrics metrics,
        Database database,
        Transaction transaction,
        List<string> sheetNames,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        // STEP 1: Get or scan model space data (CACHED - happens once per drawing)
        var modelSpaceStopwatch = Stopwatch.StartNew();
        
        var cacheEntry = _modelSpaceCache.GetOrScanModelSpace(
            metrics.DrawingPath,
            database,
            transaction,
            multileaderStyleNames,
            blockConfigurations);

        modelSpaceStopwatch.Stop();
        metrics.ModelSpaceScanTime = modelSpaceStopwatch.Elapsed;
        metrics.ModelSpaceCacheHit = modelSpaceStopwatch.Elapsed.TotalMilliseconds < 50; // Heuristic: cache hits are very fast
        metrics.TotalMultileaders = cacheEntry.Multileaders.Count;
        metrics.TotalBlocks = cacheEntry.Blocks.Count;

        _logger.LogInformation($"Model space data: {cacheEntry.Multileaders.Count} multileaders, {cacheEntry.Blocks.Count} blocks (Cache: {(metrics.ModelSpaceCacheHit ? "HIT" : "MISS")})");

        if (cacheEntry.Multileaders.Count == 0 && cacheEntry.Blocks.Count == 0)
        {
            _logger.LogInformation("No multileaders or blocks found - all sheets will have empty results");
            foreach (var sheetName in sheetNames)
            {
                metrics.SheetResults[sheetName] = new List<int>();
            }
            return;
        }

        // STEP 2: Process each sheet using cached model space data
        var viewportStopwatch = Stopwatch.StartNew();

        foreach (var sheetName in sheetNames)
        {
            try
            {
                var sheetNotes = ProcessSingleSheet(
                    database,
                    transaction,
                    sheetName,
                    cacheEntry.Multileaders,
                    cacheEntry.Blocks,
                    metrics);

                metrics.SheetResults[sheetName] = sheetNotes;
                metrics.TotalNotesFound += sheetNotes.Count;

                _logger.LogInformation($"Sheet '{sheetName}': {sheetNotes.Count} notes found: [{string.Join(", ", sheetNotes)}]");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error processing sheet '{sheetName}': {ex.Message}", ex);
                metrics.SheetResults[sheetName] = new List<int>();
            }
        }

        viewportStopwatch.Stop();
        metrics.ViewportAnalysisTime = viewportStopwatch.Elapsed;
    }

    /// <summary>
    /// Processes a single sheet using pre-scanned model space data
    /// </summary>
    private List<int> ProcessSingleSheet(
        Database database,
        Transaction transaction,
        string sheetName,
        List<MultileaderAnalyzer.MultileaderInfo> allMultileaders,
        List<BlockAnalyzer.BlockInfo> allBlocks,
        PerformanceMetrics metrics)
    {
        var noteNumbers = new List<int>();

        try
        {
            // Find the layout
            var layout = FindLayout(database, transaction, sheetName);
            if (layout == null)
            {
                _logger.LogWarning($"Layout '{sheetName}' not found");
                return noteNumbers;
            }

            // Get the layout's block table record  
            var layoutBlock = transaction.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
            if (layoutBlock == null)
            {
                _logger.LogWarning($"Could not access block table record for layout '{sheetName}'");
                return noteNumbers;
            }

            // Find all viewports in the layout (excluding paper space viewport)
            var viewports = FindViewportsInLayout(layoutBlock, transaction);
            metrics.TotalViewports += viewports.Count;

            if (viewports.Count == 0)
            {
                _logger.LogDebug($"No viewports found in layout '{sheetName}'");
                return noteNumbers;
            }

            _logger.LogDebug($"Processing {viewports.Count} viewports in layout '{sheetName}'");

            // Process each viewport
            foreach (var viewportId in viewports)
            {
                try
                {
                    var viewport = transaction.GetObject(viewportId, OpenMode.ForRead) as Viewport;
                    if (viewport == null) continue;

                    // Get viewport boundary (CACHED)
                    var boundary = _viewportBoundaryCache.GetOrCalculateViewportBoundary(
                        viewport, 
                        transaction, 
                        sheetName);

                    // Update cache statistics
                    if (boundary.Count > 0)
                    {
                        // Heuristic: if boundary calculation was very fast, it was likely a cache hit
                        var wasRecentlyCached = _viewportBoundaryCache.GetCacheStatistics()
                            .NewestCacheTime > DateTime.Now.AddSeconds(-1);
                        
                        if (wasRecentlyCached)
                        {
                            metrics.ViewportCacheHits++;
                        }
                        else
                        {
                            metrics.ViewportCacheMisses++;
                        }

                        // Filter entities within viewport boundaries
                        var viewportNotes = FilterEntitiesInViewport(
                            allMultileaders,
                            allBlocks,
                            boundary,
                            viewport,
                            transaction);

                        noteNumbers.AddRange(viewportNotes);
                    }
                    else
                    {
                        _logger.LogWarning($"Could not calculate boundary for viewport {viewport.Number} in '{sheetName}'");
                        metrics.ViewportCacheMisses++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Error processing viewport in sheet '{sheetName}': {ex.Message}");
                }
            }

            // Remove duplicates and sort
            noteNumbers = noteNumbers.Distinct().OrderBy(n => n).ToList();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error processing sheet '{sheetName}': {ex.Message}", ex);
        }

        return noteNumbers;
    }

    /// <summary>
    /// Finds all viewports in a layout (excluding paper space viewport)
    /// </summary>
    private List<ObjectId> FindViewportsInLayout(BlockTableRecord layoutBtr, Transaction transaction)
    {
        var viewports = new List<ObjectId>();
        int viewportIndex = 0;

        foreach (ObjectId entityId in layoutBtr)
        {
            var entity = transaction.GetObject(entityId, OpenMode.ForRead);
            
            if (entity is Viewport)
            {
                viewportIndex++;
                
                // Skip the first viewport (paper space viewport)
                if (viewportIndex > 1)
                {
                    viewports.Add(entityId);
                }
            }
        }

        return viewports;
    }

    /// <summary>
    /// Filters multileaders and blocks to those within viewport boundaries
    /// </summary>
    private List<int> FilterEntitiesInViewport(
        List<MultileaderAnalyzer.MultileaderInfo> allMultileaders,
        List<BlockAnalyzer.BlockInfo> allBlocks,
        Point3dCollection viewportBoundary,
        Viewport viewport,
        Transaction transaction)
    {
        var noteNumbers = new List<int>();

        try
        {
            // Filter multileaders
            var filteredMultileaders = _multileaderAnalyzer.FilterMultileadersInViewport(
                allMultileaders, 
                viewportBoundary, 
                viewport, 
                transaction);
                
            var multileaderNotes = _multileaderAnalyzer.ConsolidateNoteNumbers(filteredMultileaders);
            
            // Filter blocks  
            var filteredBlocks = _blockAnalyzer.FilterBlocksInViewport(
                allBlocks, 
                viewportBoundary, 
                viewport, 
                transaction);
                
            var blockNotes = _blockAnalyzer.ConsolidateNoteNumbers(filteredBlocks);

            // Combine results
            noteNumbers.AddRange(multileaderNotes);
            noteNumbers.AddRange(blockNotes);

            _logger.LogDebug($"Viewport filter results: {filteredMultileaders.Count} multileaders, {filteredBlocks.Count} blocks → {noteNumbers.Count} notes");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error filtering entities in viewport: {ex.Message}");
        }

        return noteNumbers;
    }

    /// <summary>
    /// Finds a layout by name in the database
    /// </summary>
    private Layout? FindLayout(Database database, Transaction transaction, string layoutName)
    {
        try
        {
            var layoutDict = transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            if (layoutDict == null || !layoutDict.Contains(layoutName))
            {
                return null;
            }

            var layoutId = layoutDict.GetAt(layoutName);
            return transaction.GetObject(layoutId, OpenMode.ForRead) as Layout;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error finding layout '{layoutName}': {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Determines the state of a drawing (active, inactive, or closed)
    /// </summary>
    private DrawingState DetermineDrawingState(string drawingPath)
    {
        try
        {
            var activeDoc = Application.DocumentManager.MdiActiveDocument;
            if (activeDoc != null && string.Equals(activeDoc.Name, drawingPath, StringComparison.OrdinalIgnoreCase))
            {
                return DrawingState.Active;
            }

            // Check if drawing is open but not active
            foreach (Document doc in Application.DocumentManager)
            {
                if (string.Equals(doc.Name, drawingPath, StringComparison.OrdinalIgnoreCase))
                {
                    return DrawingState.Inactive;
                }
            }

            return DrawingState.Closed;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error determining drawing state for '{drawingPath}': {ex.Message}");
            return DrawingState.Closed; // Default to closed if we can't determine
        }
    }

    /// <summary>
    /// Logs cache statistics for performance monitoring
    /// </summary>
    private void LogCacheStatistics()
    {
        try
        {
            var modelSpaceStats = _modelSpaceCache.GetCacheStatistics();
            var viewportStats = _viewportBoundaryCache.GetCacheStatistics();

            _logger.LogInformation($"=== CACHE PERFORMANCE ===");
            _logger.LogInformation($"Model Space: {modelSpaceStats}");
            _logger.LogInformation($"Viewport Boundaries: {viewportStats}");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error logging cache statistics: {ex.Message}");
        }
    }

    /// <summary>
    /// Clears all caches (useful for testing or memory management)
    /// </summary>
    public void ClearAllCaches()
    {
        _modelSpaceCache.ClearCache();
        _viewportBoundaryCache.ClearCache();
        _logger.LogInformation("All caches cleared");
    }

    /// <summary>
    /// Invalidates cache for a specific drawing
    /// </summary>
    public void InvalidateDrawingCache(string drawingPath)
    {
        _modelSpaceCache.InvalidateDrawing(drawingPath);
        // Viewport cache is layout-specific, so we'd need layout names to invalidate
        _logger.LogInformation($"Model space cache invalidated for: {Path.GetFileName(drawingPath)}");
    }
}

