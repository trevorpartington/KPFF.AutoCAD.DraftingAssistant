using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Optimized version of AutoNotesService that maintains backward compatibility while providing
/// significant performance improvements through caching and consolidated processing.
/// </summary>
public class OptimizedAutoNotesService
{
    private readonly ILogger _logger;
    private readonly OptimizedAutoNotesProcessor _processor;
    private static bool _useOptimizedProcessor = true; // Feature flag for rollback

    public OptimizedAutoNotesService(ILogger logger)
    {
        _logger = logger;
        _processor = new OptimizedAutoNotesProcessor(logger);
    }

    /// <summary>
    /// Gets construction note numbers automatically detected from a sheet's viewports.
    /// Now uses optimized processor with caching for improved performance.
    /// </summary>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="config">Project configuration containing multileader style settings</param>
    /// <returns>List of unique note numbers found in the sheet's viewports</returns>
    public async Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        return await Task.Run(() => GetAutoNotesForSheet(sheetName, config));
    }

    /// <summary>
    /// Synchronous implementation of auto notes detection with optimization.
    /// </summary>
    /// <param name="sheetName">Name of the sheet/layout to analyze</param>
    /// <param name="config">Project configuration containing multileader style settings</param>
    /// <returns>List of unique note numbers found in the sheet's viewports</returns>
    public List<int> GetAutoNotesForSheet(string sheetName, ProjectConfiguration config)
    {
        try
        {
            if (!_useOptimizedProcessor)
            {
                // Fallback to legacy implementation if needed
                return GetAutoNotesForSheetLegacy(sheetName, config);
            }

            _logger.LogInformation($"Starting optimized auto notes detection for sheet '{sheetName}'");

            // Get current document path for the optimized processor
            var doc = Application.DocumentManager.MdiActiveDocument;
            var drawingPath = doc?.Name ?? string.Empty;

            if (string.IsNullOrEmpty(drawingPath))
            {
                _logger.LogWarning("No active document found for auto notes detection");
                return new List<int>();
            }

            // Extract multileader styles and block configurations from config
            var multileaderStyles = config.ConstructionNotes?.MultileaderStyleNames;
            var blockConfigurations = config.ConstructionNotes?.NoteBlocks;

            // Use optimized processor
            var results = _processor.GetAutoNotesForSheet(
                drawingPath,
                sheetName,
                multileaderStyles,
                blockConfigurations);

            _logger.LogInformation($"Optimized auto notes detection completed for sheet '{sheetName}': found {results.Count} notes");
            return results;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get optimized auto notes for sheet '{sheetName}': {ex.Message}", ex);
            
            // Fallback to legacy implementation on error
            if (_useOptimizedProcessor)
            {
                _logger.LogInformation("Falling back to legacy implementation");
                return GetAutoNotesForSheetLegacy(sheetName, config);
            }
            
            return new List<int>();
        }
    }

    /// <summary>
    /// Legacy implementation for fallback scenarios or comparison testing
    /// </summary>
    private List<int> GetAutoNotesForSheetLegacy(string sheetName, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"Using legacy auto notes detection for sheet '{sheetName}'");

            // Use the original AutoNotesService implementation
            var legacyService = new AutoNotesService(_logger);
            return legacyService.GetAutoNotesForSheet(sheetName, config);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Legacy auto notes detection also failed for sheet '{sheetName}': {ex.Message}", ex);
            return new List<int>();
        }
    }

    /// <summary>
    /// Gets auto notes for multiple sheets with optimal batch processing.
    /// This method provides the biggest performance gains compared to calling single sheets repeatedly.
    /// </summary>
    /// <param name="sheetNames">List of sheet names to process</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Dictionary mapping sheet names to their auto notes</returns>
    public async Task<Dictionary<string, List<int>>> GetAutoNotesForMultipleSheetsAsync(
        List<string> sheetNames, 
        ProjectConfiguration config)
    {
        return await Task.Run(() => GetAutoNotesForMultipleSheets(sheetNames, config));
    }

    /// <summary>
    /// Synchronous batch processing for multiple sheets
    /// </summary>
    public Dictionary<string, List<int>> GetAutoNotesForMultipleSheets(
        List<string> sheetNames,
        ProjectConfiguration config)
    {
        try
        {
            if (!_useOptimizedProcessor || sheetNames.Count <= 1)
            {
                // Process individually for single sheets or if optimization is disabled
                var results = new Dictionary<string, List<int>>();
                foreach (var sheetName in sheetNames)
                {
                    results[sheetName] = GetAutoNotesForSheet(sheetName, config);
                }
                return results;
            }

            _logger.LogInformation($"Starting batch auto notes detection for {sheetNames.Count} sheets");

            // Get current document path
            var doc = Application.DocumentManager.MdiActiveDocument;
            var drawingPath = doc?.Name ?? string.Empty;

            if (string.IsNullOrEmpty(drawingPath))
            {
                _logger.LogWarning("No active document found for batch auto notes detection");
                return sheetNames.ToDictionary(name => name, _ => new List<int>());
            }

            // Extract configuration
            var multileaderStyles = config.ConstructionNotes?.MultileaderStyleNames;
            var blockConfigurations = config.ConstructionNotes?.NoteBlocks;

            // Use optimized batch processor
            var metrics = _processor.GetAutoNotesForMultipleSheets(
                drawingPath,
                sheetNames,
                multileaderStyles,
                blockConfigurations);

            _logger.LogInformation($"Batch auto notes detection completed: {metrics}");
            return metrics.SheetResults;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to get batch auto notes: {ex.Message}", ex);
            
            // Fallback to individual processing
            var results = new Dictionary<string, List<int>>();
            foreach (var sheetName in sheetNames)
            {
                results[sheetName] = GetAutoNotesForSheet(sheetName, config);
            }
            return results;
        }
    }

    /// <summary>
    /// Provides diagnostic information about multileaders and viewports for debugging.
    /// Enhanced with cache information.
    /// </summary>
    /// <param name="sheetName">Name of the sheet to diagnose</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Diagnostic information string</returns>
    public string GetDiagnosticInfo(string sheetName, ProjectConfiguration config)
    {
        try
        {
            var basicDiagnostics = GetAutoNotesForSheet(sheetName, config);
            
            var diagnosticInfo = new List<string>
            {
                $"Optimized Auto Notes Diagnostic for '{sheetName}':",
                $"  Found {basicDiagnostics.Count} notes: [{string.Join(", ", basicDiagnostics)}]",
                $"  Using optimized processor: {_useOptimizedProcessor}",
                ""
            };

            // Add cache statistics if using optimized processor
            if (_useOptimizedProcessor)
            {
                diagnosticInfo.Add("Cache Performance:");
                
                try
                {
                    // Note: Direct cache access would require public properties on OptimizedAutoNotesProcessor
                    // For now, we'll provide a summary through a public method
                    diagnosticInfo.Add($"  Optimized processor active with caching enabled");
                }
                catch (Exception cacheEx)
                {
                    diagnosticInfo.Add($"  Cache statistics unavailable: {cacheEx.Message}");
                }
                
                diagnosticInfo.Add("");
            }

            // Add fallback diagnostic using legacy service
            try
            {
                var legacyService = new AutoNotesService(_logger);
                var legacyDiagnostic = legacyService.GetDiagnosticInfo(sheetName, config);
                
                diagnosticInfo.Add("Legacy Diagnostic Comparison:");
                diagnosticInfo.Add(legacyDiagnostic);
            }
            catch (Exception legacyEx)
            {
                diagnosticInfo.Add($"Legacy diagnostic unavailable: {legacyEx.Message}");
            }

            return string.Join("\n", diagnosticInfo);
        }
        catch (Exception ex)
        {
            return $"Error getting diagnostic info: {ex.Message}";
        }
    }

    /// <summary>
    /// Clears all optimization caches (useful for testing or memory management)
    /// </summary>
    public void ClearOptimizationCaches()
    {
        try
        {
            _processor.ClearAllCaches();
            _logger.LogInformation("Optimization caches cleared");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error clearing optimization caches: {ex.Message}");
        }
    }

    /// <summary>
    /// Invalidates cache for the current drawing
    /// </summary>
    public void InvalidateCurrentDrawingCache()
    {
        try
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null && !string.IsNullOrEmpty(doc.Name))
            {
                _processor.InvalidateDrawingCache(doc.Name);
                _logger.LogInformation($"Cache invalidated for current drawing: {System.IO.Path.GetFileName(doc.Name)}");
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error invalidating current drawing cache: {ex.Message}");
        }
    }

    /// <summary>
    /// Enables or disables the optimized processor (for testing/rollback)
    /// </summary>
    /// <param name="useOptimized">True to use optimized processor, false for legacy</param>
    public static void SetOptimizationEnabled(bool useOptimized)
    {
        _useOptimizedProcessor = useOptimized;
    }

    /// <summary>
    /// Gets whether optimization is currently enabled
    /// </summary>
    public static bool IsOptimizationEnabled => _useOptimizedProcessor;

    /// <summary>
    /// Runs comprehensive performance tests comparing optimized vs legacy processing
    /// </summary>
    /// <param name="testDrawingPath">Path to test drawing</param>
    /// <param name="config">Project configuration</param>
    /// <returns>Test results summary</returns>
    public TestingSummary RunPerformanceTests(string testDrawingPath, ProjectConfiguration config)
    {
        try
        {
            var testingService = new AutoNotesTestingService(_logger);
            return testingService.RunComprehensiveTests(testDrawingPath, config);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Performance testing failed: {ex.Message}", ex);
            throw;
        }
    }
}