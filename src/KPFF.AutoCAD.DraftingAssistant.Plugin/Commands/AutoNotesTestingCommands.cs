using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System.Diagnostics;
using System.IO;
using SystemException = System.Exception;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// AutoCAD commands for testing and validating the optimized Auto Notes system.
/// These commands facilitate comprehensive testing of the performance improvements.
/// </summary>
public class AutoNotesTestingCommands
{
    private static readonly ILogger _logger = new SimpleLogger();

    [CommandMethod("KPFF", "TESTAUTONOTES", CommandFlags.Modal)]
    public static void TestAutoNotesOptimization()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== AUTO NOTES OPTIMIZATION TESTING ===\n");

            // Get the current drawing path
            var drawingPath = doc.Name;
            if (string.IsNullOrEmpty(drawingPath) || drawingPath.Contains("Drawing"))
            {
                ed.WriteMessage("Error: Please open a real drawing file for testing.\n");
                return;
            }

            ed.WriteMessage($"Testing drawing: {Path.GetFileName(drawingPath)}\n");

            // Create a test configuration
            var config = CreateTestConfiguration(ed);
            if (config == null) return;

            // Run comprehensive tests
            var testingService = new AutoNotesTestingService(_logger);
            var summary = testingService.RunComprehensiveTests(drawingPath, config);

            // Display results in AutoCAD
            DisplayTestResults(ed, summary);
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error during testing: {ex.Message}\n");
            _logger.LogError($"Auto Notes testing failed: {ex.Message}", ex);
        }
    }

    [CommandMethod("KPFF", "COMPAREAUTONOTES", CommandFlags.Modal)]
    public static void CompareOptimizedVsLegacy()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== AUTO NOTES PERFORMANCE COMPARISON ===\n");

            // Get sheet name to test
            var sheetNameResult = ed.GetString("\nEnter sheet name to test: ");
            if (sheetNameResult.Status != PromptStatus.OK)
            {
                return;
            }

            var sheetName = sheetNameResult.StringResult;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                ed.WriteMessage("Error: Sheet name cannot be empty.\n");
                return;
            }

            // Create test configuration
            var config = CreateTestConfiguration(ed);
            if (config == null) return;

            ed.WriteMessage($"\nTesting sheet: {sheetName}\n");

            // Test optimized version
            var optimizedService = new OptimizedAutoNotesService(_logger);
            var optimizedStopwatch = Stopwatch.StartNew();
            var optimizedResults = optimizedService.GetAutoNotesForSheet(sheetName, config);
            optimizedStopwatch.Stop();

            // Test legacy version
            OptimizedAutoNotesService.SetOptimizationEnabled(false);
            var legacyStopwatch = Stopwatch.StartNew();
            var legacyResults = optimizedService.GetAutoNotesForSheet(sheetName, config);
            legacyStopwatch.Stop();
            OptimizedAutoNotesService.SetOptimizationEnabled(true);

            // Compare results
            var resultsMatch = optimizedResults.SequenceEqual(legacyResults);
            var performanceGain = legacyStopwatch.Elapsed.TotalMilliseconds > 0
                ? (legacyStopwatch.Elapsed.TotalMilliseconds - optimizedStopwatch.Elapsed.TotalMilliseconds) / legacyStopwatch.Elapsed.TotalMilliseconds * 100
                : 0;

            // Display comparison
            ed.WriteMessage("\n=== COMPARISON RESULTS ===\n");
            ed.WriteMessage($"Optimized Time: {optimizedStopwatch.Elapsed.TotalMilliseconds:F0}ms\n");
            ed.WriteMessage($"Legacy Time: {legacyStopwatch.Elapsed.TotalMilliseconds:F0}ms\n");
            ed.WriteMessage($"Performance Gain: {performanceGain:F1}%\n");
            ed.WriteMessage($"Results Match: {(resultsMatch ? "✅ Yes" : "❌ No")}\n");
            ed.WriteMessage($"Optimized Results: [{string.Join(", ", optimizedResults)}]\n");
            ed.WriteMessage($"Legacy Results: [{string.Join(", ", legacyResults)}]\n");

            if (!resultsMatch)
            {
                ed.WriteMessage("⚠️  WARNING: Results do not match between optimized and legacy versions!\n");
            }
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error during comparison: {ex.Message}\n");
            _logger.LogError($"Auto Notes comparison failed: {ex.Message}", ex);
        }
    }

    [CommandMethod("KPFF", "BATCHAUTONOTES", CommandFlags.Modal)]
    public static void TestBatchAutoNotes()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== BATCH AUTO NOTES TESTING ===\n");

            // Get available sheets
            var availableSheets = GetAvailableSheets(doc);
            if (availableSheets.Count == 0)
            {
                ed.WriteMessage("Error: No layouts found in current drawing.\n");
                return;
            }

            ed.WriteMessage($"Available sheets: {string.Join(", ", availableSheets)}\n");

            // Get sheets to test
            var sheetsToTest = new List<string>();
            while (true)
            {
                var sheetResult = ed.GetString($"\nEnter sheet name to include (or press Enter to finish): ");
                if (sheetResult.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(sheetResult.StringResult))
                {
                    break;
                }

                var sheetName = sheetResult.StringResult.Trim();
                if (availableSheets.Contains(sheetName, StringComparer.OrdinalIgnoreCase))
                {
                    sheetsToTest.Add(sheetName);
                    ed.WriteMessage($"Added: {sheetName}\n");
                }
                else
                {
                    ed.WriteMessage($"Warning: Sheet '{sheetName}' not found in drawing.\n");
                }
            }

            if (sheetsToTest.Count == 0)
            {
                ed.WriteMessage("No valid sheets selected for testing.\n");
                return;
            }

            // Create test configuration
            var config = CreateTestConfiguration(ed);
            if (config == null) return;

            ed.WriteMessage($"\nProcessing {sheetsToTest.Count} sheets...\n");

            // Test batch processing
            var optimizedService = new OptimizedAutoNotesService(_logger);
            var batchStopwatch = Stopwatch.StartNew();
            var batchResults = optimizedService.GetAutoNotesForMultipleSheets(sheetsToTest, config);
            batchStopwatch.Stop();

            // Test individual processing for comparison
            var individualStopwatch = Stopwatch.StartNew();
            var individualResults = new Dictionary<string, List<int>>();
            foreach (var sheet in sheetsToTest)
            {
                individualResults[sheet] = optimizedService.GetAutoNotesForSheet(sheet, config);
            }
            individualStopwatch.Stop();

            // Display results
            ed.WriteMessage("\n=== BATCH PROCESSING RESULTS ===\n");
            ed.WriteMessage($"Batch Time: {batchStopwatch.Elapsed.TotalMilliseconds:F0}ms\n");
            ed.WriteMessage($"Individual Time: {individualStopwatch.Elapsed.TotalMilliseconds:F0}ms\n");
            
            var batchEfficiency = individualStopwatch.Elapsed.TotalMilliseconds > 0
                ? (individualStopwatch.Elapsed.TotalMilliseconds - batchStopwatch.Elapsed.TotalMilliseconds) / individualStopwatch.Elapsed.TotalMilliseconds * 100
                : 0;
            
            ed.WriteMessage($"Batch Efficiency: {batchEfficiency:F1}%\n");

            ed.WriteMessage("\nSheet Results:\n");
            foreach (var sheet in sheetsToTest)
            {
                var batchNotes = batchResults.GetValueOrDefault(sheet, new List<int>());
                var individualNotes = individualResults.GetValueOrDefault(sheet, new List<int>());
                var match = batchNotes.SequenceEqual(individualNotes) ? "✅" : "❌";
                
                ed.WriteMessage($"  {sheet}: {match} [{string.Join(", ", batchNotes)}]\n");
            }
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error during batch testing: {ex.Message}\n");
            _logger.LogError($"Batch Auto Notes testing failed: {ex.Message}", ex);
        }
    }

    [CommandMethod("KPFF", "CLEARAUTONOTES", CommandFlags.Modal)]
    public static void ClearAutoNotesCache()
    {
        var ed = Application.DocumentManager.MdiActiveDocument.Editor;

        try
        {
            ed.WriteMessage("\n=== CLEARING AUTO NOTES CACHE ===\n");

            var optimizedService = new OptimizedAutoNotesService(_logger);
            optimizedService.ClearOptimizationCaches();

            ed.WriteMessage("✅ All Auto Notes caches cleared successfully.\n");
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error clearing caches: {ex.Message}\n");
            _logger.LogError($"Cache clearing failed: {ex.Message}", ex);
        }
    }

    [CommandMethod("KPFF", "AUTONOTESDIAG", CommandFlags.Modal)]
    public static void AutoNotesDiagnostics()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\n=== AUTO NOTES DIAGNOSTICS ===\n");

            // Get sheet name
            var sheetResult = ed.GetString("\nEnter sheet name for diagnostics: ");
            if (sheetResult.Status != PromptStatus.OK)
            {
                return;
            }

            var sheetName = sheetResult.StringResult;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                ed.WriteMessage("Error: Sheet name cannot be empty.\n");
                return;
            }

            // Create test configuration
            var config = CreateTestConfiguration(ed);
            if (config == null) return;

            // Get diagnostics
            var optimizedService = new OptimizedAutoNotesService(_logger);
            var diagnostics = optimizedService.GetDiagnosticInfo(sheetName, config);

            ed.WriteMessage("\n" + diagnostics + "\n");
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error getting diagnostics: {ex.Message}\n");
            _logger.LogError($"Auto Notes diagnostics failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates a test configuration with basic multileader styles
    /// </summary>
    private static ProjectConfiguration? CreateTestConfiguration(Editor ed)
    {
        try
        {
            // Ask for multileader style names
            var stylesResult = ed.GetString("\nEnter multileader style names (comma-separated, or press Enter for 'any'): ");
            if (stylesResult.Status != PromptStatus.OK)
            {
                return null;
            }

            var multileaderStyles = new List<string>();
            if (!string.IsNullOrWhiteSpace(stylesResult.StringResult))
            {
                multileaderStyles = stylesResult.StringResult
                    .Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToList();
            }

            // Create basic configuration
            var config = new ProjectConfiguration
            {
                ConstructionNotes = new ConstructionNotesConfiguration
                {
                    MultileaderStyleNames = multileaderStyles,
                    NoteBlocks = new List<NoteBlockConfiguration>() // Add if needed
                }
            };

            ed.WriteMessage($"Using multileader styles: {(multileaderStyles.Count > 0 ? string.Join(", ", multileaderStyles) : "any")}\n");
            return config;
        }
        catch (SystemException ex)
        {
            ed.WriteMessage($"Error creating test configuration: {ex.Message}\n");
            return null;
        }
    }

    /// <summary>
    /// Gets available sheet names from the current drawing
    /// </summary>
    private static List<string> GetAvailableSheets(Document doc)
    {
        var sheets = new List<string>();

        try
        {
            var db = doc.Database;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (layoutDict != null)
                {
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                        {
                            sheets.Add(entry.Key);
                        }
                    }
                }
                tr.Commit();
            }
        }
        catch (SystemException ex)
        {
            _logger.LogError($"Error getting available sheets: {ex.Message}", ex);
        }

        return sheets;
    }

    /// <summary>
    /// Displays comprehensive test results in the AutoCAD editor
    /// </summary>
    private static void DisplayTestResults(Editor ed, TestingSummary summary)
    {
        ed.WriteMessage("\n=== COMPREHENSIVE TEST RESULTS ===\n");
        ed.WriteMessage($"Total Tests: {summary.TotalTests}\n");
        ed.WriteMessage($"Passed: {summary.PassedTests} ({summary.PassRate:F1}%)\n");
        ed.WriteMessage($"Failed: {summary.FailedTests}\n");
        ed.WriteMessage($"Average Performance Gain: {summary.AveragePerformanceGain:F1}%\n");
        ed.WriteMessage($"Best Performance Gain: {summary.BestPerformanceGain:F1}%\n");
        ed.WriteMessage($"Total Testing Time: {summary.TotalTestingTime.TotalSeconds:F1}s\n");

        ed.WriteMessage("\n=== INDIVIDUAL TEST RESULTS ===\n");
        foreach (var result in summary.Results)
        {
            var status = result.TestPassed ? "✅ PASS" : "❌ FAIL";
            var performance = result.PerformanceGain > 0 ? $"+{result.PerformanceGain:F1}%" : $"{result.PerformanceGain:F1}%";
            
            ed.WriteMessage($"{status} {result.TestName}: {performance} improvement\n");
            
            if (!result.TestPassed && !string.IsNullOrEmpty(result.ErrorMessage))
            {
                ed.WriteMessage($"    Error: {result.ErrorMessage}\n");
            }
        }

        if (summary.FailedTests > 0)
        {
            ed.WriteMessage("\n⚠️  Some tests failed. Check the logs for detailed error information.\n");
        }
        else
        {
            ed.WriteMessage("\n🎉 All tests passed! Auto Notes optimization is working correctly.\n");
        }
    }
}

/// <summary>
/// Simple logger implementation for AutoCAD commands
/// </summary>
internal class SimpleLogger : ILogger
{
    public void LogDebug(string message) => System.Diagnostics.Debug.WriteLine($"[DEBUG] {message}");
    public void LogInformation(string message) => System.Diagnostics.Debug.WriteLine($"[INFO] {message}");
    public void LogWarning(string message) => System.Diagnostics.Debug.WriteLine($"[WARN] {message}");
    public void LogError(string message) => System.Diagnostics.Debug.WriteLine($"[ERROR] {message}");
    public void LogError(string message, SystemException exception) => System.Diagnostics.Debug.WriteLine($"[ERROR] {message}: {exception.Message}");
}