using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace TestConsole;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("Testing Construction Notes Services...\n");

        var logger = new ConsoleLogger();
        
        // Test paths
        var currentDir = Directory.GetCurrentDirectory();
        var solutionRoot = FindSolutionRoot(currentDir);
        var testDataPath = Path.Combine(solutionRoot, "testdata");
        var excelPath = Path.Combine(testDataPath, "ProjectIndex.xlsx");
        var configPath = Path.Combine(testDataPath, "ProjectConfig.json");

        Console.WriteLine($"Test data path: {testDataPath}");
        Console.WriteLine($"Excel file path: {excelPath}");
        Console.WriteLine($"Config file path: {configPath}");
        Console.WriteLine();

        // Test Project Configuration Service
        await TestProjectConfiguration(logger, configPath);
        Console.WriteLine();

        // Test Excel Reader Service  
        var configService = new ProjectConfigurationService(logger);
        var config = await configService.LoadConfigurationAsync(configPath);
        if (config != null)
        {
            await TestExcelReader(logger, excelPath, config);
        }
        else
        {
            Console.WriteLine("Cannot test Excel reader - no configuration loaded");
        }
    }

    static async Task TestProjectConfiguration(ConsoleLogger logger, string configPath)
    {
        Console.WriteLine("=== Testing Project Configuration Service ===");
        
        var configService = new ProjectConfigurationService(logger);

        try
        {
            if (File.Exists(configPath))
            {
                var config = await configService.LoadConfigurationAsync(configPath);
                if (config != null)
                {
                    Console.WriteLine($"✓ Loaded configuration: {config.ProjectName}");
                    Console.WriteLine($"  - Client: {config.ClientName}");
                    Console.WriteLine($"  - Excel File: {config.ProjectIndexFilePath}");
                    Console.WriteLine($"  - Sheet Pattern: {config.SheetNaming.Pattern}");

                    // Test validation
                    var isValid = configService.ValidateConfiguration(config, out var errors);
                    Console.WriteLine($"  - Valid: {isValid}");
                    if (errors.Count > 0)
                    {
                        Console.WriteLine("  - Errors:");
                        foreach (var error in errors)
                            Console.WriteLine($"    • {error}");
                    }

                    // Test sheet name parsing
                    var testSheets = new[] { "PROJ-ABC-100", "PROJ-PV-101", "PROJ-C-001" };
                    Console.WriteLine("  - Sheet name parsing:");
                    foreach (var sheet in testSheets)
                    {
                        var parts = configService.ExtractSeriesFromSheetName(sheet, config.SheetNaming);
                        if (parts.Length >= 2)
                            Console.WriteLine($"    • {sheet} → Series: '{parts[0]}', Number: '{parts[1]}'");
                        else
                            Console.WriteLine($"    • {sheet} → Failed to parse");
                    }
                }
                else
                {
                    Console.WriteLine("✗ Failed to load configuration");
                }
            }
            else
            {
                Console.WriteLine($"✗ Config file not found: {configPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}");
        }
    }

    static async Task TestExcelReader(ConsoleLogger logger, string excelPath, ProjectConfiguration config)
    {
        Console.WriteLine("=== Testing Excel Reader Service ===");
        
        var excelReader = new ExcelReaderService(logger);

        try
        {
            // Test file existence
            var exists = await excelReader.FileExistsAsync(excelPath);
            Console.WriteLine($"File exists: {exists}");

            if (!exists)
            {
                Console.WriteLine($"✗ Excel file not found: {excelPath}");
                return;
            }

            // Test worksheet names
            var worksheets = await excelReader.GetWorksheetNamesAsync(excelPath);
            Console.WriteLine($"Worksheets ({worksheets.Length}): {string.Join(", ", worksheets)}");

            // Check what tables exist in each worksheet
            foreach (var worksheet in worksheets)
            {
                var tables = await excelReader.GetTableNamesAsync(excelPath, worksheet);
                Console.WriteLine($"  {worksheet} tables ({tables.Length}): {string.Join(", ", tables)}");
            }

            // Test reading sheet index
            try
            {
                var sheets = await excelReader.ReadSheetIndexAsync(excelPath, config);
                Console.WriteLine($"✓ Read {sheets.Count} sheets from SheetIndex table");
                
                if (sheets.Count > 0)
                {
                    var firstSheet = sheets.First();
                    Console.WriteLine($"  - First sheet: {firstSheet.SheetName}");
                    Console.WriteLine($"    Title: {firstSheet.DrawingTitle}");
                    Console.WriteLine($"    Series: {firstSheet.Series}");
                    Console.WriteLine($"    Number: {firstSheet.Number}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error reading sheet index: {ex.Message}");
            }

            // Test reading construction notes for ABC series
            try
            {
                var notes = await excelReader.ReadConstructionNotesAsync(excelPath, "ABC", config);
                Console.WriteLine($"✓ Read {notes.Count} construction notes for ABC series");
                
                foreach (var note in notes.Take(3))
                {
                    Console.WriteLine($"  - Note {note.Number}: {note.Text}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error reading construction notes: {ex.Message}");
            }

            // Test reading Excel notes mappings
            try
            {
                var mappings = await excelReader.ReadExcelNotesAsync(excelPath, config);
                Console.WriteLine($"✓ Read Excel notes mappings for {mappings.Count} sheets");
                
                foreach (var mapping in mappings.Take(3))
                {
                    Console.WriteLine($"  - {mapping.SheetName}: {string.Join(", ", mapping.NoteNumbers)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Error reading excel notes: {ex.Message}");
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}");
        }
    }

    static string FindSolutionRoot(string startPath)
    {
        var current = startPath;
        while (current != null && !File.Exists(Path.Combine(current, "KPFF.AutoCAD.DraftingAssistant.sln")))
        {
            current = Directory.GetParent(current)?.FullName;
        }
        return current ?? startPath;
    }
}

public class ConsoleLogger : IApplicationLogger
{
    public void LogInformation(string message) => Console.WriteLine($"[INFO] {message}");
    public void LogWarning(string message) => Console.WriteLine($"[WARN] {message}");
    public void LogError(string message, Exception? exception = null) => 
        Console.WriteLine($"[ERROR] {message}" + (exception != null ? $" - {exception.Message}" : ""));
    public void LogDebug(string message) => Console.WriteLine($"[DEBUG] {message}");
    public void LogCritical(string message, Exception? exception = null) => 
        Console.WriteLine($"[CRITICAL] {message}" + (exception != null ? $" - {exception.Message}" : ""));
}