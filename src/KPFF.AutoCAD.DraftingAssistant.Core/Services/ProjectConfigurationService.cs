using System.Text.Json;
using System.Text.RegularExpressions;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using Microsoft.Extensions.Logging;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

public class ProjectConfigurationService : IProjectConfigurationService
{
    private readonly IApplicationLogger _logger;
    private readonly JsonSerializerOptions _jsonOptions;

    public ProjectConfigurationService(IApplicationLogger logger)
    {
        _logger = logger;
        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };
    }

    public async Task<ProjectConfiguration?> LoadConfigurationAsync(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                _logger.LogWarning($"Configuration file not found: {filePath}");
                return null;
            }

            var json = await File.ReadAllTextAsync(filePath);
            var config = JsonSerializer.Deserialize<ProjectConfiguration>(json, _jsonOptions);
            
            _logger.LogInformation($"Loaded project configuration: {config?.ProjectName}");
            return config;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error loading configuration from {filePath}: {ex.Message}");
            return null;
        }
    }

    public async Task SaveConfigurationAsync(ProjectConfiguration configuration, string filePath)
    {
        try
        {
            var json = JsonSerializer.Serialize(configuration, _jsonOptions);
            await File.WriteAllTextAsync(filePath, json);
            
            _logger.LogInformation($"Saved project configuration: {configuration.ProjectName}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error saving configuration to {filePath}: {ex.Message}");
            throw;
        }
    }

    public ProjectConfiguration CreateDefaultConfiguration()
    {
        return new ProjectConfiguration
        {
            ProjectName = "New Project",
            ClientName = "",
            ProjectIndexFilePath = "",
            ProjectDWGFilePath = "",
            SheetNaming = new SheetNamingConfiguration
            {
                Pattern = @"([A-Z]{1,3})-?(\d{1,3})",
                SeriesGroup = 1,
                NumberGroup = 2,
                Examples = new[] { "ABC-101", "PV-101", "C-001" }
            },
            Tables = new TableConfiguration
            {
                SheetIndex = "SHEET_INDEX",
                ExcelNotes = "EXCEL_NOTES",
                NotesPattern = "{0}_NOTES"
            },
            ConstructionNotes = new ConstructionNotesConfiguration
            {
                MultileaderStyleNames = new List<string> { "KPFF_Note_Standard" },
                NoteBlockPattern = @"^NT\d{2}$",
                MaxNotesPerSheet = 24,
                Attributes = new ConstructionNoteAttributes
                {
                    NumberAttribute = "Number",
                    NoteAttribute = "Note"
                },
                VisibilityPropertyName = "Visibility"
            }
        };
    }

    public bool ValidateConfiguration(ProjectConfiguration configuration, out List<string> errors)
    {
        errors = new List<string>();

        if (string.IsNullOrEmpty(configuration.ProjectName))
            errors.Add("Project name is required");

        if (string.IsNullOrEmpty(configuration.ProjectIndexFilePath))
            errors.Add("Project index file path is required");
        else if (!File.Exists(configuration.ProjectIndexFilePath))
            errors.Add($"Project index file not found: {configuration.ProjectIndexFilePath}");

        if (string.IsNullOrEmpty(configuration.ProjectDWGFilePath))
            errors.Add("Project DWG file path is required");
        else if (!Directory.Exists(configuration.ProjectDWGFilePath))
            errors.Add($"Project DWG directory not found: {configuration.ProjectDWGFilePath}");

        if (string.IsNullOrEmpty(configuration.SheetNaming.Pattern))
            errors.Add("Sheet naming pattern is required");
        else
        {
            try
            {
                var regex = new Regex(configuration.SheetNaming.Pattern);
                // Test with examples if provided
                foreach (var example in configuration.SheetNaming.Examples)
                {
                    var match = regex.Match(example);
                    if (!match.Success)
                        errors.Add($"Sheet naming pattern does not match example: {example}");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Invalid sheet naming pattern: {ex.Message}");
            }
        }

        if (string.IsNullOrEmpty(configuration.ConstructionNotes.MultileaderStyleName))
            errors.Add("Construction notes multileader style name is required");

        if (configuration.ConstructionNotes.MaxNotesPerSheet <= 0 || configuration.ConstructionNotes.MaxNotesPerSheet > 100)
            errors.Add("Max notes per sheet must be between 1 and 100");

        return errors.Count == 0;
    }

    public string[] ExtractSeriesFromSheetName(string sheetName, SheetNamingConfiguration namingConfig)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(namingConfig.Pattern))
            return Array.Empty<string>();

        try
        {
            var regex = new Regex(namingConfig.Pattern);
            var match = regex.Match(sheetName);
            
            if (match.Success && match.Groups.Count > namingConfig.SeriesGroup)
            {
                var series = match.Groups[namingConfig.SeriesGroup].Value;
                var number = match.Groups.Count > namingConfig.NumberGroup 
                    ? match.Groups[namingConfig.NumberGroup].Value 
                    : "";
                
                return new[] { series, number };
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Error extracting series from sheet name '{sheetName}': {ex.Message}");
        }

        return Array.Empty<string>();
    }
}