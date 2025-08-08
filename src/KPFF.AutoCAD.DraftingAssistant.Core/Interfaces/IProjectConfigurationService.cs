using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IProjectConfigurationService
{
    Task<ProjectConfiguration?> LoadConfigurationAsync(string filePath);
    Task SaveConfigurationAsync(ProjectConfiguration configuration, string filePath);
    ProjectConfiguration CreateDefaultConfiguration();
    bool ValidateConfiguration(ProjectConfiguration configuration, out List<string> errors);
    string[] ExtractSeriesFromSheetName(string sheetName, SheetNamingConfiguration namingConfig);
}