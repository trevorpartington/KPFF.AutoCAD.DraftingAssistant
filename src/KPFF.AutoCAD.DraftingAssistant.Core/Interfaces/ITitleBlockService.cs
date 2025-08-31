using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface ITitleBlockService
{
    Task<TitleBlockMapping?> GetTitleBlockMappingForSheetAsync(string sheetName, ProjectConfiguration config);
    Task UpdateTitleBlockAsync(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config);
    Task<bool> ValidateTitleBlockExistsAsync(string sheetName, ProjectConfiguration config);
    Task<Dictionary<string, string>> GetTitleBlockAttributesAsync(string sheetName, ProjectConfiguration config);
}