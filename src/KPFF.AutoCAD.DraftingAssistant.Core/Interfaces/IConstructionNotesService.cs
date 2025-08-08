using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IConstructionNotesService
{
    Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config);
    Task<List<int>> GetExcelNotesForSheetAsync(string sheetName, ProjectConfiguration config);
    Task UpdateConstructionNoteBlocksAsync(string sheetName, List<int> noteNumbers, ProjectConfiguration config);
    Task<List<ConstructionNote>> GetNotesForSeriesAsync(string series, ProjectConfiguration config);
    Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config);
    Task CreateNoteBlocksForSheetAsync(string sheetName, ProjectConfiguration config);
}