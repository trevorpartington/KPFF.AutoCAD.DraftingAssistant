using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IExcelReader : IDisposable
{
    Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config);
    Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config);
    Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config);
    Task<bool> FileExistsAsync(string filePath);
    Task<string[]> GetWorksheetNamesAsync(string filePath);
    Task<string[]> GetTableNamesAsync(string filePath, string worksheetName);
}