using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IExcelReader
{
    Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath);
    Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series);
    Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath);
    Task<bool> FileExistsAsync(string filePath);
    Task<string[]> GetWorksheetNamesAsync(string filePath);
    Task<string[]> GetTableNamesAsync(string filePath, string worksheetName);
}