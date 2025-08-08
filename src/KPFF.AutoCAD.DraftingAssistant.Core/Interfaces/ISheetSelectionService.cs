using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface ISheetSelectionService
{
    Task<List<SheetInfo>> GetAvailableSheetsAsync(ProjectConfiguration config);
    List<SheetInfo> GetSelectedSheets();
    void SetSelectedSheets(List<SheetInfo> sheets);
    void ClearSelection();
    bool HasSelection();
}