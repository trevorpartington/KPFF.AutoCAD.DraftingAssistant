using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IMultileaderAnalyzer
{
    Task<List<MultileaderInfo>> AnalyzeLayoutMultileadersAsync(string layoutName, ProjectConfiguration config);
    Task<List<ViewportInfo>> GetLayoutViewportsAsync(string layoutName);
    bool IsMultileaderInViewport(MultileaderInfo multileader, ViewportInfo viewport);
    bool IsValidNoteNumber(string text, out int noteNumber);
}