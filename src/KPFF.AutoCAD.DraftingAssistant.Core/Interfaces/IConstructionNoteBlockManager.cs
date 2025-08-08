using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

public interface IConstructionNoteBlockManager
{
    Task<List<ConstructionNoteBlock>> GetExistingNoteBlocksAsync(string layoutName, ProjectConfiguration config);
    Task UpdateNoteBlockAsync(string layoutName, ConstructionNoteBlock noteBlock, ProjectConfiguration config);
    Task ClearAllNoteBlocksAsync(string layoutName, ProjectConfiguration config);
    Task<bool> DoNoteBlocksExistAsync(string layoutName, ProjectConfiguration config);
    Task CreateNoteBlocksAsync(string layoutName, ProjectConfiguration config);
    string GenerateNoteBlockName(string sheetNumber, int noteIndex, ProjectConfiguration config);
}