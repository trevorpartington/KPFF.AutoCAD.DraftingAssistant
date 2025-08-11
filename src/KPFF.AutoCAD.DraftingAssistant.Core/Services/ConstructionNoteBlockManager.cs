using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Stub implementation of IConstructionNoteBlockManager to prevent crashes
/// Will be replaced with full Model Space implementation
/// </summary>
public class ConstructionNoteBlockManager : IConstructionNoteBlockManager
{
    private readonly ILogger _logger;

    public ConstructionNoteBlockManager(ILogger logger)
    {
        _logger = logger;
    }

    public Task<List<ConstructionNoteBlock>> GetExistingNoteBlocksAsync(string layoutName, ProjectConfiguration config)
    {
        _logger.LogDebug($"GetExistingNoteBlocksAsync called for {layoutName} - returning empty list (stub implementation)");
        return Task.FromResult(new List<ConstructionNoteBlock>());
    }

    public Task UpdateNoteBlockAsync(string layoutName, ConstructionNoteBlock noteBlock, ProjectConfiguration config)
    {
        _logger.LogDebug($"UpdateNoteBlockAsync called for {layoutName} - no operation (stub implementation)");
        return Task.CompletedTask;
    }

    public Task ClearAllNoteBlocksAsync(string layoutName, ProjectConfiguration config)
    {
        _logger.LogDebug($"ClearAllNoteBlocksAsync called for {layoutName} - no operation (stub implementation)");
        return Task.CompletedTask;
    }

    public Task<bool> DoNoteBlocksExistAsync(string layoutName, ProjectConfiguration config)
    {
        _logger.LogDebug($"DoNoteBlocksExistAsync called for {layoutName} - returning false (stub implementation)");
        return Task.FromResult(false);
    }

    public Task CreateNoteBlocksAsync(string layoutName, ProjectConfiguration config)
    {
        _logger.LogDebug($"CreateNoteBlocksAsync called for {layoutName} - no operation (stub implementation)");
        return Task.CompletedTask;
    }

    public string GenerateNoteBlockName(string sheetNumber, int noteIndex, ProjectConfiguration config)
    {
        string blockName = $"NT{noteIndex:D2}";
        _logger.LogDebug($"GenerateNoteBlockName called for sheet {sheetNumber}, index {noteIndex} - returning {blockName} (simplified naming)");
        return blockName;
    }
}