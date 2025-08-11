using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Stub implementation of IConstructionNotesService to prevent crashes
/// Will be replaced with full Model Space implementation
/// </summary>
public class ConstructionNotesService : IConstructionNotesService
{
    private readonly ILogger _logger;

    public ConstructionNotesService(ILogger logger)
    {
        _logger = logger;
    }

    public Task<List<int>> GetAutoNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        _logger.LogDebug($"GetAutoNotesForSheetAsync called for {sheetName} - returning empty list (stub implementation)");
        return Task.FromResult(new List<int>());
    }

    public Task<List<int>> GetExcelNotesForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        _logger.LogDebug($"GetExcelNotesForSheetAsync called for {sheetName} - returning empty list (stub implementation)");
        return Task.FromResult(new List<int>());
    }

    public Task UpdateConstructionNoteBlocksAsync(string sheetName, List<int> noteNumbers, ProjectConfiguration config)
    {
        _logger.LogDebug($"UpdateConstructionNoteBlocksAsync called for {sheetName} with {noteNumbers.Count} notes - no operation (stub implementation)");
        return Task.CompletedTask;
    }

    public Task<List<ConstructionNote>> GetNotesForSeriesAsync(string series, ProjectConfiguration config)
    {
        _logger.LogDebug($"GetNotesForSeriesAsync called for series {series} - returning empty list (stub implementation)");
        return Task.FromResult(new List<ConstructionNote>());
    }

    public Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config)
    {
        _logger.LogDebug($"ValidateNoteBlocksExistAsync called for {sheetName} - returning false (stub implementation)");
        return Task.FromResult(false);
    }

    public Task CreateNoteBlocksForSheetAsync(string sheetName, ProjectConfiguration config)
    {
        _logger.LogDebug($"CreateNoteBlocksForSheetAsync called for {sheetName} - no operation (stub implementation)");
        return Task.CompletedTask;
    }
}