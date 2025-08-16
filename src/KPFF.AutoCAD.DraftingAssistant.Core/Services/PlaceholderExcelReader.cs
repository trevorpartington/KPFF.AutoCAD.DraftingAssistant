using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Placeholder implementation of IExcelReader that returns empty results
/// Used as a temporary stub until Phase 3 Excel functionality is implemented
/// </summary>
public class PlaceholderExcelReader : IExcelReader
{
    private readonly IApplicationLogger _logger;
    private bool _disposed = false;

    public PlaceholderExcelReader(IApplicationLogger logger)
    {
        _logger = logger;
    }

    public Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
    {
        _logger.LogDebug($"PlaceholderExcelReader.ReadSheetIndexAsync called for {filePath} - returning empty list");
        return Task.FromResult(new List<SheetInfo>());
    }

    public Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config)
    {
        _logger.LogDebug($"PlaceholderExcelReader.ReadConstructionNotesAsync called for series {series} - returning empty list");
        return Task.FromResult(new List<ConstructionNote>());
    }

    public Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config)
    {
        _logger.LogDebug($"PlaceholderExcelReader.ReadExcelNotesAsync called for {filePath} - returning empty list");
        return Task.FromResult(new List<SheetNoteMapping>());
    }

    public Task<bool> FileExistsAsync(string filePath)
    {
        _logger.LogDebug($"PlaceholderExcelReader.FileExistsAsync called for {filePath} - returning false");
        return Task.FromResult(false);
    }

    public Task<string[]> GetWorksheetNamesAsync(string filePath)
    {
        _logger.LogDebug($"PlaceholderExcelReader.GetWorksheetNamesAsync called for {filePath} - returning empty array");
        return Task.FromResult(Array.Empty<string>());
    }

    public Task<string[]> GetTableNamesAsync(string filePath, string worksheetName)
    {
        _logger.LogDebug($"PlaceholderExcelReader.GetTableNamesAsync called for {filePath}/{worksheetName} - returning empty array");
        return Task.FromResult(Array.Empty<string>());
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _logger?.LogDebug("PlaceholderExcelReader disposed");
            }
            _disposed = true;
        }
    }
}