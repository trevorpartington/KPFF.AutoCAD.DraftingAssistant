using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using Microsoft.Extensions.Logging;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Excel reader service that delegates to out-of-process Excel reader
/// This prevents EPPlus from causing AutoCAD freezing issues
/// </summary>
public class ExcelReaderService : IExcelReader
{
    private readonly IApplicationLogger _logger;
    private readonly NamedPipeExcelClient _client;
    private bool _disposed = false;

    public ExcelReaderService(IApplicationLogger logger)
    {
        _logger = logger;
        _client = new NamedPipeExcelClient(logger);
    }

    public async Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"Reading sheet index from {filePath} via out-of-process Excel reader");
            return await _client.ReadSheetIndexAsync(filePath, config);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading sheet index from {filePath}: {ex.Message}");
            throw;
        }
    }

    public async Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"Reading construction notes for series {series} from {filePath} via out-of-process Excel reader");
            return await _client.ReadConstructionNotesAsync(filePath, series, config);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading construction notes for series {series}: {ex.Message}");
            throw;
        }
    }

    public async Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config)
    {
        try
        {
            _logger.LogInformation($"Reading Excel notes from {filePath} via out-of-process Excel reader");
            return await _client.ReadExcelNotesAsync(filePath, config);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading excel notes from {filePath}: {ex.Message}");
            throw;
        }
    }

    public async Task<bool> FileExistsAsync(string filePath)
    {
        return await _client.FileExistsAsync(filePath);
    }

    public async Task<string[]> GetWorksheetNamesAsync(string filePath)
    {
        return await _client.GetWorksheetNamesAsync(filePath);
    }

    public async Task<string[]> GetTableNamesAsync(string filePath, string worksheetName)
    {
        return await _client.GetTableNamesAsync(filePath, worksheetName);
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
                _logger?.LogInformation("Disposing ExcelReaderService and terminating Excel reader process");
                _client?.Dispose();
            }
            _disposed = true;
        }
    }
}