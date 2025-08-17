using ClosedXML.Excel;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Excel reader service implementation using ClosedXML library
/// Provides async Excel reading functionality with comprehensive error handling
/// </summary>
public class ExcelReaderService : IExcelReader
{
    private readonly IApplicationLogger _logger;
    private bool _disposed = false;

    public ExcelReaderService(IApplicationLogger logger)
    {
        _logger = logger;
    }

    public async Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
    {
        return await Task.Run(() =>
        {
            _logger.LogDebug($"Reading sheet index from {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    var errorMsg = $"Excel file not found: {filePath}";
                    _logger.LogError(errorMsg);
                    return new List<SheetInfo>();
                }

                using var workbook = new XLWorkbook(filePath);
                
                // Find SHEET_INDEX table across all worksheets
                var sheetIndexTable = FindTableByName(workbook, config.Tables.SheetIndex);
                if (sheetIndexTable == null)
                {
                    var errorMsg = $"Table '{config.Tables.SheetIndex}' not found in workbook";
                    _logger.LogError(errorMsg);
                    return new List<SheetInfo>();
                }

                var sheets = new List<SheetInfo>();
                var dataRange = sheetIndexTable.DataRange;
                
                _logger.LogDebug($"Found {dataRange.RowCount()} rows in SHEET_INDEX table");

                foreach (var row in dataRange.Rows())
                {
                    try
                    {
                        var sheetName = row.Cell(1).GetString().Trim();
                        var fileName = row.Cell(2).GetString().Trim();
                        var title = row.Cell(3).GetString().Trim();

                        if (string.IsNullOrEmpty(sheetName))
                            continue;

                        // Extract series and number from sheet name using configuration
                        var configService = new ProjectConfigurationService(_logger);
                        var parts = configService.ExtractSeriesFromSheetName(sheetName, config.SheetNaming);
                        
                        var sheetInfo = new SheetInfo
                        {
                            SheetName = sheetName,
                            DWGFileName = fileName,
                            DrawingTitle = title
                        };

                        sheets.Add(sheetInfo);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Failed to parse sheet row: {ex.Message}");
                        continue;
                    }
                }

                _logger.LogInformation($"Successfully read {sheets.Count} sheets from index");
                return sheets;
            }
            catch (Exception ex)
            {
                var errorMsg = $"Failed to read sheet index from {filePath}: {ex.Message}";
                _logger.LogError(errorMsg, ex);
                return new List<SheetInfo>();
            }
        });
    }

    public async Task<List<ConstructionNote>> ReadConstructionNotesAsync(string filePath, string series, ProjectConfiguration config)
    {
        return await Task.Run(() =>
        {
            _logger.LogDebug($"Reading construction notes for series '{series}' from {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    var errorMsg = $"Excel file not found: {filePath}";
                    _logger.LogError(errorMsg);
                    return new List<ConstructionNote>();
                }

                using var workbook = new XLWorkbook(filePath);
                
                // Build table name dynamically: e.g., "ABC_NOTES"
                var tableName = string.Format(config.Tables.NotesPattern, series);
                var notesTable = FindTableByName(workbook, tableName);
                
                if (notesTable == null)
                {
                    var errorMsg = $"Table '{tableName}' not found in workbook";
                    _logger.LogError(errorMsg);
                    return new List<ConstructionNote>();
                }

                var notes = new List<ConstructionNote>();
                var dataRange = notesTable.DataRange;
                
                _logger.LogDebug($"Found {dataRange.RowCount()} rows in {tableName} table");

                foreach (var row in dataRange.Rows())
                {
                    try
                    {
                        var numberCell = row.Cell(1).GetString().Trim();
                        var noteText = row.Cell(2).GetString().Trim();

                        if (string.IsNullOrEmpty(numberCell) || string.IsNullOrEmpty(noteText))
                            continue;

                        if (!int.TryParse(numberCell, out var noteNumber))
                        {
                            _logger.LogWarning($"Unable to parse note number from '{numberCell}' in series {series}");
                            continue;
                        }

                        var note = new ConstructionNote
                        {
                            Number = noteNumber,
                            Text = noteText,
                            Series = series,
                            IsActive = true,
                            LastUpdated = DateTime.Now
                        };

                        notes.Add(note);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Failed to parse construction note row: {ex.Message}");
                        continue;
                    }
                }

                _logger.LogInformation($"Successfully read {notes.Count} construction notes for series {series}");
                return notes;
            }
            catch (Exception ex)
            {
                var errorMsg = $"Failed to read construction notes for series '{series}' from {filePath}: {ex.Message}";
                _logger.LogError(errorMsg, ex);
                return new List<ConstructionNote>();
            }
        });
    }

    public async Task<List<SheetNoteMapping>> ReadExcelNotesAsync(string filePath, ProjectConfiguration config)
    {
        return await Task.Run(() =>
        {
            _logger.LogDebug($"Reading Excel notes mappings from {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    var errorMsg = $"Excel file not found: {filePath}";
                    _logger.LogError(errorMsg);
                    return new List<SheetNoteMapping>();
                }

                using var workbook = new XLWorkbook(filePath);
                
                // Find EXCEL_NOTES table
                var excelNotesTable = FindTableByName(workbook, config.Tables.ExcelNotes);
                if (excelNotesTable == null)
                {
                    var errorMsg = $"Table '{config.Tables.ExcelNotes}' not found in workbook";
                    _logger.LogError(errorMsg);
                    return new List<SheetNoteMapping>();
                }

                var dataRange = excelNotesTable.DataRange;
                var columnCount = dataRange.ColumnCount();
                var maxExpectedColumns = config.ConstructionNotes.MaxNotesPerSheet + 1; // +1 for sheet name column

                // Validate column count
                if (columnCount > maxExpectedColumns)
                {
                    var errorMsg = $"EXCEL_NOTES table has {columnCount} columns but max is {maxExpectedColumns} (1 sheet + {config.ConstructionNotes.MaxNotesPerSheet} notes)";
                    _logger.LogError(errorMsg);
                    return new List<SheetNoteMapping>();
                }

                var mappings = new List<SheetNoteMapping>();
                
                _logger.LogDebug($"Found {dataRange.RowCount()} rows in EXCEL_NOTES table with {columnCount} columns");

                foreach (var row in dataRange.Rows())
                {
                    try
                    {
                        var sheetName = row.Cell(1).GetString().Trim();
                        if (string.IsNullOrEmpty(sheetName))
                            continue;

                        var noteNumbers = new List<int>();

                        // Read note numbers from columns 2 onwards
                        for (int col = 2; col <= columnCount; col++)
                        {
                            var cellValue = row.Cell(col).GetString().Trim();
                            if (string.IsNullOrEmpty(cellValue))
                                continue;

                            if (int.TryParse(cellValue, out var noteNumber))
                            {
                                // Validate note number is within valid range
                                if (noteNumber >= 1 && noteNumber <= config.ConstructionNotes.MaxNotesPerSheet)
                                {
                                    noteNumbers.Add(noteNumber);
                                }
                                else
                                {
                                    _logger.LogWarning($"Note number {noteNumber} is outside valid range 1-{config.ConstructionNotes.MaxNotesPerSheet} for sheet {sheetName}");
                                }
                            }
                            else
                            {
                                _logger.LogWarning($"Unable to parse note number from column {col}: '{cellValue}' for sheet {sheetName}");
                            }
                        }

                        // Consolidate duplicates (e.g., [4, 4, 7] → [4, 7])
                        var uniqueNoteNumbers = noteNumbers.Distinct().OrderBy(n => n).ToList();
                        
                        if (uniqueNoteNumbers.Count != noteNumbers.Count)
                        {
                            _logger.LogDebug($"Consolidated duplicates for sheet {sheetName}: {noteNumbers.Count} → {uniqueNoteNumbers.Count} unique notes");
                        }

                        if (uniqueNoteNumbers.Count > 0)
                        {
                            var mapping = new SheetNoteMapping
                            {
                                SheetName = sheetName,
                                NoteNumbers = uniqueNoteNumbers
                            };

                            mappings.Add(mapping);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Failed to parse Excel notes row: {ex.Message}");
                        continue;
                    }
                }

                _logger.LogInformation($"Successfully read Excel notes mappings for {mappings.Count} sheets");
                return mappings;
            }
            catch (Exception ex)
            {
                var errorMsg = $"Failed to read Excel notes from {filePath}: {ex.Message}";
                _logger.LogError(errorMsg, ex);
                return new List<SheetNoteMapping>();
            }
        });
    }

    public async Task<bool> FileExistsAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            var exists = File.Exists(filePath);
            _logger.LogDebug($"File existence check for {filePath}: {exists}");
            return exists;
        });
    }

    public async Task<string[]> GetWorksheetNamesAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            _logger.LogDebug($"Getting worksheet names from {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Excel file not found: {filePath}");
                    return Array.Empty<string>();
                }

                using var workbook = new XLWorkbook(filePath);
                var worksheetNames = workbook.Worksheets.Select(ws => ws.Name).ToArray();
                
                _logger.LogDebug($"Found {worksheetNames.Length} worksheets: {string.Join(", ", worksheetNames)}");
                return worksheetNames;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to get worksheet names from {filePath}: {ex.Message}", ex);
                return Array.Empty<string>();
            }
        });
    }

    public async Task<string[]> GetTableNamesAsync(string filePath, string worksheetName)
    {
        return await Task.Run(() =>
        {
            _logger.LogDebug($"Getting table names from worksheet '{worksheetName}' in {filePath}");
            
            try
            {
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Excel file not found: {filePath}");
                    return Array.Empty<string>();
                }

                using var workbook = new XLWorkbook(filePath);
                
                if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                {
                    _logger.LogError($"Worksheet '{worksheetName}' not found in {filePath}");
                    return Array.Empty<string>();
                }

                var tableNames = worksheet.Tables.Select(t => t.Name).ToArray();
                
                _logger.LogDebug($"Found {tableNames.Length} tables in worksheet '{worksheetName}': {string.Join(", ", tableNames)}");
                return tableNames;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to get table names from worksheet '{worksheetName}' in {filePath}: {ex.Message}", ex);
                return Array.Empty<string>();
            }
        });
    }

    /// <summary>
    /// Finds a named table across all worksheets in the workbook
    /// </summary>
    private static IXLTable? FindTableByName(XLWorkbook workbook, string tableName)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            if (worksheet.Tables.TryGetTable(tableName, out var table))
            {
                return table;
            }
        }
        return null;
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
                _logger?.LogDebug("ExcelReaderService disposed");
            }
            _disposed = true;
        }
    }
}