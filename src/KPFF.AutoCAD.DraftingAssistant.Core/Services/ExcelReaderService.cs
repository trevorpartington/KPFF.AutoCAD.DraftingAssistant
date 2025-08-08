using OfficeOpenXml;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using Microsoft.Extensions.Logging;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

public class ExcelReaderService : IExcelReader
{
    private readonly IApplicationLogger _logger;

    public ExcelReaderService(IApplicationLogger logger)
    {
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
    {
        try
        {
            var sheets = new List<SheetInfo>();
            
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[config.Worksheets.Sheets];
            if (worksheet == null)
            {
                _logger.LogWarning($"Worksheet '{config.Worksheets.Sheets}' not found in {filePath}");
                return sheets;
            }

            var table = worksheet.Tables[config.Tables.SheetIndex];
            if (table == null)
            {
                _logger.LogWarning($"Table '{config.Tables.SheetIndex}' not found in worksheet '{config.Worksheets.Sheets}'");
                return sheets;
            }

            var headers = table.Columns.Select(c => c.Name).ToArray();
            var startRow = table.Address.Start.Row + 1; // Skip header
            var endRow = table.Address.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                var sheet = new SheetInfo();
                
                for (int col = 0; col < headers.Length; col++)
                {
                    var headerName = headers[col];
                    var cellValue = worksheet.Cells[row, table.Address.Start.Column + col].Text;

                    switch (headerName.ToLowerInvariant())
                    {
                        case "sheet":
                        case "sheet name":
                        case "sheetname":
                            sheet.SheetName = cellValue;
                            break;
                        case "file":
                        case "filename":
                        case "dwg file":
                            sheet.DWGFileName = cellValue;
                            break;
                        case "title":
                        case "drawing title":
                        case "drawingtitle":
                            sheet.DrawingTitle = cellValue;
                            break;
                        case "project number":
                        case "projectnumber":
                        case "project":
                            sheet.ProjectNumber = cellValue;
                            break;
                        case "scale":
                            sheet.Scale = cellValue;
                            break;
                        case "sheet type":
                        case "sheettype":
                        case "type":
                            sheet.SheetType = cellValue;
                            break;
                        case "designed by":
                        case "designedby":
                        case "designer":
                            sheet.DesignedBy = cellValue;
                            break;
                        case "checked by":
                        case "checkedby":
                        case "checker":
                            sheet.CheckedBy = cellValue;
                            break;
                        case "drawn by":
                        case "drawnby":
                        case "drafter":
                            sheet.DrawnBy = cellValue;
                            break;
                        case "issue date":
                        case "issuedate":
                        case "date":
                            if (DateTime.TryParse(cellValue, out var date))
                                sheet.IssueDate = date;
                            break;
                        default:
                            sheet.AdditionalProperties[headerName] = cellValue;
                            break;
                    }
                }

                if (!string.IsNullOrEmpty(sheet.SheetName))
                    sheets.Add(sheet);
            }

            _logger.LogInformation($"Read {sheets.Count} sheets from {filePath}");
            return sheets;
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
            var notes = new List<ConstructionNote>();
            var tableName = string.Format(config.Tables.NotesPattern, series.ToUpperInvariant());

            using var package = new ExcelPackage(new FileInfo(filePath));
            
            // Try to find worksheet with notes - series-specific first, then config
            var worksheet = package.Workbook.Worksheets[$"{series} Notes"] ?? 
                           package.Workbook.Worksheets[series] ??
                           package.Workbook.Worksheets[config.Worksheets.Notes];
            
            if (worksheet == null)
            {
                _logger.LogWarning($"Notes worksheet '{config.Worksheets.Notes}' not found for series {series}");
                return notes;
            }

            var table = worksheet.Tables[tableName];
            if (table == null)
            {
                _logger.LogWarning($"Table '{tableName}' not found in notes worksheet");
                return notes;
            }

            var startRow = table.Address.Start.Row + 1; // Skip header
            var endRow = table.Address.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                var numberText = worksheet.Cells[row, table.Address.Start.Column].Text;
                var noteText = worksheet.Cells[row, table.Address.Start.Column + 1].Text;

                if (int.TryParse(numberText, out var number) && !string.IsNullOrEmpty(noteText))
                {
                    notes.Add(new ConstructionNote
                    {
                        Number = number,
                        Text = noteText,
                        Series = series,
                        IsActive = true
                    });
                }
            }

            _logger.LogInformation($"Read {notes.Count} construction notes for series {series}");
            return notes;
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
            var mappings = new List<SheetNoteMapping>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            // Excel notes are in a separate worksheet, not the series-specific notes worksheet
            var worksheet = package.Workbook.Worksheets["Excel Notes"];
            if (worksheet == null)
            {
                _logger.LogWarning($"Excel Notes worksheet not found in {filePath}");
                return mappings;
            }

            var table = worksheet.Tables[config.Tables.ExcelNotes];
            if (table == null)
            {
                _logger.LogWarning($"Table '{config.Tables.ExcelNotes}' not found in notes worksheet");
                return mappings;
            }

            var startRow = table.Address.Start.Row + 1; // Skip header
            var endRow = table.Address.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                var sheetName = worksheet.Cells[row, table.Address.Start.Column].Text;
                if (string.IsNullOrEmpty(sheetName))
                    continue;

                var mapping = new SheetNoteMapping { SheetName = sheetName };

                // Read note numbers from columns 2-25 (up to 24 notes)
                for (int col = 1; col < Math.Min(25, table.Address.Columns); col++)
                {
                    var noteText = worksheet.Cells[row, table.Address.Start.Column + col].Text;
                    if (int.TryParse(noteText, out var noteNumber))
                        mapping.NoteNumbers.Add(noteNumber);
                }

                if (mapping.NoteNumbers.Count > 0)
                    mappings.Add(mapping);
            }

            _logger.LogInformation($"Read excel notes mappings for {mappings.Count} sheets");
            return mappings;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading excel notes from {filePath}: {ex.Message}");
            throw;
        }
    }

    public async Task<bool> FileExistsAsync(string filePath)
    {
        return File.Exists(filePath);
    }

    public async Task<string[]> GetWorksheetNamesAsync(string filePath)
    {
        try
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            return package.Workbook.Worksheets.Select(ws => ws.Name).ToArray();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting worksheet names from {filePath}: {ex.Message}");
            return Array.Empty<string>();
        }
    }

    public async Task<string[]> GetTableNamesAsync(string filePath, string worksheetName)
    {
        try
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[worksheetName];
            return worksheet?.Tables.Select(t => t.Name).ToArray() ?? Array.Empty<string>();
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting table names from {filePath}:{worksheetName}: {ex.Message}");
            return Array.Empty<string>();
        }
    }
}