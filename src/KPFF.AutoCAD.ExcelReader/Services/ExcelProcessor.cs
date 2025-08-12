using OfficeOpenXml;
using KPFF.AutoCAD.ExcelReader.Models;
using System.Text.Json;

namespace KPFF.AutoCAD.ExcelReader.Services;

public class ExcelProcessor
{
    public ExcelProcessor()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<ExcelResponse> ProcessRequestAsync(ExcelRequest request)
    {
        try
        {
            Console.WriteLine($"=== Processing request: {request.Operation} ===");
            Console.WriteLine($"File path: {request.FilePath}");
            Console.WriteLine($"Parameters count: {request.Parameters.Count}");
            
            var response = request.Operation.ToLowerInvariant() switch
            {
                "readsheetindex" => await ReadSheetIndexAsync(request),
                "readconstructionnotes" => await ReadConstructionNotesAsync(request),
                "readexcelnotes" => await ReadExcelNotesAsync(request),
                "ping" => new ExcelResponse { Success = true, Data = "pong" },
                _ => new ExcelResponse { Success = false, Error = $"Unknown operation: {request.Operation}" }
            };
            
            Console.WriteLine($"=== Request completed: {request.Operation} ===");
            Console.WriteLine($"Success: {response.Success}");
            Console.WriteLine($"Error: {response.Error ?? "None"}");
            Console.WriteLine($"Data type: {response.Data?.GetType().Name ?? "Null"}");
            
            return response;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"!!! CRITICAL ERROR in ProcessRequestAsync !!!");
            Console.WriteLine($"Exception type: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            
            return new ExcelResponse 
            { 
                Success = false, 
                Error = $"Critical error processing {request.Operation}: {ex.Message}" 
            };
        }
    }

    private async Task<ExcelResponse> ReadSheetIndexAsync(ExcelRequest request)
    {
        if (!File.Exists(request.FilePath))
            return new ExcelResponse { Success = false, Error = $"File not found: {request.FilePath}" };

        var sheets = new List<SheetInfo>();

        using var package = new ExcelPackage(new FileInfo(request.FilePath));
        
        // Find SHEET_INDEX table across all worksheets
        var table = package.Workbook.Worksheets
            .SelectMany(ws => ws.Tables)
            .FirstOrDefault(t => t.Name.Equals("SHEET_INDEX", StringComparison.OrdinalIgnoreCase));
            
        if (table == null)
            return new ExcelResponse { Success = false, Error = "SHEET_INDEX table not found" };
        
        var worksheet = table.WorkSheet;
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
                }
            }

            if (!string.IsNullOrEmpty(sheet.SheetName))
                sheets.Add(sheet);
        }

        return new ExcelResponse { Success = true, Data = sheets };
    }

    private async Task<ExcelResponse> ReadConstructionNotesAsync(ExcelRequest request)
    {
        try
        {
            Console.WriteLine($"ReadConstructionNotesAsync - Starting...");
            Console.WriteLine($"File path: {request.FilePath}");
            Console.WriteLine($"Parameters count: {request.Parameters.Count}");
            
            if (!File.Exists(request.FilePath))
            {
                var error = $"File not found: {request.FilePath}";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }

            foreach (var param in request.Parameters)
            {
                Console.WriteLine($"  Parameter: '{param.Key}' = '{param.Value}' (Type: {param.Value?.GetType().Name})");
            }
            
            if (!request.Parameters.TryGetValue("series", out var seriesObj))
            {
                var error = "Series parameter required";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }
                
            string? series = null;
            if (seriesObj is string str)
                series = str;
            else if (seriesObj is JsonElement element && element.ValueKind == JsonValueKind.String)
                series = element.GetString();
                
            if (string.IsNullOrEmpty(series))
            {
                var error = "Series parameter must be a non-empty string";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }

            Console.WriteLine($"Series parameter resolved to: '{series}'");

            var notes = new List<ConstructionNote>();
            var tableName = $"{series.ToUpperInvariant()}_NOTES";
            Console.WriteLine($"Looking for table: '{tableName}'");

            Console.WriteLine("Opening Excel package...");
            using var package = new ExcelPackage(new FileInfo(request.FilePath));
            Console.WriteLine("Excel package opened successfully");
            
            // List all available tables for debugging
            Console.WriteLine("Available tables in workbook:");
            foreach (var ws in package.Workbook.Worksheets)
            {
                Console.WriteLine($"  Worksheet: {ws.Name}");
                foreach (var t in ws.Tables)
                {
                    Console.WriteLine($"    Table: '{t.Name}' (Range: {t.Address})");
                }
            }
            
            // Find series-specific notes table
            var table = package.Workbook.Worksheets
                .SelectMany(ws => ws.Tables)
                .FirstOrDefault(t => t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase));
                
            if (table == null)
            {
                var error = $"Table '{tableName}' not found!";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }
            
            Console.WriteLine($"Found table '{table.Name}' in worksheet '{table.WorkSheet.Name}'");
            Console.WriteLine($"Table range: {table.Address}");
            
            var worksheet = table.WorkSheet;
            var startRow = table.Address.Start.Row + 1; // Skip header
            var endRow = table.Address.End.Row;
            
            Console.WriteLine($"Reading rows {startRow} to {endRow}");

            for (int row = startRow; row <= endRow; row++)
            {
                var numberText = worksheet.Cells[row, table.Address.Start.Column].Text;
                var noteText = worksheet.Cells[row, table.Address.Start.Column + 1].Text;
                
                Console.WriteLine($"Row {row}: Number='{numberText}', Note='{noteText}'");

                if (int.TryParse(numberText, out var number) && !string.IsNullOrEmpty(noteText))
                {
                    notes.Add(new ConstructionNote
                    {
                        Number = number,
                        Text = noteText,
                        Series = series
                    });
                    Console.WriteLine($"  Added note #{number}: {noteText}");
                }
                else
                {
                    Console.WriteLine($"  Skipped row {row} - invalid data");
                }
            }

            Console.WriteLine($"Total notes found: {notes.Count}");
            
            // DEBUG: Show each note in detail before returning
            Console.WriteLine("\nDEBUG: Final notes being returned:");
            foreach (var note in notes)
            {
                Console.WriteLine($"  Note #{note.Number} (Type: {note.Number.GetType().Name}): '{note.Text}' (Series: '{note.Series}')");
            }
            
            Console.WriteLine("ReadConstructionNotesAsync - Completed successfully");
            return new ExcelResponse { Success = true, Data = notes };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"!!! ERROR in ReadConstructionNotesAsync !!!");
            Console.WriteLine($"Exception type: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            
            return new ExcelResponse 
            { 
                Success = false, 
                Error = $"Error reading construction notes: {ex.Message}" 
            };
        }
    }

    private async Task<ExcelResponse> ReadExcelNotesAsync(ExcelRequest request)
    {
        try
        {
            Console.WriteLine($"ReadExcelNotesAsync - Starting...");
            Console.WriteLine($"File path: {request.FilePath}");
            
            if (!File.Exists(request.FilePath))
            {
                var error = $"File not found: {request.FilePath}";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }

            var mappings = new List<SheetNoteMapping>();

            using var package = new ExcelPackage(new FileInfo(request.FilePath));
            Console.WriteLine("Excel package opened successfully");
            
            // Find EXCEL_NOTES table
            var table = package.Workbook.Worksheets
                .SelectMany(ws => ws.Tables)
                .FirstOrDefault(t => t.Name.Equals("EXCEL_NOTES", StringComparison.OrdinalIgnoreCase));
                
            if (table == null)
            {
                var error = "EXCEL_NOTES table not found";
                Console.WriteLine($"ERROR: {error}");
                return new ExcelResponse { Success = false, Error = error };
            }
            
            Console.WriteLine($"Found EXCEL_NOTES table in worksheet '{table.WorkSheet.Name}'");
            Console.WriteLine($"Table range: {table.Address}");
            
            var worksheet = table.WorkSheet;
            var startRow = table.Address.Start.Row + 1; // Skip header
            var endRow = table.Address.End.Row;
            
            Console.WriteLine($"Reading rows {startRow} to {endRow}");

            for (int row = startRow; row <= endRow; row++)
            {
                var sheetName = worksheet.Cells[row, table.Address.Start.Column].Text;
                Console.WriteLine($"Row {row}: Sheet name = '{sheetName}'");
                
                if (string.IsNullOrEmpty(sheetName))
                {
                    Console.WriteLine($"  Skipped row {row} - empty sheet name");
                    continue;
                }

                var mapping = new SheetNoteMapping { SheetName = sheetName };
                Console.WriteLine($"  Created mapping for sheet: '{mapping.SheetName}'");

                // Read note numbers from columns 2-25 (up to 24 notes)
                for (int col = 1; col < Math.Min(25, table.Address.Columns); col++)
                {
                    var noteText = worksheet.Cells[row, table.Address.Start.Column + col].Text;
                    if (int.TryParse(noteText, out var noteNumber))
                    {
                        mapping.NoteNumbers.Add(noteNumber);
                        Console.WriteLine($"    Added note number: {noteNumber}");
                    }
                }

                if (mapping.NoteNumbers.Count > 0)
                {
                    mappings.Add(mapping);
                    Console.WriteLine($"  Added mapping with {mapping.NoteNumbers.Count} notes");
                }
                else
                {
                    Console.WriteLine($"  Skipped mapping - no note numbers found");
                }
            }

            Console.WriteLine($"Total mappings found: {mappings.Count}");
            foreach (var mapping in mappings)
            {
                Console.WriteLine($"  Sheet: '{mapping.SheetName}' -> Notes: [{string.Join(", ", mapping.NoteNumbers)}]");
            }
            
            Console.WriteLine("ReadExcelNotesAsync - Completed successfully");
            return new ExcelResponse { Success = true, Data = mappings };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"!!! ERROR in ReadExcelNotesAsync !!!");
            Console.WriteLine($"Exception type: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            
            return new ExcelResponse 
            { 
                Success = false, 
                Error = $"Error reading Excel notes: {ex.Message}" 
            };
        }
    }
}