using System.Text.Json.Serialization;

namespace KPFF.AutoCAD.ExcelReader.Models;

// Request/Response models for named pipe communication
public class ExcelRequest
{
    [JsonPropertyName("operation")]
    public string Operation { get; set; } = string.Empty;
    
    [JsonPropertyName("filePath")]
    public string FilePath { get; set; } = string.Empty;
    
    [JsonPropertyName("parameters")]
    public Dictionary<string, object> Parameters { get; set; } = new();
}

public class ExcelResponse
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }
    
    [JsonPropertyName("data")]
    public object? Data { get; set; }
    
    [JsonPropertyName("error")]
    public string? Error { get; set; }
}

// Excel data models matching the Core project
public class SheetInfo
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;
    
    [JsonPropertyName("dwgFileName")]
    public string DWGFileName { get; set; } = string.Empty;
    
    [JsonPropertyName("drawingTitle")]
    public string DrawingTitle { get; set; } = string.Empty;
}

public class ConstructionNote
{
    [JsonPropertyName("number")]
    public int Number { get; set; }
    
    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;
    
    [JsonPropertyName("series")]
    public string Series { get; set; } = string.Empty;
}

public class SheetNoteMapping
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;
    
    [JsonPropertyName("noteNumbers")]
    public List<int> NoteNumbers { get; set; } = new();
}