using System.Text.Json.Serialization;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class ConstructionNote
{
    [JsonPropertyName("number")]
    public int Number { get; set; }
    
    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;
    
    [JsonPropertyName("series")]
    public string Series { get; set; } = string.Empty;
    
    public string Category { get; set; } = string.Empty;
    public bool IsActive { get; set; } = true;
    public DateTime? LastUpdated { get; set; }
}


public class ConstructionNoteBlock
{
    public string BlockName { get; set; } = string.Empty;
    public int Number { get; set; }
    public string Note { get; set; } = string.Empty;
    public bool IsVisible { get; set; }
    public string SheetName { get; set; } = string.Empty;
    public string Series { get; set; } = string.Empty;
}

public enum ConstructionNoteMode
{
    AutoNotes,
    ExcelNotes
}