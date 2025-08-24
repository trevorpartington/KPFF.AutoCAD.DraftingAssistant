using System.Text.Json.Serialization;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class ProjectConfiguration
{
    public string ProjectName { get; set; } = string.Empty;
    public string ClientName { get; set; } = string.Empty;
    public string ProjectIndexFilePath { get; set; } = string.Empty;
    public string ProjectDWGFilePath { get; set; } = string.Empty;
    public SheetNamingConfiguration SheetNaming { get; set; } = new();
    public TableConfiguration Tables { get; set; } = new();
    public ConstructionNotesConfiguration ConstructionNotes { get; set; } = new();
    
    /// <summary>
    /// Runtime-only property for selected sheets. Not serialized to JSON.
    /// </summary>
    [JsonIgnore]
    public List<SheetInfo> SelectedSheets { get; set; } = new();
}

public class SheetNamingConfiguration
{
    public string Pattern { get; set; } = @"([A-Z]{1,3})-?(\d{1,3})";
    public int SeriesGroup { get; set; } = 1;
    public int NumberGroup { get; set; } = 2;
    public string[] Examples { get; set; } = Array.Empty<string>();
}


public class TableConfiguration
{
    public string SheetIndex { get; set; } = "SHEET_INDEX";
    public string ExcelNotes { get; set; } = "EXCEL_NOTES";
    public string NotesPattern { get; set; } = "{0}_NOTES";
}

public class ConstructionNotesConfiguration
{
    public string MultileaderStyleName { get; set; } = string.Empty;
    public string NoteBlockPattern { get; set; } = "{0}-NT{1:D2}";
    public int MaxNotesPerSheet { get; set; } = 24;
    public ConstructionNoteAttributes Attributes { get; set; } = new();
    public string VisibilityPropertyName { get; set; } = "Visibility";
}

public class ConstructionNoteAttributes
{
    public string NumberAttribute { get; set; } = "Number";
    public string NoteAttribute { get; set; } = "Note";
}