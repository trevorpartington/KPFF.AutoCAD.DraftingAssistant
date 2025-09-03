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
    public TitleBlockConfiguration TitleBlocks { get; set; } = new();
    public PlottingConfiguration Plotting { get; set; } = new();
    
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
    
    /// <summary>
    /// Manual series length override. 0 = Auto Detect (default)
    /// </summary>
    public int SeriesLength { get; set; } = 0;
}


public class TableConfiguration
{
    public string SheetIndex { get; set; } = "SHEET_INDEX";
    public string ExcelNotes { get; set; } = "EXCEL_NOTES";
    public string NotesPattern { get; set; } = "{0}_NOTES";
}

public class ConstructionNotesConfiguration
{
    public List<string> MultileaderStyleNames { get; set; } = new() { "ML-STYLE-01" };
    public List<NoteBlockConfiguration> NoteBlocks { get; set; } = new() { new NoteBlockConfiguration() };
    public string NoteBlockPattern { get; set; } = @"^NT\d{2}$";
    public int MaxNotesPerSheet { get; set; } = 24;
    public ConstructionNoteAttributes Attributes { get; set; } = new();
    public string VisibilityPropertyName { get; set; } = "Visibility";
    public string NoteBlockFilePath { get; set; } = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\Blocks\NTXX.dwg";
    
    // Legacy property for backward compatibility
    [System.Text.Json.Serialization.JsonIgnore]
    public string MultileaderStyleName 
    { 
        get => MultileaderStyleNames.FirstOrDefault() ?? string.Empty;
        set 
        { 
            if (!string.IsNullOrEmpty(value))
            {
                MultileaderStyleNames = new List<string> { value };
            }
        }
    }
}

public class ConstructionNoteAttributes
{
    public string NumberAttribute { get; set; } = "Number";
    public string NoteAttribute { get; set; } = "Note";
}

public class NoteBlockConfiguration
{
    public string BlockName { get; set; } = "_TagCircle";
    public string AttributeName { get; set; } = "TAGNUMBER";
}

public class PlottingConfiguration
{
    /// <summary>
    /// Output directory for plot files
    /// </summary>
    public string OutputDirectory { get; set; } = string.Empty;
    
}


