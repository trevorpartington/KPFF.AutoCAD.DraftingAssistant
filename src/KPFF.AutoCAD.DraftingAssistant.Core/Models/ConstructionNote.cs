namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class ConstructionNote
{
    public int Number { get; set; }
    public string Text { get; set; } = string.Empty;
    public string Series { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public bool IsActive { get; set; } = true;
    public DateTime? LastUpdated { get; set; }
}

public class SheetNoteMapping
{
    public string SheetName { get; set; } = string.Empty;
    public List<int> NoteNumbers { get; set; } = new();
    public string Series => ExtractSeries(SheetName);

    private string ExtractSeries(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return string.Empty;

        var dashIndex = sheetName.IndexOf('-');
        if (dashIndex > 0)
            return sheetName[..dashIndex];

        var digitIndex = sheetName.ToList().FindIndex(char.IsDigit);
        return digitIndex > 0 ? sheetName[..digitIndex] : sheetName;
    }
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