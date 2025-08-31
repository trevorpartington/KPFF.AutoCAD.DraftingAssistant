using System.Text.Json.Serialization;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class SheetNoteMapping
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;
    
    [JsonPropertyName("noteNumbers")]
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