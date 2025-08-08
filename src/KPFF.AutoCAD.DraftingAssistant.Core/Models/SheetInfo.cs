namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class SheetInfo
{
    public string SheetName { get; set; } = string.Empty;
    public string DrawingTitle { get; set; } = string.Empty;
    public string ProjectNumber { get; set; } = string.Empty;
    public string Scale { get; set; } = string.Empty;
    public string SheetType { get; set; } = string.Empty;
    public DateTime? IssueDate { get; set; }
    public string DesignedBy { get; set; } = string.Empty;
    public string CheckedBy { get; set; } = string.Empty;
    public string DrawnBy { get; set; } = string.Empty;
    public Dictionary<string, object> AdditionalProperties { get; set; } = new();

    public string Series => ExtractSeries(SheetName);
    public string Number => ExtractNumber(SheetName);

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

    private string ExtractNumber(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return string.Empty;

        var dashIndex = sheetName.IndexOf('-');
        if (dashIndex > 0)
            return sheetName[(dashIndex + 1)..];

        var digitIndex = sheetName.ToList().FindIndex(char.IsDigit);
        return digitIndex >= 0 ? sheetName[digitIndex..] : string.Empty;
    }
}