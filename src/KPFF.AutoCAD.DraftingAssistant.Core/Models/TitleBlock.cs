using System.Text.Json.Serialization;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class TitleBlockInfo
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;
    
    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;
    
    [JsonPropertyName("attributes")]
    public Dictionary<string, string> Attributes { get; set; } = new();
    
    public bool IsActive { get; set; } = true;
    public DateTime? LastUpdated { get; set; }
}

public class TitleBlockConfiguration
{
    /// <summary>
    /// Regex pattern to match title block block names (e.g., "^TB_ATT$")
    /// </summary>
    public string TitleBlockPattern { get; set; } = @"^TB_ATT$";
    
    /// <summary>
    /// Maximum number of attributes to process per title block
    /// </summary>
    public int MaxAttributesPerTitleBlock { get; set; } = 50;
    
    /// <summary>
    /// Property name for visibility control (if applicable)
    /// </summary>
    public string VisibilityPropertyName { get; set; } = "Visibility";
    
    /// <summary>
    /// Whether to log attribute mapping details during processing
    /// </summary>
    public bool LogAttributeMapping { get; set; } = false;
}

public class TitleBlockMapping
{
    public string SheetName { get; set; } = string.Empty;
    public Dictionary<string, string> AttributeValues { get; set; } = new();
    
    public TitleBlockMapping() { }
    
    public TitleBlockMapping(string sheetName, Dictionary<string, string> attributeValues)
    {
        SheetName = sheetName;
        AttributeValues = attributeValues ?? new();
    }
}