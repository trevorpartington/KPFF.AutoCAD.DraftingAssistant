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
    /// Path to the title block DWG file for insertion. This file IS the block (created with WBLOCK).
    /// The system will use whatever block name exists in this file.
    /// </summary>
    public string TitleBlockFilePath { get; set; } = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\Blocks\DBRT-TTLB-ATT.dwg";
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

/// <summary>
/// Represents title block attribute data for external drawing updates
/// Similar to ConstructionNoteData but for title block attributes
/// </summary>
public class TitleBlockAttributeData
{
    public string AttributeName { get; set; } = string.Empty;
    public string AttributeValue { get; set; } = string.Empty;

    public TitleBlockAttributeData() { }

    public TitleBlockAttributeData(string attributeName, string attributeValue)
    {
        AttributeName = attributeName;
        AttributeValue = attributeValue;
    }
}