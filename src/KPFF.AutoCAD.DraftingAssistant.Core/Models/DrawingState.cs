namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

/// <summary>
/// Represents the different states a drawing can be in relative to AutoCAD
/// </summary>
public enum DrawingState
{
    /// <summary>Drawing is currently active in AutoCAD</summary>
    Active,
    
    /// <summary>Drawing is open but not the active document</summary>
    Inactive,
    
    /// <summary>Drawing file exists but is not currently open</summary>
    Closed,
    
    /// <summary>Drawing file was not found at the expected location</summary>
    NotFound,
    
    /// <summary>An error occurred while trying to determine the drawing state</summary>
    Error
}