using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for drawing operations including construction note block management
/// Abstracts AutoCAD-specific operations for better testability and separation of concerns
/// </summary>
public interface IDrawingOperations : IDisposable
{
    /// <summary>
    /// Gets all construction note blocks for a specific layout/sheet
    /// </summary>
    Task<List<ConstructionNoteBlock>> GetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config);

    /// <summary>
    /// Updates a single construction note block with note number and text
    /// </summary>
    Task<bool> UpdateConstructionNoteBlockAsync(string sheetName, int blockIndex, int noteNumber, string noteText, ProjectConfiguration config);


    /// <summary>
    /// Validates that construction note blocks exist for the specified sheet
    /// </summary>
    Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config);

    /// <summary>
    /// Sets construction notes using a dictionary-based clear-and-fill approach
    /// Clears all blocks first, then fills with provided note data in sequential order
    /// </summary>
    Task<bool> SetConstructionNotesAsync(string sheetName, Dictionary<int, string> noteData, ProjectConfiguration config);

    /// <summary>
    /// Resets all construction note blocks for a sheet to invisible/empty state
    /// </summary>
    Task<bool> ResetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config, CurrentDrawingBlockManager? blockManager = null);

    /// <summary>
    /// Updates title block attributes for a specific layout/sheet
    /// </summary>
    Task UpdateTitleBlockAsync(string sheetName, TitleBlockMapping mapping, ProjectConfiguration config);

    /// <summary>
    /// Validates that title block exists for the specified sheet
    /// </summary>
    Task<bool> ValidateTitleBlockExistsAsync(string sheetName, ProjectConfiguration config);

    /// <summary>
    /// Gets current title block attributes for a sheet
    /// </summary>
    Task<Dictionary<string, string>> GetTitleBlockAttributesAsync(string sheetName, ProjectConfiguration config);
}