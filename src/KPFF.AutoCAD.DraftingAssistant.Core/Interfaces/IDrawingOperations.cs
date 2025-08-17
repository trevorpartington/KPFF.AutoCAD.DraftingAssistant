using KPFF.AutoCAD.DraftingAssistant.Core.Models;

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
    /// Updates multiple construction note blocks for a sheet
    /// Resets all blocks first, then sets the specified ones
    /// </summary>
    Task<bool> UpdateConstructionNoteBlocksAsync(string sheetName, List<int> noteNumbers, List<ConstructionNote> notes, ProjectConfiguration config);

    /// <summary>
    /// Validates that construction note blocks exist for the specified sheet
    /// </summary>
    Task<bool> ValidateNoteBlocksExistAsync(string sheetName, ProjectConfiguration config);

    /// <summary>
    /// Resets all construction note blocks for a sheet to invisible/empty state
    /// </summary>
    Task<bool> ResetConstructionNoteBlocksAsync(string sheetName, ProjectConfiguration config);
}