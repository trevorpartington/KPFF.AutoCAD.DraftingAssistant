namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Manages AutoCAD document context to prevent crashes when operating on closed drawings
/// </summary>
public interface IDocumentContextManager
{
    /// <summary>
    /// Ensures document context exists for closed drawing operations.
    /// Creates a protection document if needed when no documents are open (start tab scenario).
    /// </summary>
    /// <returns>Context token that should be disposed when operations are complete</returns>
    IDocumentContextToken? EnsureContextForClosedDrawingOperations();
    
    /// <summary>
    /// Checks if document context protection is needed for closed drawing operations.
    /// </summary>
    /// <returns>True if protection is needed, false otherwise</returns>
    bool NeedsContextProtection();
}

/// <summary>
/// Represents a document context protection that should be disposed when operations complete
/// </summary>
public interface IDocumentContextToken : IDisposable
{
    /// <summary>
    /// The name of the protection document, if any
    /// </summary>
    string? DocumentName { get; }
}