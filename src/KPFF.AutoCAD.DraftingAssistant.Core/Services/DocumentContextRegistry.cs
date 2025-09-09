using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Registry for document context management services.
/// Allows the Plugin project to register the AutoCAD-specific implementation 
/// and the Core services to access it without circular dependencies.
/// </summary>
public static class DocumentContextRegistry
{
    private static IDocumentContextManager? _instance;

    /// <summary>
    /// Registers the document context manager implementation.
    /// Should be called once during plugin initialization.
    /// </summary>
    /// <param name="manager">The document context manager implementation</param>
    public static void Register(IDocumentContextManager manager)
    {
        _instance = manager;
    }

    /// <summary>
    /// Gets the registered document context manager instance.
    /// Returns null if no implementation has been registered.
    /// </summary>
    public static IDocumentContextManager? Instance => _instance;

    /// <summary>
    /// Unregisters the current document context manager.
    /// Should be called during plugin cleanup if needed.
    /// </summary>
    public static void Unregister()
    {
        _instance = null;
    }
}