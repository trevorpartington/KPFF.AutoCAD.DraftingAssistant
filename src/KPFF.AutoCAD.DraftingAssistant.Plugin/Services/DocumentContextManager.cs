using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific implementation of document context management
/// </summary>
public class DocumentContextManager : IDocumentContextManager
{
    private readonly ILogger _logger;
    private static string? _activeProtectionDocumentName = null;
    private static readonly object _protectionLock = new object();

    public DocumentContextManager(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Ensures document context exists for closed drawing operations.
    /// Creates a protection document if needed when no documents are open (start tab scenario).
    /// Uses singleton pattern to prevent multiple protection documents.
    /// </summary>
    /// <returns>Context token that should be disposed when operations are complete</returns>
    public IDocumentContextToken? EnsureContextForClosedDrawingOperations()
    {
        lock (_protectionLock)
        {
            try
            {
                var docManager = Application.DocumentManager;
                
                // Check if we already have a valid protection document
                if (!string.IsNullOrEmpty(_activeProtectionDocumentName))
                {
                    var existingDoc = FindDocumentByName(_activeProtectionDocumentName);
                    if (existingDoc != null)
                    {
                        _logger.LogDebug($"Using existing protection document '{_activeProtectionDocumentName}'");
                        return new DocumentContextToken(_activeProtectionDocumentName, _logger, false); // Don't track disposal for reused documents
                    }
                    else
                    {
                        _logger.LogWarning($"Protection document '{_activeProtectionDocumentName}' no longer exists, will create new one");
                        _activeProtectionDocumentName = null;
                    }
                }
                
                // Only create protection document if no documents are open (start tab scenario)
                if (docManager.Count == 0)
                {
                    _logger.LogDebug("No documents open (start tab) - creating protection document for closed drawing operations");
                    var doc = docManager.Add("acad.dwt");
                    
                    if (doc != null)
                    {
                        // Make the protection document active to ensure proper context
                        docManager.MdiActiveDocument = doc;
                        _activeProtectionDocumentName = doc.Name;
                        
                        _logger.LogInformation($"Created and activated protection document '{doc.Name}' to prevent context loss during closed drawing operations");
                        return new DocumentContextToken(doc.Name, _logger, true); // Track disposal for new documents
                    }
                    else
                    {
                        _logger.LogWarning("DocumentManager.Add returned null");
                        return null;
                    }
                }
                
                _logger.LogDebug("Document context already exists - no protection needed");
                return null;
            }
            catch (System.Exception ex)
            {
                _logger.LogError($"Failed to create document protection: {ex.Message}", ex);
                return null;
            }
        }
    }

    /// <summary>
    /// Checks if document context protection is needed for closed drawing operations.
    /// Protection needed when no documents are open (start tab scenario).
    /// </summary>
    /// <returns>True if protection is needed, false otherwise</returns>
    public bool NeedsContextProtection()
    {
        try
        {
            var docManager = Application.DocumentManager;
            return docManager.Count == 0;
        }
        catch (System.Exception)
        {
            return true; // Err on the side of caution
        }
    }

    /// <summary>
    /// Finds a document by name in the current document manager
    /// </summary>
    /// <param name="documentName">Name of the document to find</param>
    /// <returns>Document if found, null otherwise</returns>
    private static Document? FindDocumentByName(string documentName)
    {
        try
        {
            var docManager = Application.DocumentManager;
            foreach (Document doc in docManager)
            {
                if (string.Equals(doc.Name, documentName, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }
            return null;
        }
        catch (System.Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Clears the active protection document tracking when plugin terminates
    /// </summary>
    public static void ClearProtectionTracking()
    {
        lock (_protectionLock)
        {
            _activeProtectionDocumentName = null;
        }
    }
}

/// <summary>
/// Token representing a document context protection
/// </summary>
public class DocumentContextToken : IDocumentContextToken
{
    private readonly ILogger _logger;
    private readonly bool _shouldTrackDisposal;
    private bool _disposed = false;

    public string? DocumentName { get; }

    public DocumentContextToken(string? documentName, ILogger logger, bool shouldTrackDisposal = true)
    {
        DocumentName = documentName;
        _logger = logger;
        _shouldTrackDisposal = shouldTrackDisposal;
    }

    public void Dispose()
    {
        if (!_disposed && DocumentName != null && _shouldTrackDisposal)
        {
            _logger.LogDebug($"Protection document '{DocumentName}' context token disposed - document remains open for user");
            // Document will remain open for user as intended - no cleanup needed
            // But we track that this particular token has been disposed
            _disposed = true;
        }
    }
}