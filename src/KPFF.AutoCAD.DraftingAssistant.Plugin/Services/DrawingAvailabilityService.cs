using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific implementation of drawing availability service
/// Ensures a drawing is available before sheet processing operations
/// </summary>
public class DrawingAvailabilityService : IDrawingAvailabilityService
{
    private readonly ILogger _logger;

    public DrawingAvailabilityService(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Ensures a drawing is available for operations. Creates a new drawing if needed.
    /// </summary>
    public bool EnsureDrawingAvailable(bool isPlottingOperation = false)
    {
        try
        {
            var docManager = Application.DocumentManager;
            
            // Scenario 1: No drawings open - always create one
            if (docManager.Count == 0)
            {
                _logger.LogInformation("No drawings open - creating new drawing for operations");
                return CreateNewDrawing();
            }
            
            // Scenario 2: Only Drawing1.dwg open during plotting - create new drawing
            if (isPlottingOperation && docManager.Count == 1)
            {
                var singleDoc = GetSingleDocument();
                if (singleDoc != null && IsDefaultDrawing(singleDoc.Name))
                {
                    _logger.LogInformation($"Found only default drawing '{singleDoc.Name}' during plotting operation - creating new drawing");
                    return CreateNewDrawing();
                }
            }
            
            // Drawing context already available
            _logger.LogDebug($"Drawing context available - {docManager.Count} document(s) open");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to ensure drawing availability: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Checks if a drawing is currently available for operations
    /// </summary>
    public bool IsDrawingAvailable()
    {
        try
        {
            var docManager = Application.DocumentManager;
            return docManager.Count > 0;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to check drawing availability: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Creates a new drawing using the default template
    /// </summary>
    private bool CreateNewDrawing()
    {
        try
        {
            var docManager = Application.DocumentManager;
            var newDoc = docManager.Add("acad.dwt");
            
            if (newDoc != null)
            {
                // Make the new document active to ensure proper context
                docManager.MdiActiveDocument = newDoc;
                _logger.LogInformation($"Successfully created and activated new drawing '{newDoc.Name}'");
                return true;
            }
            else
            {
                _logger.LogWarning("DocumentManager.Add returned null - failed to create new drawing");
                return false;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to create new drawing: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets the single document if only one is open
    /// </summary>
    private Document? GetSingleDocument()
    {
        try
        {
            var docManager = Application.DocumentManager;
            if (docManager.Count == 1)
            {
                foreach (Document doc in docManager)
                {
                    return doc;
                }
            }
            return null;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Checks if the given document name is Drawing1.dwg specifically.
    /// Only Drawing1.dwg is automatically removed by AutoCAD, other Drawing[n].dwg files persist.
    /// </summary>
    private bool IsDefaultDrawing(string documentName)
    {
        if (string.IsNullOrEmpty(documentName))
            return false;
            
        // Only check for "Drawing1.dwg" specifically - this is the only one AutoCAD removes automatically
        var fileName = System.IO.Path.GetFileNameWithoutExtension(documentName);
        return string.Equals(fileName, "Drawing1", StringComparison.OrdinalIgnoreCase);
    }
}