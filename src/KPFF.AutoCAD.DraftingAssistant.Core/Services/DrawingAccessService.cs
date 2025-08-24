using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for detecting and managing access to drawings in different states
/// Handles Active (current), Inactive (open but not current), and Closed drawings
/// </summary>
public class DrawingAccessService
{
    private readonly ILogger _logger;

    public DrawingAccessService(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Determines the current state of a drawing file
    /// </summary>
    /// <param name="dwgFilePath">Full path to the DWG file</param>
    /// <returns>The current state of the drawing</returns>
    public DrawingState GetDrawingState(string dwgFilePath)
    {
        try
        {
            _logger.LogDebug($"Checking drawing state for: {dwgFilePath}");

            if (!File.Exists(dwgFilePath))
            {
                _logger.LogWarning($"Drawing file not found: {dwgFilePath}");
                return DrawingState.NotFound;
            }

            // Check if any documents are open in AutoCAD
            var docManager = Application.DocumentManager;
            if (docManager == null)
            {
                _logger.LogDebug("No DocumentManager available, treating as closed drawing");
                return DrawingState.Closed;
            }

            // Get the active document
            var activeDoc = docManager.MdiActiveDocument;
            if (activeDoc != null && IsSameDwgFile(activeDoc.Name, dwgFilePath))
            {
                _logger.LogDebug($"Drawing is currently active: {Path.GetFileName(dwgFilePath)}");
                return DrawingState.Active;
            }

            // Check if the drawing is open but inactive
            foreach (Document doc in docManager)
            {
                if (IsSameDwgFile(doc.Name, dwgFilePath))
                {
                    _logger.LogDebug($"Drawing is open but inactive: {Path.GetFileName(dwgFilePath)}");
                    return DrawingState.Inactive;
                }
            }

            _logger.LogDebug($"Drawing is closed: {Path.GetFileName(dwgFilePath)}");
            return DrawingState.Closed;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error determining drawing state for {dwgFilePath}: {ex.Message}", ex);
            return DrawingState.Error;
        }
    }

    /// <summary>
    /// Gets the Document object for an open drawing (Active or Inactive)
    /// </summary>
    /// <param name="dwgFilePath">Full path to the DWG file</param>
    /// <returns>Document if found, null otherwise</returns>
    public Document? GetOpenDocument(string dwgFilePath)
    {
        try
        {
            var docManager = Application.DocumentManager;
            if (docManager == null) return null;

            // Check active document first
            var activeDoc = docManager.MdiActiveDocument;
            if (activeDoc != null && IsSameDwgFile(activeDoc.Name, dwgFilePath))
            {
                return activeDoc;
            }

            // Check all open documents
            foreach (Document doc in docManager)
            {
                if (IsSameDwgFile(doc.Name, dwgFilePath))
                {
                    return doc;
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting open document for {dwgFilePath}: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Gets all currently open drawings with their states
    /// </summary>
    /// <returns>Dictionary of file paths to their drawing states</returns>
    public Dictionary<string, DrawingState> GetAllOpenDrawings()
    {
        var openDrawings = new Dictionary<string, DrawingState>();

        try
        {
            var docManager = Application.DocumentManager;
            if (docManager == null) return openDrawings;

            var activeDoc = docManager.MdiActiveDocument;
            
            foreach (Document doc in docManager)
            {
                var filePath = doc.Name;
                var state = (doc == activeDoc) ? DrawingState.Active : DrawingState.Inactive;
                openDrawings[filePath] = state;
                
                _logger.LogDebug($"Found open drawing: {Path.GetFileName(filePath)} ({state})");
            }

            _logger.LogInformation($"Found {openDrawings.Count} open drawings");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting all open drawings: {ex.Message}", ex);
        }

        return openDrawings;
    }

    /// <summary>
    /// Attempts to make a drawing active if it's currently open but inactive
    /// </summary>
    /// <param name="dwgFilePath">Full path to the DWG file</param>
    /// <returns>True if successful, false otherwise</returns>
    public bool TryMakeDrawingActive(string dwgFilePath)
    {
        try
        {
            var state = GetDrawingState(dwgFilePath);
            if (state != DrawingState.Inactive)
            {
                _logger.LogDebug($"Drawing {Path.GetFileName(dwgFilePath)} is not inactive (state: {state}), cannot make active");
                return state == DrawingState.Active; // Already active is considered success
            }

            var doc = GetOpenDocument(dwgFilePath);
            if (doc == null)
            {
                _logger.LogWarning($"Could not find open document for {dwgFilePath}");
                return false;
            }

            // Make the document active
            Application.DocumentManager.MdiActiveDocument = doc;
            _logger.LogInformation($"Successfully made drawing active: {Path.GetFileName(dwgFilePath)}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error making drawing active {dwgFilePath}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets the full file path for a drawing given a sheet name and project configuration
    /// </summary>
    /// <param name="sheetName">Name of the sheet (e.g., "ABC-101")</param>
    /// <param name="config">Project configuration</param>
    /// <param name="sheetInfos">List of sheet information from Excel</param>
    /// <returns>Full path to the DWG file, or null if not found</returns>
    public string? GetDrawingFilePath(string sheetName, ProjectConfiguration config, List<SheetInfo> sheetInfos)
    {
        try
        {
            // Find the sheet info for this sheet name
            var sheetInfo = sheetInfos.FirstOrDefault(s => 
                s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheetInfo == null)
            {
                _logger.LogWarning($"Sheet '{sheetName}' not found in sheet index");
                return null;
            }

            if (string.IsNullOrEmpty(sheetInfo.DWGFileName))
            {
                _logger.LogWarning($"No DWG file name specified for sheet '{sheetName}'");
                return null;
            }

            // Combine with project DWG file path
            var dwgFileName = sheetInfo.DWGFileName;
            if (!dwgFileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                dwgFileName += ".dwg";
            }

            var fullPath = Path.Combine(config.ProjectDWGFilePath, dwgFileName);
            
            _logger.LogDebug($"Resolved sheet '{sheetName}' to file path: {fullPath}");
            return fullPath;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting drawing file path for sheet {sheetName}: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Compares two file paths to determine if they refer to the same DWG file
    /// Handles different path formats and case sensitivity
    /// </summary>
    private bool IsSameDwgFile(string path1, string path2)
    {
        try
        {
            if (string.IsNullOrEmpty(path1) || string.IsNullOrEmpty(path2))
                return false;

            // Normalize the paths
            var normalizedPath1 = Path.GetFullPath(path1).ToLowerInvariant();
            var normalizedPath2 = Path.GetFullPath(path2).ToLowerInvariant();

            return normalizedPath1.Equals(normalizedPath2);
        }
        catch
        {
            // If path normalization fails, fall back to simple string comparison
            return string.Equals(path1, path2, StringComparison.OrdinalIgnoreCase);
        }
    }
}

/// <summary>
/// Enumeration of possible drawing states
/// </summary>
public enum DrawingState
{
    /// <summary>Drawing is currently active in AutoCAD</summary>
    Active,
    
    /// <summary>Drawing is open but not the active document</summary>
    Inactive,
    
    /// <summary>Drawing file exists but is not currently open</summary>
    Closed,
    
    /// <summary>Drawing file was not found</summary>
    NotFound,
    
    /// <summary>Error occurred while checking drawing state</summary>
    Error
}