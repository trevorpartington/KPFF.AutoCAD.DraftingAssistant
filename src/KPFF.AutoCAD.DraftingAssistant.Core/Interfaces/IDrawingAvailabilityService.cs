namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Service to ensure a drawing is available before sheet processing operations
/// Prevents AutoCAD API errors when no drawing context exists
/// </summary>
public interface IDrawingAvailabilityService
{
    /// <summary>
    /// Ensures a drawing is available for operations. Creates a new drawing if needed.
    /// Scenarios:
    /// - No drawings open: Creates new drawing
    /// - Only Drawing1.dwg open during plotting: Creates new drawing
    /// </summary>
    /// <returns>True if drawing is now available, false if creation failed</returns>
    bool EnsureDrawingAvailable(bool isPlottingOperation = false);

    /// <summary>
    /// Checks if a drawing is currently available for operations
    /// </summary>
    /// <returns>True if at least one drawing is open</returns>
    bool IsDrawingAvailable();
}