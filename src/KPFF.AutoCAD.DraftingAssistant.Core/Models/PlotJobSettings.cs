namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

/// <summary>
/// Settings for a plot job including optional pre-plot updates
/// </summary>
public class PlotJobSettings
{
    /// <summary>
    /// Whether to update construction notes before plotting
    /// </summary>
    public bool UpdateConstructionNotes { get; set; } = false;

    /// <summary>
    /// Whether to update title blocks before plotting
    /// </summary>
    public bool UpdateTitleBlocks { get; set; } = false;

    /// <summary>
    /// Whether to apply operations to current sheet only
    /// </summary>
    public bool ApplyToCurrentSheetOnly { get; set; } = false;

    /// <summary>
    /// Construction notes mode (true = Auto Notes, false = Excel Notes)
    /// </summary>
    public bool IsAutoNotesMode { get; set; } = true;

    /// <summary>
    /// Output directory for plot files (uses project configuration if not specified)
    /// </summary>
    public string? OutputDirectory { get; set; }
}

/// <summary>
/// Result of a plot operation
/// </summary>
public class PlotResult
{
    /// <summary>
    /// Overall success status
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// List of sheets that were successfully plotted
    /// </summary>
    public List<string> SuccessfulSheets { get; set; } = new();

    /// <summary>
    /// List of sheets that failed to plot with error messages
    /// </summary>
    public List<SheetPlotError> FailedSheets { get; set; } = new();

    /// <summary>
    /// General error message if the entire operation failed
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Total number of sheets processed
    /// </summary>
    public int TotalSheets => SuccessfulSheets.Count + FailedSheets.Count;

    /// <summary>
    /// Success rate as percentage
    /// </summary>
    public double SuccessRate => TotalSheets == 0 ? 0 : (double)SuccessfulSheets.Count / TotalSheets * 100;
}

/// <summary>
/// Error information for a failed sheet plot
/// </summary>
public class SheetPlotError
{
    /// <summary>
    /// Name of the sheet that failed
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Error message
    /// </summary>
    public string ErrorMessage { get; set; } = string.Empty;

    /// <summary>
    /// Exception details if available
    /// </summary>
    public string? ExceptionDetails { get; set; }
}

/// <summary>
/// Progress information for plot operations
/// </summary>
public class PlotProgress
{
    /// <summary>
    /// Current sheet being processed
    /// </summary>
    public string CurrentSheet { get; set; } = string.Empty;

    /// <summary>
    /// Current operation being performed
    /// </summary>
    public string CurrentOperation { get; set; } = string.Empty;

    /// <summary>
    /// Progress percentage (0-100)
    /// </summary>
    public int ProgressPercentage { get; set; }

    /// <summary>
    /// Number of sheets completed
    /// </summary>
    public int CompletedSheets { get; set; }

    /// <summary>
    /// Total number of sheets to process
    /// </summary>
    public int TotalSheets { get; set; }
}

/// <summary>
/// Validation result for plot operations
/// </summary>
public class PlotValidationResult
{
    /// <summary>
    /// Whether all sheets passed validation
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// List of validation issues found
    /// </summary>
    public List<PlotValidationIssue> Issues { get; set; } = new();

    /// <summary>
    /// Sheets that can be plotted
    /// </summary>
    public List<string> ValidSheets { get; set; } = new();

    /// <summary>
    /// Sheets that have issues preventing plotting
    /// </summary>
    public List<string> InvalidSheets { get; set; } = new();
}

/// <summary>
/// Individual validation issue
/// </summary>
public class PlotValidationIssue
{
    /// <summary>
    /// Sheet name with the issue
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Type of issue (Missing, Invalid, etc.)
    /// </summary>
    public PlotValidationIssueType IssueType { get; set; }

    /// <summary>
    /// Description of the issue
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Whether this is a warning (can proceed) or error (blocks plotting)
    /// </summary>
    public bool IsWarning { get; set; }
}

/// <summary>
/// Types of plot validation issues
/// </summary>
public enum PlotValidationIssueType
{
    /// <summary>
    /// Drawing file not found
    /// </summary>
    MissingDrawing,

    /// <summary>
    /// Layout not found in drawing
    /// </summary>
    MissingLayout,

    /// <summary>
    /// Invalid plot configuration
    /// </summary>
    InvalidPlotConfig,

    /// <summary>
    /// Plot device not available
    /// </summary>
    MissingPlotDevice,

    /// <summary>
    /// Other general issues
    /// </summary>
    Other
}

/// <summary>
/// Default plot settings for a specific sheet
/// </summary>
public class SheetPlotSettings
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Layout name in the drawing
    /// </summary>
    public string LayoutName { get; set; } = string.Empty;

    /// <summary>
    /// Plot device/printer name
    /// </summary>
    public string PlotDevice { get; set; } = string.Empty;

    /// <summary>
    /// Paper size
    /// </summary>
    public string PaperSize { get; set; } = string.Empty;

    /// <summary>
    /// Plot scale
    /// </summary>
    public string PlotScale { get; set; } = string.Empty;

    /// <summary>
    /// Plot area (Layout, Extents, etc.)
    /// </summary>
    public string PlotArea { get; set; } = string.Empty;

    /// <summary>
    /// Whether plot is centered
    /// </summary>
    public bool PlotCentered { get; set; }

    /// <summary>
    /// Drawing file path
    /// </summary>
    public string DrawingPath { get; set; } = string.Empty;
}