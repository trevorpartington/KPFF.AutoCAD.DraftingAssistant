namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

/// <summary>
/// Settings for batch operations across multiple drawings
/// </summary>
public class BatchOperationSettings
{
    /// <summary>
    /// Whether to update construction notes
    /// </summary>
    public bool UpdateConstructionNotes { get; set; } = false;

    /// <summary>
    /// Whether to update title blocks
    /// </summary>
    public bool UpdateTitleBlocks { get; set; } = false;

    /// <summary>
    /// Whether to plot sheets to PDF
    /// </summary>
    public bool PlotToPdf { get; set; } = false;

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

    /// <summary>
    /// Returns true if no operations are selected
    /// </summary>
    public bool IsEmpty => !UpdateConstructionNotes && !UpdateTitleBlocks && !PlotToPdf;

    /// <summary>
    /// Creates batch settings from plot job settings for backward compatibility
    /// </summary>
    public static BatchOperationSettings FromPlotJobSettings(PlotJobSettings plotSettings)
    {
        return new BatchOperationSettings
        {
            UpdateConstructionNotes = plotSettings.UpdateConstructionNotes,
            UpdateTitleBlocks = plotSettings.UpdateTitleBlocks,
            PlotToPdf = true, // Always true when coming from plot settings
            ApplyToCurrentSheetOnly = plotSettings.ApplyToCurrentSheetOnly,
            IsAutoNotesMode = plotSettings.IsAutoNotesMode,
            OutputDirectory = plotSettings.OutputDirectory
        };
    }
}

/// <summary>
/// Result of a batch operation across multiple drawings
/// </summary>
public class BatchOperationResult
{
    /// <summary>
    /// Overall success status
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// List of sheets that were successfully processed
    /// </summary>
    public List<BatchOperationSheetResult> SuccessfulSheets { get; set; } = new();

    /// <summary>
    /// List of sheets that failed processing with error messages
    /// </summary>
    public List<BatchOperationSheetError> FailedSheets { get; set; } = new();

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

    /// <summary>
    /// Summary of operations performed
    /// </summary>
    public BatchOperationSummary Summary { get; set; } = new();
}

/// <summary>
/// Result information for a successfully processed sheet
/// </summary>
public class BatchOperationSheetResult
{
    /// <summary>
    /// Name of the sheet
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Drawing file path
    /// </summary>
    public string DrawingPath { get; set; } = string.Empty;

    /// <summary>
    /// Drawing state when processed
    /// </summary>
    public DrawingState DrawingState { get; set; }

    /// <summary>
    /// Operations that were completed successfully
    /// </summary>
    public List<BatchOperationType> CompletedOperations { get; set; } = new();

    /// <summary>
    /// Number of construction notes updated (if applicable)
    /// </summary>
    public int ConstructionNotesUpdated { get; set; }

    /// <summary>
    /// Number of title block attributes updated (if applicable)
    /// </summary>
    public int TitleBlockAttributesUpdated { get; set; }

    /// <summary>
    /// Whether plotting was successful (if applicable)
    /// </summary>
    public bool PlottingSuccessful { get; set; }
}

/// <summary>
/// Error information for a failed sheet operation
/// </summary>
public class BatchOperationSheetError
{
    /// <summary>
    /// Name of the sheet that failed
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Drawing file path (if known)
    /// </summary>
    public string DrawingPath { get; set; } = string.Empty;

    /// <summary>
    /// Operation that failed
    /// </summary>
    public BatchOperationType FailedOperation { get; set; }

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
/// Types of batch operations
/// </summary>
public enum BatchOperationType
{
    /// <summary>
    /// Construction notes update
    /// </summary>
    ConstructionNotes,

    /// <summary>
    /// Title block update
    /// </summary>
    TitleBlocks,

    /// <summary>
    /// PDF plotting
    /// </summary>
    Plotting,

    /// <summary>
    /// Validation or setup
    /// </summary>
    Validation
}

/// <summary>
/// Summary of operations performed in a batch operation
/// </summary>
public class BatchOperationSummary
{
    /// <summary>
    /// Number of sheets where construction notes were updated
    /// </summary>
    public int ConstructionNotesUpdated { get; set; }

    /// <summary>
    /// Number of sheets where title blocks were updated
    /// </summary>
    public int TitleBlocksUpdated { get; set; }

    /// <summary>
    /// Number of sheets that were plotted successfully
    /// </summary>
    public int SheetsPlotted { get; set; }

    /// <summary>
    /// Total construction notes processed across all sheets
    /// </summary>
    public int TotalConstructionNotes { get; set; }

    /// <summary>
    /// Total title block attributes processed across all sheets
    /// </summary>
    public int TotalTitleBlockAttributes { get; set; }

    /// <summary>
    /// Drawing states encountered during processing
    /// </summary>
    public Dictionary<DrawingState, int> DrawingStateBreakdown { get; set; } = new();
}

/// <summary>
/// Progress information for batch operations
/// </summary>
public class BatchOperationProgress
{
    /// <summary>
    /// Current sheet being processed
    /// </summary>
    public string CurrentSheet { get; set; } = string.Empty;

    /// <summary>
    /// Current operation being performed
    /// </summary>
    public BatchOperationType CurrentOperation { get; set; }

    /// <summary>
    /// Human-readable description of current operation
    /// </summary>
    public string CurrentOperationDescription { get; set; } = string.Empty;

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

    /// <summary>
    /// Phase of the batch operation (Validation, PreUpdates, Plotting, etc.)
    /// </summary>
    public string Phase { get; set; } = string.Empty;
}

/// <summary>
/// Validation result for batch operations
/// </summary>
public class BatchOperationValidationResult
{
    /// <summary>
    /// Whether all sheets passed validation for all requested operations
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// List of validation issues found
    /// </summary>
    public List<BatchOperationValidationIssue> Issues { get; set; } = new();

    /// <summary>
    /// Sheets that can be processed for all requested operations
    /// </summary>
    public List<string> ValidSheets { get; set; } = new();

    /// <summary>
    /// Sheets that have issues preventing some or all operations
    /// </summary>
    public List<string> InvalidSheets { get; set; } = new();

    /// <summary>
    /// Operations that can be performed
    /// </summary>
    public List<BatchOperationType> ValidOperations { get; set; } = new();

    /// <summary>
    /// Operations that cannot be performed due to validation issues
    /// </summary>
    public List<BatchOperationType> InvalidOperations { get; set; } = new();
}

/// <summary>
/// Individual validation issue for batch operations
/// </summary>
public class BatchOperationValidationIssue
{
    /// <summary>
    /// Sheet name with the issue
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Operation affected by the issue
    /// </summary>
    public BatchOperationType AffectedOperation { get; set; }

    /// <summary>
    /// Type of issue
    /// </summary>
    public BatchOperationValidationIssueType IssueType { get; set; }

    /// <summary>
    /// Description of the issue
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Whether this is a warning (can proceed) or error (blocks operation)
    /// </summary>
    public bool IsWarning { get; set; }
}

/// <summary>
/// Types of batch operation validation issues
/// </summary>
public enum BatchOperationValidationIssueType
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
    /// Sheet not found in project index
    /// </summary>
    MissingSheetInfo,

    /// <summary>
    /// Construction notes configuration issue
    /// </summary>
    ConstructionNotesConfig,

    /// <summary>
    /// Title block configuration issue
    /// </summary>
    TitleBlockConfig,

    /// <summary>
    /// Plotting configuration issue
    /// </summary>
    PlottingConfig,

    /// <summary>
    /// Drawing access issue
    /// </summary>
    DrawingAccess,

    /// <summary>
    /// Other general issues
    /// </summary>
    Other
}