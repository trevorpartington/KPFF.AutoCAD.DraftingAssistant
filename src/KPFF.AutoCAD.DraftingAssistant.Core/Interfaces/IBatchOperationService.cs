using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for coordinated batch operations across multiple drawings
/// Handles any combination of Construction Notes, Title Blocks, and PDF plotting
/// Supports Active, Inactive, and Closed drawing states seamlessly
/// </summary>
public interface IBatchOperationService
{
    /// <summary>
    /// Performs a coordinated batch operation with any combination of updates and plotting
    /// </summary>
    /// <param name="sheetNames">List of sheet names to process</param>
    /// <param name="config">Project configuration</param>
    /// <param name="batchSettings">Settings specifying which operations to perform</param>
    /// <param name="progressCallback">Optional callback for progress updates</param>
    /// <returns>Batch operation result with detailed success/failure information</returns>
    Task<BatchOperationResult> ExecuteBatchOperationAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings,
        IProgress<BatchOperationProgress>? progressCallback = null);

    /// <summary>
    /// Updates construction notes only across multiple drawings
    /// </summary>
    /// <param name="sheetNames">List of sheet names to update</param>
    /// <param name="config">Project configuration</param>
    /// <param name="isAutoNotesMode">True for Auto Notes mode, false for Excel Notes mode</param>
    /// <param name="applyToCurrentSheetOnly">If true, only process the current active sheet</param>
    /// <param name="progressCallback">Optional callback for progress updates</param>
    /// <returns>Construction notes update result</returns>
    Task<BatchOperationResult> UpdateConstructionNotesAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        bool isAutoNotesMode,
        bool applyToCurrentSheetOnly = false,
        IProgress<BatchOperationProgress>? progressCallback = null);

    /// <summary>
    /// Updates title blocks only across multiple drawings
    /// </summary>
    /// <param name="sheetNames">List of sheet names to update</param>
    /// <param name="config">Project configuration</param>
    /// <param name="applyToCurrentSheetOnly">If true, only process the current active sheet</param>
    /// <param name="progressCallback">Optional callback for progress updates</param>
    /// <returns>Title block update result</returns>
    Task<BatchOperationResult> UpdateTitleBlocksAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        bool applyToCurrentSheetOnly = false,
        IProgress<BatchOperationProgress>? progressCallback = null);

    /// <summary>
    /// Validates that the specified sheets can be processed for the given operations
    /// </summary>
    /// <param name="sheetNames">List of sheet names to validate</param>
    /// <param name="config">Project configuration</param>
    /// <param name="batchSettings">Settings specifying which operations will be performed</param>
    /// <returns>Validation result with any issues found</returns>
    Task<BatchOperationValidationResult> ValidateSheetsForBatchOperationAsync(
        List<string> sheetNames,
        ProjectConfiguration config,
        BatchOperationSettings batchSettings);
}