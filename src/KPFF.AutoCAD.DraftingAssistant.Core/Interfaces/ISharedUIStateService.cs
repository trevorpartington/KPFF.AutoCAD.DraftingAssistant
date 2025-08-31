namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Interface for managing shared UI state across different tabs and controls
/// </summary>
public interface ISharedUIStateService
{
    /// <summary>
    /// Gets or sets whether operations should apply to the current sheet only
    /// This setting persists across all tabs and controls
    /// </summary>
    bool ApplyToCurrentSheetOnly { get; set; }

    /// <summary>
    /// Event raised when the ApplyToCurrentSheetOnly setting changes
    /// Subscribe to this event to update UI controls when the setting changes from other tabs
    /// </summary>
    event Action<bool>? OnApplyToCurrentSheetOnlyChanged;
}