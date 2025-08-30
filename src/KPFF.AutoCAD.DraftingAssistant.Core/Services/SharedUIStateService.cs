using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for managing shared UI state across different tabs and controls
/// Provides a singleton pattern for persisting user preferences like checkbox states
/// </summary>
public class SharedUIStateService : ISharedUIStateService
{
    private static readonly Lazy<SharedUIStateService> _instance = new(() => new SharedUIStateService());
    
    /// <summary>
    /// Gets the singleton instance of the SharedUIStateService
    /// </summary>
    public static SharedUIStateService Instance => _instance.Value;

    private bool _applyToCurrentSheetOnly = false;

    /// <summary>
    /// Gets or sets whether operations should apply to the current sheet only
    /// This setting persists across all tabs and controls
    /// </summary>
    public bool ApplyToCurrentSheetOnly
    {
        get => _applyToCurrentSheetOnly;
        set
        {
            if (_applyToCurrentSheetOnly != value)
            {
                _applyToCurrentSheetOnly = value;
                OnApplyToCurrentSheetOnlyChanged?.Invoke(value);
            }
        }
    }

    /// <summary>
    /// Event raised when the ApplyToCurrentSheetOnly setting changes
    /// Subscribe to this event to update UI controls when the setting changes from other tabs
    /// </summary>
    public event Action<bool>? OnApplyToCurrentSheetOnlyChanged;

    private SharedUIStateService()
    {
        // Private constructor for singleton pattern
    }
}