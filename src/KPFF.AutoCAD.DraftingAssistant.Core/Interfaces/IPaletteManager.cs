namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Manages the drafting assistant palette lifecycle and visibility
/// </summary>
public interface IPaletteManager
{
    void Initialize();
    void Show();
    void Hide();
    void Toggle();
    void Cleanup();
    bool IsVisible { get; }
    bool IsInitialized { get; }
}