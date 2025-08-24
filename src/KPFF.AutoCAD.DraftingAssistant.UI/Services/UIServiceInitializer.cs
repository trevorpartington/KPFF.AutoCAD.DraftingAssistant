using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.UI.Controls;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Services;

/// <summary>
/// Initializes all services for the WPF UI layer
/// </summary>
public static class UIServiceInitializer
{
    private static bool _isInitialized = false;

    /// <summary>
    /// Initialize all application services for the UI environment
    /// </summary>
    public static void Initialize()
    {
        if (_isInitialized) return;

        try
        {
            // UI services are now initialized through dependency injection at the plugin level
            // This method is kept for backward compatibility but is essentially a no-op
            _isInitialized = true;
        }
        catch (Exception ex)
        {
            // Log initialization failure using fallback logger
            var fallbackLogger = new DebugLogger();
            fallbackLogger.LogCritical($"Failed to initialize UI services: {ex.Message}", ex);
            _isInitialized = true;
        }
    }

    /// <summary>
    /// Check if services are initialized
    /// </summary>
    public static bool IsInitialized => _isInitialized;

    /// <summary>
    /// Reset initialization state - for testing purposes only
    /// </summary>
    internal static void Reset()
    {
        _isInitialized = false;
        // ApplicationServices has been removed - initialization now handled by plugin layer
    }
}