using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// Initializes all services for the AutoCAD plugin
/// </summary>
public static class PluginServiceInitializer
{
    private static bool _isInitialized = false;

    /// <summary>
    /// Initialize all application services for the plugin environment
    /// Service initialization is now handled directly by DraftingAssistantExtensionApplication
    /// </summary>
    public static void Initialize()
    {
        if (_isInitialized) return;

        try
        {
            // Service initialization is now handled by DraftingAssistantExtensionApplication
            // This method is kept for backward compatibility but is essentially a no-op
            _isInitialized = true;
        }
        catch (Exception ex)
        {
            // Log initialization failure using fallback logger
            var fallbackLogger = new DebugLogger();
            fallbackLogger.LogCritical($"Failed to initialize plugin services: {ex.Message}", ex);
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
        // ApplicationServices has been removed - service initialization now handled by extension application
    }
}