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
            // Create and configure service registration
            var serviceRegistration = new ApplicationServiceRegistration();
            
            // Register UI-specific services
            serviceRegistration.RegisterNotificationService<WpfNotificationService>();
            
            // Initialize the application services
            ApplicationServices.Initialize(serviceRegistration);
            
            _isInitialized = true;
        }
        catch (Exception ex)
        {
            // Log initialization failure using fallback logger
            var fallbackLogger = new DebugLogger();
            fallbackLogger.LogCritical($"Failed to initialize UI services: {ex.Message}", ex);
            
            // Initialize with minimal services as fallback
            ApplicationServices.Initialize();
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
        ApplicationServices.Reset();
    }
}