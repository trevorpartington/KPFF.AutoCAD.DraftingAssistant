using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using System.Windows.Controls;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

/// <summary>
/// Base class for all user controls in the drafting assistant
/// </summary>
public abstract class BaseUserControl : UserControl
{
    protected ILogger Logger { get; }
    protected INotificationService NotificationService { get; }

    protected BaseUserControl() : this(null, null)
    {
    }

    protected BaseUserControl(ILogger? logger, INotificationService? notificationService)
    {
        // Use dependency injection if available, otherwise fallback to service locator
        Logger = logger ?? GetLoggerService();
        NotificationService = notificationService ?? GetNotificationService();
    }

    private static ILogger GetLoggerService()
    {
        return ApplicationServices.IsInitialized 
            ? ApplicationServices.GetService<ILogger>()
            : new DebugLogger();
    }

    private static INotificationService GetNotificationService()
    {
        return ApplicationServices.IsInitialized
            ? ApplicationServices.GetService<INotificationService>()
            : new Controls.WpfNotificationService();
    }

    /// <summary>
    /// Shows a "Coming Soon" notification for features under development
    /// </summary>
    /// <param name="featureName">Name of the feature</param>
    protected void ShowComingSoonNotification(string featureName)
    {
        Logger.LogInformation($"Coming soon feature accessed: {featureName}");
        NotificationService.ShowInformation(
            ApplicationConstants.ApplicationName,
            $"Coming Soon - {featureName}\n\nThis feature will be implemented in a future update.");
    }

    /// <summary>
    /// Updates a text block with feature information and coming soon message
    /// </summary>
    /// <param name="textBlock">The text block to update</param>
    /// <param name="featureName">Name of the feature</param>
    /// <param name="featureDescription">Description of what the feature will provide</param>
    protected static void UpdateTextBlockWithComingSoon(TextBlock textBlock, string featureName, string featureDescription)
    {
        textBlock.Text = $"{featureName} button clicked!\n\n" +
                        "This feature is coming soon. It will provide:\n" +
                        featureDescription + "\n\n" +
                        "Status: Feature in development...";
    }
}