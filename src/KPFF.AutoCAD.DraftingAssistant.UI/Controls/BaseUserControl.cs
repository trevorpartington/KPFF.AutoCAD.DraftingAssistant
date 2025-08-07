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
    protected ILogger Logger { get; private set; } = null!;
    protected INotificationService NotificationService { get; private set; } = null!;

    protected BaseUserControl()
    {
        // Initialize on UI thread
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => InitializeServices());
        }
        else
        {
            InitializeServices();
        }
    }

    private void InitializeServices()
    {
        // Use service container if available, otherwise create fallback services
        try
        {
            var container = ServiceContainer.Instance;
            Logger = container.IsRegistered<ILogger>() 
                ? container.Resolve<ILogger>() 
                : new DebugLogger();
            
            NotificationService = container.IsRegistered<INotificationService>()
                ? container.Resolve<INotificationService>()
                : new WpfNotificationService();
        }
        catch
        {
            // Fallback if service container is not available
            Logger = new DebugLogger();
            NotificationService = new WpfNotificationService();
        }
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