using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using System.Text;
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
        // CRASH FIX: Never access ApplicationServices during UI initialization
        // This prevents binding errors during AutoCAD document context switching
        return new DebugLogger();
    }

    private static INotificationService GetNotificationService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization  
        // This prevents binding errors during AutoCAD document context switching
        return new Controls.WpfNotificationService();
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

    /// <summary>
    /// Builds a standardized three-section readout display
    /// </summary>
    /// <param name="activeProject">Name of the active project, or null if none</param>
    /// <param name="statusMessages">List of status/information messages to display</param>
    /// <param name="selectedSheets">List of selected sheets to display</param>
    /// <param name="additionalInfo">Optional additional information to show in status section</param>
    /// <returns>Formatted readout string</returns>
    protected static string BuildStandardReadout(
        string? activeProject,
        List<string>? statusMessages = null,
        List<SheetInfo>? selectedSheets = null,
        string? additionalInfo = null)
    {
        var sb = new StringBuilder();

        // Section 1: Active Project
        sb.AppendLine($"Active Project: {activeProject ?? "No project selected"}");
        sb.AppendLine();
        sb.AppendLine();

        // Section 2: Status Messages
        if (statusMessages != null && statusMessages.Count > 0)
        {
            foreach (var message in statusMessages)
            {
                sb.AppendLine(message);
            }
        }
        else if (!string.IsNullOrEmpty(additionalInfo))
        {
            sb.AppendLine(additionalInfo);
        }
        else
        {
            sb.AppendLine("Ready for operations.");
        }

        sb.AppendLine();
        sb.AppendLine();
        
        // Section 3: Selected Sheets
        if (selectedSheets != null && selectedSheets.Count > 0)
        {
            sb.AppendLine($"Selected Sheets: {selectedSheets.Count}");
            foreach (var sheet in selectedSheets)
            {
                var title = !string.IsNullOrEmpty(sheet.DrawingTitle) ? $" - {sheet.DrawingTitle}" : "";
                sb.AppendLine($"• {sheet.SheetName}{title}");
            }
        }
        else
        {
            sb.AppendLine("Selected Sheets: 0");
            sb.AppendLine("No sheets selected for processing");
        }

        return sb.ToString();
    }
}