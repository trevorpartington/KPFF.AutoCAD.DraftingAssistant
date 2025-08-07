using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using Autodesk.AutoCAD.ApplicationServices.Core;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific notification service implementation
/// </summary>
public class AutoCadNotificationService : INotificationService
{
    public void ShowInformation(string title, string message)
    {
        Application.ShowAlertDialog($"{title}\n\n{message}");
    }

    public void ShowWarning(string title, string message)
    {
        Application.ShowAlertDialog($"{title}\n\n{message}");
    }

    public void ShowError(string title, string message)
    {
        Application.ShowAlertDialog($"Error - {title}\n\n{message}");
    }

    public bool ShowConfirmation(string title, string message)
    {
        // AutoCAD doesn't have a built-in confirmation dialog, so we'll use ShowAlertDialog
        // In a more complete implementation, you might create a custom WPF dialog
        ShowInformation(title, $"{message}\n\n(Feature requires confirmation dialog implementation)");
        return true; // Default to true for now
    }
}