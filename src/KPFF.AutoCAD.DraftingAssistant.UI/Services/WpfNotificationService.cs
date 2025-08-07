using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

/// <summary>
/// WPF-based notification service implementation using MessageBox
/// </summary>
public class WpfNotificationService : INotificationService
{
    public void ShowInformation(string title, string message)
    {
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
    }

    public void ShowWarning(string title, string message)
    {
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public void ShowError(string title, string message)
    {
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public bool ShowConfirmation(string title, string message)
    {
        var result = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
        return result == MessageBoxResult.Yes;
    }
}