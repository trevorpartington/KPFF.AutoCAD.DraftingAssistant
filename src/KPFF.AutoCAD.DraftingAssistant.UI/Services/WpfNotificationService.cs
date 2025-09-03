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
        System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
    }

    public void ShowWarning(string title, string message)
    {
        System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
    }

    public void ShowError(string title, string message)
    {
        System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
    }

    public bool ShowConfirmation(string title, string message)
    {
        var result = System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question);
        return result == System.Windows.MessageBoxResult.Yes;
    }
}