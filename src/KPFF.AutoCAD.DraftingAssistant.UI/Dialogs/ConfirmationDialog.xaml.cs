using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

/// <summary>
/// Confirmation dialog for user yes/no choices
/// </summary>
public partial class ConfirmationDialog : Window
{
    public bool Result { get; private set; } = false;

    public ConfirmationDialog()
    {
        InitializeComponent();
    }

    public ConfirmationDialog(string title, string message) : this()
    {
        TitleTextBlock.Text = title;
        MessageTextBlock.Text = message;
        Title = title;
    }

    private void YesButton_Click(object sender, RoutedEventArgs e)
    {
        Result = true;
        DialogResult = true;
        Close();
    }

    private void NoButton_Click(object sender, RoutedEventArgs e)
    {
        Result = false;
        DialogResult = false;
        Close();
    }

    /// <summary>
    /// Shows a confirmation dialog and returns the result
    /// </summary>
    /// <param name="owner">Owner window (can be null)</param>
    /// <param name="title">Dialog title</param>
    /// <param name="message">Confirmation message</param>
    /// <returns>True if user clicked Yes, false otherwise</returns>
    public static bool ShowConfirmation(Window? owner, string title, string message)
    {
        try
        {
            var dialog = new ConfirmationDialog(title, message);
            
            if (owner != null)
            {
                dialog.Owner = owner;
            }
            
            dialog.ShowDialog();
            return dialog.Result;
        }
        catch
        {
            // Fallback to MessageBox if custom dialog fails
            var result = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            return result == MessageBoxResult.Yes;
        }
    }
}