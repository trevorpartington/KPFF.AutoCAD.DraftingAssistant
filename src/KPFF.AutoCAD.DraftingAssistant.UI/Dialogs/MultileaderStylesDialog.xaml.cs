using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class MultileaderStylesDialog : Window
{
    public List<string> MultileaderStyles { get; private set; } = new();

    public MultileaderStylesDialog(List<string> currentStyles)
    {
        InitializeComponent();
        
        // Load current styles into the text box as comma-separated values
        if (currentStyles?.Count > 0)
        {
            StylesTextBox.Text = string.Join(", ", currentStyles);
        }
        
        StylesTextBox.Focus();
        StylesTextBox.SelectAll();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var inputText = StylesTextBox.Text?.Trim() ?? string.Empty;
            
            if (string.IsNullOrWhiteSpace(inputText))
            {
                MessageBox.Show("Please enter at least one multileader style name.", 
                               "Input Required", MessageBoxButton.OK, MessageBoxImage.Information);
                StylesTextBox.Focus();
                return;
            }

            // Parse comma-separated values
            var styles = inputText
                .Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (styles.Count == 0)
            {
                MessageBox.Show("Please enter at least one valid multileader style name.", 
                               "Input Required", MessageBoxButton.OK, MessageBoxImage.Information);
                StylesTextBox.Focus();
                return;
            }

            MultileaderStyles = styles;
            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error parsing multileader styles: {ex.Message}", 
                           "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}