using System.Windows;
using System.Windows.Controls;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConstructionNoteControl : UserControl
{
    public ConstructionNoteControl()
    {
        InitializeComponent();
    }

    private void UpdateNotesButton_Click(object sender, RoutedEventArgs e)
    {
        ShowComingSoonWindow("Update Notes");
        
        // Update text readout
        NotesTextBlock.Text = @"Update Notes button clicked!

This feature is coming soon. It will provide:
- Automatic construction note detection
- Smart note updating from external sources
- Validation and error checking
- Batch processing capabilities
- Integration with KPFF note databases

Status: Feature in development...";
    }

    private void ShowComingSoonWindow(string featureName)
    {
        MessageBox.Show($"Coming Soon - {featureName}\n\nThis feature will be implemented in a future update.", 
                       "KPFF Drafting Assistant", 
                       MessageBoxButton.OK, 
                       MessageBoxImage.Information);
    }
}