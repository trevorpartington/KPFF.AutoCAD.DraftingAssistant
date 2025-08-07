using System.Windows;
using System.Windows.Controls;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class TitleBlockControl : UserControl
{
    public TitleBlockControl()
    {
        InitializeComponent();
    }

    private void UpdateTitleBlocksButton_Click(object sender, RoutedEventArgs e)
    {
        ShowComingSoonWindow("Update Title Blocks");
        
        // Update text readout
        TitleBlockTextBlock.Text = @"Update Title Blocks button clicked!

This feature is coming soon. It will provide:
- Automatic title block detection and updates
- Sheet set integration and synchronization  
- Project information propagation
- Attribute validation and correction
- Batch processing across multiple sheets
- Custom title block template management

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