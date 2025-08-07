using System.Windows;
using System.Windows.Controls;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConfigurationControl : UserControl
{
    public ConfigurationControl()
    {
        InitializeComponent();
    }

    private void ConfigureButton_Click(object sender, RoutedEventArgs e)
    {
        ShowComingSoonWindow("Configure");
        
        // Update text readout
        ConfigurationTextBlock.Text = "Configure button clicked!" + Environment.NewLine + Environment.NewLine + 
                                     "This feature is coming soon. It will allow you to:" + Environment.NewLine +
                                     "- Set project parameters" + Environment.NewLine +
                                     "- Configure default settings" + Environment.NewLine +
                                     "- Manage user preferences" + Environment.NewLine +
                                     "- Set up project templates";
    }

    private void ShowComingSoonWindow(string featureName)
    {
        MessageBox.Show($"Coming Soon - {featureName}\n\nThis feature will be implemented in a future update.", 
                       "KPFF Drafting Assistant", 
                       MessageBoxButton.OK, 
                       MessageBoxImage.Information);
    }
}