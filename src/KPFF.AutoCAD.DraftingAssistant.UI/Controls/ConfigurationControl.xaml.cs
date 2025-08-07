using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConfigurationControl : BaseUserControl
{
    public ConfigurationControl()
    {
        InitializeComponent();
    }

    private void ConfigureButton_Click(object sender, RoutedEventArgs e)
    {
        const string featureName = "Configure";
        ShowComingSoonNotification(featureName);
        
        const string featureDescription = 
            "- Set project parameters\n" +
            "- Configure default settings\n" +
            "- Manage user preferences\n" +
            "- Set up project templates";
            
        UpdateTextBlockWithComingSoon(ConfigurationTextBlock, featureName, featureDescription);
    }
}