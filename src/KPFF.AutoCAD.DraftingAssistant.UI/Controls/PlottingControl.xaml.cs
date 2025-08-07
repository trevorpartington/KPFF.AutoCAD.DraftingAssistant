using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class PlottingControl : BaseUserControl
{
    public PlottingControl()
    {
        InitializeComponent();
    }

    private void PlotButton_Click(object sender, RoutedEventArgs e)
    {
        const string featureName = "Plot";
        ShowComingSoonNotification(featureName);
        
        const string featureDescription = 
            "- Intelligent batch plotting capabilities\n" +
            "- Custom plot configuration management\n" +
            "- Quality control and validation checks\n" +
            "- Multiple output format support (PDF, DWF, etc.)\n" +
            "- Automated file naming and organization\n" +
            "- Integration with KPFF plotting standards";
            
        UpdateTextBlockWithComingSoon(PlottingTextBlock, featureName, featureDescription);
    }
}