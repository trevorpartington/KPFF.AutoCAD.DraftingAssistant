using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class TitleBlockControl : BaseUserControl
{
    public TitleBlockControl() : this(null, null)
    {
    }

    public TitleBlockControl(
        ILogger? logger,
        INotificationService? notificationService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
    }

    private void UpdateTitleBlocksButton_Click(object sender, RoutedEventArgs e)
    {
        const string featureName = "Update Title Blocks";
        ShowComingSoonNotification(featureName);
        
        const string featureDescription = 
            "- Automatic title block detection and updates\n" +
            "- Sheet set integration and synchronization\n" +
            "- Project information propagation\n" +
            "- Attribute validation and correction\n" +
            "- Batch processing across multiple sheets\n" +
            "- Custom title block template management";
            
        UpdateTextBlockWithComingSoon(TitleBlockTextBlock, featureName, featureDescription);
    }
}