using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConstructionNoteControl : BaseUserControl
{
    public ConstructionNoteControl() : this(null, null)
    {
    }

    public ConstructionNoteControl(
        ILogger? logger,
        INotificationService? notificationService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
    }

    private void UpdateNotesButton_Click(object sender, RoutedEventArgs e)
    {
        bool isAutoNotesMode = AutoNotesRadioButton.IsChecked == true;
        string selectedMode = isAutoNotesMode ? "Auto Notes" : "Excel Notes";
        
        const string featureName = "Update Notes";
        ShowComingSoonNotification(featureName);
        
        string featureDescription = isAutoNotesMode 
            ? "Auto Notes Mode:\n" +
              "- Automatic construction note detection from viewports\n" +
              "- Bubble multileader analysis\n" +
              "- Model space boundary calculation\n" +
              "- Smart note number extraction"
            : "Excel Notes Mode:\n" +
              "- Manual sheet-to-notes mapping from Excel file\n" +
              "- EXCEL_NOTES table processing\n" +
              "- Series-specific note lookup\n" +
              "- Direct construction note block updates";
            
        UpdateTextBlockWithComingSoon(NotesTextBlock, $"{featureName} ({selectedMode})", featureDescription);
    }
}