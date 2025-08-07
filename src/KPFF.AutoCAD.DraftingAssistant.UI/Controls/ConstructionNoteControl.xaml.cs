using System.Windows;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConstructionNoteControl : BaseUserControl
{
    public ConstructionNoteControl()
    {
        InitializeComponent();
    }

    private void UpdateNotesButton_Click(object sender, RoutedEventArgs e)
    {
        const string featureName = "Update Notes";
        ShowComingSoonNotification(featureName);
        
        const string featureDescription = 
            "- Automatic construction note detection\n" +
            "- Smart note updating from external sources\n" +
            "- Validation and error checking\n" +
            "- Batch processing capabilities\n" +
            "- Integration with KPFF note databases";
            
        UpdateTextBlockWithComingSoon(NotesTextBlock, featureName, featureDescription);
    }
}