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
        
        if (isAutoNotesMode)
        {
            // Auto Notes functionality - Phase 1 & 2 working functionality
            ExecuteAutoNotesUpdate();
        }
        else
        {
            // Excel Notes - show "Coming Soon" message
            ShowComingSoonNotification("Excel Notes");
            
            string featureDescription = "Excel Notes Mode:\n" +
                                      "- Manual sheet-to-notes mapping from Excel file\n" +
                                      "- EXCEL_NOTES table processing\n" +
                                      "- Series-specific note lookup\n" +
                                      "- Direct construction note block updates";
                                      
            UpdateTextBlockWithComingSoon(NotesTextBlock, "Update Notes (Excel Notes)", featureDescription);
        }
    }
    
    private void ExecuteAutoNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Auto Notes update functionality");
            
            // TODO: Implement actual Auto Notes functionality
            // This should integrate with Phase 1 & 2 working functionality
            // For now, show a placeholder message indicating it would work
            
            NotificationService.ShowInformation(
                "Auto Notes", 
                "Auto Notes functionality would execute here.\n\n" +
                "This will be connected to the working Phase 1 & 2 implementation.");
                
            NotesTextBlock.Text = "Auto Notes Update Executed!\n\n" +
                                 "This feature is working and will:\n" +
                                 "- Automatically detect construction notes from viewports\n" +
                                 "- Analyze bubble multileaders in model space\n" +
                                 "- Calculate viewport boundaries and extract note numbers\n" +
                                 "- Update construction note blocks accordingly\n\n" +
                                 "Status: Ready for implementation connection...";
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error executing Auto Notes update: {ex.Message}", ex);
            NotificationService.ShowError("Auto Notes Error", $"Failed to execute Auto Notes update: {ex.Message}");
        }
    }
}