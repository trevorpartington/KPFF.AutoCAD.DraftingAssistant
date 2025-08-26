using System.Collections.ObjectModel;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class NoteBlocksDialog : Window
{
    private ObservableCollection<NoteBlockConfiguration> _blockConfigurations;
    
    public List<NoteBlockConfiguration> NoteBlocks { get; private set; }

    public NoteBlocksDialog(List<NoteBlockConfiguration> noteBlocks)
    {
        InitializeComponent();
        
        // Create copies to avoid modifying the original list until OK is clicked
        _blockConfigurations = new ObservableCollection<NoteBlockConfiguration>(
            noteBlocks.Select(nb => new NoteBlockConfiguration 
            { 
                BlockName = nb.BlockName, 
                AttributeName = nb.AttributeName 
            }));
        
        NoteBlocks = noteBlocks;
        
        BlocksDataGrid.ItemsSource = _blockConfigurations;
        BlocksDataGrid.SelectionChanged += BlocksDataGrid_SelectionChanged;
        
        UpdateButtonStates();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        // Validate configurations
        var validationErrors = ValidateBlockConfigurations();
        if (validationErrors.Count > 0)
        {
            var errorMessage = "Please fix the following errors:\n\n" + string.Join("\n", validationErrors);
            MessageBox.Show(errorMessage, "Validation Errors", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Update the result with valid configurations
        NoteBlocks = _blockConfigurations
            .Where(bc => !string.IsNullOrWhiteSpace(bc.BlockName) && !string.IsNullOrWhiteSpace(bc.AttributeName))
            .ToList();

        DialogResult = true;
        Close();
    }

    private List<string> ValidateBlockConfigurations()
    {
        var errors = new List<string>();
        var validConfigs = _blockConfigurations
            .Where(bc => !string.IsNullOrWhiteSpace(bc.BlockName) || !string.IsNullOrWhiteSpace(bc.AttributeName))
            .ToList();

        for (int i = 0; i < validConfigs.Count; i++)
        {
            var config = validConfigs[i];
            
            if (string.IsNullOrWhiteSpace(config.BlockName))
            {
                errors.Add($"Row {i + 1}: Block name is required");
            }
            
            if (string.IsNullOrWhiteSpace(config.AttributeName))
            {
                errors.Add($"Row {i + 1}: Attribute name is required");
            }
        }

        // Check for duplicates
        var duplicates = validConfigs
            .GroupBy(bc => new { bc.BlockName, bc.AttributeName })
            .Where(g => g.Count() > 1)
            .Select(g => g.Key);

        foreach (var duplicate in duplicates)
        {
            errors.Add($"Duplicate configuration: Block '{duplicate.BlockName}' with attribute '{duplicate.AttributeName}'");
        }

        return errors;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }


    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (BlocksDataGrid.SelectedItem is NoteBlockConfiguration selectedBlock)
        {
            var result = MessageBox.Show(
                $"Are you sure you want to delete the block configuration:\n\n" +
                $"Block Name: {selectedBlock.BlockName}\n" +
                $"Attribute Name: {selectedBlock.AttributeName}",
                "Confirm Delete", 
                MessageBoxButton.YesNo, 
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _blockConfigurations.Remove(selectedBlock);
                UpdateButtonStates();
            }
        }
    }

    private void BlocksDataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        DeleteButton.IsEnabled = BlocksDataGrid.SelectedItem != null;
    }
}