using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class SheetSelectionDialog : Window
{
    public ObservableCollection<SheetItemViewModel> Sheets { get; } = new();
    public List<SheetInfo> SelectedSheets => Sheets.Where(s => s.IsSelected).Select(s => s.Sheet).ToList();

    public SheetSelectionDialog(List<SheetInfo> availableSheets, List<SheetInfo>? currentSelection = null)
    {
        InitializeComponent();
        
        var selectedSheetNames = currentSelection?.Select(s => s.SheetName).ToHashSet() ?? new HashSet<string>();
        
        foreach (var sheet in availableSheets)
        {
            Sheets.Add(new SheetItemViewModel(sheet, selectedSheetNames.Contains(sheet.SheetName)));
        }
        
        SheetsItemsControl.ItemsSource = Sheets;
    }

    private void SelectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var sheet in Sheets)
        {
            sheet.IsSelected = true;
        }
    }

    private void ClearAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var sheet in Sheets)
        {
            sheet.IsSelected = false;
        }
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void SelectSeriesButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Get currently selected series from sheets
            var currentlySelectedSeries = GetSelectedSeries();
            
            // Get all available sheets for series selection
            var allSheets = Sheets.Select(s => s.Sheet).ToList();
            
            var seriesDialog = new SeriesSelectionDialog(allSheets, currentlySelectedSeries)
            {
                Owner = this
            };
            
            if (seriesDialog.ShowDialog() == true)
            {
                var selectedSeries = seriesDialog.SelectedSeries.ToHashSet();
                
                // Update sheet selections based on selected series
                UpdateSheetSelectionsBySeries(selectedSeries);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening series selection: {ex.Message}", 
                           "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private List<string> GetSelectedSeries()
    {
        var selectedSeries = new HashSet<string>();
        var sheetNamePattern = new Regex(@"^([A-Z]{1,3})", RegexOptions.IgnoreCase);
        
        foreach (var sheet in Sheets.Where(s => s.IsSelected))
        {
            var match = sheetNamePattern.Match(sheet.Sheet.SheetName);
            if (match.Success)
            {
                selectedSeries.Add(match.Groups[1].Value.ToUpper());
            }
        }
        
        return selectedSeries.ToList();
    }

    private void UpdateSheetSelectionsBySeries(HashSet<string> selectedSeries)
    {
        var sheetNamePattern = new Regex(@"^([A-Z]{1,3})", RegexOptions.IgnoreCase);
        
        foreach (var sheet in Sheets)
        {
            var match = sheetNamePattern.Match(sheet.Sheet.SheetName);
            if (match.Success)
            {
                var seriesName = match.Groups[1].Value.ToUpper();
                sheet.IsSelected = selectedSeries.Contains(seriesName);
            }
            else
            {
                // If sheet name doesn't match pattern, leave it unselected
                sheet.IsSelected = false;
            }
        }
    }
}

public class SheetItemViewModel : INotifyPropertyChanged
{
    private bool _isSelected;
    
    public SheetInfo Sheet { get; }
    
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected != value)
            {
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }
    }
    
    public string DisplayText => 
        $"{Sheet.SheetName} - {Sheet.DrawingTitle}" + 
        (string.IsNullOrEmpty(Sheet.Scale) ? "" : $" (Scale: {Sheet.Scale})");

    public SheetItemViewModel(SheetInfo sheet, bool isSelected = false)
    {
        Sheet = sheet;
        _isSelected = isSelected;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}