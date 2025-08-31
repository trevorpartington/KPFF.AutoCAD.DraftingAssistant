using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class SeriesSelectionDialog : Window
{
    public ObservableCollection<SeriesItemViewModel> Series { get; } = new();
    public List<string> SelectedSeries => Series.Where(s => s.IsSelected).Select(s => s.SeriesName).ToList();

    public SeriesSelectionDialog(List<SheetInfo> availableSheets, List<string>? currentSelection = null)
    {
        InitializeComponent();
        
        var selectedSeriesNames = currentSelection?.ToHashSet() ?? new HashSet<string>();
        
        // Extract series from sheet names and group them
        var seriesGroups = ExtractSeriesFromSheets(availableSheets);
        
        foreach (var seriesGroup in seriesGroups.OrderBy(kvp => kvp.Key))
        {
            var seriesName = seriesGroup.Key;
            var sheetsInSeries = seriesGroup.Value;
            var isSelected = selectedSeriesNames.Contains(seriesName);
            
            Series.Add(new SeriesItemViewModel(seriesName, sheetsInSeries.Count, isSelected));
        }
        
        SeriesItemsControl.ItemsSource = Series;
    }

    private Dictionary<string, List<SheetInfo>> ExtractSeriesFromSheets(List<SheetInfo> sheets)
    {
        var seriesGroups = new Dictionary<string, List<SheetInfo>>();
        
        // Regex pattern to extract series from sheet names like ABC-101, AB-12, A-1
        var sheetNamePattern = new Regex(@"^([A-Z]{1,3})", RegexOptions.IgnoreCase);
        
        foreach (var sheet in sheets)
        {
            var match = sheetNamePattern.Match(sheet.SheetName);
            if (match.Success)
            {
                var seriesName = match.Groups[1].Value.ToUpper();
                
                if (!seriesGroups.ContainsKey(seriesName))
                {
                    seriesGroups[seriesName] = new List<SheetInfo>();
                }
                
                seriesGroups[seriesName].Add(sheet);
            }
        }
        
        return seriesGroups;
    }

    private void SelectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var series in Series)
        {
            series.IsSelected = true;
        }
    }

    private void ClearAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var series in Series)
        {
            series.IsSelected = false;
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
}

public class SeriesItemViewModel : INotifyPropertyChanged
{
    private bool _isSelected;
    
    public string SeriesName { get; }
    public int SheetCount { get; }
    public string SheetCountText => $"({SheetCount} sheet{(SheetCount == 1 ? "" : "s")})";
    
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
    
    public SeriesItemViewModel(string seriesName, int sheetCount, bool isSelected = false)
    {
        SeriesName = seriesName;
        SheetCount = sheetCount;
        _isSelected = isSelected;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}