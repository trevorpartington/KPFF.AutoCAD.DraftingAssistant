using System.Windows;
using Microsoft.Win32;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConfigurationControl : BaseUserControl
{
    private readonly IProjectConfigurationService _configService;
    private readonly IExcelReader _excelReader;
    private ProjectConfiguration? _currentProject;
    private List<SheetInfo> _availableSheets = new();
    private List<SheetInfo> _selectedSheets = new();

    public ConfigurationControl()
    {
        InitializeComponent();
        
        // TODO: Replace with proper dependency injection
        var logger = new DebugLogger();
        _configService = new ProjectConfigurationService(logger);
        _excelReader = new ExcelReaderService(logger);
    }

    private async void SelectProjectButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new ProjectSelectionDialog
            {
                Owner = Window.GetWindow(this)
            };

            if (dialog.ShowDialog() == true && dialog.SelectedProject != null)
            {
                _currentProject = dialog.SelectedProject;
                ActiveProjectTextBlock.Text = _currentProject.ProjectName;
                ActiveProjectTextBlock.FontStyle = FontStyles.Normal;
                
                await LoadProjectDetails();
            }
        }
        catch (Exception ex)
        {
            ShowErrorNotification($"Error loading project: {ex.Message}");
        }
    }

    private async void SelectSheetsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentProject == null)
        {
            ShowWarningNotification("Please select a project first.");
            return;
        }

        try
        {
            // Load available sheets from Excel file
            if (System.IO.File.Exists(_currentProject.ProjectIndexFilePath))
            {
                _availableSheets = await _excelReader.ReadSheetIndexAsync(_currentProject.ProjectIndexFilePath, _currentProject);
                
                if (_availableSheets.Count == 0)
                {
                    ShowWarningNotification("No sheets found in project index file.");
                    return;
                }

                // Show sheet selection dialog
                var dialog = new SheetSelectionDialog(_availableSheets, _selectedSheets)
                {
                    Owner = Window.GetWindow(this)
                };
                
                if (dialog.ShowDialog() == true)
                {
                    _selectedSheets = dialog.SelectedSheets;
                    UpdateConfigurationDisplay($"Selected {_selectedSheets.Count} of {_availableSheets.Count} sheets:\n\n" + 
                                             string.Join("\n", _selectedSheets.Select(s => $"• {s.SheetName} - {s.DrawingTitle}")));
                }
            }
            else
            {
                ShowErrorNotification($"Project index file not found: {_currentProject.ProjectIndexFilePath}");
            }
        }
        catch (Exception ex)
        {
            ShowErrorNotification($"Error loading sheets: {ex.Message}");
        }
    }


    private async Task LoadProjectDetails()
    {
        if (_currentProject == null) return;

        try
        {
            var details = new List<string>
            {
                $"Project: {_currentProject.ProjectName}",
                $"Client: {_currentProject.ClientName}",
                $"Excel File: {_currentProject.ProjectIndexFilePath}",
                $"DWG Path: {_currentProject.ProjectDWGFilePath}",
                "",
                "Configuration Details:",
                $"• Sheet Pattern: {_currentProject.SheetNaming.Pattern}",
                $"• Series Group: {_currentProject.SheetNaming.SeriesGroup}, Number Group: {_currentProject.SheetNaming.NumberGroup}",
                $"• Sheets Worksheet: {_currentProject.Worksheets.Sheets}",
                $"• Notes Worksheet: {_currentProject.Worksheets.Notes}",
                $"• Max Notes per Sheet: {_currentProject.ConstructionNotes.MaxNotesPerSheet}",
                "",
            };

            // Validate configuration
            var isValid = _configService.ValidateConfiguration(_currentProject, out var errors);
            if (isValid)
            {
                details.Add("✓ Configuration is valid");
                
                // Try to load sheet count
                if (System.IO.File.Exists(_currentProject.ProjectIndexFilePath))
                {
                    var sheets = await _excelReader.ReadSheetIndexAsync(_currentProject.ProjectIndexFilePath, _currentProject);
                    details.Add($"✓ Found {sheets.Count} sheets in index");
                    _availableSheets = sheets;
                    
                    if (_selectedSheets.Count > 0)
                    {
                        details.Add($"✓ {_selectedSheets.Count} sheets selected for processing");
                    }
                }
            }
            else
            {
                details.Add("✗ Configuration has errors:");
                details.AddRange(errors.Select(e => $"  • {e}"));
            }

            UpdateConfigurationDisplay(string.Join("\n", details));
        }
        catch (Exception ex)
        {
            UpdateConfigurationDisplay($"Error loading project details: {ex.Message}");
        }
    }

    private void UpdateConfigurationDisplay(string text)
    {
        ConfigurationTextBlock.Text = text;
    }

    private void ShowErrorNotification(string message)
    {
        // TODO: Implement proper error notification
        UpdateConfigurationDisplay($"ERROR: {message}");
    }

    private void ShowWarningNotification(string message)
    {
        // TODO: Implement proper warning notification
        UpdateConfigurationDisplay($"WARNING: {message}");
    }
}