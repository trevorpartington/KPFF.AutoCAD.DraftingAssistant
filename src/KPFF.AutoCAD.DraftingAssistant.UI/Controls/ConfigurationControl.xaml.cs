using System.IO;
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
    private string? _currentProjectFilePath;
    private List<SheetInfo> _availableSheets = new();
    private List<SheetInfo> _selectedSheets = new();

    /// <summary>
    /// Gets the current project configuration with selected sheets
    /// </summary>
    public ProjectConfiguration? CurrentConfiguration => _currentProject;

    public ConfigurationControl() : this(null, null, null, null)
    {
    }

    public ConfigurationControl(
        IProjectConfigurationService? configService,
        IExcelReader? excelReader,
        ILogger? logger,
        INotificationService? notificationService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
        
        // Use constructor injection or fall back to service locator
        _configService = configService ?? GetConfigurationService();
        _excelReader = excelReader ?? GetExcelReaderService();
        
        // Load default project configuration
        _ = LoadDefaultProjectAsync();
    }

    private static IProjectConfigurationService GetConfigurationService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        // This prevents heap corruption from AutoCAD object disposal issues
        var logger = new DebugLogger();
        return new ProjectConfigurationService(logger);
    }

    private static IExcelReader GetExcelReaderService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        // This prevents heap corruption from AutoCAD object disposal issues
        var logger = new DebugLogger();
        return new ExcelReaderService(logger);
    }

    private async Task LoadDefaultProjectAsync()
    {
        await ExceptionHandler.TryExecuteAsync(async () =>
        {
            // Get the solution directory and construct path to testdata
            var currentDirectory = Directory.GetCurrentDirectory();
            var solutionRoot = FindSolutionRoot(currentDirectory);
            if (solutionRoot != null)
            {
                var defaultConfigPath = Path.Combine(solutionRoot, "testdata", "DBRT Test", "DBRT_Config.json");
                if (File.Exists(defaultConfigPath))
                {
                    _currentProject = await _configService.LoadConfigurationAsync(defaultConfigPath);
                    if (_currentProject != null)
                    {
                        _currentProjectFilePath = defaultConfigPath;
                        ActiveProjectTextBlock.Text = _currentProject.ProjectName;
                        ActiveProjectTextBlock.FontStyle = FontStyles.Normal;
                        
                        await LoadProjectDetails();
                        await LoadAndSelectAllSheetsAsync();
                        
                    }
                }
            }
        }, Logger, NotificationService, "LoadDefaultProjectAsync", showUserMessage: false);
    }

    private static string? FindSolutionRoot(string startPath)
    {
        var directory = new DirectoryInfo(startPath);
        while (directory != null)
        {
            if (directory.GetFiles("*.sln").Length > 0 || 
                directory.GetDirectories("testdata").Length > 0)
            {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
        return null;
    }

    private async Task LoadAndSelectAllSheetsAsync()
    {
        if (_currentProject == null) return;

        try
        {
            // Load available sheets from Excel file
            if (File.Exists(_currentProject.ProjectIndexFilePath))
            {
                _availableSheets = await _excelReader.ReadSheetIndexAsync(_currentProject.ProjectIndexFilePath, _currentProject);
                
                if (_availableSheets.Count > 0)
                {
                    // Select all sheets by default
                    _selectedSheets = new List<SheetInfo>(_availableSheets);
                    _currentProject.SelectedSheets = new List<SheetInfo>(_selectedSheets);
                    
                    // Update display to show all sheets are selected
                    var displayText = $"Default Configuration Loaded\n\n" +
                                    $"Project: {_currentProject.ProjectName}\n" +
                                    $"Client: {_currentProject.ClientName}\n\n" +
                                    $"Automatically selected all {_selectedSheets.Count} sheets:\n\n" +
                                    string.Join("\n", _selectedSheets.Take(10).Select(s => $"• {s.SheetName} - {s.DrawingTitle}"));
                    
                    if (_selectedSheets.Count > 10)
                    {
                        displayText += $"\n... and {_selectedSheets.Count - 10} more sheets";
                    }
                    
                    UpdateConfigurationDisplay(displayText);
                }
            }
        }
        catch (Exception ex)
        {
            UpdateConfigurationDisplay($"Error loading sheets for auto-selection: {ex.Message}");
        }
    }

    private async void SelectProjectButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new ProjectSelectionDialog(_currentProject, _currentProjectFilePath)
            {
                Owner = Window.GetWindow(this)
            };

            if (dialog.ShowDialog() == true && dialog.SelectedProject != null)
            {
                _currentProject = dialog.SelectedProject;
                _currentProjectFilePath = dialog.SelectedProjectFilePath;
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
                    _currentProject.SelectedSheets = new List<SheetInfo>(_selectedSheets);
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



    private void ShowInfoNotification(string message)
    {
        Logger.LogInformation(message);
        NotificationService.ShowInformation("Configuration", message);
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
                $"• Sheet Index Table: {_currentProject.Tables.SheetIndex}",
                $"• Excel Notes Table: {_currentProject.Tables.ExcelNotes}",
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
        Logger.LogError(message);
        NotificationService.ShowError("Configuration Error", message);
        UpdateConfigurationDisplay($"ERROR: {message}");
    }

    private void ShowWarningNotification(string message)
    {
        Logger.LogWarning(message);
        NotificationService.ShowWarning("Configuration Warning", message);
        UpdateConfigurationDisplay($"WARNING: {message}");
    }

}