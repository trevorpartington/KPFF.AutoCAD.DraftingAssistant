using System.IO;
using System.Linq;
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
    private bool _hasInitiallyLoadedSheets = false;

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
        
        // Load default project configuration when the control is fully loaded
        this.Loaded += ConfigurationControl_Loaded;
        
        // Display will be initialized when Loaded event fires
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
                        
                        // Ensure UI updates happen on the UI thread
                        Dispatcher.BeginInvoke(() =>
                        {
                            ActiveProjectTextBlock.Text = _currentProject.ProjectName;
                            ActiveProjectTextBlock.FontStyle = FontStyles.Normal;
                        });
                        
                        await LoadProjectDetails();
                        RefreshDisplay();
                        
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
                
                // Reset the sheet loading flag when switching projects
                _hasInitiallyLoadedSheets = false;
                _selectedSheets.Clear();
                
                await LoadProjectDetails();
                RefreshDisplay();
                
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
                    
                    // Mark that we've loaded sheets (to preserve this selection)
                    if (_selectedSheets.Count > 0)
                    {
                        _hasInitiallyLoadedSheets = true;
                    }
                    
                    RefreshDisplay();
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
                
                // Try to load sheet count and preserve selection
                if (System.IO.File.Exists(_currentProject.ProjectIndexFilePath))
                {
                    // Save current selection to restore later
                    var savedSelection = new List<SheetInfo>(_selectedSheets);
                    
                    var sheets = await _excelReader.ReadSheetIndexAsync(_currentProject.ProjectIndexFilePath, _currentProject);
                    details.Add($"✓ Found {sheets.Count} sheets in index");
                    _availableSheets = sheets;
                    
                    // Restore selection or auto-select all on first load
                    if (!_hasInitiallyLoadedSheets && sheets.Count > 0)
                    {
                        // First time loading - select all sheets
                        _selectedSheets = new List<SheetInfo>(sheets);
                        _hasInitiallyLoadedSheets = true;
                        details.Add($"✓ Automatically selected all {_selectedSheets.Count} sheets (initial load)");
                    }
                    else if (savedSelection.Count > 0)
                    {
                        // Restore previous selection by matching sheet names
                        _selectedSheets = sheets
                            .Where(sheet => savedSelection.Any(saved => saved.SheetName == sheet.SheetName))
                            .ToList();
                        details.Add($"✓ Restored {_selectedSheets.Count} sheets from previous selection");
                    }
                    
                    // Update project configuration with current selection
                    _currentProject.SelectedSheets = new List<SheetInfo>(_selectedSheets);
                    
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

            RefreshDisplay(details);
            
        }
        catch (Exception ex)
        {
            RefreshDisplay(new List<string> { $"Error loading project details: {ex.Message}" });
        }
    }


    private void UpdateConfigurationDisplay(string text)
    {
        ConfigurationTextBlock.Text = text;
    }

    private void RefreshDisplay(List<string>? statusMessages = null)
    {
        try
        {
            var readout = BuildStandardReadout(
                _currentProject?.ProjectName,
                statusMessages,
                _selectedSheets.Count > 0 ? _selectedSheets : null
            );
            ConfigurationTextBlock.Text = readout;
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error refreshing configuration display: {ex.Message}", ex);
            // Fallback to a simple display
            ConfigurationTextBlock.Text = $"Active Project: {_currentProject?.ProjectName ?? "No project selected"}\n" +
                                         "────────────────────────────────────────────────────────\n\n" +
                                         "Ready for configuration.\n\n" +
                                         "────────────────────────────────────────────────────────\n" +
                                         "Selected Sheets: 0\n" +
                                         "────────────────────────────────────────────────────────\n" +
                                         "No sheets selected for processing";
        }
    }

    private void ShowErrorNotification(string message)
    {
        Logger.LogError(message);
        NotificationService.ShowError("Configuration Error", message);
        RefreshDisplay(new List<string> { $"ERROR: {message}" });
    }

    private void ShowWarningNotification(string message)
    {
        Logger.LogWarning(message);
        NotificationService.ShowWarning("Configuration Warning", message);
        RefreshDisplay(new List<string> { $"WARNING: {message}" });
    }

    private async void ConfigurationControl_Loaded(object sender, RoutedEventArgs e)
    {
        // Only load default project after the control is fully loaded and palette is created
        // This prevents any potential issues during palette initialization
        RefreshDisplay(); // Initialize standard display format
        await LoadDefaultProjectAsync();
    }

}