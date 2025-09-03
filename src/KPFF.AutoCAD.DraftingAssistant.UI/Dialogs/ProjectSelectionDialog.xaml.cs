using System.Windows;
using Microsoft.Win32;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class ProjectSelectionDialog : Window
{
    private readonly IProjectConfigurationService _configService;
    private readonly IExcelReader _excelReader;
    private ProjectConfiguration? _loadedProject;
    private string? _configurationFilePath;

    public ProjectConfiguration? SelectedProject => _loadedProject;
    public string? SelectedProjectFilePath => _configurationFilePath;

    public ProjectSelectionDialog() : this(null, null)
    {
    }

    public ProjectSelectionDialog(ProjectConfiguration? currentProject, string? currentProjectFilePath)
    {
        InitializeComponent();
        
        // TODO: Replace with proper dependency injection
        var logger = new DebugLogger();
        _configService = new ProjectConfigurationService(logger);
        _excelReader = new ExcelReaderService(logger);
        
        // Pre-populate with current project if provided
        if (currentProject != null && !string.IsNullOrEmpty(currentProjectFilePath))
        {
            _loadedProject = currentProject;
            _configurationFilePath = currentProjectFilePath;
            ConfigFilePathTextBox.Text = currentProjectFilePath;
            LoadProjectButton.IsEnabled = true;
            
            // Auto-load the project details and enable configure button
            _ = AutoLoadCurrentProjectAsync();
        }
    }

    private async Task AutoLoadCurrentProjectAsync()
    {
        try
        {
            await DisplayProjectDetails();
            ConfigureProjectButton.IsEnabled = true;
            OkButton.IsEnabled = true;
        }
        catch (Exception ex)
        {
            ShowError($"Error loading current project: {ex.Message}");
        }
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Project Configuration",
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
            DefaultExt = "json"
        };

        if (dialog.ShowDialog() == true)
        {
            ConfigFilePathTextBox.Text = dialog.FileName;
            _configurationFilePath = dialog.FileName;
            LoadProjectButton.IsEnabled = true;
        }
    }

    private async void LoadProjectButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrEmpty(ConfigFilePathTextBox.Text))
            {
                ShowError("Please select a configuration file.");
                return;
            }

            var config = await _configService.LoadConfigurationAsync(ConfigFilePathTextBox.Text);
            if (config != null)
            {
                _loadedProject = config;
                await DisplayProjectDetails();
                
                ConfigureProjectButton.IsEnabled = true;
                OkButton.IsEnabled = true;
            }
            else
            {
                ShowError("Failed to load project configuration.");
            }
        }
        catch (Exception ex)
        {
            ShowError($"Error loading project: {ex.Message}");
        }
    }

    private async void ConfigureProjectButton_Click(object sender, RoutedEventArgs e)
    {
        if (_loadedProject == null || string.IsNullOrEmpty(_configurationFilePath))
        {
            ShowError("Please load a project first.");
            return;
        }

        try
        {
            var dialog = new ProjectConfigWindow(_loadedProject, _configurationFilePath)
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true && dialog.ConfigurationSaved)
            {
                // Configuration was saved, reload the project details
                _loadedProject = dialog.UpdatedConfiguration;
                if (_loadedProject != null)
                {
                    await DisplayProjectDetails();
                    ProjectDetailsTextBlock.Text += "\n\n✓ Project configuration updated successfully.";
                }
            }
        }
        catch (Exception ex)
        {
            ShowError($"Error opening configuration editor: {ex.Message}");
        }
    }

    private async Task DisplayProjectDetails()
    {
        if (_loadedProject == null) return;

        try
        {
            var details = new List<string>
            {
                $"Project: {_loadedProject.ProjectName}",
                $"Client: {_loadedProject.ClientName}",
                $"Excel File: {_loadedProject.ProjectIndexFilePath}",
                $"DWG Path: {_loadedProject.ProjectDWGFilePath}",
                "",
                "Configuration Details:",
                $"• Sheet Index Table: {_loadedProject.Tables.SheetIndex}",
                $"• Excel Notes Table: {_loadedProject.Tables.ExcelNotes}",
                $"• Max Notes per Sheet: {_loadedProject.ConstructionNotes.MaxNotesPerSheet}",
                "",
            };

            // Validate configuration
            var isValid = _configService.ValidateConfiguration(_loadedProject, out var errors);
            if (isValid)
            {
                details.Add("✓ Configuration is valid");
                
                // Try to load sheet count
                if (System.IO.File.Exists(_loadedProject.ProjectIndexFilePath))
                {
                    var sheets = await _excelReader.ReadSheetIndexAsync(_loadedProject.ProjectIndexFilePath, _loadedProject);
                    details.Add($"✓ Found {sheets.Count} sheets in index");
                }
            }
            else
            {
                details.Add("✗ Configuration has errors:");
                details.AddRange(errors.Select(e => $"  • {e}"));
            }

            ProjectDetailsTextBlock.Text = string.Join("\n", details);
        }
        catch (Exception ex)
        {
            ProjectDetailsTextBlock.Text = $"Error loading project details: {ex.Message}";
        }
    }

    private void ShowError(string message)
    {
        ProjectDetailsTextBlock.Text = $"ERROR: {message}";
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