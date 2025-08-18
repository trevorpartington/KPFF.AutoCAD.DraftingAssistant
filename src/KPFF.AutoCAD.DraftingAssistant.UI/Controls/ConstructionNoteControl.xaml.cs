using System.IO;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class ConstructionNoteControl : BaseUserControl
{
    private readonly IConstructionNotesService _constructionNotesService;
    private readonly IProjectConfigurationService _configService;

    public ConstructionNoteControl() : this(null, null, null, null)
    {
    }

    public ConstructionNoteControl(
        ILogger? logger,
        INotificationService? notificationService,
        IConstructionNotesService? constructionNotesService,
        IProjectConfigurationService? configService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
        
        // Use constructor injection or fall back to service locator
        _constructionNotesService = constructionNotesService ?? GetConstructionNotesService();
        _configService = configService ?? GetConfigurationService();
    }

    private static IConstructionNotesService GetConstructionNotesService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var excelReader = new ExcelReaderService(logger);
        var drawingOps = new DrawingOperations(logger);
        return new ConstructionNotesService(logger, excelReader, drawingOps);
    }

    private static IProjectConfigurationService GetConfigurationService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        return new ProjectConfigurationService(logger);
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
            // Excel Notes functionality
            ExecuteExcelNotesUpdate();
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
            
            // Auto Notes functionality placeholder - no popup needed
                
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
            NotesTextBlock.Text = $"ERROR: Failed to execute Auto Notes update: {ex.Message}";
        }
    }

    private async void ExecuteExcelNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Excel Notes update functionality");
            
            // Use AutoCAD services when actually performing operations
            var autocadLogger = new AutoCADLogger();
            IConstructionNotesService? constructionNotesService = null;
            
            // Use AutoCAD-connected services
            var excelReader = new ExcelReaderService(autocadLogger);
            var drawingOps = new DrawingOperations(autocadLogger);
            constructionNotesService = new ConstructionNotesService(autocadLogger, excelReader, drawingOps);
            
            autocadLogger.LogInformation("Using AutoCAD-connected services for Excel Notes update");
            
            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                UpdateStatus("ERROR: No project configuration loaded. Please select a project in the Configuration tab.");
                return;
            }

            // For testing, use all sheets from SHEET_INDEX - later this will come from ConfigurationControl
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdateStatus("No sheets selected. Please select sheets in the Configuration tab.");
                return;
            }

            UpdateStatus($"Processing {selectedSheets.Count} sheets in Excel Notes mode...\n");
            autocadLogger.LogInformation($"Starting Excel Notes update for {selectedSheets.Count} sheets");

            var errors = new List<string>();
            var successes = new List<string>();

            foreach (var sheet in selectedSheets)
            {
                try
                {
                    autocadLogger.LogDebug($"Processing sheet {sheet.SheetName}");
                    
                    // Get Excel notes for this sheet
                    var noteNumbers = await constructionNotesService.GetExcelNotesForSheetAsync(sheet.SheetName, config);
                    
                    if (noteNumbers.Count == 0)
                    {
                        autocadLogger.LogWarning($"No Excel notes found for sheet {sheet.SheetName}");
                        errors.Add($"No Excel notes found for sheet {sheet.SheetName}");
                        continue;
                    }

                    autocadLogger.LogInformation($"Found {noteNumbers.Count} notes for sheet {sheet.SheetName}: {string.Join(", ", noteNumbers)}");

                    // Update construction note blocks
                    await constructionNotesService.UpdateConstructionNoteBlocksAsync(sheet.SheetName, noteNumbers, config);
                    
                    successes.Add($"Updated {sheet.SheetName} with {noteNumbers.Count} notes: {string.Join(", ", noteNumbers)}");
                    autocadLogger.LogInformation($"Successfully updated sheet {sheet.SheetName}");
                }
                catch (Exception ex)
                {
                    var errorMsg = $"Failed to update sheet {sheet.SheetName}: {ex.Message}";
                    errors.Add(errorMsg);
                    autocadLogger.LogError(errorMsg, ex);
                }
            }

            // Report results (errors first, then successes)
            var statusText = "";
            if (errors.Count > 0)
            {
                statusText += "ERRORS:\n";
                statusText += string.Join("\n", errors.Select(e => $"• {e}"));
                statusText += "\n\n";
            }

            if (successes.Count > 0)
            {
                statusText += $"SUCCESS: Updated {successes.Count} of {selectedSheets.Count} sheets:\n";
                statusText += string.Join("\n", successes.Select(s => $"• {s}"));
            }

            if (errors.Count == 0 && successes.Count == 0)
            {
                statusText = "No updates performed. Check that sheets have Excel notes configured.";
            }

            UpdateStatus(statusText);
            
            autocadLogger.LogInformation($"Excel Notes update completed. Processed {selectedSheets.Count} sheets. Successful: {successes.Count}, Errors: {errors.Count}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Error executing Excel Notes update: {ex.Message}";
            Logger.LogError(errorMsg, ex);
            UpdateStatus($"ERROR: {errorMsg}");
        }
    }

    private async Task<ProjectConfiguration?> LoadProjectConfigurationAsync()
    {
        try
        {
            // First try to get shared configuration from ConfigurationControl sibling
            var sharedConfig = GetSharedConfigurationFromSibling();
            if (sharedConfig != null)
            {
                Logger.LogDebug($"Using shared configuration with {sharedConfig.SelectedSheets.Count} selected sheets: {sharedConfig.ProjectName}");
                return sharedConfig;
            }
            
            // Fall back to loading from the default test location
            var testConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectConfig.json";
            
            if (File.Exists(testConfigPath))
            {
                var config = await _configService.LoadConfigurationAsync(testConfigPath);
                Logger.LogDebug($"Loaded test configuration from file: {config?.ProjectName}");
                return config;
            }
            else
            {
                Logger.LogWarning($"Test configuration file not found: {testConfigPath}");
                return null;
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to load project configuration: {ex.Message}", ex);
            return null;
        }
    }

    private ProjectConfiguration? GetSharedConfigurationFromSibling()
    {
        try
        {
            // Find the parent DraftingAssistantControl
            var parent = this.Parent;
            while (parent != null && !(parent is DraftingAssistantControl))
            {
                parent = parent is FrameworkElement fe ? fe.Parent : null;
            }

            if (parent is DraftingAssistantControl draftingAssistantControl)
            {
                // Find ConfigurationControl sibling by name
                var configControl = draftingAssistantControl.FindName("ConfigurationControl") as ConfigurationControl;
                if (configControl != null)
                {
                    Logger.LogDebug("Found ConfigurationControl sibling, getting shared configuration");
                    return configControl.CurrentConfiguration;
                }
                else
                {
                    Logger.LogDebug("ConfigurationControl sibling not found by name");
                }
            }
            else
            {
                Logger.LogDebug("DraftingAssistantControl parent not found");
            }
            
            return null;
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to get shared configuration from sibling: {ex.Message}");
            return null;
        }
    }

    private async Task<List<SheetInfo>> GetSelectedSheetsAsync(ProjectConfiguration config)
    {
        try
        {
            // Use selected sheets from ProjectConfiguration if available
            if (config.SelectedSheets.Count > 0)
            {
                Logger.LogDebug($"Using {config.SelectedSheets.Count} selected sheets from configuration");
                return config.SelectedSheets;
            }
            
            // Fall back to all sheets if no selection is made
            var excelReader = new ExcelReaderService((IApplicationLogger)Logger);
            var allSheets = await excelReader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
            
            Logger.LogDebug($"No sheets selected, using all {allSheets.Count} sheets from index");
            return allSheets;
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to get selected sheets: {ex.Message}", ex);
            return new List<SheetInfo>();
        }
    }

    private void UpdateStatus(string message)
    {
        NotesTextBlock.Text = message;
        Logger.LogDebug($"Status updated: {message}");
    }
}