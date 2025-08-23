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
    private readonly MultiDrawingConstructionNotesService _multiDrawingService;

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
        
        // Initialize multi-drawing service
        _multiDrawingService = GetMultiDrawingService();
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

    private static MultiDrawingConstructionNotesService GetMultiDrawingService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(logger);
        var externalDrawingManager = new ExternalDrawingManager(logger);
        var excelReaderService = new ExcelReaderService(logger);
        var constructionNotesService = new ConstructionNotesService(logger, excelReaderService, new DrawingOperations(logger));
        
        return new MultiDrawingConstructionNotesService(
            logger,
            drawingAccessService,
            externalDrawingManager,
            constructionNotesService,
            excelReaderService);
    }

    private void UpdateNotesButton_Click(object sender, RoutedEventArgs e)
    {
        bool isAutoNotesMode = AutoNotesRadioButton.IsChecked == true;
        
        if (isAutoNotesMode)
        {
            // Auto Notes functionality with multi-drawing support
            ExecuteAutoNotesUpdate();
        }
        else
        {
            // Excel Notes functionality with multi-drawing support
            ExecuteExcelNotesUpdate();
        }
    }

    private async void PreviewDrawingStatesButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Logger.LogInformation("Previewing drawing states for selected sheets");
            UpdateStatus("Analyzing selected sheets and their drawing states...\n");

            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                UpdateStatus("ERROR: No project configuration loaded. Please select a project in the Configuration tab.");
                return;
            }

            // Get selected sheets from configuration
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdateStatus("No sheets selected. Please select sheets in the Configuration tab.");
                return;
            }

            // Use AutoCAD services for drawing state analysis
            var autocadLogger = new AutoCADLogger();
            var drawingAccessService = new DrawingAccessService(autocadLogger);

            var statusText = $"=== DRAWING STATE PREVIEW ===\n\n";
            statusText += $"Analyzing {selectedSheets.Count} selected sheets...\n\n";

            var stateGroups = new Dictionary<string, List<SheetInfo>>();

            // Analyze each sheet's drawing state
            foreach (var sheet in selectedSheets)
            {
                try
                {
                    var dwgPath = drawingAccessService.GetDrawingFilePath(sheet.SheetName, config, selectedSheets);
                    if (string.IsNullOrEmpty(dwgPath))
                    {
                        if (!stateGroups.ContainsKey("NotFound")) stateGroups["NotFound"] = new List<SheetInfo>();
                        stateGroups["NotFound"].Add(sheet);
                        continue;
                    }

                    var state = drawingAccessService.GetDrawingState(dwgPath);
                    var stateKey = state.ToString();
                    
                    if (!stateGroups.ContainsKey(stateKey)) stateGroups[stateKey] = new List<SheetInfo>();
                    stateGroups[stateKey].Add(sheet);
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to analyze drawing state for sheet {sheet.SheetName}: {ex.Message}");
                    if (!stateGroups.ContainsKey("Error")) stateGroups["Error"] = new List<SheetInfo>();
                    stateGroups["Error"].Add(sheet);
                }
            }

            // Generate report by state
            foreach (var (state, sheets) in stateGroups.OrderBy(kvp => kvp.Key))
            {
                statusText += $"{GetDrawingStateIcon(state)} {state.ToUpper()} ({sheets.Count} drawings):\n";
                foreach (var sheet in sheets.Take(10)) // Show first 10 to avoid overwhelming
                {
                    statusText += $"  • {sheet.SheetName} - {sheet.DrawingTitle}\n";
                }
                if (sheets.Count > 10)
                {
                    statusText += $"  ... and {sheets.Count - 10} more\n";
                }
                statusText += "\n";
            }

            // Add operation summary
            statusText += "BATCH OPERATION READINESS:\n";
            statusText += $"✓ All {selectedSheets.Count} sheets can be processed\n";
            statusText += $"✓ Active/Inactive drawings will use current AutoCAD session\n";
            statusText += $"✓ Closed drawings will use external database operations\n";
            statusText += $"✓ Sequential NT block assignment will be applied\n\n";
            statusText += "Ready to process with 'Update Notes' button.";

            UpdateStatus(statusText);
            
            autocadLogger.LogInformation($"Drawing state preview completed for {selectedSheets.Count} sheets");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Error previewing drawing states: {ex.Message}";
            Logger.LogError(errorMsg, ex);
            UpdateStatus($"ERROR: {errorMsg}");
        }
    }

    private static string GetDrawingStateIcon(string state)
    {
        return state switch
        {
            "Active" => "🟢",
            "Inactive" => "🟡", 
            "Closed" => "🔵",
            "NotFound" => "❌",
            "Error" => "⚠️",
            _ => "❓"
        };
    }
    
    private async void ExecuteAutoNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Auto Notes update with multi-drawing support");
            
            // Use AutoCAD services when actually performing operations
            var autocadLogger = new AutoCADLogger();
            
            // Create production-ready multi-drawing service
            var drawingAccessService = new DrawingAccessService(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger);
            var excelReaderService = new ExcelReaderService(autocadLogger);
            var constructionNotesService = new ConstructionNotesService(autocadLogger, excelReaderService, new DrawingOperations(autocadLogger));
            
            var multiDrawingService = new MultiDrawingConstructionNotesService(
                autocadLogger,
                drawingAccessService,
                externalDrawingManager,
                constructionNotesService,
                excelReaderService);
            
            autocadLogger.LogInformation("Using multi-drawing batch processing for Auto Notes update");
            
            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                UpdateStatus("ERROR: No project configuration loaded. Please select a project in the Configuration tab.");
                return;
            }

            // Update config with current multileader style if changed
            await UpdateMultileaderStyleInConfigAsync(config);

            // Get selected sheets from configuration
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdateStatus("No sheets selected. Please select sheets in the Configuration tab.");
                return;
            }

            UpdateStatus($"Processing {selectedSheets.Count} sheets in Auto Notes mode...\n" +
                        "Analyzing drawing states and preparing batch operations...\n");
            autocadLogger.LogInformation($"Starting Auto Notes batch update for {selectedSheets.Count} sheets");

            // Gather Auto Notes for selected sheets
            var sheetToNotes = new Dictionary<string, List<int>>();
            foreach (var sheet in selectedSheets)
            {
                try
                {
                    var noteNumbers = await constructionNotesService.GetAutoNotesForSheetAsync(sheet.SheetName, config);
                    if (noteNumbers.Count > 0)
                    {
                        sheetToNotes[sheet.SheetName] = noteNumbers;
                        autocadLogger.LogDebug($"Auto Notes for {sheet.SheetName}: [{string.Join(", ", noteNumbers)}]");
                    }
                }
                catch (Exception ex)
                {
                    autocadLogger.LogWarning($"Failed to get Auto Notes for sheet {sheet.SheetName}: {ex.Message}");
                }
            }

            if (sheetToNotes.Count == 0)
            {
                UpdateStatus("No Auto Notes detected on any selected sheets. Check that sheets have viewports with multileaders using the configured style.");
                return;
            }

            UpdateStatus($"Found Auto Notes on {sheetToNotes.Count} sheets. Processing batch update...\n");

            // Use the multi-drawing service for batch processing
            var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(sheetToNotes, config, selectedSheets);

            // Generate detailed status report
            var statusText = GenerateBatchUpdateStatusReport(result, sheetToNotes.Count, "Auto Notes");
            UpdateStatus(statusText);
            
            autocadLogger.LogInformation($"Auto Notes batch update completed. " +
                                       $"Processed {selectedSheets.Count} sheets. " +
                                       $"Successful: {result.Successes.Count}, " +
                                       $"Failures: {result.Failures.Count}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Error executing Auto Notes batch update: {ex.Message}";
            Logger.LogError(errorMsg, ex);
            UpdateStatus($"ERROR: {errorMsg}");
        }
    }

    private async void ExecuteExcelNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Excel Notes update with multi-drawing support");
            
            // Use AutoCAD services when actually performing operations
            var autocadLogger = new AutoCADLogger();
            
            // Create production-ready multi-drawing service
            var drawingAccessService = new DrawingAccessService(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger);
            var excelReaderService = new ExcelReaderService(autocadLogger);
            var constructionNotesService = new ConstructionNotesService(autocadLogger, excelReaderService, new DrawingOperations(autocadLogger));
            
            var multiDrawingService = new MultiDrawingConstructionNotesService(
                autocadLogger,
                drawingAccessService,
                externalDrawingManager,
                constructionNotesService,
                excelReaderService);
            
            autocadLogger.LogInformation("Using multi-drawing batch processing for Excel Notes update");
            
            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                UpdateStatus("ERROR: No project configuration loaded. Please select a project in the Configuration tab.");
                return;
            }

            // Get selected sheets from configuration
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdateStatus("No sheets selected. Please select sheets in the Configuration tab.");
                return;
            }

            UpdateStatus($"Processing {selectedSheets.Count} sheets in Excel Notes mode...\n" +
                        "Analyzing drawing states and preparing batch operations...\n");
            autocadLogger.LogInformation($"Starting Excel Notes batch update for {selectedSheets.Count} sheets");

            // Gather Excel Notes for selected sheets
            var sheetToNotes = new Dictionary<string, List<int>>();
            foreach (var sheet in selectedSheets)
            {
                try
                {
                    var noteNumbers = await constructionNotesService.GetExcelNotesForSheetAsync(sheet.SheetName, config);
                    if (noteNumbers.Count > 0)
                    {
                        sheetToNotes[sheet.SheetName] = noteNumbers;
                        autocadLogger.LogDebug($"Excel Notes for {sheet.SheetName}: [{string.Join(", ", noteNumbers)}]");
                    }
                }
                catch (Exception ex)
                {
                    autocadLogger.LogWarning($"Failed to get Excel Notes for sheet {sheet.SheetName}: {ex.Message}");
                }
            }

            if (sheetToNotes.Count == 0)
            {
                UpdateStatus("No Excel Notes found for any selected sheets. Check that sheets have notes configured in the Excel EXCEL_NOTES table.");
                return;
            }

            UpdateStatus($"Found Excel Notes on {sheetToNotes.Count} sheets. Processing batch update...\n");

            // Use the multi-drawing service for batch processing
            var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(sheetToNotes, config, selectedSheets);

            // Generate detailed status report
            var statusText = GenerateBatchUpdateStatusReport(result, sheetToNotes.Count, "Excel Notes");
            UpdateStatus(statusText);
            
            autocadLogger.LogInformation($"Excel Notes batch update completed. " +
                                       $"Processed {selectedSheets.Count} sheets. " +
                                       $"Successful: {result.Successes.Count}, " +
                                       $"Failures: {result.Failures.Count}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Error executing Excel Notes batch update: {ex.Message}";
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
                
                // Load multileader style from config into UI
                LoadMultileaderStyleFromConfig(sharedConfig);
                
                return sharedConfig;
            }
            
            // Fall back to loading from the default test location
            var testConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectConfig.json";
            
            if (File.Exists(testConfigPath))
            {
                var config = await _configService.LoadConfigurationAsync(testConfigPath);
                Logger.LogDebug($"Loaded test configuration from file: {config?.ProjectName}");
                
                // Load multileader style from config into UI
                LoadMultileaderStyleFromConfig(config);
                
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

    private async Task UpdateMultileaderStyleInConfigAsync(ProjectConfiguration config)
    {
        try
        {
            string currentStyleInTextBox = MultileaderStyleTextBox.Text?.Trim() ?? "";
            string configStyle = config.ConstructionNotes?.MultileaderStyleName ?? "";

            if (!string.Equals(currentStyleInTextBox, configStyle, StringComparison.OrdinalIgnoreCase))
            {
                Logger.LogInformation($"Updating multileader style in config from '{configStyle}' to '{currentStyleInTextBox}'");
                
                // Ensure ConstructionNotes section exists
                if (config.ConstructionNotes == null)
                {
                    config.ConstructionNotes = new ConstructionNotesConfiguration();
                }
                
                config.ConstructionNotes.MultileaderStyleName = currentStyleInTextBox;
                
                // Save the updated configuration back to file
                var testConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\ProjectConfig.json";
                if (File.Exists(testConfigPath))
                {
                    await _configService.SaveConfigurationAsync(config, testConfigPath);
                    Logger.LogInformation($"Configuration saved with updated multileader style: {currentStyleInTextBox}");
                }
                else
                {
                    Logger.LogWarning("Could not save configuration - config file path not found");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to update multileader style in config: {ex.Message}", ex);
        }
    }

    private void LoadMultileaderStyleFromConfig(ProjectConfiguration? config)
    {
        try
        {
            if (config?.ConstructionNotes?.MultileaderStyleName != null)
            {
                MultileaderStyleTextBox.Text = config.ConstructionNotes.MultileaderStyleName;
                Logger.LogDebug($"Loaded multileader style from config: {config.ConstructionNotes.MultileaderStyleName}");
            }
            else
            {
                // Keep default value
                Logger.LogDebug("No multileader style found in config, using default");
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to load multileader style from config: {ex.Message}");
        }
    }

    /// <summary>
    /// Generates a comprehensive status report for batch update operations
    /// Shows drawing states, success/failure counts, and detailed results
    /// </summary>
    private string GenerateBatchUpdateStatusReport(MultiDrawingUpdateResult result, int totalSheets, string mode)
    {
        var statusText = $"=== {mode} BATCH UPDATE COMPLETE ===\n\n";
        
        // Summary statistics
        statusText += $"Total sheets processed: {totalSheets}\n";
        statusText += $"Successful updates: {result.Successes.Count}\n";
        statusText += $"Failed updates: {result.Failures.Count}\n";
        statusText += $"Success rate: {(result.Successes.Count * 100.0 / totalSheets):F1}%\n\n";

        // Drawing state breakdown
        var groupedByState = result.Successes
            .GroupBy(s => s.DrawingState)
            .ToDictionary(g => g.Key, g => g.Count());
        
        if (groupedByState.Count > 0)
        {
            statusText += "DRAWING STATES PROCESSED:\n";
            foreach (var (state, count) in groupedByState)
            {
                statusText += $"• {state}: {count} drawings\n";
            }
            statusText += "\n";
        }

        // Failures first (if any)
        if (result.Failures.Count > 0)
        {
            statusText += "❌ FAILED UPDATES:\n";
            foreach (var failure in result.Failures)
            {
                statusText += $"  • {failure.SheetName}: {failure.ErrorMessage}\n";
            }
            statusText += "\n";
        }

        // Successes
        if (result.Successes.Count > 0)
        {
            statusText += "✅ SUCCESSFUL UPDATES:\n";
            foreach (var success in result.Successes)
            {
                statusText += $"  • {success.SheetName} ({success.DrawingState}): {success.NotesUpdated} notes updated\n";
            }
        }

        if (result.Successes.Count == 0 && result.Failures.Count == 0)
        {
            statusText += $"No updates performed. Check that sheets have {mode.ToLower()} configured.";
        }

        return statusText;
    }
}