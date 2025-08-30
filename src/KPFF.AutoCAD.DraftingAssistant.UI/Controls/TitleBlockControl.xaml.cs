using System.IO;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class TitleBlockControl : BaseUserControl
{
    private readonly ITitleBlockService _titleBlockService;
    private readonly IProjectConfigurationService _configService;
    private readonly MultiDrawingTitleBlockService _multiDrawingService;

    public TitleBlockControl() : this(null, null, null, null)
    {
    }

    public TitleBlockControl(
        ILogger? logger,
        INotificationService? notificationService,
        ITitleBlockService? titleBlockService,
        IProjectConfigurationService? configService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
        
        // Use constructor injection or fall back to service locator
        _titleBlockService = titleBlockService ?? GetTitleBlockService();
        _configService = configService ?? GetConfigurationService();
        
        // Initialize multi-drawing service
        _multiDrawingService = GetMultiDrawingService();
        
        // Load initial display information
        _ = LoadInitialDisplayAsync();
    }

    private static ITitleBlockService GetTitleBlockService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var excelReader = new ExcelReaderService(logger);
        var drawingOps = new DrawingOperations(logger);
        return new TitleBlockService(logger, excelReader, drawingOps);
    }

    private static IProjectConfigurationService GetConfigurationService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        return new ProjectConfigurationService(logger);
    }

    private static MultiDrawingTitleBlockService GetMultiDrawingService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(logger);
        var backupCleanupService = new BackupCleanupService(logger);
        var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService);
        var excelReaderService = new ExcelReaderService(logger);
        var titleBlockService = new TitleBlockService(logger, excelReaderService, new DrawingOperations(logger));
        
        return new MultiDrawingTitleBlockService(
            logger,
            drawingAccessService,
            externalDrawingManager,
            titleBlockService,
            excelReaderService);
    }

    private async void UpdateTitleBlocksButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteTitleBlocksUpdate();
    }

    private void InsertTitleBlockButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Logger.LogInformation("Insert Title Block button clicked - executing KPFFINSERTTITLEBLOCK command");
            
            // Execute the AutoCAD command to insert title block
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager?.MdiActiveDocument;
            if (doc != null)
            {
                // Send the command to AutoCAD for execution with newline to complete the command
                doc.SendStringToExecute("KPFFINSERTTITLEBLOCK\n", true, false, false);
                
                UpdateStatus("Insert Title Block command executed. Title block will be inserted at 0,0.");
            }
            else
            {
                Logger.LogError("No active document found for Insert Title Block command");
                UpdateStatus("ERROR: No active drawing found. Please open a drawing before inserting blocks.");
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error executing Insert Title Block command: {ex.Message}", ex);
            UpdateStatus($"ERROR: Failed to execute Insert Title Block command - {ex.Message}");
        }
    }
    
    private async Task ExecuteTitleBlocksUpdate()
    {
        try
        {
            Logger.LogInformation("Executing title blocks update with multi-drawing support");
            
            // Use AutoCAD services when actually performing operations
            var autocadLogger = new AutoCADLogger();
            
            // Create production-ready multi-drawing service
            var drawingAccessService = new DrawingAccessService(autocadLogger);
            var backupCleanupService = new BackupCleanupService(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger, backupCleanupService);
            var excelReaderService = new ExcelReaderService(autocadLogger);
            var titleBlockService = new TitleBlockService(autocadLogger, excelReaderService, new DrawingOperations(autocadLogger));
            
            var multiDrawingService = new MultiDrawingTitleBlockService(
                autocadLogger,
                drawingAccessService,
                externalDrawingManager,
                titleBlockService,
                excelReaderService);
            
            autocadLogger.LogInformation("Using multi-drawing batch processing for title blocks update");
            
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

            UpdateStatus($"Processing {selectedSheets.Count} sheets for title blocks update...\n" +
                        "Analyzing drawing states and preparing batch operations...\n");
            autocadLogger.LogInformation($"Starting title blocks batch update for {selectedSheets.Count} sheets");

            // Get sheet names
            var selectedSheetNames = selectedSheets.Select(s => s.SheetName).ToList();

            // Execute the multi-drawing update
            var result = await multiDrawingService.UpdateTitleBlocksAcrossDrawingsAsync(selectedSheetNames, config, selectedSheets);

            // Update UI with results
            var successMessage = result.Successes.Count > 0 
                ? $"✓ Successfully updated title blocks for {result.Successes.Count} sheets:\n  - {string.Join("\n  - ", result.Successes.Select(s => s.SheetName))}\n\n"
                : "";
            
            var failureMessage = result.Failures.Count > 0
                ? $"✗ Failed to update title blocks for {result.Failures.Count} sheets:\n{string.Join("\n", result.Failures.Select(f => $"  - {f.SheetName}: {f.ErrorMessage}"))}\n"
                : "";

            var statusMessage = $"{successMessage}{failureMessage}";
            
            if (result.Successes.Count > 0 && result.Failures.Count == 0)
            {
                statusMessage += "Title blocks update completed successfully!";
            }
            else if (result.Successes.Count > 0 && result.Failures.Count > 0)
            {
                statusMessage += "Title blocks update completed with some failures.";
            }
            else
            {
                statusMessage += "Title blocks update failed for all selected sheets.";
            }

            UpdateStatus(statusMessage);
            autocadLogger.LogInformation($"Title blocks batch update completed. Success: {result.Successes.Count}, Failed: {result.Failures.Count}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Failed to update title blocks: {ex.Message}";
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
            var testConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\DBRT_Config.json";
            
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
            Logger.LogError($"Error getting shared configuration from sibling: {ex.Message}", ex);
            return null;
        }
    }

    private async Task<List<SheetInfo>> GetSelectedSheetsAsync(ProjectConfiguration config)
    {
        try
        {
            // Check if "Apply to current sheet only" is checked
            if (ApplyToCurrentSheetCheckBox.IsChecked == true)
            {
                var currentLayoutName = GetCurrentLayoutName();
                if (!string.IsNullOrEmpty(currentLayoutName))
                {
                    // Get all sheets to find the one that matches the current layout
                    List<SheetInfo> availableSheets;
                    if (config.SelectedSheets.Count > 0)
                    {
                        availableSheets = config.SelectedSheets;
                    }
                    else
                    {
                        var reader = new ExcelReaderService((IApplicationLogger)Logger);
                        availableSheets = await reader.ReadSheetIndexAsync(config.ProjectIndexFilePath, config);
                    }
                    
                    // Find the sheet that matches the current layout
                    var currentSheet = availableSheets.FirstOrDefault(s => s.SheetName == currentLayoutName);
                    if (currentSheet != null)
                    {
                        Logger.LogInformation($"Applying to current sheet only: {currentLayoutName}");
                        return new List<SheetInfo> { currentSheet };
                    }
                    else
                    {
                        Logger.LogWarning($"Current layout '{currentLayoutName}' not found in sheet list");
                        return new List<SheetInfo>();
                    }
                }
                else
                {
                    Logger.LogWarning("Could not determine current layout name");
                    return new List<SheetInfo>();
                }
            }
            
            // Return the selected sheets from configuration
            return await Task.FromResult(config.SelectedSheets ?? new List<SheetInfo>());
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to get selected sheets: {ex.Message}", ex);
            return new List<SheetInfo>();
        }
    }

    private async Task LoadInitialDisplayAsync()
    {
        try
        {
            var config = await LoadProjectConfigurationAsync();
            if (config != null)
            {
                var selectedSheets = await GetSelectedSheetsAsync(config);
                UpdateInfoDisplay(config, selectedSheets);
            }
            else
            {
                await Task.Run(() =>
                {
                    // Update initial status
                    Dispatcher.BeginInvoke(() =>
                    {
                        UpdateStatus("Ready to update title blocks.\n\n" +
                                   "1. Configure your project in the Configuration tab\n" +
                                   "2. Select the sheets you want to update\n" +
                                   "3. Click 'Update Title Blocks' to apply changes\n\n" +
                                   "The system will read title block data from the SHEET_INDEX table\n" +
                                   "and update title blocks across all selected drawings.");
                    });
                });
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to load initial display: {ex.Message}", ex);
        }
    }

    private async void ApplyToCurrentSheetCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        await RefreshSheetSelectionDisplay();
    }

    private async Task RefreshSheetSelectionDisplay()
    {
        try
        {
            var config = await LoadProjectConfigurationAsync();
            if (config != null)
            {
                var selectedSheets = await GetSelectedSheetsAsync(config);
                UpdateInfoDisplay(config, selectedSheets);
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error refreshing sheet selection display: {ex.Message}", ex);
        }
    }

    private void UpdateInfoDisplay(ProjectConfiguration? config, List<SheetInfo> selectedSheets)
    {
        var displayText = $"Title Block Management\n\n";
        
        if (config != null)
        {
            displayText += $"Project: {config.ProjectName}\n\n";
        }
        
        // Show sheet selection mode
        if (ApplyToCurrentSheetCheckBox.IsChecked == true)
        {
            var currentLayoutName = GetCurrentLayoutName();
            displayText += $"Mode: Apply to current sheet only\n";
            if (!string.IsNullOrEmpty(currentLayoutName))
            {
                displayText += $"Current Sheet: {currentLayoutName}\n\n";
                if (selectedSheets.Count > 0)
                {
                    var sheet = selectedSheets[0];
                    displayText += $"Target Sheet: {sheet.SheetName} - {sheet.DrawingTitle}\n";
                }
                else
                {
                    displayText += $"WARNING: Current sheet '{currentLayoutName}' not found in project index\n";
                }
            }
            else
            {
                displayText += "WARNING: Could not determine current sheet\n";
            }
        }
        else
        {
            displayText += $"Mode: Apply to selected sheets ({selectedSheets.Count} sheets)\n";
            displayText += $"Selected Sheet(s):\n";
            
            if (selectedSheets.Count > 0)
            {
                var displayLimit = 5; // Show first 5 sheets
                foreach (var sheet in selectedSheets.Take(displayLimit))
                {
                    displayText += $"  • {sheet.SheetName} - {sheet.DrawingTitle}\n";
                }
                
                if (selectedSheets.Count > displayLimit)
                {
                    displayText += $"  ... and {selectedSheets.Count - displayLimit} more sheets\n";
                }
            }
            else
            {
                displayText += "  No sheets selected\n";
            }
        }
        
        displayText += "\nTitle block data will be read from the SHEET_INDEX table\n" +
                      "and applied across all target drawings.";
        
        UpdateStatus(displayText);
    }

    private void UpdateStatus(string message)
    {
        if (TitleBlockTextBlock != null)
        {
            TitleBlockTextBlock.Text = message;
            TitleBlockTextBlock.UpdateLayout();
        }
    }

    private string? GetCurrentLayoutName()
    {
        try
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager?.MdiActiveDocument;
            if (doc != null)
            {
                var layoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                return layoutManager.CurrentLayout;
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to get current layout name: {ex.Message}");
        }
        return null;
    }
}