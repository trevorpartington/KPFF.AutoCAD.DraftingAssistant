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
    private readonly IBatchOperationService _batchOperationService;
    private readonly SharedUIStateService _sharedUIState;

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
        
        // Initialize batch operation service
        _batchOperationService = GetBatchOperationService();
        
        // Initialize shared UI state
        _sharedUIState = SharedUIStateService.Instance;
        
        // Subscribe to shared state changes
        _sharedUIState.OnApplyToCurrentSheetOnlyChanged += OnSharedApplyToCurrentSheetOnlyChanged;
        
        // Initialize checkbox with shared state
        ApplyToCurrentSheetCheckBox.IsChecked = _sharedUIState.ApplyToCurrentSheetOnly;
        
        // Load initial display information when the control is fully loaded
        this.Loaded += TitleBlockControl_Loaded;
        
        // Display will be initialized when Loaded event fires
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

    private static IBatchOperationService GetBatchOperationService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(logger);
        var backupCleanupService = new BackupCleanupService(logger);
        var multileaderAnalyzer = new MultileaderAnalyzer(logger);
        var blockAnalyzer = new BlockAnalyzer(logger);
        var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
        var excelReaderService = new ExcelReaderService(logger);
        var constructionNotesService = new ConstructionNotesService(logger, excelReaderService, new DrawingOperations(logger));
        var titleBlockService = new TitleBlockService(logger, excelReaderService, new DrawingOperations(logger));
        
        // Create multi-drawing services
        var multiDrawingConstructionNotesService = new MultiDrawingConstructionNotesService(
            logger, drawingAccessService, externalDrawingManager, constructionNotesService, excelReaderService);
        var multiDrawingTitleBlockService = new MultiDrawingTitleBlockService(
            logger, drawingAccessService, externalDrawingManager, titleBlockService, excelReaderService);
        var drawingOperations = new DrawingOperations(logger);
        var plottingService = new PlottingService(logger, constructionNotesService, drawingOperations, excelReaderService, 
            multiDrawingConstructionNotesService, multiDrawingTitleBlockService);
        
        return new BatchOperationService(
            logger,
            drawingAccessService,
            multiDrawingConstructionNotesService,
            multiDrawingTitleBlockService,
            plottingService,
            constructionNotesService,
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
            Logger.LogInformation("Executing title blocks update using BatchOperationService");
            
            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                UpdateStatus("ERROR: No project configuration loaded. Please select a project in the Configure tab.");
                return;
            }

            // Get selected sheets from configuration
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdateStatus("No sheets selected. Please select sheets in the Configure tab.");
                return;
            }

            UpdateStatus($"Processing {selectedSheets.Count} sheets for title blocks update...\n" +
                        "Starting batch operation...\n");

            // Use BatchOperationService for proper async handling
            var sheetNames = selectedSheets.Select(s => s.SheetName).ToList();
            var applyToCurrentOnly = ApplyToCurrentSheetCheckBox.IsChecked == true;
            var progress = new Progress<BatchOperationProgress>(p => 
            {
                UpdateStatus($"[{p.ProgressPercentage}%] {p.Phase}: {p.CurrentOperationDescription}\n" +
                           $"Processing: {p.CompletedSheets}/{p.TotalSheets} sheets");
            });

            var result = await _batchOperationService.UpdateTitleBlocksAsync(
                sheetNames, config, applyToCurrentOnly, progress);

            // Generate status report
            var statusText = GenerateBatchOperationStatusReport(result, "Title Blocks");
            UpdateStatus(statusText);
            
            Logger.LogInformation($"Title blocks batch update completed via BatchOperationService. " +
                                $"Success: {result.Success}");
        }
        catch (Exception ex)
        {
            var errorMsg = $"Failed to update title blocks: {ex.Message}";
            Logger.LogError(errorMsg, ex);
            UpdateStatus($"ERROR: {errorMsg}");
        }
    }

    /// <summary>
    /// Generates a comprehensive status report for batch operation results
    /// Shows success/failure counts and detailed results from BatchOperationService
    /// </summary>
    private string GenerateBatchOperationStatusReport(BatchOperationResult result, string mode)
    {
        var statusText = $"=== {mode} BATCH UPDATE COMPLETE ===\n\n";
        
        // Summary statistics
        statusText += $"Operation Status: {(result.Success ? "SUCCESS" : "FAILED")}\n";
        
        if (!string.IsNullOrEmpty(result.ErrorMessage))
        {
            statusText += $"Error: {result.ErrorMessage}\n\n";
        }
        
        if (result.TotalSheets > 0)
        {
            statusText += $"Total sheets processed: {result.TotalSheets}\n";
            statusText += $"Successful updates: {result.SuccessfulSheets.Count}\n";
            statusText += $"Failed updates: {result.FailedSheets.Count}\n";
            statusText += $"Success rate: {result.SuccessRate:F1}%\n\n";

            // Failures first (if any)
            if (result.FailedSheets.Count > 0)
            {
                statusText += "❌ FAILED UPDATES:\n";
                foreach (var failure in result.FailedSheets)
                {
                    statusText += $"  • {failure.SheetName}: {failure.ErrorMessage}\n";
                }
                statusText += "\n";
            }

            // Successes
            if (result.SuccessfulSheets.Count > 0)
            {
                statusText += "✅ SUCCESSFUL UPDATES:\n";
                foreach (var success in result.SuccessfulSheets)
                {
                    statusText += $"  • {success.SheetName}: Update completed successfully\n";
                }
            }
        }
        else
        {
            statusText += $"No sheets processed. Check that sheets have {mode.ToLower()} configured.";
        }

        return statusText;
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
                    // Get selected sheets to find the one that matches the current layout
                    if (config.SelectedSheets.Count == 0)
                    {
                        Logger.LogWarning("No sheets selected, cannot apply to current sheet");
                        return new List<SheetInfo>();
                    }
                    
                    var availableSheets = config.SelectedSheets;
                    
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
                RefreshDisplay(BuildInfoMessages(config));
            }
            else
            {
                await Task.Run(() =>
                {
                    // Update initial status
                    Dispatcher.BeginInvoke(() =>
                    {
                        UpdateStatus("Ready to update title blocks.\n\n" +
                                   "1. Configure your project in the Configure tab\n" +
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
        // Update shared state - this will notify other tabs
        _sharedUIState.ApplyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true;
        await RefreshSheetSelectionDisplay();
    }

    private async void OnSharedApplyToCurrentSheetOnlyChanged(bool applyToCurrentSheetOnly)
    {
        // Update checkbox when shared state changes from other tabs
        if (ApplyToCurrentSheetCheckBox.IsChecked != applyToCurrentSheetOnly)
        {
            ApplyToCurrentSheetCheckBox.IsChecked = applyToCurrentSheetOnly;
            await RefreshSheetSelectionDisplay();
        }
    }

    private async Task RefreshSheetSelectionDisplay()
    {
        try
        {
            var config = await LoadProjectConfigurationAsync();
            if (config != null)
            {
                RefreshDisplay(BuildInfoMessages(config));
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error refreshing sheet selection display: {ex.Message}", ex);
        }
    }


    private void UpdateStatus(string message)
    {
        RefreshDisplay(new List<string> { message });
    }

    private async void RefreshDisplay(List<string>? statusMessages = null)
    {
        try
        {
            var config = GetSharedConfigurationFromSibling();
            var selectedSheets = await GetSelectedSheetsForDisplay(config);
            
            var readout = BuildStandardReadout(
                config?.ProjectName,
                statusMessages,
                selectedSheets
            );
            
            if (TitleBlockTextBlock != null)
            {
                TitleBlockTextBlock.Text = readout;
                TitleBlockTextBlock.UpdateLayout();
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error refreshing title block display: {ex.Message}", ex);
            // Fallback to a simple display
            if (TitleBlockTextBlock != null)
            {
                TitleBlockTextBlock.Text = "Active Project: No project selected\n" +
                                          "────────────────────────────────────────────────────────\n\n" +
                                          "Ready for title block operations.\n\n" +
                                          "────────────────────────────────────────────────────────\n" +
                                          "Selected Sheets: 0\n" +
                                          "────────────────────────────────────────────────────────\n" +
                                          "No sheets selected for processing";
                TitleBlockTextBlock.UpdateLayout();
            }
        }
    }

    private async Task<List<SheetInfo>> GetSelectedSheetsForDisplay(ProjectConfiguration? config)
    {
        if (config == null) return new List<SheetInfo>();
        
        try
        {
            return await GetSelectedSheetsAsync(config);
        }
        catch
        {
            return new List<SheetInfo>();
        }
    }

    private List<string> BuildInfoMessages(ProjectConfiguration? config)
    {
        var messages = new List<string>();
        
        // Show sheet selection mode
        if (ApplyToCurrentSheetCheckBox.IsChecked == true)
        {
            var currentLayoutName = GetCurrentLayoutName();
            messages.Add("Mode: Apply to current sheet only");
            if (!string.IsNullOrEmpty(currentLayoutName))
            {
                messages.Add($"Current Sheet: {currentLayoutName}");
            }
            else
            {
                messages.Add("WARNING: Could not determine current sheet");
            }
        }
        else
        {
            messages.Add("Mode: Apply to selected sheets");
        }
        
        return messages;
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

    private async void TitleBlockControl_Loaded(object sender, RoutedEventArgs e)
    {
        // Only load initial display after the control is fully loaded and palette is created
        // This prevents AutoCAD API access during palette initialization
        RefreshDisplay(); // Initialize standard display format
        await LoadInitialDisplayAsync();
    }
}