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
    private readonly IBatchOperationService _batchOperationService;
    private readonly SharedUIStateService _sharedUIState;

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
        
        // Initialize batch operation service
        _batchOperationService = GetBatchOperationService();
        
        // Initialize shared UI state
        _sharedUIState = SharedUIStateService.Instance;
        
        // Subscribe to shared state changes
        _sharedUIState.OnApplyToCurrentSheetOnlyChanged += OnSharedApplyToCurrentSheetOnlyChanged;
        _sharedUIState.OnConstructionNotesModeChanged += OnSharedConstructionNotesModeChanged;
        
        // Initialize checkbox with shared state
        ApplyToCurrentSheetCheckBox.IsChecked = _sharedUIState.ApplyToCurrentSheetOnly;
        
        // Initialize radio buttons with shared state
        AutoNotesRadioButton.IsChecked = _sharedUIState.IsAutoNotesMode;
        ExcelNotesRadioButton.IsChecked = !_sharedUIState.IsAutoNotesMode;
        
        // Load initial display information when the control is fully loaded
        this.Loaded += ConstructionNoteControl_Loaded;
        
        // Display will be initialized when Loaded event fires
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
        
        // Create multi-drawing services
        var multiDrawingConstructionNotesService = new MultiDrawingConstructionNotesService(
            logger, drawingAccessService, externalDrawingManager, constructionNotesService, excelReaderService);
        var titleBlockService = new TitleBlockService(logger, excelReaderService, new DrawingOperations(logger));
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

    private void InsertBlocksButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Logger.LogInformation("Insert Blocks button clicked - executing KPFFINSERTBLOCKS command");
            
            // Execute the AutoCAD command to insert construction note blocks
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager?.MdiActiveDocument;
            if (doc != null)
            {
                // Send the command to AutoCAD for execution
                doc.SendStringToExecute("KPFFINSERTBLOCKS ", true, false, false);
                
                UpdateStatus("Insert Blocks command executed. Follow the prompts in AutoCAD to select insertion point.");
            }
            else
            {
                Logger.LogError("No active document found for Insert Blocks command");
                UpdateStatus("ERROR: No active drawing found. Please open a drawing before inserting blocks.");
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error executing Insert Blocks command: {ex.Message}", ex);
            UpdateStatus($"ERROR: Failed to execute Insert Blocks command - {ex.Message}");
        }
    }

    
    private async void ExecuteAutoNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Auto Notes update using BatchOperationService");
            
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

            UpdateStatus($"Processing {selectedSheets.Count} sheets in Auto Notes mode...\n" +
                        "Starting batch operation...\n");

            // Create batch operation settings
            var batchSettings = new BatchOperationSettings
            {
                UpdateConstructionNotes = true,
                IsAutoNotesMode = true,
                ApplyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true
            };

            // Use BatchOperationService for proper async handling
            var sheetNames = selectedSheets.Select(s => s.SheetName).ToList();
            var progress = new Progress<BatchOperationProgress>(p => 
            {
                UpdateStatus($"[{p.ProgressPercentage}%] {p.Phase}: {p.CurrentOperationDescription}\n" +
                           $"Processing: {p.CompletedSheets}/{p.TotalSheets} sheets");
            });

            var result = await _batchOperationService.UpdateConstructionNotesAsync(
                sheetNames, config, batchSettings.IsAutoNotesMode, batchSettings.ApplyToCurrentSheetOnly, progress);

            // Generate status report
            var statusText = GenerateBatchOperationStatusReport(result, "Auto Notes");
            UpdateStatus(statusText);
            
            Logger.LogInformation($"Auto Notes batch update completed via BatchOperationService. " +
                                $"Success: {result.Success}");
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
            Logger.LogInformation("Executing Excel Notes update using BatchOperationService");
            
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
                        "Starting batch operation...\n");

            // Create batch operation settings
            var batchSettings = new BatchOperationSettings
            {
                UpdateConstructionNotes = true,
                IsAutoNotesMode = false, // Excel Notes mode
                ApplyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true
            };

            // Use BatchOperationService for proper async handling
            var sheetNames = selectedSheets.Select(s => s.SheetName).ToList();
            var progress = new Progress<BatchOperationProgress>(p => 
            {
                UpdateStatus($"[{p.ProgressPercentage}%] {p.Phase}: {p.CurrentOperationDescription}\n" +
                           $"Processing: {p.CompletedSheets}/{p.TotalSheets} sheets");
            });

            var result = await _batchOperationService.UpdateConstructionNotesAsync(
                sheetNames, config, batchSettings.IsAutoNotesMode, batchSettings.ApplyToCurrentSheetOnly, progress);

            // Generate status report
            var statusText = GenerateBatchOperationStatusReport(result, "Excel Notes");
            UpdateStatus(statusText);
            
            Logger.LogInformation($"Excel Notes batch update completed via BatchOperationService. " +
                                $"Success: {result.Success}");
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
            Logger.LogWarning($"Failed to get shared configuration from sibling: {ex.Message}");
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
        RefreshDisplay(new List<string> { message });
        Logger.LogDebug($"Status updated: {message}");
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
            NotesTextBlock.Text = readout;
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error refreshing construction notes display: {ex.Message}", ex);
            // Fallback to a simple display
            NotesTextBlock.Text = "Active Project: No project selected\n" +
                                 "────────────────────────────────────────────────────────\n\n" +
                                 "Ready for construction notes operations.\n\n" +
                                 "────────────────────────────────────────────────────────\n" +
                                 "Selected Sheets: 0\n" +
                                 "────────────────────────────────────────────────────────\n" +
                                 "No sheets selected for processing";
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

    private async Task LoadInitialDisplayAsync()
    {
        await ExceptionHandler.TryExecuteAsync(async () =>
        {
            var config = await LoadProjectConfigurationAsync();
            RefreshDisplay(BuildInfoMessages(config));
        }, Logger, NotificationService, "LoadInitialDisplayAsync", showUserMessage: false);
    }

    private List<string> BuildInfoMessages(ProjectConfiguration? config)
    {
        var messages = new List<string>();
        
        // Show sheet selection mode
        if (ApplyToCurrentSheetCheckBox?.IsChecked == true)
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

    private void RadioButton_CheckedChanged(object sender, RoutedEventArgs e)
    {
        // Update shared state when radio buttons change
        // Null check needed because XAML can trigger this before constructor completes
        if (_sharedUIState != null)
        {
            var isAutoMode = AutoNotesRadioButton.IsChecked == true;
            _sharedUIState.IsAutoNotesMode = isAutoMode;
        }
    }

    private void OnSharedConstructionNotesModeChanged(bool isAutoNotesMode)
    {
        // Update radio buttons if they don't match shared state
        if (AutoNotesRadioButton.IsChecked != isAutoNotesMode)
        {
            AutoNotesRadioButton.IsChecked = isAutoNotesMode;
            ExcelNotesRadioButton.IsChecked = !isAutoNotesMode;
        }
    }

    private async void ConstructionNoteControl_Loaded(object sender, RoutedEventArgs e)
    {
        // Only load initial display after the control is fully loaded and palette is created
        // This prevents AutoCAD API access during palette initialization
        RefreshDisplay(); // Initialize standard display format
        await LoadInitialDisplayAsync();
    }
}