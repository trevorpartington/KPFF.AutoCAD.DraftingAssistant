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
        
        // Load initial display information
        _ = LoadInitialDisplayAsync();
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
        var backupCleanupService = new BackupCleanupService(logger);
        var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService);
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

    
    private async void ExecuteAutoNotesUpdate()
    {
        try
        {
            Logger.LogInformation("Executing Auto Notes update with multi-drawing support");
            
            // Use AutoCAD services when actually performing operations
            var autocadLogger = new AutoCADLogger();
            
            // Create production-ready multi-drawing service
            var drawingAccessService = new DrawingAccessService(autocadLogger);
            var backupCleanupService = new BackupCleanupService(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger, backupCleanupService);
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

            // Gather Auto Notes for selected sheets using smart drawing state handling
            var sheetToNotes = new Dictionary<string, List<int>>();
            
            foreach (var sheet in selectedSheets)
            {
                try
                {
                    var dwgPath = drawingAccessService.GetDrawingFilePath(sheet.SheetName, config, selectedSheets);
                    if (string.IsNullOrEmpty(dwgPath))
                    {
                        autocadLogger.LogWarning($"Could not find drawing file for sheet {sheet.SheetName}");
                        continue;
                    }

                    var drawingState = drawingAccessService.GetDrawingState(dwgPath);
                    List<int> noteNumbers;

                    if (drawingState == Core.Services.DrawingState.Closed)
                    {
                        // Use ExternalDrawingManager for closed drawings
                        autocadLogger.LogDebug($"Using external drawing analysis for closed sheet {sheet.SheetName}");
                        var styleNames = config.ConstructionNotes?.MultileaderStyleNames ?? new List<string> { "ML-STYLE-01" };
                        noteNumbers = externalDrawingManager.GetAutoNotesForClosedDrawing(dwgPath, sheet.SheetName, styleNames);
                    }
                    else
                    {
                        // For active/inactive drawings, use the standard approach
                        if (drawingState == Core.Services.DrawingState.Inactive)
                        {
                            // Make inactive drawing active temporarily
                            drawingAccessService.TryMakeDrawingActive(dwgPath);
                        }
                        
                        noteNumbers = await constructionNotesService.GetAutoNotesForSheetAsync(sheet.SheetName, config);
                    }

                    if (noteNumbers.Count > 0)
                    {
                        sheetToNotes[sheet.SheetName] = noteNumbers;
                        autocadLogger.LogDebug($"Auto Notes for {sheet.SheetName}: [{string.Join(", ", noteNumbers)}]");
                    }
                    else
                    {
                        autocadLogger.LogDebug($"No Auto Notes detected for {sheet.SheetName}");
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
            var backupCleanupService = new BackupCleanupService(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger, backupCleanupService);
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
                    sheetToNotes[sheet.SheetName] = noteNumbers;
                    if (noteNumbers.Count > 0)
                    {
                        autocadLogger.LogDebug($"Excel Notes for {sheet.SheetName}: [{string.Join(", ", noteNumbers)}]");
                    }
                    else
                    {
                        autocadLogger.LogDebug($"No Excel Notes for {sheet.SheetName}, will clear construction note blocks");
                    }
                }
                catch (Exception ex)
                {
                    autocadLogger.LogWarning($"Failed to get Excel Notes for sheet {sheet.SheetName}: {ex.Message}");
                }
            }

            var sheetsWithNotes = sheetToNotes.Values.Count(notes => notes.Count > 0);
            UpdateStatus($"Processing {sheetToNotes.Count} sheets in Excel Notes mode. " +
                        $"{sheetsWithNotes} sheets have notes, {sheetToNotes.Count - sheetsWithNotes} will be cleared...\n");

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

    private async Task LoadInitialDisplayAsync()
    {
        await ExceptionHandler.TryExecuteAsync(async () =>
        {
            var config = await LoadProjectConfigurationAsync();
            var selectedSheets = config != null ? await GetSelectedSheetsAsync(config) : new List<SheetInfo>();
            
            UpdateInfoDisplay(config, selectedSheets);
        }, Logger, NotificationService, "LoadInitialDisplayAsync", showUserMessage: false);
    }

    private void UpdateInfoDisplay(ProjectConfiguration? config, List<SheetInfo> selectedSheets)
    {
        var multileaderStyle = config?.ConstructionNotes?.MultileaderStyleName ?? "Not configured";
        
        var displayText = $"Construction Notes Information\n\n" +
                         $"Multileader Style(s): {multileaderStyle}\n" +
                         $"Selected Sheet(s):\n";
        
        if (selectedSheets.Count > 0)
        {
            foreach (var sheet in selectedSheets)
            {
                displayText += $"  • {sheet.SheetName} - {sheet.DrawingTitle}\n";
            }
        }
        else
        {
            displayText += "  No sheets selected\n";
        }
        
        UpdateStatus(displayText);
    }
}