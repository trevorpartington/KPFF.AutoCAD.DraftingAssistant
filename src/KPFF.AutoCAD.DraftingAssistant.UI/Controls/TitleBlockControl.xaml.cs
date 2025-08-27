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
            Logger.LogError($"Error getting shared configuration from sibling: {ex.Message}", ex);
            return null;
        }
    }

    private async Task<List<SheetInfo>> GetSelectedSheetsAsync(ProjectConfiguration config)
    {
        try
        {
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
                               "and update TB_ATT blocks across all selected drawings.");
                });
            });
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to load initial display: {ex.Message}", ex);
        }
    }

    private void UpdateStatus(string message)
    {
        if (TitleBlockTextBlock != null)
        {
            TitleBlockTextBlock.Text = message;
            TitleBlockTextBlock.UpdateLayout();
        }
    }
}