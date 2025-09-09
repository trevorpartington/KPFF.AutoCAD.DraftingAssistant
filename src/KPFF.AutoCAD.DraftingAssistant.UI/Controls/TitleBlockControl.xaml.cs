using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.UI.Utilities;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class TitleBlockControl : BaseUserControl
{
    private readonly ITitleBlockService _titleBlockService;
    private readonly IProjectConfigurationService _configService;
    private readonly MultiDrawingTitleBlockService _multiDrawingService;
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
        
        // Initialize multi-drawing service
        _multiDrawingService = GetMultiDrawingService();
        
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

    private static MultiDrawingTitleBlockService GetMultiDrawingService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var logger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(logger);
        var backupCleanupService = new BackupCleanupService(logger);
        var multileaderAnalyzer = new MultileaderAnalyzer(logger);
        var blockAnalyzer = new BlockAnalyzer(logger);
        var externalDrawingManager = new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
        var excelReaderService = new ExcelReaderService(logger);
        var titleBlockService = new TitleBlockService(logger, excelReaderService, new DrawingOperations(logger));
        
        return new MultiDrawingTitleBlockService(
            logger,
            drawingAccessService,
            externalDrawingManager,
            titleBlockService,
            excelReaderService,
            drawingAvailabilityService: null);
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
            var multileaderAnalyzer = new MultileaderAnalyzer(autocadLogger);
            var blockAnalyzer = new BlockAnalyzer(autocadLogger);
            var externalDrawingManager = new ExternalDrawingManager(autocadLogger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
            var excelReaderService = new ExcelReaderService(autocadLogger);
            var titleBlockService = new TitleBlockService(autocadLogger, excelReaderService, new DrawingOperations(autocadLogger));
            
            // Try to get DrawingAvailabilityService from the composition root
            IDrawingAvailabilityService? drawingAvailabilityService = null;
            try
            {
                var pluginAssembly = System.AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name == "KPFF.AutoCAD.DraftingAssistant.Plugin");
                if (pluginAssembly != null)
                {
                    var extensionAppType = pluginAssembly.GetType("KPFF.AutoCAD.DraftingAssistant.Plugin.DraftingAssistantExtensionApplication");
                    if (extensionAppType != null)
                    {
                        var compositionRootProperty = extensionAppType.GetProperty("CompositionRoot", 
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                        var compositionRoot = compositionRootProperty?.GetValue(null);
                        if (compositionRoot != null)
                        {
                            var getServiceMethod = compositionRoot.GetType().GetMethod("GetOptionalService");
                            if (getServiceMethod != null)
                            {
                                var genericMethod = getServiceMethod.MakeGenericMethod(typeof(IDrawingAvailabilityService));
                                drawingAvailabilityService = (IDrawingAvailabilityService?)genericMethod.Invoke(compositionRoot, null);
                                autocadLogger.LogDebug($"DrawingAvailabilityService resolved for title blocks: {drawingAvailabilityService != null}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                autocadLogger.LogDebug($"Failed to resolve DrawingAvailabilityService for title blocks: {ex.Message}");
            }
            
            var multiDrawingService = new MultiDrawingTitleBlockService(
                autocadLogger,
                drawingAccessService,
                externalDrawingManager,
                titleBlockService,
                excelReaderService,
                drawingAvailabilityService);
            
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

            UpdateStatus("Analyzing drawing states and preparing batch operations...");
            autocadLogger.LogInformation($"Starting title blocks batch update for {selectedSheets.Count} sheets");

            // Get sheet names
            var selectedSheetNames = selectedSheets.Select(s => s.SheetName).ToList();

            // Add progress reporting during processing
            for (int i = 1; i <= selectedSheetNames.Count; i++)
            {
                var progressMessage = MessageFormatHelper.CreateProgressMessage("Title Block Attributes", i, selectedSheetNames.Count);
                UpdateStatus(progressMessage);
                // Short delay to allow UI to update
                await Task.Delay(100);
            }
            
            // Execute the multi-drawing update
            var result = await multiDrawingService.UpdateTitleBlocksAcrossDrawingsAsync(selectedSheetNames, config, selectedSheets);

            // Generate standardized completion message
            var successfulSheets = result.Successes.Select(s => s.SheetName).ToList();
            var failedSheets = result.Failures.ToDictionary(f => f.SheetName, f => f.ErrorMessage);
            var completionMessage = MessageFormatHelper.CreateCompletionMessage("Title Block Attributes", successfulSheets, failedSheets);
            UpdateStatus(completionMessage);
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