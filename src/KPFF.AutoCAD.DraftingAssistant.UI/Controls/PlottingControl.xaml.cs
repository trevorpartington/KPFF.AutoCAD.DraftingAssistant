using System.IO;
using System.Linq;
using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.UI.Utilities;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class PlottingControl : BaseUserControl
{
    private readonly IPlottingService _plottingService;
    private readonly IProjectConfigurationService _configService;
    private readonly SharedUIStateService _sharedUIState;
    
    // Lazy initialization to avoid AutoCAD API access during UI construction
    private MultiDrawingConstructionNotesService? _multiDrawingConstructionNotesService;
    private MultiDrawingTitleBlockService? _multiDrawingTitleBlockService;
    private IConstructionNotesService? _constructionNotesService;
    
    public PlottingControl() : this(null, null, null, null)
    {
    }

    public PlottingControl(
        ILogger? logger,
        INotificationService? notificationService,
        IPlottingService? plottingService,
        IProjectConfigurationService? configService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
        
        // Use constructor injection or fall back to service creation
        _plottingService = plottingService ?? GetPlottingService();
        _configService = configService ?? GetConfigurationService();
        
        // Initialize shared UI state - services will be lazily initialized when needed
        _sharedUIState = SharedUIStateService.Instance;
        
        // Subscribe to shared state changes
        _sharedUIState.OnApplyToCurrentSheetOnlyChanged += OnSharedApplyToCurrentSheetOnlyChanged;
        
        // Initialize display when control is loaded
        this.Loaded += PlottingControl_Loaded;
        
        // Wire up event handlers and initialize state after control is fully loaded
        this.Loaded += (s, e) => {
            // Initialize checkbox with shared state
            ApplyToCurrentSheetCheckBox.IsChecked = _sharedUIState.ApplyToCurrentSheetOnly;
            
            // Wire up event handlers
            UpdateConstructionNotesCheckBox.Checked += UpdateConstructionNotesCheckBox_CheckedChanged;
            UpdateConstructionNotesCheckBox.Unchecked += UpdateConstructionNotesCheckBox_CheckedChanged;
            UpdateTitleBlocksCheckBox.Checked += UpdateTitleBlocksCheckBox_CheckedChanged;
            UpdateTitleBlocksCheckBox.Unchecked += UpdateTitleBlocksCheckBox_CheckedChanged;
            ApplyToCurrentSheetCheckBox.Checked += ApplyToCurrentSheetCheckBox_CheckedChanged;
            ApplyToCurrentSheetCheckBox.Unchecked += ApplyToCurrentSheetCheckBox_CheckedChanged;
        };
    }

    private static IConstructionNotesService GetConstructionNotesService()
    {
        var debugLogger = new DebugLogger();
        var excelReader = new ExcelReaderService(debugLogger);
        var drawingOps = new DrawingOperations(debugLogger);
        return new ConstructionNotesService(debugLogger, excelReader, drawingOps);
    }

    private static IPlottingService GetPlottingService()
    {
        var debugLogger = new DebugLogger();
        var excelReader = new ExcelReaderService(debugLogger);
        var drawingOps = new DrawingOperations(debugLogger);
        var constructionNotesService = new ConstructionNotesService(debugLogger, excelReader, drawingOps);
        var multiDrawingConstructionNotesService = GetMultiDrawingConstructionNotesService();
        var multiDrawingTitleBlockService = GetMultiDrawingTitleBlockService();
        
        // Try to get DrawingAvailabilityService from the composition root
        IDrawingAvailabilityService? drawingAvailabilityService = null;
        try
        {
            // Access the composition root using reflection to avoid compile-time dependency
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
                        }
                    }
                }
            }
        }
        catch
        {
            // Ignore - service will be null and PlottingService will handle gracefully
        }
        
        return new PlottingService(debugLogger, constructionNotesService, drawingOps, excelReader, 
            multiDrawingConstructionNotesService, multiDrawingTitleBlockService, plotManager: null, drawingAvailabilityService);
    }
    
    private static IProjectConfigurationService GetConfigurationService()
    {
        var logger = new DebugLogger();
        return new ProjectConfigurationService(logger);
    }
    
    private static MultiDrawingConstructionNotesService GetMultiDrawingConstructionNotesService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var debugLogger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(debugLogger);
        var backupCleanupService = new BackupCleanupService(debugLogger);
        var multileaderAnalyzer = new MultileaderAnalyzer(debugLogger);
        var blockAnalyzer = new BlockAnalyzer(debugLogger);
        var externalDrawingManager = new ExternalDrawingManager(debugLogger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
        var excelReaderService = new ExcelReaderService(debugLogger);
        var constructionNotesService = new ConstructionNotesService(debugLogger, excelReaderService, new DrawingOperations(debugLogger));
        
        return new MultiDrawingConstructionNotesService(
            debugLogger,
            drawingAccessService,
            externalDrawingManager,
            constructionNotesService,
            excelReaderService);
    }
    
    private static MultiDrawingTitleBlockService GetMultiDrawingTitleBlockService()
    {
        // CRASH FIX: Never access ApplicationServices during UI initialization
        var debugLogger = new DebugLogger();
        var drawingAccessService = new DrawingAccessService(debugLogger);
        var backupCleanupService = new BackupCleanupService(debugLogger);
        var multileaderAnalyzer = new MultileaderAnalyzer(debugLogger);
        var blockAnalyzer = new BlockAnalyzer(debugLogger);
        var externalDrawingManager = new ExternalDrawingManager(debugLogger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
        var excelReaderService = new ExcelReaderService(debugLogger);
        var titleBlockService = new TitleBlockService(debugLogger, excelReaderService, new DrawingOperations(debugLogger));
        
        return new MultiDrawingTitleBlockService(
            debugLogger,
            drawingAccessService,
            externalDrawingManager,
            titleBlockService,
            excelReaderService);
    }
    
    /// <summary>
    /// Lazy initialization of construction notes service
    /// </summary>
    private IConstructionNotesService GetConstructionNotesServiceLazy()
    {
        return _constructionNotesService ??= GetConstructionNotesService();
    }
    
    /// <summary>
    /// Lazy initialization of multi-drawing construction notes service
    /// </summary>
    private MultiDrawingConstructionNotesService GetMultiDrawingConstructionNotesServiceLazy()
    {
        return _multiDrawingConstructionNotesService ??= GetMultiDrawingConstructionNotesService();
    }
    
    /// <summary>
    /// Lazy initialization of multi-drawing title block service
    /// </summary>
    private MultiDrawingTitleBlockService GetMultiDrawingTitleBlockServiceLazy()
    {
        return _multiDrawingTitleBlockService ??= GetMultiDrawingTitleBlockService();
    }
    
    
    private async void PlotButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            UpdatePlottingDisplay("🔄 Starting plot operation...");
            
            // Load project configuration
            var config = await LoadProjectConfigurationAsync();
            if (config == null)
            {
                var message = "No project configuration loaded. Please configure a project first.";
                UpdatePlottingDisplay($"❌ {message}");
                NotificationService?.ShowError("Plot", message);
                return;
            }
            
            // Get selected sheets
            var selectedSheets = await GetSelectedSheetsAsync(config);
            if (selectedSheets.Count == 0)
            {
                UpdatePlottingDisplay("No sheets selected. Please select sheets in the Configure tab.");
                return;
            }

            UpdatePlottingDisplay($"📋 Processing {selectedSheets.Count} sheets...");
            Logger?.LogInformation($"Starting plot operation for {selectedSheets.Count} sheets");

            // Get checkbox states
            var updateConstructionNotes = UpdateConstructionNotesCheckBox.IsChecked == true;
            var updateTitleBlocks = UpdateTitleBlocksCheckBox.IsChecked == true;
            var applyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true;

            Logger?.LogDebug($"Plot settings - Construction Notes: {updateConstructionNotes}, Title Blocks: {updateTitleBlocks}, Current Sheet Only: {applyToCurrentSheetOnly}");

            // Perform pre-plot updates if requested
            if (updateConstructionNotes || updateTitleBlocks)
            {
                UpdatePlottingDisplay("🔄 Performing pre-plot updates...");
                await PerformPrePlotUpdatesAsync(selectedSheets, config, updateConstructionNotes, updateTitleBlocks);
            }

            // Build PlotJobSettings
            var plotSettings = new PlotJobSettings
            {
                UpdateConstructionNotes = updateConstructionNotes,
                UpdateTitleBlocks = updateTitleBlocks,
                ApplyToCurrentSheetOnly = applyToCurrentSheetOnly,
                IsAutoNotesMode = _sharedUIState.IsAutoNotesMode,
                OutputDirectory = config.Plotting?.OutputDirectory
            };

            // Convert SheetInfo to sheet names for plotting
            var sheetNames = selectedSheets.Select(s => s.SheetName).ToList();

            UpdatePlottingDisplay("📄 Starting plot operation...");

            // Create production-ready services with AutoCAD API access for plotting
            Logger?.LogInformation("Creating production plotting services with AutoCAD API access");
            var autocadLogger = new AutoCADLogger();
            var excelReader = new ExcelReaderService(autocadLogger);
            var drawingOps = new DrawingOperations(autocadLogger);
            
            // Get multi-drawing services for proper cross-drawing operations
            var multiDrawingConstructionNotesService = GetMultiDrawingConstructionNotesServiceLazy();
            var multiDrawingTitleBlockService = GetMultiDrawingTitleBlockServiceLazy();
            autocadLogger.LogInformation("Using multi-drawing services for cross-drawing operations");
            
            // Create PlotManager using multiple fallback approaches
            IPlotManager? plotManager = null;
            
            // Approach 1: Try reflection with current assembly location
            try
            {
                var currentAssemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var currentDir = Path.GetDirectoryName(currentAssemblyLocation);
                var pluginDllPath = Path.Combine(currentDir!, "KPFF.AutoCAD.DraftingAssistant.Plugin.dll");
                
                autocadLogger.LogDebug($"Looking for Plugin DLL at: {pluginDllPath}");
                
                if (File.Exists(pluginDllPath))
                {
                    var pluginAssembly = System.Reflection.Assembly.LoadFrom(pluginDllPath);
                    var plotManagerType = pluginAssembly.GetType("KPFF.AutoCAD.DraftingAssistant.Plugin.Services.PlotManager");
                    if (plotManagerType != null)
                    {
                        plotManager = (IPlotManager?)Activator.CreateInstance(plotManagerType, autocadLogger);
                        autocadLogger.LogInformation("Successfully created PlotManager via reflection");
                    }
                }
            }
            catch (Exception ex)
            {
                autocadLogger.LogDebug($"Reflection approach 1 failed: {ex.Message}");
            }
            
            // Approach 2: Try reflection by looking in loaded assemblies
            if (plotManager == null)
            {
                try
                {
                    var pluginAssembly = System.AppDomain.CurrentDomain.GetAssemblies()
                        .FirstOrDefault(a => a.GetName().Name == "KPFF.AutoCAD.DraftingAssistant.Plugin");
                    
                    if (pluginAssembly != null)
                    {
                        var plotManagerType = pluginAssembly.GetType("KPFF.AutoCAD.DraftingAssistant.Plugin.Services.PlotManager");
                        if (plotManagerType != null)
                        {
                            plotManager = (IPlotManager?)Activator.CreateInstance(plotManagerType, autocadLogger);
                            autocadLogger.LogInformation("Successfully created PlotManager from loaded assembly");
                        }
                    }
                }
                catch (Exception ex)
                {
                    autocadLogger.LogDebug($"Loaded assembly approach failed: {ex.Message}");
                }
            }
            
            if (plotManager == null)
            {
                autocadLogger.LogWarning("Could not create PlotManager - plotting will use fallback mode");
            }
            
            // Execute pre-plot updates using multi-drawing services if checkboxes are checked
            if (plotSettings.UpdateConstructionNotes)
            {
                await PerformMultiDrawingConstructionNotesUpdates(selectedSheets, config, plotSettings.IsAutoNotesMode, multiDrawingConstructionNotesService, autocadLogger);
            }
            
            if (plotSettings.UpdateTitleBlocks)
            {
                await PerformMultiDrawingTitleBlockUpdates(selectedSheets, config, multiDrawingTitleBlockService, autocadLogger);
            }

            // Try to get DrawingAvailabilityService from the composition root for runtime plotting
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
                                autocadLogger.LogDebug($"DrawingAvailabilityService resolved: {drawingAvailabilityService != null}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                autocadLogger.LogDebug($"Failed to resolve DrawingAvailabilityService: {ex.Message}");
            }
            
            // Create ConstructionNotesService with DrawingAvailabilityService for pre-plot operations
            var basicConstructionNotesService = new ConstructionNotesService(autocadLogger, excelReader, drawingOps, drawingAvailabilityService);
            
            var productionPlottingService = new PlottingService(autocadLogger, basicConstructionNotesService, drawingOps, excelReader, 
                multiDrawingConstructionNotesService, multiDrawingTitleBlockService, plotManager, drawingAvailabilityService);
            
            // Create plot settings without pre-plot updates since we already did them
            var plotOnlySettings = new PlotJobSettings
            {
                IsAutoNotesMode = plotSettings.IsAutoNotesMode,
                UpdateConstructionNotes = false, // Already done with multi-drawing services
                UpdateTitleBlocks = false        // Already done with multi-drawing services
            };

            // Execute plotting with production services
            var progress = new Progress<PlotProgress>(OnPlotProgress);
            var plotResult = await productionPlottingService.PlotSheetsAsync(sheetNames, config, plotOnlySettings, progress);

            // Generate standardized completion message using helper
            var failedSheets = plotResult.FailedSheets.ToDictionary(f => f.SheetName, f => f.ErrorMessage);
            var completionMessage = MessageFormatHelper.CreateCompletionMessage(
                "Plotting",
                plotResult.SuccessfulSheets.ToList(), 
                failedSheets);

            UpdatePlottingDisplay(completionMessage);
            
            // No popup notification needed - user can see results in the display
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Plot operation failed: {ex.Message}", ex);
            NotificationService?.ShowError("Plot", $"Plot operation failed: {ex.Message}");
            UpdatePlottingDisplay($"❌ Plot operation failed: {ex.Message}");
        }
    }
    
    
    
    /// <summary>
    /// Performs pre-plot updates for construction notes and/or title blocks
    /// </summary>
    private async Task PerformPrePlotUpdatesAsync(List<SheetInfo> selectedSheets, ProjectConfiguration config, bool updateConstructionNotes, bool updateTitleBlocks)
    {
        try
        {
            if (updateConstructionNotes)
            {
                UpdatePlottingDisplay("📝 Updating construction notes...");
                Logger?.LogInformation("Performing construction notes updates before plotting");

                // Determine construction notes mode from shared state
                var isAutoMode = _sharedUIState.IsAutoNotesMode;
                Logger?.LogDebug($"Construction notes mode: {(isAutoMode ? "Auto Notes" : "Excel Notes")}");

                if (isAutoMode)
                {
                    // Use Auto Notes mode
                    var sheetToNotes = new Dictionary<string, List<int>>();
                    
                    // For auto notes, we need to analyze each sheet for bubble multileaders
                    foreach (var sheet in selectedSheets)
                    {
                        try
                        {
                            // This would normally scan viewports for multileaders
                            // For now, use the existing construction notes service
                            var noteNumbers = await GetConstructionNotesServiceLazy().GetAutoNotesForSheetAsync(sheet.SheetName, config);
                            sheetToNotes[sheet.SheetName] = noteNumbers;
                        }
                        catch (Exception ex)
                        {
                            Logger?.LogWarning($"Failed to get auto notes for sheet {sheet.SheetName}: {ex.Message}");
                            sheetToNotes[sheet.SheetName] = new List<int>();
                        }
                    }

                    // Update construction notes using multi-drawing service
                    await GetMultiDrawingConstructionNotesServiceLazy().UpdateConstructionNotesAcrossDrawingsAsync(
                        sheetToNotes, config, selectedSheets);
                }
                else
                {
                    // Use Excel Notes mode  
                    var selectedSheetNames = selectedSheets.Select(s => s.SheetName).ToList();
                    
                    // Read Excel notes from EXCEL_NOTES table and get actual note numbers
                    var sheetToNotes = new Dictionary<string, List<int>>();
                    foreach (var sheet in selectedSheets)
                    {
                        try
                        {
                            var noteNumbers = await GetConstructionNotesServiceLazy().GetExcelNotesForSheetAsync(sheet.SheetName, config);
                            sheetToNotes[sheet.SheetName] = noteNumbers;
                        }
                        catch (Exception ex)
                        {
                            Logger?.LogWarning($"Failed to get Excel notes for sheet {sheet.SheetName}: {ex.Message}");
                            sheetToNotes[sheet.SheetName] = new List<int>();
                        }
                    }

                    await GetMultiDrawingConstructionNotesServiceLazy().UpdateConstructionNotesAcrossDrawingsAsync(
                        sheetToNotes, config, selectedSheets);
                }

                Logger?.LogInformation("Construction notes updates completed");
            }

            if (updateTitleBlocks)
            {
                UpdatePlottingDisplay("📋 Updating title blocks...");
                Logger?.LogInformation("Performing title block updates before plotting");

                // Get the sheet names for title block updates
                var selectedSheetNames = selectedSheets.Select(s => s.SheetName).ToList();

                // Update title blocks using multi-drawing service
                await GetMultiDrawingTitleBlockServiceLazy().UpdateTitleBlocksAcrossDrawingsAsync(
                    selectedSheetNames, config, selectedSheets);

                Logger?.LogInformation("Title block updates completed");
            }
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Pre-plot updates failed: {ex.Message}", ex);
            UpdatePlottingDisplay($"⚠️ Pre-plot updates had errors: {ex.Message}");
            // Continue with plotting even if pre-plot updates fail
        }
    }

    /// <summary>
    /// Handles progress updates from the plotting service
    /// </summary>
    private void OnPlotProgress(PlotProgress progress)
    {
        // Don't show batch plotting progress - AutoCAD provides its own progress window
        // This prevents the "stuck on 1/Y" issue while AutoCAD shows real progress
    }

    private void UpdatePlottingDisplay(string message)
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
                statusMessages ?? BuildInfoMessages(config),
                selectedSheets,
                GetOutputDirectoryInfo(config)
            );
            
            Dispatcher.Invoke(() =>
            {
                PlottingTextBlock.Text = readout;
            });
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Error refreshing plotting display: {ex.Message}", ex);
            // Fallback to a simple display
            Dispatcher.Invoke(() =>
            {
                PlottingTextBlock.Text = "Active Project: No project selected\n" +
                                        "────────────────────────────────────────────────────────\n\n" +
                                        "Ready for plotting operations.\n\n" +
                                        "────────────────────────────────────────────────────────\n" +
                                        "Selected Sheets: 0\n" +
                                        "────────────────────────────────────────────────────────\n" +
                                        "No sheets selected for processing";
            });
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
        
        // Only show sheet selection mode
        if (ApplyToCurrentSheetCheckBox?.IsChecked == true)
        {
            messages.Add("Mode: Apply to current sheet only");
        }
        else if (ApplyToCurrentSheetCheckBox != null)
        {
            messages.Add("Mode: Apply to selected sheets");
        }
        
        return messages;
    }
    
    private string? GetOutputDirectoryInfo(ProjectConfiguration? config)
    {
        if (!string.IsNullOrEmpty(config?.Plotting?.OutputDirectory))
        {
            return $"Output Directory: {config.Plotting.OutputDirectory}";
        }
        return null;
    }
    
    private void ApplyToCurrentSheetCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        // Update shared state
        _sharedUIState.ApplyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true;
        RefreshDisplay();
    }
    
    private void UpdateConstructionNotesCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        RefreshDisplay();
    }
    
    private void UpdateTitleBlocksCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        RefreshDisplay();
    }
    
    private void OnSharedApplyToCurrentSheetOnlyChanged(bool applyToCurrentSheetOnly)
    {
        // Update checkbox if it doesn't match shared state
        if (ApplyToCurrentSheetCheckBox.IsChecked != applyToCurrentSheetOnly)
        {
            ApplyToCurrentSheetCheckBox.IsChecked = applyToCurrentSheetOnly;
        }
        RefreshDisplay();
    }

    /// <summary>
    /// Gets the selected sheets from the project configuration, applying the current sheet filter if enabled
    /// </summary>
    private async Task<List<SheetInfo>> GetSelectedSheetsAsync(ProjectConfiguration config)
    {
        try
        {
            if (_sharedUIState.ApplyToCurrentSheetOnly)
            {
                Logger?.LogDebug("ApplyToCurrentSheetOnly is enabled, filtering to current sheet only");
                
                try
                {
                    // Get current layout name from AutoCAD
                    var layoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                    var currentLayoutName = layoutManager?.CurrentLayout;
                    
                    if (!string.IsNullOrEmpty(currentLayoutName))
                    {
                        Logger?.LogDebug($"Current layout name: {currentLayoutName}");
                        
                        // Get all sheets to find the one that matches the current layout
                        List<SheetInfo> availableSheets;
                        if (config.SelectedSheets.Count > 0)
                        {
                            availableSheets = config.SelectedSheets;
                        }
                        else
                        {
                            // No sheets selected, cannot apply to current sheet
                            Logger?.LogWarning("No sheets selected, cannot apply to current sheet");
                            return new List<SheetInfo>();
                        }
                        
                        // Find the sheet that matches the current layout
                        var currentSheet = availableSheets.FirstOrDefault(s => s.SheetName.Equals(currentLayoutName, StringComparison.OrdinalIgnoreCase));
                        
                        if (currentSheet != null)
                        {
                            Logger?.LogDebug($"Found current sheet: {currentSheet.SheetName}");
                            return new List<SheetInfo> { currentSheet };
                        }
                        else
                        {
                            Logger?.LogWarning($"Current layout '{currentLayoutName}' not found in selected sheets");
                        }
                    }
                    else
                    {
                        Logger?.LogWarning("Could not determine current layout name");
                    }
                }
                catch (Exception ex)
                {
                    Logger?.LogError($"Error getting current layout: {ex.Message}", ex);
                }
                
                // If we get here, we couldn't get the current layout, so return empty
                return new List<SheetInfo>();
            }
            
            // Use selected sheets from ProjectConfiguration if available
            if (config.SelectedSheets.Count > 0)
            {
                Logger?.LogDebug($"Using {config.SelectedSheets.Count} selected sheets from configuration");
                return config.SelectedSheets;
            }
            
            // Return empty list if no sheets are selected
            Logger?.LogDebug("No sheets selected, returning empty list");
            return new List<SheetInfo>();
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Error getting selected sheets: {ex.Message}", ex);
            return new List<SheetInfo>();
        }
    }

    /// <summary>
    /// Loads the current project configuration from the shared configuration manager
    /// </summary>
    private async Task<ProjectConfiguration?> LoadProjectConfigurationAsync()
    {
        try
        {
            // First try to get shared configuration from ConfigurationControl sibling
            var sharedConfig = GetSharedConfigurationFromSibling();
            if (sharedConfig != null)
            {
                Logger?.LogDebug($"Using shared configuration with {sharedConfig.SelectedSheets.Count} selected sheets: {sharedConfig.ProjectName}");
                return sharedConfig;
            }
            
            // Fall back to loading from the default test location
            var testConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\DBRT_Config.json";
            
            if (File.Exists(testConfigPath))
            {
                var config = await _configService.LoadConfigurationAsync(testConfigPath);
                Logger?.LogDebug($"Loaded test configuration from file: {config?.ProjectName}");
                return config;
            }
            
            Logger?.LogWarning("No project configuration found");
            return null;
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Error loading project configuration: {ex.Message}", ex);
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
                    Logger?.LogDebug("Found ConfigurationControl sibling, getting shared configuration");
                    return configControl.CurrentConfiguration;
                }
                else
                {
                    Logger?.LogDebug("ConfigurationControl sibling not found by name");
                }
            }
            else
            {
                Logger?.LogDebug("DraftingAssistantControl parent not found");
            }
            
            return null;
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Error getting shared configuration from sibling: {ex.Message}", ex);
            return null;
        }
    }

    private async Task PerformMultiDrawingConstructionNotesUpdates(List<SheetInfo> selectedSheets, ProjectConfiguration config, bool isAutoNotesMode, MultiDrawingConstructionNotesService multiDrawingService, IApplicationLogger logger)
    {
        try
        {
            logger.LogInformation($"Performing multi-drawing construction notes updates for {selectedSheets.Count} sheets using {(isAutoNotesMode ? "Auto Notes" : "Excel Notes")} mode");

            // Build sheet-to-notes mapping using multi-drawing approach
            var sheetToNotes = new Dictionary<string, List<int>>();

            foreach (var sheet in selectedSheets)
            {
                try
                {
                    List<int> noteNumbers;
                    if (isAutoNotesMode)
                    {
                        // Perform proper Auto Notes detection based on drawing state
                        noteNumbers = await GetAutoNotesForSheet(sheet, config, logger);
                        logger.LogDebug($"Getting auto notes for sheet {sheet.SheetName} from {sheet.DWGFileName}: {string.Join(", ", noteNumbers)}");
                    }
                    else
                    {
                        // Get Excel notes - this works regardless of drawing state
                        var excelReader = new ExcelReaderService(logger);
                        var mappings = await excelReader.ReadExcelNotesAsync(config.ProjectIndexFilePath, config);
                        var sheetMapping = mappings.FirstOrDefault(m => m.SheetName.Equals(sheet.SheetName, StringComparison.OrdinalIgnoreCase));
                        noteNumbers = sheetMapping?.NoteNumbers ?? new List<int>();
                        logger.LogDebug($"Getting Excel notes for sheet {sheet.SheetName}: {string.Join(", ", noteNumbers)}");
                    }
                    
                    sheetToNotes[sheet.SheetName] = noteNumbers;
                }
                catch (Exception ex)
                {
                    logger.LogError($"Failed to get notes for sheet {sheet.SheetName}: {ex.Message}");
                    sheetToNotes[sheet.SheetName] = new List<int>();
                }
            }

            // Execute multi-drawing update
            var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(sheetToNotes, config, selectedSheets);
            
            logger.LogInformation($"Multi-drawing construction notes update completed. Success: {result.Successes.Count}, Failed: {result.Failures.Count}");
        }
        catch (Exception ex)
        {
            logger.LogError($"Failed to perform multi-drawing construction notes updates: {ex.Message}");
        }
    }

    private async Task PerformMultiDrawingTitleBlockUpdates(List<SheetInfo> selectedSheets, ProjectConfiguration config, MultiDrawingTitleBlockService multiDrawingService, IApplicationLogger logger)
    {
        try
        {
            logger.LogInformation($"Performing multi-drawing title block updates for {selectedSheets.Count} sheets");

            // Get sheet names for title block update
            var selectedSheetNames = selectedSheets.Select(s => s.SheetName).ToList();

            // Execute multi-drawing title block update
            var result = await multiDrawingService.UpdateTitleBlocksAcrossDrawingsAsync(selectedSheetNames, config, selectedSheets);
            
            logger.LogInformation($"Multi-drawing title block update completed. Success: {result.Successes.Count}, Failed: {result.Failures.Count}");
        }
        catch (Exception ex)
        {
            logger.LogError($"Failed to perform multi-drawing title block updates: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets Auto Notes for a sheet based on drawing state - uses appropriate method for active, inactive, or closed drawings
    /// </summary>
    private async Task<List<int>> GetAutoNotesForSheet(SheetInfo sheet, ProjectConfiguration config, IApplicationLogger logger)
    {
        try
        {
            // Get drawing file path and state
            var drawingAccessService = new DrawingAccessService(logger);
            var dwgPath = drawingAccessService.GetDrawingFilePath(sheet.SheetName, config, new List<SheetInfo> { sheet });
            
            if (string.IsNullOrEmpty(dwgPath))
            {
                logger.LogWarning($"Could not resolve DWG file path for sheet {sheet.SheetName}");
                return new List<int>();
            }

            var drawingState = drawingAccessService.GetDrawingState(dwgPath);
            logger.LogDebug($"Sheet {sheet.SheetName} drawing state: {drawingState}");

            switch (drawingState)
            {
                case DrawingState.Active:
                case DrawingState.Inactive:
                    // For active/inactive drawings, use AutoNotesService (works with current AutoCAD session)
                    var autoNotesService = new AutoNotesService(logger);
                    
                    // For inactive drawings, try to make active first (better detection accuracy)
                    if (drawingState == DrawingState.Inactive)
                    {
                        if (drawingAccessService.TryMakeDrawingActive(dwgPath))
                        {
                            logger.LogDebug($"Successfully made drawing active for better Auto Notes detection: {sheet.SheetName}");
                        }
                        else
                        {
                            logger.LogDebug($"Could not make drawing active, proceeding with inactive detection: {sheet.SheetName}");
                        }
                    }
                    
                    return await autoNotesService.GetAutoNotesForSheetAsync(sheet.SheetName, config);

                case DrawingState.Closed:
                    // For closed drawings, use ExternalDrawingManager's specialized method
                    var externalDrawingManager = GetExternalDrawingManager(logger);
                    var multileaderStyles = config.ConstructionNotes?.MultileaderStyleNames ?? new List<string>();
                    var blockConfigurations = config.ConstructionNotes?.NoteBlocks ?? new List<NoteBlockConfiguration>();
                    
                    return externalDrawingManager.GetAutoNotesForClosedDrawing(dwgPath, sheet.SheetName, multileaderStyles, blockConfigurations);

                case DrawingState.NotFound:
                    logger.LogError($"Drawing file not found for sheet {sheet.SheetName}: {dwgPath}");
                    return new List<int>();

                case DrawingState.Error:
                    logger.LogError($"Error accessing drawing for sheet {sheet.SheetName}: {dwgPath}");
                    return new List<int>();

                default:
                    logger.LogWarning($"Unknown drawing state {drawingState} for sheet {sheet.SheetName}");
                    return new List<int>();
            }
        }
        catch (Exception ex)
        {
            logger.LogError($"Error getting Auto Notes for sheet {sheet.SheetName}: {ex.Message}", ex);
            return new List<int>();
        }
    }

    /// <summary>
    /// Creates ExternalDrawingManager with proper dependencies
    /// </summary>
    private ExternalDrawingManager GetExternalDrawingManager(IApplicationLogger logger)
    {
        var backupCleanupService = new BackupCleanupService(logger);
        var multileaderAnalyzer = new MultileaderAnalyzer(logger);
        var blockAnalyzer = new BlockAnalyzer(logger);
        
        return new ExternalDrawingManager(logger, backupCleanupService, multileaderAnalyzer, blockAnalyzer);
    }
    
    private void PlottingControl_Loaded(object sender, RoutedEventArgs e)
    {
        // Initialize standard display format when control is loaded
        RefreshDisplay();
    }
}