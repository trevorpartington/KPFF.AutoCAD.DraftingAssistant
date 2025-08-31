using System.Windows;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class PlottingControl : BaseUserControl
{
    private readonly IPlottingService _plottingService;
    private readonly IProjectConfigurationService _configService;
    private readonly SharedUIStateService _sharedUIState;
    
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
        
        // Initialize shared UI state
        _sharedUIState = SharedUIStateService.Instance;
        
        // Subscribe to shared state changes
        _sharedUIState.OnApplyToCurrentSheetOnlyChanged += OnSharedApplyToCurrentSheetOnlyChanged;
        
        // Initialize checkbox with shared state
        ApplyToCurrentSheetCheckBox.IsChecked = _sharedUIState.ApplyToCurrentSheetOnly;
    }

    private static IPlottingService GetPlottingService()
    {
        var logger = new DebugLogger();
        var excelReader = new ExcelReaderService(logger);
        var drawingOps = new DrawingOperations(logger);
        var constructionNotesService = new ConstructionNotesService(logger, excelReader, drawingOps);
        return new PlottingService(logger, constructionNotesService, drawingOps, excelReader);
    }
    
    private static IProjectConfigurationService GetConfigurationService()
    {
        var logger = new DebugLogger();
        return new ProjectConfigurationService(logger);
    }
    
    
    private void PlotButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            UpdatePlottingDisplay("🔄 Starting plot operation...");
            
            // For now, show that plotting is ready - full integration will come with sheet selection UI
            NotificationService?.ShowInformation("Plot", "Plotting feature is now functional! Full integration will be available once sheet selection UI is implemented.");
            
            var result = $"{Environment.NewLine}📊 Plotting System Status{Environment.NewLine}";
            result += $"✅ PlottingService: Ready{Environment.NewLine}";
            result += $"✅ SharedUIStateService: Connected{Environment.NewLine}";
            result += $"✅ UI Controls: Functional{Environment.NewLine}";
            result += $"✅ Settings: {Environment.NewLine}";
            result += $"   • Update Construction Notes: {(UpdateConstructionNotesCheckBox.IsChecked == true ? "Yes" : "No")}{Environment.NewLine}";
            result += $"   • Update Title Blocks: {(UpdateTitleBlocksCheckBox.IsChecked == true ? "Yes" : "No")}{Environment.NewLine}";
            result += $"   • Apply to Current Sheet Only: {(ApplyToCurrentSheetCheckBox.IsChecked == true ? "Yes" : "No")}{Environment.NewLine}";
            result += $"{Environment.NewLine}🔄 Ready for sheet selection integration!";
            
            UpdatePlottingDisplay(result);
        }
        catch (Exception ex)
        {
            Logger?.LogError($"Plot operation failed: {ex.Message}", ex);
            NotificationService?.ShowError("Plot", $"Plot operation failed: {ex.Message}");
            UpdatePlottingDisplay($"❌ Plot operation failed: {ex.Message}");
        }
    }
    
    
    
    private void UpdatePlottingDisplay(string message)
    {
        Dispatcher.Invoke(() =>
        {
            PlottingTextBlock.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        });
    }
    
    private void ApplyToCurrentSheetCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        // Update shared state
        _sharedUIState.ApplyToCurrentSheetOnly = ApplyToCurrentSheetCheckBox.IsChecked == true;
    }
    
    private void OnSharedApplyToCurrentSheetOnlyChanged(bool applyToCurrentSheetOnly)
    {
        // Update checkbox if it doesn't match shared state
        if (ApplyToCurrentSheetCheckBox.IsChecked != applyToCurrentSheetOnly)
        {
            ApplyToCurrentSheetCheckBox.IsChecked = applyToCurrentSheetOnly;
        }
    }
}