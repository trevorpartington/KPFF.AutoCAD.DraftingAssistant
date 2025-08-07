using Autodesk.AutoCAD.ApplicationServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// Manages the startup sequence and initialization state of the plugin
/// </summary>
public class PluginStartupManager
{
    private readonly ILogger _logger;
    private readonly IPaletteManager _paletteManager;
    private bool _isInitialized = false;
    private bool _isInitializing = false;
    private int _initAttempts = 0;
    private const int MaxInitAttempts = 3;
    private const int InitDelayMs = 2000;

    public PluginStartupManager(ILogger logger, IPaletteManager paletteManager)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _paletteManager = paletteManager ?? throw new ArgumentNullException(nameof(paletteManager));
    }

    public bool IsInitialized => _isInitialized;
    public bool IsInitializing => _isInitializing;

    /// <summary>
    /// Starts the initialization process with multiple fallback mechanisms
    /// </summary>
    public void BeginInitialization()
    {
        if (_isInitialized || _isInitializing)
        {
            _logger.LogDebug("Initialization already in progress or completed");
            return;
        }

        _logger.LogInformation("Beginning plugin initialization sequence");
        
        // Try multiple initialization strategies
        TryImmediateInitialization();
        ScheduleDelayedInitialization();
        RegisterEventBasedInitialization();
    }

    /// <summary>
    /// Attempts immediate initialization if conditions are favorable
    /// </summary>
    private void TryImmediateInitialization()
    {
        try
        {
            if (IsAutoCadReady())
            {
                _logger.LogInformation("AutoCAD appears ready - attempting immediate initialization");
                PerformInitialization();
            }
            else
            {
                _logger.LogInformation("AutoCAD not ready for immediate initialization");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogWarning($"Immediate initialization failed, will retry later: {ex.Message}");
        }
    }

    /// <summary>
    /// Schedules delayed initialization using a timer
    /// </summary>
    private void ScheduleDelayedInitialization()
    {
        var timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(InitDelayMs)
        };
        
        timer.Tick += (sender, e) =>
        {
            timer.Stop();
            if (!_isInitialized)
            {
                _logger.LogInformation("Attempting delayed initialization");
                TryInitializationWithRetry();
            }
        };
        
        timer.Start();
        _logger.LogDebug($"Scheduled delayed initialization in {InitDelayMs}ms");
    }

    /// <summary>
    /// Registers for AutoCAD events to trigger initialization
    /// </summary>
    private void RegisterEventBasedInitialization()
    {
        try
        {
            Application.Idle += OnApplicationIdle;
            
            // Register for document lifecycle events
            var docManager = Application.DocumentManager;
            docManager.DocumentCreated += OnDocumentCreated;
            docManager.DocumentToBeActivated += OnDocumentToBeActivated;
            docManager.DocumentActivated += OnDocumentActivated;
            docManager.DocumentToBeDestroyed += OnDocumentToBeDestroyed;
            
            _logger.LogDebug("Registered for AutoCAD application and document events");
        }
        catch (System.Exception ex)
        {
            _logger.LogWarning($"Failed to register for AutoCAD events: {ex.Message}");
        }
    }

    private void OnApplicationIdle(object? sender, System.EventArgs e)
    {
        if (!_isInitialized)
        {
            _logger.LogInformation("AutoCAD is idle - attempting initialization");
            TryInitializationWithRetry();
        }
        
        // Unsubscribe after successful initialization or max attempts
        if (_isInitialized || _initAttempts >= MaxInitAttempts)
        {
            Application.Idle -= OnApplicationIdle;
        }
    }

    private void OnDocumentCreated(object? sender, DocumentCollectionEventArgs e)
    {
        try
        {
            _logger.LogDebug($"Document created: {e.Document?.Name ?? "Unknown"}");
            
            if (!_isInitialized)
            {
                _logger.LogInformation("Document created - attempting initialization");
                TryInitializationWithRetry();
            }
            else
            {
                // Ensure palette is available for new documents
                ValidatePaletteForDocument(e.Document);
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error handling document created event", ex);
        }
    }

    private void OnDocumentToBeActivated(object? sender, DocumentCollectionEventArgs e)
    {
        try
        {
            _logger.LogDebug($"Document about to be activated: {e.Document?.Name ?? "Unknown"}");
            
            // Prepare for document switch - suspend any active operations
            if (_isInitialized)
            {
                PrepareForDocumentSwitch(e.Document);
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error handling document to be activated event", ex);
        }
    }

    private void OnDocumentActivated(object? sender, DocumentCollectionEventArgs e)
    {
        try
        {
            _logger.LogDebug($"Document activated: {e.Document?.Name ?? "Unknown"}");
            
            if (_isInitialized)
            {
                // Ensure services are available for the activated document
                ValidatePaletteForDocument(e.Document);
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error handling document activated event", ex);
        }
    }

    private void OnDocumentToBeDestroyed(object? sender, DocumentCollectionEventArgs e)
    {
        try
        {
            _logger.LogDebug($"Document about to be destroyed: {e.Document?.Name ?? "Unknown"}");
            
            if (_isInitialized)
            {
                // Clean up any document-specific resources
                CleanupDocumentResources(e.Document);
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Error handling document to be destroyed event", ex);
        }
    }

    /// <summary>
    /// Attempts initialization with retry logic
    /// </summary>
    private void TryInitializationWithRetry()
    {
        if (_isInitialized || _isInitializing)
            return;

        _initAttempts++;
        _logger.LogInformation($"Initialization attempt {_initAttempts}/{MaxInitAttempts}");

        try
        {
            if (IsAutoCadReady())
            {
                PerformInitialization();
            }
            else
            {
                _logger.LogWarning($"AutoCAD not ready for initialization (attempt {_initAttempts})");
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Initialization attempt {_initAttempts} failed", ex);
        }

        // If we've exceeded max attempts, give up
        if (_initAttempts >= MaxInitAttempts && !_isInitialized)
        {
            _logger.LogError("Maximum initialization attempts exceeded - plugin may not function correctly");
            CleanupEventHandlers();
        }
    }

    /// <summary>
    /// Checks if AutoCAD is in a state ready for initialization
    /// </summary>
    private static bool IsAutoCadReady()
    {
        try
        {
            // Check if application is available
            var docManager = Application.DocumentManager;
            if (docManager == null) return false;

            // Check if we can access basic AutoCAD functionality
            var docCount = docManager.Count; // This will throw if AutoCAD isn't ready
            
            // Additional checks for document readiness
            if (docCount > 0)
            {
                var currentDoc = docManager.MdiActiveDocument;
                if (currentDoc != null)
                {
                    // Try to access document name to ensure it's fully loaded
                    var _ = currentDoc.Name;
                }
            }

            // Additional validation - check if we can access main application services
            try
            {
                var _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.Version;
            }
            catch
            {
                return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Performs the actual initialization
    /// </summary>
    private void PerformInitialization()
    {
        if (_isInitializing) return;

        _isInitializing = true;
        try
        {
            _logger.LogInformation("Performing plugin initialization");
            
            // Initialize the palette manager
            _paletteManager.Initialize();
            
            _isInitialized = true;
            _logger.LogInformation("Plugin initialization completed successfully");
            
            // Clean up event handlers
            CleanupEventHandlers();
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Plugin initialization failed", ex);
            throw;
        }
        finally
        {
            _isInitializing = false;
        }
    }

    /// <summary>
    /// Validates that palette services are ready for the specified document
    /// </summary>
    private void ValidatePaletteForDocument(Document? document)
    {
        try
        {
            if (document == null)
            {
                _logger.LogWarning("Cannot validate palette for null document");
                return;
            }

            _logger.LogDebug($"Validating palette for document: {document.Name}");
            
            // Check if palette manager is still functional
            if (!_paletteManager.IsInitialized)
            {
                _logger.LogWarning("Palette manager not initialized - attempting re-initialization");
                _paletteManager.Initialize();
            }
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error validating palette for document {document?.Name}", ex);
        }
    }

    /// <summary>
    /// Prepares for document switching by suspending active operations
    /// </summary>
    private void PrepareForDocumentSwitch(Document? document)
    {
        try
        {
            if (document == null) return;

            _logger.LogDebug($"Preparing for document switch: {document.Name}");
            
            // Could add logic here to suspend any active operations
            // or save state before document switches
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error preparing for document switch: {document?.Name}", ex);
        }
    }

    /// <summary>
    /// Cleans up resources associated with a document being destroyed
    /// </summary>
    private void CleanupDocumentResources(Document? document)
    {
        try
        {
            if (document == null) return;

            _logger.LogDebug($"Cleaning up resources for document: {document.Name}");
            
            // Could add logic here to clean up document-specific resources
            // if we stored any document-specific state
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error cleaning up document resources: {document?.Name}", ex);
        }
    }

    /// <summary>
    /// Cleans up event handlers
    /// </summary>
    private void CleanupEventHandlers()
    {
        try
        {
            Application.Idle -= OnApplicationIdle;
            
            var docManager = Application.DocumentManager;
            docManager.DocumentCreated -= OnDocumentCreated;
            docManager.DocumentToBeActivated -= OnDocumentToBeActivated;
            docManager.DocumentActivated -= OnDocumentActivated;
            docManager.DocumentToBeDestroyed -= OnDocumentToBeDestroyed;
            
            _logger.LogDebug("Cleaned up all event handlers");
        }
        catch (System.Exception ex)
        {
            _logger.LogWarning($"Error cleaning up event handlers: {ex.Message}");
        }
    }

    /// <summary>
    /// Forces initialization if not already completed
    /// </summary>
    public void ForceInitialization()
    {
        if (_isInitialized) return;

        _logger.LogInformation("Forcing plugin initialization");
        try
        {
            PerformInitialization();
        }
        catch (System.Exception ex)
        {
            _logger.LogError("Forced initialization failed", ex);
            throw;
        }
    }
}