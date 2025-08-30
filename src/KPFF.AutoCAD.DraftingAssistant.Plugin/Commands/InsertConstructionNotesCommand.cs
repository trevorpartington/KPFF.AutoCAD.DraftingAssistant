using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for inserting construction note blocks from external DWG files
/// </summary>
public class InsertConstructionNotesCommand : ICommandHandler
{
    private readonly ILogger _logger;
    private readonly BlockInsertionService _blockInsertionService;
    private readonly IProjectConfigurationService _configService;

    public InsertConstructionNotesCommand(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _blockInsertionService = new BlockInsertionService(_logger);
        _configService = new ProjectConfigurationService(new AutoCADLogger());
    }

    public string CommandName => "KPFFINSERTBLOCKS";
    public string Description => "Insert construction note blocks from external DWG files";

    public void Execute()
    {
        ExceptionHandler.TryExecute(
            action: async () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                
                // Path to the construction note block file
                var blockFilePath = await GetConstructionNoteBlockPathAsync();
                
                if (string.IsNullOrEmpty(blockFilePath))
                {
                    _logger.LogError("Could not locate construction note block file");
                    return;
                }

                // Insert all 24 construction note blocks (NT01 through NT24)
                var success = _blockInsertionService.InsertConstructionNoteBlockStack(blockFilePath);
                
                if (success)
                {
                    _logger.LogInformation("Construction note block stack inserted successfully");
                }
                else
                {
                    _logger.LogError("Failed to insert construction note block stack");
                }
            },
            logger: _logger,
            context: $"Command Execution: {CommandName}",
            showUserMessage: true
        );
    }

    /// <summary>
    /// Gets the path to the construction note block DWG file from project configuration
    /// </summary>
    private async Task<string> GetConstructionNoteBlockPathAsync()
    {
        try
        {
            // Try to load the default project configuration
            var defaultConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\DBRT_Config.json";
            
            if (File.Exists(defaultConfigPath))
            {
                _logger.LogInformation($"Loading configuration from: {defaultConfigPath}");
                var projectConfig = await _configService.LoadConfigurationAsync(defaultConfigPath);
                
                if (projectConfig != null && 
                    !string.IsNullOrWhiteSpace(projectConfig.ConstructionNotes.NoteBlockFilePath))
                {
                    var configuredPath = projectConfig.ConstructionNotes.NoteBlockFilePath;
                    _logger.LogInformation($"Found configured construction note block path: {configuredPath}");
                    
                    if (File.Exists(configuredPath))
                    {
                        _logger.LogInformation($"Using configured construction note block path: {configuredPath}");
                        return configuredPath;
                    }
                    else
                    {
                        _logger.LogError($"Configured construction note block file does not exist: {configuredPath}");
                        return string.Empty;
                    }
                }
                else
                {
                    _logger.LogWarning("Configuration loaded but noteBlockFilePath is empty or null");
                }
            }
            else
            {
                _logger.LogError($"Configuration file does not exist: {defaultConfigPath}");
            }
            
            // No configuration available
            _logger.LogError("Construction note block file path not configured. Please configure the NTXX.dwg file location in Project Settings.");
            return string.Empty;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting construction note block path: {ex.Message}", ex);
            return string.Empty;
        }
    }

}