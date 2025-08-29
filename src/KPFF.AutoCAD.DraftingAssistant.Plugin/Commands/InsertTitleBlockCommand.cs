using System;
using System.IO;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command to insert title block (TB_ATT.dwg) at origin (0,0)
/// </summary>
public class InsertTitleBlockCommand : ICommandHandler
{
    private readonly ILogger _logger;
    private readonly BlockInsertionService _blockInsertionService;
    private readonly IProjectConfigurationService _configService;

    public InsertTitleBlockCommand(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _blockInsertionService = new BlockInsertionService(_logger);
        _configService = new ProjectConfigurationService(new AutoCADLogger());
    }

    public string CommandName => "KPFFINSERTTITLEBLOCK";
    public string Description => "Insert title block (TB_ATT.dwg) at origin (0,0)";

    public void Execute()
    {
        ExceptionHandler.TryExecute(
            action: async () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                
                // Path to the title block file
                var blockFilePath = await GetTitleBlockPathAsync();
                
                if (string.IsNullOrEmpty(blockFilePath))
                {
                    _logger.LogError("Could not locate title block file");
                    return;
                }

                // Insert the TB_ATT block at origin (0,0)
                var success = _blockInsertionService.InsertTitleBlock(blockFilePath);
                
                if (success)
                {
                    _logger.LogInformation("Title block inserted successfully at origin (0,0)");
                }
                else
                {
                    _logger.LogError("Failed to insert title block");
                }
            },
            logger: _logger,
            context: $"Command Execution: {CommandName}",
            showUserMessage: true
        );
    }

    /// <summary>
    /// Gets the path to the title block DWG file from project configuration
    /// </summary>
    private async Task<string?> GetTitleBlockPathAsync()
    {
        try
        {
            // Try to load the default project configuration
            var defaultConfigPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\DBRT Test\DBRT_Config.json";
            
            if (File.Exists(defaultConfigPath))
            {
                var projectConfig = await _configService.LoadConfigurationAsync(defaultConfigPath);
                
                if (projectConfig != null && 
                    !string.IsNullOrWhiteSpace(projectConfig.TitleBlocks.TitleBlockFilePath))
                {
                    var configuredPath = projectConfig.TitleBlocks.TitleBlockFilePath;
                    
                    if (File.Exists(configuredPath))
                    {
                        _logger.LogDebug($"Using configured title block path: {configuredPath}");
                        return configuredPath;
                    }
                    else
                    {
                        _logger.LogError($"Configured title block file does not exist: {configuredPath}");
                        return null;
                    }
                }
            }
            
            // No configuration available
            _logger.LogError("Title block file path not configured. Please configure the title block file location in Project Settings.");
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting title block path: {ex.Message}", ex);
            return null;
        }
    }
}