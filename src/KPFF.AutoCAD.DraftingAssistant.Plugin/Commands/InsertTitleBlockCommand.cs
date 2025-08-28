using System;
using System.IO;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command to insert title block (TB_ATT.dwg) at origin (0,0)
/// </summary>
public class InsertTitleBlockCommand : ICommandHandler
{
    private readonly ILogger _logger;
    private readonly BlockInsertionService _blockInsertionService;

    public InsertTitleBlockCommand(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _blockInsertionService = new BlockInsertionService(_logger);
    }

    public string CommandName => "KPFFINSERTTITLEBLOCK";
    public string Description => "Insert title block (TB_ATT.dwg) at origin (0,0)";

    public void Execute()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                
                // Path to the title block file
                var blockFilePath = GetTitleBlockPath();
                
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
    /// Gets the path to the title block DWG file
    /// </summary>
    private string? GetTitleBlockPath()
    {
        try
        {
            // Get the path relative to the project directory
            var projectDir = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location!))));
            
            if (projectDir != null)
            {
                var blockFilePath = Path.Combine(projectDir, "testdata", "Blocks", "TB_ATT.dwg");
                
                _logger.LogDebug($"Looking for title block at: {blockFilePath}");
                
                if (File.Exists(blockFilePath))
                {
                    _logger.LogDebug($"Found title block at: {blockFilePath}");
                    return blockFilePath;
                }
                else
                {
                    _logger.LogError($"Title block file not found at: {blockFilePath}");
                }
            }
            
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting title block path: {ex.Message}", ex);
            return null;
        }
    }
}