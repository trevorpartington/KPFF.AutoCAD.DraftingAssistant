using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for inserting construction note blocks from external DWG files
/// </summary>
public class InsertConstructionNotesCommand : ICommandHandler
{
    private readonly ILogger _logger;
    private readonly BlockInsertionService _blockInsertionService;

    public InsertConstructionNotesCommand(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _blockInsertionService = new BlockInsertionService(_logger);
    }

    public string CommandName => "KPFFINSERTBLOCKS";
    public string Description => "Insert construction note blocks from external DWG files";

    public void Execute()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                
                // Path to the construction note block file
                var blockFilePath = GetConstructionNoteBlockPath();
                
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
    /// Gets the path to the construction note block DWG file
    /// </summary>
    private string GetConstructionNoteBlockPath()
    {
        try
        {
            // Get the plugin assembly location to find the testdata folder
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            
            // Navigate to the solution root and find testdata/Blocks
            var solutionRoot = FindSolutionRoot(assemblyDirectory);
            if (!string.IsNullOrEmpty(solutionRoot))
            {
                var blockPath = Path.Combine(solutionRoot, "testdata", "Blocks", "NTXX.dwg");
                if (File.Exists(blockPath))
                {
                    _logger.LogDebug($"Found construction note block at: {blockPath}");
                    return blockPath;
                }
            }

            // Fallback: try the hardcoded path from the user's request
            var fallbackPath = @"C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\testdata\Blocks\NTXX.dwg";
            if (File.Exists(fallbackPath))
            {
                _logger.LogDebug($"Using fallback path: {fallbackPath}");
                return fallbackPath;
            }

            _logger.LogError($"Construction note block file not found. Searched: {(solutionRoot != null ? Path.Combine(solutionRoot, "testdata", "Blocks", "NTXX.dwg") : "solution root not found")}, {fallbackPath}");
            return string.Empty;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error locating construction note block path: {ex.Message}", ex);
            return string.Empty;
        }
    }

    /// <summary>
    /// Finds the solution root directory by looking for the .sln file
    /// </summary>
    private string? FindSolutionRoot(string? startDirectory)
    {
        if (string.IsNullOrEmpty(startDirectory))
            return null;

        var directory = new DirectoryInfo(startDirectory);
        while (directory != null)
        {
            if (directory.GetFiles("*.sln").Length > 0)
            {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
        return null;
    }
}