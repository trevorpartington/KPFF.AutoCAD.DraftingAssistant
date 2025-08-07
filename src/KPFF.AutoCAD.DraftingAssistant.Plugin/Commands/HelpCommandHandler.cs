using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for displaying help information
/// </summary>
public class HelpCommandHandler : ICommandHandler
{
    private readonly ILogger _logger;

    public HelpCommandHandler(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public string CommandName => CommandNames.KpffHelp;
    public string Description => "Displays help information for all available commands";

    public void Execute()
    {
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                ed.WriteMessage($"\n=== {ApplicationConstants.ApplicationName} Help ===\n");
                ed.WriteMessage("Available Commands:\n");

                // Define available commands statically since we know them
                var commands = new[]
                {
                    new { Name = CommandNames.DraftingAssistant, Description = "Shows the KPFF Drafting Assistant palette" },
                    new { Name = CommandNames.HideDraftingAssistant, Description = "Hides the KPFF Drafting Assistant palette" },
                    new { Name = CommandNames.ToggleDraftingAssistant, Description = "Toggles the KPFF Drafting Assistant palette visibility" },
                    new { Name = CommandNames.KpffStart, Description = "Initializes and shows the KPFF Drafting Assistant" },
                    new { Name = CommandNames.KpffHelp, Description = "Displays this help information" }
                };

                foreach (var command in commands.OrderBy(c => c.Name))
                {
                    ed.WriteMessage($"  {command.Name.PadRight(30)} - {command.Description}\n");
                }

                ed.WriteMessage("\nFor more information, visit: https://kpff.com\n");
                ed.WriteMessage("============================================\n");
            },
            logger: _logger,
            context: $"Command Execution: {CommandName}"
        );
    }
}