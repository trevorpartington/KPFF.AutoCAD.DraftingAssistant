using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for displaying help information
/// </summary>
public class HelpCommandHandler : ICommandHandler
{
    private readonly ILogger _logger;
    private readonly IEnumerable<ICommandHandler> _availableCommands;

    public HelpCommandHandler(ILogger logger, IEnumerable<ICommandHandler> availableCommands)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _availableCommands = availableCommands ?? throw new ArgumentNullException(nameof(availableCommands));
    }

    public string CommandName => CommandNames.KpffHelp;
    public string Description => "Displays help information for all available commands";

    public void Execute()
    {
        try
        {
            _logger.LogInformation($"Executing command: {CommandName}");
            
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage($"\n=== {ApplicationConstants.ApplicationName} Help ===\n");
            ed.WriteMessage("Available Commands:\n");

            foreach (var command in _availableCommands.OrderBy(c => c.CommandName))
            {
                ed.WriteMessage($"  {command.CommandName.PadRight(25)} - {command.Description}\n");
            }

            ed.WriteMessage("============================================\n");
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error executing command {CommandName}", ex);
            throw;
        }
    }
}