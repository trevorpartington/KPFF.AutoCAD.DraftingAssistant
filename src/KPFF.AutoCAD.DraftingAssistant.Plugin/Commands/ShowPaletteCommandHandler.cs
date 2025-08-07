using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for showing the drafting assistant palette
/// </summary>
public class ShowPaletteCommandHandler : ICommandHandler
{
    private readonly IPaletteManager _paletteManager;
    private readonly ILogger _logger;

    public ShowPaletteCommandHandler(IPaletteManager paletteManager, ILogger logger)
    {
        _paletteManager = paletteManager ?? throw new ArgumentNullException(nameof(paletteManager));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public string CommandName => CommandNames.DraftingAssistant;
    public string Description => "Shows the KPFF Drafting Assistant palette";

    public void Execute()
    {
        try
        {
            _logger.LogInformation($"Executing command: {CommandName}");
            _paletteManager.Show();
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error executing command {CommandName}", ex);
            throw;
        }
    }
}