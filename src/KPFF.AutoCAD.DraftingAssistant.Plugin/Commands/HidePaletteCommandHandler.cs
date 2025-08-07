using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for hiding the drafting assistant palette
/// </summary>
public class HidePaletteCommandHandler : ICommandHandler
{
    private readonly IPaletteManager _paletteManager;
    private readonly ILogger _logger;

    public HidePaletteCommandHandler(IPaletteManager paletteManager, ILogger logger)
    {
        _paletteManager = paletteManager ?? throw new ArgumentNullException(nameof(paletteManager));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public string CommandName => CommandNames.HideDraftingAssistant;
    public string Description => "Hides the KPFF Drafting Assistant palette";

    public void Execute()
    {
        try
        {
            _logger.LogInformation($"Executing command: {CommandName}");
            _paletteManager.Hide();
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error executing command {CommandName}", ex);
            throw;
        }
    }
}