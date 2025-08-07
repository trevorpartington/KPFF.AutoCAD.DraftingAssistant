using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;

/// <summary>
/// Command handler for toggling the drafting assistant palette visibility
/// </summary>
public class TogglePaletteCommandHandler : ICommandHandler
{
    private readonly IPaletteManager _paletteManager;
    private readonly ILogger _logger;

    public TogglePaletteCommandHandler(IPaletteManager paletteManager, ILogger logger)
    {
        _paletteManager = paletteManager ?? throw new ArgumentNullException(nameof(paletteManager));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public string CommandName => CommandNames.ToggleDraftingAssistant;
    public string Description => "Toggles the KPFF Drafting Assistant palette visibility";

    public void Execute()
    {
        try
        {
            _logger.LogInformation($"Executing command: {CommandName}");
            _paletteManager.Toggle();
        }
        catch (System.Exception ex)
        {
            _logger.LogError($"Error executing command {CommandName}", ex);
            throw;
        }
    }
}