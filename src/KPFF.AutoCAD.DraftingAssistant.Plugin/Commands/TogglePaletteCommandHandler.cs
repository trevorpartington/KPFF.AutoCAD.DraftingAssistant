using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

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
        ExceptionHandler.TryExecute(
            action: () =>
            {
                _logger.LogInformation($"Executing command: {CommandName}");
                var wasVisible = _paletteManager.IsVisible;
                _paletteManager.Toggle();
                var newState = _paletteManager.IsVisible ? "visible" : "hidden";
                _logger.LogDebug($"Palette toggled from {(wasVisible ? "visible" : "hidden")} to {newState}");
            },
            logger: _logger,
            context: $"Command Execution: {CommandName}"
        );
    }
}