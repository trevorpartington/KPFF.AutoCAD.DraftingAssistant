using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Controls;

public partial class DraftingAssistantControl : BaseUserControl
{
    public DraftingAssistantControl() : this(null, null)
    {
    }

    public DraftingAssistantControl(
        ILogger? logger,
        INotificationService? notificationService) 
        : base(logger, notificationService)
    {
        InitializeComponent();
    }
}