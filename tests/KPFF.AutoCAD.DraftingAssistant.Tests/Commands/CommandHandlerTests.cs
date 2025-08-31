using KPFF.AutoCAD.DraftingAssistant.Core.Constants;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Plugin.Commands;
using Moq;

namespace KPFF.AutoCAD.DraftingAssistant.Tests.Commands;

/// <summary>
/// Unit tests for command handlers
/// </summary>
public class CommandHandlerTests
{
    private readonly Mock<IPaletteManager> _mockPaletteManager;
    private readonly Mock<ILogger> _mockLogger;

    public CommandHandlerTests()
    {
        _mockPaletteManager = new Mock<IPaletteManager>();
        _mockLogger = new Mock<ILogger>();
    }

    [Fact]
    public void TogglePaletteCommandHandler_Should_Have_Correct_Properties()
    {
        // Arrange
        var handler = new TogglePaletteCommandHandler(_mockPaletteManager.Object, _mockLogger.Object);

        // Assert
        Assert.Equal(CommandNames.Kpff, handler.CommandName);
        Assert.False(string.IsNullOrEmpty(handler.Description));
    }

    [Fact]
    public void TogglePaletteCommandHandler_Execute_Should_Call_PaletteManager_Toggle()
    {
        // Arrange
        var handler = new TogglePaletteCommandHandler(_mockPaletteManager.Object, _mockLogger.Object);

        // Act
        handler.Execute();

        // Assert
        _mockPaletteManager.Verify(pm => pm.Toggle(), Times.Once);
        _mockLogger.Verify(l => l.LogInformation(It.IsAny<string>()), Times.Once);
    }

    [Fact]
    public void TogglePaletteCommandHandler_Should_Throw_ArgumentNullException_For_Null_Dependencies()
    {
        // Assert
        Assert.Throws<ArgumentNullException>(() => 
            new TogglePaletteCommandHandler(null!, _mockLogger.Object));
        Assert.Throws<ArgumentNullException>(() => 
            new TogglePaletteCommandHandler(_mockPaletteManager.Object, null!));
    }
}