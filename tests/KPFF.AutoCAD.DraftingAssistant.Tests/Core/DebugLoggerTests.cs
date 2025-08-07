using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using System.Diagnostics;

namespace KPFF.AutoCAD.DraftingAssistant.Tests.Core;

/// <summary>
/// Unit tests for the DebugLogger implementation
/// </summary>
public class DebugLoggerTests
{
    private readonly DebugLogger _logger;

    public DebugLoggerTests()
    {
        _logger = new DebugLogger();
    }

    [Fact]
    public void LogInformation_Should_Not_Throw()
    {
        // Act & Assert
        var exception = Record.Exception(() => _logger.LogInformation("Test information message"));
        Assert.Null(exception);
    }

    [Fact]
    public void LogWarning_Should_Not_Throw()
    {
        // Act & Assert
        var exception = Record.Exception(() => _logger.LogWarning("Test warning message"));
        Assert.Null(exception);
    }

    [Fact]
    public void LogError_With_Exception_Should_Not_Throw()
    {
        // Arrange
        var testException = new System.Exception("Test exception");

        // Act & Assert
        var exception = Record.Exception(() => _logger.LogError("Test error message", testException));
        Assert.Null(exception);
    }

    [Fact]
    public void LogError_Without_Exception_Should_Not_Throw()
    {
        // Act & Assert
        var exception = Record.Exception(() => _logger.LogError("Test error message"));
        Assert.Null(exception);
    }

    [Fact]
    public void LogDebug_Should_Not_Throw()
    {
        // Act & Assert
        var exception = Record.Exception(() => _logger.LogDebug("Test debug message"));
        Assert.Null(exception);
    }
}