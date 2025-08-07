using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;

namespace KPFF.AutoCAD.DraftingAssistant.Tests.Core;

/// <summary>
/// Unit tests for the ServiceContainer dependency injection implementation
/// </summary>
public class ServiceContainerTests : IDisposable
{
    private readonly ServiceContainer _container;

    public ServiceContainerTests()
    {
        _container = ServiceContainer.Instance;
        _container.Clear(); // Start with clean container for each test
    }

    public void Dispose()
    {
        _container.Clear(); // Clean up after each test
    }

    [Fact]
    public void Register_And_Resolve_Service_Should_Work()
    {
        // Arrange
        var logger = new DebugLogger();

        // Act
        _container.Register<ILogger>(logger);
        var resolved = _container.Resolve<ILogger>();

        // Assert
        Assert.Same(logger, resolved);
    }

    [Fact]
    public void Register_With_Factory_Should_Work()
    {
        // Arrange & Act
        _container.Register<ILogger>(() => new DebugLogger());
        var resolved = _container.Resolve<ILogger>();

        // Assert
        Assert.NotNull(resolved);
        Assert.IsType<DebugLogger>(resolved);
    }

    [Fact]
    public void IsRegistered_Should_Return_True_For_Registered_Service()
    {
        // Arrange
        _container.Register<ILogger>(new DebugLogger());

        // Act
        var isRegistered = _container.IsRegistered<ILogger>();

        // Assert
        Assert.True(isRegistered);
    }

    [Fact]
    public void IsRegistered_Should_Return_False_For_Unregistered_Service()
    {
        // Act
        var isRegistered = _container.IsRegistered<ILogger>();

        // Assert
        Assert.False(isRegistered);
    }

    [Fact]
    public void Resolve_Unregistered_Service_Should_Throw_Exception()
    {
        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => _container.Resolve<ILogger>());
    }

    [Fact]
    public void Clear_Should_Remove_All_Services()
    {
        // Arrange
        _container.Register<ILogger>(new DebugLogger());
        Assert.True(_container.IsRegistered<ILogger>());

        // Act
        _container.Clear();

        // Assert
        Assert.False(_container.IsRegistered<ILogger>());
    }
}