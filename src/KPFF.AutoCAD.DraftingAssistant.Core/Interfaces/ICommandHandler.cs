namespace KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

/// <summary>
/// Base interface for command handlers
/// </summary>
public interface ICommandHandler
{
    string CommandName { get; }
    string Description { get; }
    void Execute();
}

/// <summary>
/// Generic command handler with parameter support
/// </summary>
/// <typeparam name="T">Command parameter type</typeparam>
public interface ICommandHandler<in T> : ICommandHandler
{
    void Execute(T parameter);
}