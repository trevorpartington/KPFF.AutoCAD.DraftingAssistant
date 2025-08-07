// This file maintains the assembly attributes and redirects to the new architecture
// for backward compatibility during refactoring

using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(KPFF.AutoCAD.DraftingAssistant.Plugin.DraftingAssistantCommands))]
[assembly: ExtensionApplication(typeof(KPFF.AutoCAD.DraftingAssistant.Plugin.DraftingAssistantExtensionApplication))]

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

// This file serves as the entry point for AutoCAD assembly discovery
// The actual implementation has been moved to:
// - DraftingAssistantExtensionApplication.cs (for IExtensionApplication)
// - DraftingAssistantCommands.cs (for CommandClass)
// - Services/ folder for service implementations
// - Commands/ folder for command handlers