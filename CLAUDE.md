# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development Commands

### Building the Solution
- `dotnet build` - Build all projects in the solution
- `dotnet build -c Release` - Build in Release configuration
- `dotnet restore` - Restore NuGet packages

### Running Tests  
- `dotnet test` - Run all xUnit tests
- `dotnet test --verbosity normal` - Run tests with detailed output
- `dotnet test --collect:"XPlat Code Coverage"` - Run tests with coverage

### Individual Project Operations
- `dotnet build src/KPFF.AutoCAD.DraftingAssistant.Core` - Build specific project
- `dotnet test tests/KPFF.AutoCAD.DraftingAssistant.Tests` - Run specific test project

## Architecture Overview

This is a **4-project .NET solution** for an AutoCAD plugin system:

### Project Dependencies
```
Core (net8.0-windows) ← Plugin (depends on Core)
                      ← UI (WPF, depends on Core) 
                      ← Tests (net8.0, xUnit, references Plugin + Core)
```

### Key Components
- **Core Library**: Shared business logic, models, interfaces, and utilities
- **Plugin**: AutoCAD API integration and command implementations  
- **UI**: WPF-based user interface components
- **Tests**: xUnit test suite with project references

### Technology Stack
- .NET 8.0 (Windows) for core components
- .NET 8.0 for tests
- WPF for UI framework
- xUnit for testing framework
- AutoCAD API integration
- Clean architecture with dependency injection

### Project Structure
- Solution configured for multiple platforms: Any CPU, x64, x86
- Both Debug and Release configurations available
- Modern C# with nullable reference types enabled and implicit usings

## Development Notes

- The test project uses xUnit framework (not MSTest as mentioned in README)
- All projects have nullable reference types enabled
- The solution is designed for AutoCAD 2020+ compatibility
- UI project uses WPF with App.xaml/MainWindow.xaml structure

## AutoCAD Plugin Debugging

### Prerequisites
- AutoCAD 2025 with Civil 3D installed
- Visual Studio 2022 with the plugin project loaded

### Debugging Setup
1. Set the startup project to `KPFF.AutoCAD.DraftingAssistant.Plugin`
2. Select the "Debug Civil 3D 2025" launch profile
3. Press F5 to start debugging

### How It Works
- Visual Studio launches AutoCAD 2025 with Civil 3D Imperial template
- AutoCAD executes `start.scr` which loads the plugin DLL
- Plugin commands are automatically executed to verify functionality
- Set breakpoints in the plugin code for debugging

### Available Plugin Commands
- `KPFFSTART` - Initialize and start the drafting assistant
- `KPFFHELP` - Display available commands and help information

### Debugging Commands
- `dotnet build src/KPFF.AutoCAD.DraftingAssistant.Plugin` - Build plugin project
- `dotnet build src/KPFF.AutoCAD.DraftingAssistant.Plugin -c Release` - Build for release

### Troubleshooting
- Ensure AutoCAD 2025 path is correct in launch profile
- Verify DLL path in start.scr matches build output location
- Check AutoCAD command line for plugin loading errors
- Use `NETUNLOAD` in AutoCAD to unload plugin before rebuilding