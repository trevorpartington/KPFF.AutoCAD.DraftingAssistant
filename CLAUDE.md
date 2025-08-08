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

## Construction Notes Automation System

### Overview
The construction notes automation system provides dual-mode functionality:
- **Auto Notes**: Automatically scans drawing viewports to detect bubble multileaders and extract note numbers
- **Excel Notes**: Uses manual Excel-based configuration for note assignment to sheets

### File Structure

#### Project Configuration
- **ProjectConfig.json**: Main project settings including:
  - Project name, client information
  - Sheet naming convention patterns (regex-based)
  - File paths to Excel index file
  - Multileader style specifications for Auto Notes

#### Excel File Structure (ProjectIndex.xlsx)
- **SHEET_INDEX table**: Complete sheet list with title block information
- **{SERIES}_NOTES tables**: Construction note definitions per drawing series
  - Example: PV_NOTES, C_NOTES, UCP_NOTES
  - Columns: Note Number, Note Text
- **EXCEL_NOTES table**: Manual sheet-to-notes mapping for fallback mode
  - Column 1: Sheet Name (e.g., "PV-101")
  - Columns 2-25: Note numbers present on that sheet

### User Interface

#### Configuration Tab Layout
```
[Select Sheets] [Select Project] [Configure Project]
─────────────────────────────────────────────────────
Active Project: [Project Name or "No project selected"]
─────────────────────────────────────────────────────
[Configuration details display area]
```

#### Construction Notes Tab
- Radio button selection: Auto Notes (default) | Excel Notes
- "Update Notes" button processes selected sheets based on chosen mode

### Technical Implementation

#### Sheet Naming Convention
Projects use consistent naming patterns that are configurable:
- Examples: "PV-101", "C001", "UCP-12"
- Regex patterns defined in project configuration
- Auto-detection with user validation fallback

#### Construction Note Blocks
Each layout contains 24 dynamic blocks with naming pattern: `{sheet-number}-NTXX`
- Examples: "101-NT01", "1-NT24"
- Dynamic block attributes:
  - **Number**: Note number (1, 2, 4, etc.)
  - **Note**: Full note text from series-specific table
  - **Visibility**: ON/OFF state for displaying/hiding blocks
- Empty blocks start with Visibility=OFF to remain invisible

#### Auto Notes Processing
1. Identify viewports in layout
2. Calculate model space boundaries for each viewport
3. Use ray casting to determine which bubble multileaders fall within boundaries
4. Filter multileaders by designated style (from project config)
5. Extract note numbers from bubble text content
6. Combine results from all viewports into master list

#### Excel Notes Processing
1. Read EXCEL-NOTES table for target sheet
2. Extract note numbers from columns 2-25
3. Map note numbers to note text using series-specific tables
4. Update construction note blocks directly

#### Dynamic Block Handling
```csharp
// Access pattern for dynamic blocks
if (blockRef.IsDynamicBlock)
{
    var baseBlockName = blockRef.DynamicBlockTableRecord.Name;
    var properties = blockRef.DynamicBlockReferencePropertyCollection;
    // Modify Number, Note attributes and Visibility property
}
```

### Development Dependencies
- **EPPlus**: Excel file reading (handles files open in Excel)
- **AutoCAD .NET API**: Drawing analysis, block manipulation, viewport processing
- **System.Text.Json**: Project configuration serialization

### Future Enhancements
- Automated construction note block creation for new sheets
- Enhanced viewport geometry support for complex layouts
- Template-based note management system