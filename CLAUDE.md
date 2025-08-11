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
  - Columns: Sheet, File, Title
  - Example: "ABC-101", "PROJ-ABC-100", "ABC PLAN"
- **{SERIES}_NOTES tables**: Construction note definitions per drawing series
  - Example: ABC_NOTES, PV_NOTES, C_NOTES
  - Columns: Number, Note
  - Example: Number=1, Note="CONSTRUCT CURB"
- **EXCEL_NOTES table**: Manual sheet-to-notes mapping for Excel Notes mode
  - Column 1: Sheet Name (e.g., "ABC-101")  
  - Columns 2-25: Note numbers present on that sheet
  - Example: "ABC-101", "1", "1", "", "" (means two instances of note 1 on sheet ABC-101)

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
- Pattern: Series + Number (e.g., "ABC-101" where ABC=series, 101=number)
- Supports 1-3 letter series (A, AB, ABC) and 1-3 digit numbers (1, 12, 123)  
- Examples: "A-1", "AB-12", "ABC-123"
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
1. Read EXCEL_NOTES table for target sheet
2. Extract note numbers from columns 2-25
3. Map note numbers to note text using series-specific tables (e.g., ABC_NOTES)
4. Update construction note blocks directly

**Example Workflow:**
- Sheet ABC-101 has entries in EXCEL_NOTES: column 1="ABC-101", column 2="1", column 3="1"
- This means sheet ABC-101 has two instances of note 1
- Note 1 definition comes from ABC_NOTES table: Number=1, Note="CONSTRUCT CURB"
- System updates blocks 101-NT01 and 101-NT02 with Number="1", Note="CONSTRUCT CURB", Visibility="ON"

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

### Configuration Updates (Recent)
The system has been simplified to remove worksheet dependencies:
- Tables are now accessed directly by name across the entire workbook
- Removed WorksheetConfiguration class and worksheet-based lookups
- Configuration now only requires table names: SHEET_INDEX, EXCEL_NOTES, {SERIES}_NOTES
- More robust and works regardless of which worksheet contains the tables

### Future Enhancements
- Automated construction note block creation for new sheets
- Enhanced viewport geometry support for complex layouts
- Template-based note management system

## Documentation References

### Overview
This project includes comprehensive documentation and reference materials in the `docs/` folder to assist with development and understanding of the AutoCAD plugin architecture.

### AutoCAD .NET API Reference
The `docs/api-reference/` folder contains extracted documentation from the Autodesk ObjectARX for AutoCAD 2025: Managed .NET Reference Guide, organized by namespace:

#### Available Namespaces
- **Autodesk.AutoCAD.ApplicationServices**: Core application classes (Application, Document)
- **Autodesk.AutoCAD.ApplicationServices.Core**: Application management and events
- **Autodesk.AutoCAD.DatabaseServices**: Database objects (BlockReference, Layout, MLeader, Viewport, Transaction, Database)
- **Autodesk.AutoCAD.DatabaseServices.Filters**: Spatial and layer filtering (SpatialFilter, LayerFilter, Index classes)
- **Autodesk.AutoCAD.EditorInput**: Editor interaction and input handling
- **Autodesk.AutoCAD.Geometry**: Geometric primitives and transformations

#### Key Classes Documented
- **BlockReference**: Dynamic block handling, attributes, transformations
- **Layout**: Paper space layouts, viewport management
- **MLeader**: Multileader creation and manipulation
- **Viewport**: Viewport properties and model space mapping
- **Transaction/TransactionManager**: Database transaction handling
- **SpatialFilter**: Spatial querying and filtering
- **Editor**: User input and selection

**Note**: If you need documentation for any AutoCAD API methods, properties, or classes not currently in the reference folder, please request them specifically and they can be extracted from the CHM documentation.

### LISP Reference Implementation
The `docs/lisp/` folder contains the original LISP implementation of the construction notes system, which serves as an algorithmic reference for the C# implementation:

#### Files
- **DBRT_UNB_DWG.lsp**: Auto Notes implementation
  - Uses COGO points to define viewport boundaries (naming pattern: `{layoutNum}-{index}`, e.g., "101-1", "101-2")
  - Implements ray-casting algorithm for point-in-polygon detection
  - Filters multileaders by style "Arw-Straight-Hex_Anno_WSDOT"
  - Exports results to CSV files
- **DBRT_UNB_EXCEL.lsp**: Excel Notes implementation
  - Reads construction notes from CSV files at hardcoded paths
  - Uses blocks with "HEX" in naming convention (`{layoutNum}-HEX-NT##`)
  - Updates block visibility states and attributes
- **_DBRT_NOTES_FROM_DWG.scr**: AutoCAD script to load and execute Auto Notes
- **_DBRT_NOTES_FROM_EXCEL.scr**: AutoCAD script to load and execute Excel Notes

#### Key Differences from C# Implementation
| Aspect | LISP Implementation | C# Implementation |
|--------|-------------------|-------------------|
| **Viewport Detection** | Manual COGO points setup | Direct viewport geometry calculation |
| **Block Naming** | `{layoutNum}-HEX-NT##` | `{layoutNum}-NT##` |
| **Multileader Style** | Hardcoded "Arw-Straight-Hex_Anno_WSDOT" | Configurable via ProjectConfig.json |
| **Data Source** | CSV files at fixed paths | Excel files with EPPlus library |
| **Polygon Detection** | Ray-casting with COGO points | Viewport boundary calculation |

### Developer Guidance

#### Requesting Additional API Documentation
When working with AutoCAD API methods or properties not documented in the reference folder:
1. Identify the specific class, method, or property needed
2. Request extraction from the CHM documentation
3. The content will be added as a text file in the appropriate namespace folder

#### Using the LISP Reference
The LISP files demonstrate working algorithms for:
- Point-in-polygon detection using ray-casting
- Dynamic block visibility state manipulation
- Multileader content extraction
- Construction note block attribute updates

While the algorithmic approaches are valuable references, the C# implementation uses more modern .NET patterns and AutoCAD API features for improved maintainability and performance.

#### API Documentation Format
Each API reference file contains:
- Class hierarchy and inheritance
- Method signatures in C# and VB.NET
- Parameter descriptions
- Property accessor information
- Links to related classes and namespaces