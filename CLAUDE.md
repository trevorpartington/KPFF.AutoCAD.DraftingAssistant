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

## Construction Notes Automation System

### Overview
The construction notes automation system provides dual-mode functionality:
- **Auto Notes**: Automatically scans drawing viewports to detect bubble multileaders and blocks and extract note numbers
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
  - Columns: Sheet, File, Title, and title block attributes
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
[Sheets] [Project]
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
Each layout contains 24 dynamic blocks with simplified naming pattern: `NTXX`
- Examples: "NT01", "NT02", "NT24"
- Same blocks used across all layouts with different attribute values
- Dynamic block attributes:
  - **Number**: Note number (1, 2, 4, etc.)
  - **Note**: Full note text from series-specific table
  - **Visibility**: ON/OFF state for displaying/hiding blocks
- Empty blocks start with Visibility=OFF to remain invisible

#### Auto Notes Processing
1. Identify viewports in layout
2. Calculate model space boundaries for each viewport
3. Use ray casting to determine which bubble multileaders and blocks fall within boundaries
4. Filter multileaders and blocks by designated style (from project config)
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
- System updates blocks NT01 and NT02 with Number="1", Note="CONSTRUCT CURB", Visibility="ON"

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
- **ClosedXML**: Excel file reading (handles files open in Excel)
- **AutoCAD .NET API**: Drawing analysis, block manipulation, viewport processing
- **System.Text.Json**: Project configuration serialization

### Configuration Updates (Recent)
The system has been simplified to remove worksheet dependencies:
- Tables are now accessed directly by name across the entire workbook
- Removed WorksheetConfiguration class and worksheet-based lookups
- Configuration now only requires table names: SHEET_INDEX, EXCEL_NOTES, {SERIES}_NOTES
- More robust and works regardless of which worksheet contains the tables


## References
- **API Documentation**: `docs/api-reference/` - AutoCAD .NET API reference by namespace
- **LISP Reference**: `docs/lisp/` - Original LISP implementation algorithms

### Multi-Drawing Support
The construction notes system supports all drawing scenarios:
- **Active**: Drawing is open and currently active in AutoCAD
- **Inactive**: Drawing is open but not the active document  
- **Closed**: Drawing file exists on disk but is not open

Both Auto Notes and Excel Notes work seamlessly across all three drawing states with proper attribute alignment and viewport protection.

#### Auto Notes Drawing State Handling

**Critical Architectural Pattern**: Auto Notes detection requires access to viewports, which are only available in the active drawing context. The system handles drawing states during the Auto Notes collection phase:

**Collection Phase Strategy (BatchOperationService)**:
1. **Group by Drawing**: Group sheets by their drawing file paths
2. **Sequential Processing**: Process each drawing group sequentially  
3. **State Management**: Make inactive drawings active temporarily for Auto Notes detection
4. **Active Drawing Restoration**: Restore original active drawing when complete

**Drawing State Processing**:
- **Active**: Process sheets directly (drawing already active)
- **Inactive**: Make active using `TryMakeDrawingActive()`, then process sheets
- **Closed**: Cannot perform Auto Notes (skip with empty note lists - viewport access impossible)

**Key Classes**:
- **AutoNotesService**: Enhanced with database-specific overloads for explicit drawing context
- **BatchOperationService**: Orchestrates drawing state changes during collection phase
- **DrawingAccessService**: Provides drawing state detection and activation methods

**Updated AutoNotesService API**:
```csharp
// Standard method (uses current active drawing)
public List<int> GetAutoNotesForSheet(string sheetName, ProjectConfiguration config)

// Database-specific method (for explicit drawing context)  
public List<int> GetAutoNotesForSheet(string sheetName, ProjectConfiguration config, Database database, string drawingPath)
```

**Logging Improvements**:
- All Auto Notes operations now log which drawing file is being searched
- Clear distinction between "Layout not found" and "Drawing not accessible"
- Better error messages specify exact drawing context

**Excel Notes Difference**: Excel Notes mode does not require viewport access, so it works with all drawing states without state changes.

## Plotting System

### Overview
The plotting system provides seamless batch plotting capabilities that integrate with the existing construction notes and title block features. Users can plot selected sheets using each sheet's default page setup with optional pre-plot updates.

### Key Features
- **Simple PDF Plotting**: Uses each sheet's existing page setup, outputs to `{sheetname}.pdf`
- **Optional Pre-Plot Updates**: Checkbox options to update construction notes and/or title blocks before plotting
- **Multi-Drawing Support**: Works with active, inactive, and closed drawings
- **Synchronized UI State**: "Apply to current sheet only" checkbox state persists across all tabs
- **Comprehensive Error Handling**: Failed plots are skipped and logged with detailed error information

### File Structure

#### Project Configuration
- **ProjectConfig.json**: Plotting configuration section includes:
  - Output directory path for generated PDFs
  - Default plot format settings
  - Enable/disable plotting for project

### User Interface

#### Plotting Tab Layout
```
[☐ Update Construction Notes] [☐ Update Title Blocks]
[☐ Apply to current sheet only]
[Plot] button
─────────────────────────────────────────────────────
[Results and progress display area]
```

### Technical Implementation

#### Publisher API Architecture
The plotting system uses AutoCAD's Publisher API with Drawing Set Description (DSD) files for batch plotting operations. This approach provides superior performance and reliability compared to individual plot operations.

**Key Components:**
- **Publisher API**: AutoCAD's native batch plotting system
- **DsdEntry**: Represents a single sheet to plot (drawing path, layout name, page setup)
- **DsdData**: Collection of DsdEntry objects with output settings
- **DsdEntryCollection**: Container for multiple sheets to plot together

#### Plot Architecture Decision
- **Batch Operations (Multiple Sheets)**: Uses Publisher API with DsdData for optimal performance
- **Single Sheet Operations**: Uses Publisher API internally for consistency

#### Plot Workflow
1. **Sheet Filtering Phase**:
   - Apply "Apply to current sheet only" filter if enabled
   - Validate selected sheets (layout exists, drawing accessible)

2. **Pre-Plot Phase**:
   - Update construction notes (if checkbox enabled) - uses Construction Notes tab radio button selection (Auto/Excel)
   - Update title blocks (if checkbox enabled)
   - Perform pre-plot updates for all sheets before plotting begins

3. **Plot Phase**:
   - **Batch Mode** (Multiple sheets): Use `PublishSheetsToPdfAsync` with Publisher API
   - **Individual Mode** (Fallback): Use legacy plotting for single sheets or fallback scenarios
   - Honor each layout's saved page setup automatically
   - Skip sheets that fail validation or plotting
   - Track progress and errors in real-time

4. **Post-Plot Phase**:
   - Clean up temporary files
   - Display comprehensive results with success/failure summary
   - Log all errors for troubleshooting

#### Publisher API Benefits
- **No Layout Switching Required**: Eliminates `eInvalidPlotInfo` and `eSetFailed` errors
- **Handles All Drawing States**: Works with open/active, open/inactive, and closed drawings seamlessly
- **Better Performance**: Native batch plotting without UI flicker
- **Automatic Page Setup**: Uses each layout's saved plot settings without override
- **Robust Error Handling**: Built-in validation and error recovery

#### Implementation Details
```csharp
// Build DSD entries for batch plotting
var dsdEntries = new List<DsdEntry>();
foreach (var sheet in sheets)
{
    var entry = new DsdEntry
    {
        DwgName = sheet.DWGFileName,           // Full path to drawing
        Layout = sheet.SheetName,             // Layout name to plot
        Title = sheet.DrawingTitle,           // Display title
        Nps = string.Empty,                   // Use layout's saved page setup
        NpsSourceDwg = string.Empty
    };
    dsdEntries.Add(entry);
}

// Execute batch plot
using (var dsdData = new DsdData())
{
    dsdData.ProjectPath = outputDirectory;
    dsdData.SheetType = SheetType.SinglePdf;  // Individual PDFs
    dsdData.SetDsdEntryCollection(dsdCollection);
    
    Application.Publisher.PublishExecute(dsdData, plotConfig);
}
```

#### Error Handling Strategy
- **Missing Plot Device**: Skip sheet, log error in job summary
- **Drawing Access Issues**: Skip sheet, log specific error
- **Page Setup Problems**: Skip sheet, continue with remaining sheets
- **File Overwriting**: Always overwrite existing PDFs with same name

#### Integration Points
- **Construction Notes**: Checks Construction Notes tab radio button state (Auto Notes vs Excel Notes)
- **Title Blocks**: Uses existing title block update service
- **Shared UI State**: Synchronizes "Apply to current sheet only" across all tabs
- **External Drawing Manager**: Handles closed/inactive drawing operations

### Development Dependencies
- **AutoCAD .NET PlottingServices**: Plot device and validation (`PlotConfig`, `PlotConfigManager`)
- **AutoCAD .NET Publishing**: Publisher API for batch operations (`Publisher`, `DsdData`, `DsdEntry`)
- **Existing Services**: Integrates with `ConstructionNotesService` and title block operations
- **Project Configuration**: Extends existing configuration system with plotting settings

### Current Status
- ✅ Publisher API implementation completed
- ✅ Batch plotting with DSD entries
- ✅ Pre-plot operations (Construction Notes and Title Blocks)
- ✅ "Apply to current sheet only" functionality
- ✅ Multi-drawing support (open/active, open/inactive, closed)
- ✅ Error handling and progress reporting
- ✅ Backward compatibility with existing interfaces

## Multi-Drawing Operations Architecture

### Overview
The system uses a three-layer architecture for handling operations across multiple drawings regardless of their state (Active, Inactive, Closed):

#### 1. **BatchOperationService** - Central Orchestrator
- **Purpose**: Main coordinator for all multi-drawing operations
- **Location**: `Core/Services/BatchOperationService.cs`
- **Responsibilities**:
  - Coordinates Construction Notes, Title Blocks, and Plotting operations
  - Manages phased execution with progress reporting (0-100%)
  - Handles validation, error handling, and result aggregation
  - Provides async operation management to prevent UI freezing
  - Supports "Apply to current sheet only" functionality

#### 2. **Multi-Drawing Services** - Specialized Workers
- **Purpose**: Handle specific operation types across drawing states
- **Services**:
  - `MultiDrawingConstructionNotesService`: Auto Notes and Excel Notes
  - `MultiDrawingTitleBlockService`: Title block attribute updates
  - `PlottingService`: PDF generation with pre-plot operations
- **Responsibilities**:
  - Drawing state detection and routing (Active/Inactive/Closed)
  - External drawing operations via ExternalDrawingManager
  - Transaction management and error recovery

#### 3. **Drawing Access Layer** - State Management
- **Services**:
  - `DrawingAccessService`: Drawing state detection and access
  - `ExternalDrawingManager`: Closed drawing operations
  - Various analyzers and utilities for each drawing state

### UI Integration Requirements

#### ✅ **CORRECT PATTERN**: Use BatchOperationService
All UI controls should use BatchOperationService for multi-drawing operations:

```csharp
// ✅ CORRECT - Use BatchOperationService
var batchSettings = new BatchOperationSettings
{
    UpdateConstructionNotes = true,
    IsAutoNotesMode = true,
    ApplyToCurrentSheetOnly = applyToCurrentOnly
};

var progress = new Progress<BatchOperationProgress>(p => 
{
    UpdateStatus($"[{p.OverallProgress}%] {p.CurrentOperation}: {p.StatusMessage}");
});

var result = await _batchOperationService.UpdateConstructionNotesAsync(
    sheetNames, config, applyToCurrentOnly, progress);
```

#### ❌ **INCORRECT PATTERN**: Direct Multi-Drawing Service Usage
Do NOT instantiate multi-drawing services directly in UI controls:

```csharp
// ❌ WRONG - Causes UI freezing and duplicate code
var multiDrawingService = new MultiDrawingConstructionNotesService(...);
var result = await multiDrawingService.UpdateConstructionNotesAcrossDrawingsAsync(...);
```

### Implementation Status

#### ✅ Correctly Implemented
- **PlottingControl**: Uses PlottingService which integrates with BatchOperationService
- **ConstructionNoteControl**: Refactored to use BatchOperationService
- **TitleBlockControl**: Refactored to use BatchOperationService

#### Architecture Benefits
- **No UI Freezing**: Proper async operation management
- **Unified Progress Reporting**: Consistent 0-100% progress with status messages  
- **Centralized Error Handling**: Single point for validation and error management
- **Reduced Code Duplication**: Eliminates redundant service instantiations
- **Better Performance**: Optimized drawing state handling and batch processing

### Key Principles
1. **UI Controls**: Should never directly instantiate multi-drawing services
2. **BatchOperationService**: Always use for multi-sheet operations
3. **Progress Reporting**: Use provided progress callbacks for user feedback
4. **Error Handling**: Let BatchOperationService handle validation and errors
5. **Async Operations**: All multi-drawing operations must be properly async

This architecture ensures consistent behavior, prevents UI freezing, and provides proper error handling across all multi-drawing operations in the system.