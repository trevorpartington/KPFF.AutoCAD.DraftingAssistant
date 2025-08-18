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

### Excel Notes Implementation Strategy ✅ **COMPLETE**

The Excel Notes functionality has been successfully implemented with a **resilient, production-ready approach** that handles real-world scenarios:

#### Key Implementation Features ✅
- **Partial Block Set Support**: Gracefully handles layouts with fewer than 24 construction note blocks
- **Smart Reset Logic**: Only resets blocks that actually exist in the drawing
- **Comprehensive Error Handling**: Clear error messages and graceful degradation
- **Real-time Logging**: AutoCAD command line shows detailed progress and results
- **Dynamic Block Handling**: Correctly accesses dynamic block properties using `DynamicBlockTableRecord`
- **Production Testing**: Successfully tested with 2-block layouts, updating ABC-101 (1 note) and ABC-102 (2 notes)

### Implementation Status Summary

#### ✅ **COMPLETED PHASES**
- **Phase 1**: Read-only block discovery with safe AutoCAD object access
- **Phase 2**: Single block update with proper transaction handling
- **Phase 3**: Complete Excel Notes integration with ClosedXML and resilient block handling

#### 🔄 **IN PROGRESS**
- Enhanced UI for sheet selection and user experience improvements

#### 📋 **PLANNED FUTURE ENHANCEMENTS**
- **Auto Notes**: Automatic detection from drawing viewports
- **Phase 4**: Closed drawing operations for batch processing
- **Advanced Features**: Template-based note management, automated block creation

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

#### Phase 1: Read-Only Block Discovery ✅ **COMPLETE**
- **Goal**: Safely find and read construction note blocks without modifications
- **Implementation**: `CurrentDrawingBlockManager` with read-only operations
- **Key Features**:
  - Pattern matching for NT## blocks (NT01, NT02, etc.)
  - Safe AutoCAD object access with extensive error handling
  - Block attribute reading (Number, Note)
  - Dynamic block visibility state detection
- **Status**: ✅ **Working** - All test commands operational

#### Phase 2: Single Block Update ✅ **COMPLETE**
- **Goal**: Safely modify one construction note block in active drawing
- **Implementation**: `UpdateConstructionNoteBlock()` method updates:
  - NUMBER and NOTE attributes
  - Dynamic block visibility state (OFF → ON)
  - Proper transaction handling with rollback on failure
- **Status**: ✅ **Working** - Successfully updates individual blocks

#### Phase 3: Excel Notes Integration ✅ **COMPLETE**
- **Goal**: Complete Excel Notes functionality with real-time Excel reading
- **Implementation**: Full pipeline from Excel to AutoCAD block updates
- **Key Features**:
  - ClosedXML-based Excel file reading (SHEET_INDEX, EXCEL_NOTES, {SERIES}_NOTES tables)
  - AutoCADLogger for real-time command line output during operations
  - Resilient block handling for partial block sets (handles 2 blocks vs 24)
  - Comprehensive logging and error reporting throughout the pipeline
  - Smart reset logic that only resets blocks that actually exist
  - Graceful degradation: continues with partial updates rather than failing
- **Status**: ✅ **Working** - Successfully processes sheets and updates construction note blocks

#### Remaining Implementation Tasks

##### Next Priority: Enhanced UI and Sheet Selection
- **Expose selected sheets property from ConfigurationControl**
  - Allow users to select specific sheets instead of processing all sheets
  - Integrate with Excel Notes processing to only update selected sheets
- **Default all sheets to selected for testing purposes**
  - Improve user experience during testing and validation

##### Auto Notes Implementation (Future)
- **Goal**: Automatically detect construction notes from drawing viewports
- **Implementation Required**:
  - Viewport boundary calculation and model space mapping
  - Multileader detection and filtering by style (configurable)
  - Ray-casting or geometric containment for note extraction
  - Integration with existing construction note block update pipeline
- **Reference**: LISP implementation in `docs/lisp/DBRT_UNB_DWG.lsp` provides algorithmic guidance

##### Future Enhancements
- **Template-based note management system**
- **Enhanced viewport geometry support for complex layouts**
- **Automated construction note block creation for new sheets**

## Excel Integration with ClosedXML

### Overview
Phase 3 implements comprehensive Excel integration using ClosedXML 0.104.0 for real-time reading of construction notes and sheet mappings. The implementation provides robust error handling and async processing to maintain AutoCAD stability.

### Technology Stack
- **ClosedXML 0.104.0**: Excel file reading and table processing (latest stable version)
- **.NET 8.0-windows**: Target framework for AutoCAD 2025 compatibility
- **Task.Run() Pattern**: Background thread processing to avoid AutoCAD UI freezing
- **Read-only Access**: Files opened in read-only mode to prevent conflicts

### Implementation Architecture

#### ExcelReaderService
- **Location**: `Core/Services/ExcelReaderService.cs`
- **Interface**: Implements `IExcelReader` with async methods
- **Dependency**: Registered as transient service in DI container
- **Thread Safety**: All Excel operations wrapped in `Task.Run()` for background execution

#### Excel Table Structure
The implementation expects specific named tables in the Excel workbook:

1. **SHEET_INDEX**: Master sheet listing
   - Columns: Sheet, File, Title
   - Purpose: Complete drawing inventory with metadata

2. **{SERIES}_NOTES**: Construction note definitions per series
   - Pattern: `ABC_NOTES`, `PV_NOTES`, `C_NOTES` (configurable via `{0}_NOTES`)
   - Columns: Number, Note
   - Purpose: Note number to text mapping for each drawing series

3. **EXCEL_NOTES**: Manual sheet-to-notes mapping
   - Column 1: Sheet Name (e.g., "ABC-101")
   - Columns 2-25: Note numbers (max 24 per `maxNotesPerSheet` config)
   - Purpose: Defines which notes appear on each sheet

#### Data Validation
- **Column Count**: EXCEL_NOTES limited to 25 columns (1 sheet + 24 notes max)
- **Note Range**: Note numbers validated within 1-24 range
- **Duplicate Consolidation**: Duplicate note numbers automatically removed (e.g., [4, 4, 7] → [4, 7])
- **Format Validation**: Number parsing with descriptive error messages for invalid data

#### Error Handling Strategy
- **Descriptive Errors**: Specific messages for missing tables, invalid formats, locked files
- **Graceful Degradation**: Returns empty collections on failure (never null)
- **No Retry Logic**: Single attempt with clear error reporting to user
- **Comprehensive Logging**: All operations logged via `IApplicationLogger`

#### Examples of Error Messages
- `"Table 'SHEET_INDEX' not found in workbook"`
- `"EXCEL_NOTES table has 30 columns but max is 25 (1 sheet + 24 notes)"`
- `"Unable to open Excel file - may be locked by another process"`
- `"Unable to parse note number from column 3: 'ABC' for sheet ABC-101"`

### Async Processing Pattern
```csharp
public async Task<List<SheetInfo>> ReadSheetIndexAsync(string filePath, ProjectConfiguration config)
{
    return await Task.Run(() =>
    {
        using var workbook = new XLWorkbook(filePath);
        // Excel processing logic here
    });
}
```

### Benefits Over Previous Approaches
- **Real-time Updates**: Direct Excel file access when user clicks "Update Notes"
- **No File Dependencies**: Eliminates need for intermediate CSV/JSON exports
- **Rich Validation**: Comprehensive data validation with helpful error messages
- **AutoCAD Stable**: Background processing prevents UI freezing
- **Memory Efficient**: Files opened read-only with minimal tracking

### Testing Strategy
- **Unit Tests**: Mock `IExcelReader` for fast, reliable testing of plugin logic
- **Integration Tests**: Real Excel file testing with `testdata/ProjectIndex.xlsx`
- **Manual Testing**: TestConsole application validates full Excel reading pipeline
- **Error Scenario Testing**: Corrupted files, missing tables, invalid data formats

### Performance Considerations
- **First Access Penalty**: ClosedXML builds DOM on first read (can be slow for large files)
- **Memory Usage**: Entire workbook loaded into memory (not streaming)
- **Optimization**: Files opened in read-only mode for faster processing
- **Recommendation**: If files exceed 50MB, consider hybrid approach with ExcelDataReader for streaming

### Future Considerations
If AutoCAD freezing occurs during debugging (due to COM/OLE conflicts), the architecture supports migration to out-of-process Excel reading while maintaining the same `IExcelReader` interface.

#### Phase 4: Closed Drawing Operations (Future Enhancement)
**Current Status**: Excel Notes currently works with active/open drawings only

**Future Production Enhancement**: 
- **Side Database Pattern**: Use `Database.ReadDwgFile()` to access closed drawings
- **External File Access**: Open external DWG files in memory without displaying them
- **Batch Processing**: Update multiple drawings (both open and closed) in sequence  
- **File Management**: Safe opening, updating, and closing of external files
- **Advanced Error Handling**: File locking, permission issues, corrupted files

**Note**: This would support the workflow: "select any number of sheets and make updates to both drawings they have opened on their computer and drawings that are closed". Currently, users can update sheets in the active drawing, which covers the primary use case.

### Implementation Architecture

#### Simplified Block Architecture
- Block names: `NT01`, `NT02`, ... `NT24`
- Same blocks used across all layouts with different attribute values
- Result: Only 24 total block definitions (vs N layouts × 24 blocks)

#### Critical Safety Measures
**AutoCAD Crash Prevention:**
1. **Deferred Initialization**: No AutoCAD object access during plugin load
2. **Safe Object Access**: All DocumentManager access wrapped in try-catch
3. **Lazy Service Creation**: CurrentDrawingBlockManager not in DI container
4. **Defensive Programming**: Null checks and graceful fallbacks throughout

**Transaction Safety:**
- Read-only transactions for discovery operations
- Proper transaction disposal and error handling
- No premature commits that could corrupt drawings

**Drawing State Protection:**
- Never access viewports or other entities accidentally
- Only target construction note blocks with precise pattern matching
- Extensive logging for debugging and audit trails

#### Excel Integration Data Flow ✅ **IMPLEMENTED**
1. **File Validation**: Check Excel file exists and is accessible ✅
2. **Table Discovery**: Find named tables (SHEET_INDEX, EXCEL_NOTES, {SERIES}_NOTES) across all worksheets ✅
3. **Data Validation**: Validate table structure, column counts, and data formats ✅
4. **Excel Reading**: ClosedXML reads table data with comprehensive error handling ✅
5. **Data Processing**: Parse note numbers, consolidate duplicates, validate ranges ✅
6. **Data Mapping**: Map note numbers to series-specific text from {SERIES}_NOTES tables ✅
7. **Block Discovery**: Discover which NT## blocks actually exist in each layout ✅
8. **Block Reset**: Reset only existing blocks (clear attributes, set visibility OFF) ✅
9. **Block Updates**: Update available NT## blocks with mapped data, handling partial block sets gracefully ✅
10. **Result Validation**: Verify all updates completed successfully with detailed logging ✅

**Resilient Block Handling**: The system now handles layouts with any number of construction note blocks (e.g., 2 blocks instead of 24) and provides clear warnings when notes cannot be placed due to missing blocks.

#### Testing Commands ✅ **ALL WORKING**
- **TESTPHASE1**: Read-only block discovery test ✅
- **TESTPHASE1DETAIL**: Detailed block discovery with attribute display ✅
- **TESTPHASE2**: Single block update test ✅
- **CLEARNOTES**: Clear all construction note blocks in all layouts ✅
- **TESTLAYOUTS**: Layout enumeration test ✅

#### Production Commands ✅ **WORKING**
- **DRAFTINGASSISTANT**: Launch the main UI palette ✅
- **Excel Notes Tab**: Complete workflow from Excel reading to block updates ✅

**System Status**: All critical functionality operational. The system ensures **reliable, crash-free operation** while maintaining **clean, maintainable code** that follows AutoCAD .NET API best practices.

## Remaining Development Tasks

### Immediate Priorities
1. **Enhanced Sheet Selection UI**
   - Expose selected sheets property from ConfigurationControl
   - Allow users to select specific sheets instead of processing all sheets automatically
   - Default all sheets to selected for improved testing experience

2. **Auto Notes Implementation**
   - Implement viewport boundary calculation
   - Add multileader detection and filtering
   - Integrate with existing block update pipeline
   - Reference LISP algorithms in `docs/lisp/` for implementation guidance

### Future Enhancements
- Closed drawing operations (Phase 4) for batch processing
- Template-based construction note management
- Enhanced viewport geometry support
- Automated block creation for new sheets

### Current System Capabilities ✅
- **Excel Notes**: Complete end-to-end functionality working
- **Configuration Management**: Project loading and sheet discovery
- **Block Operations**: Safe reading and updating of construction note blocks
- **Error Handling**: Comprehensive logging and graceful failure handling
- **Partial Block Support**: Works with any number of available construction note blocks
- **Real-time Feedback**: AutoCAD command line shows detailed operation progress