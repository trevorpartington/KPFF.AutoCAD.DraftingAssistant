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


## References
- **API Documentation**: `docs/api-reference/` - AutoCAD .NET API reference by namespace
- **LISP Reference**: `docs/lisp/` - Original LISP implementation algorithms

## Multi-Scenario Drawing Update System 🚧 **IN PROGRESS**

### Overview
The construction notes system is being extended to handle three drawing scenarios seamlessly:

#### Drawing States
1. **Active**: Drawing is open and currently active in AutoCAD
2. **Inactive**: Drawing is open but not the active document  
3. **Closed**: Drawing file exists on disk but is not open

### Implementation Plan

#### ✅ Task 1: Drawing Access Service (COMPLETE)
**File**: `DrawingAccessService.cs`
- **DrawingState Detection**: Enumerates open documents, detects Active/Inactive/Closed states
- **Document Access**: Provides unified interface for accessing drawings in different states
- **File Path Resolution**: Maps sheet names to DWG file paths using Excel SHEET_INDEX table
- **Path Normalization**: Handles different path formats and case sensitivity

**Key Methods**:
- `GetDrawingState(dwgFilePath)`: Returns DrawingState enum (Active/Inactive/Closed/NotFound/Error)
- `GetOpenDocument(dwgFilePath)`: Gets Document object for open drawings
- `GetAllOpenDrawings()`: Lists all open drawings with their states
- `TryMakeDrawingActive(dwgFilePath)`: Attempts to activate inactive drawings
- `GetDrawingFilePath(sheetName, config, sheetInfos)`: Resolves sheet name to full file path

#### 📋 Remaining Tasks

##### Task 2: External Drawing Manager
**New File**: `ExternalDrawingManager.cs`
- Handle closed drawings using `Database.ReadDwgFile()`
- Implement side database operations
- Apply proven Justify + AdjustAlignment approach for attribute centering

##### Task 3: Extend CurrentDrawingBlockManager ⚠️ **LESSONS LEARNED**
**Update**: `CurrentDrawingBlockManager.cs`
- **ISSUE**: Aggressive refactoring broke working TESTCLOSEDUPDATE functionality
- **ROOT CAUSE**: External drawing operations have subtle requirements that differ from current drawing operations
- **LESSON**: Preserve working external drawing code even if it means some code duplication
- **APPROACH**: Conservative extension - add external database support without changing existing methods
- **PRIORITY**: Working functionality > Perfect SOLID principles for complex AutoCAD operations

##### Task 4: Multi-Drawing Operations Service
**New File**: `MultiDrawingOperations.cs`
- Orchestrate updates across multiple drawings
- Route to appropriate manager based on drawing state
- Batch process sheet updates efficiently

##### Task 5: Update DrawingOperations Service
**Update**: `DrawingOperations.cs`
- Integrate with DrawingAccessService
- Determine drawing state before operations
- Route to appropriate handler (Current/External)

##### Task 6: Construction Notes Service Integration
**Update**: `ConstructionNotesService.cs`
- Use SheetInfo.DWGFileName to locate drawing files
- Combine with projectDWGFilePath from config
- Handle file not found scenarios

##### Task 7: Testing & Validation
- Create test commands for each scenario
- Test ATTSYNC behavior on dynamic blocks
- Verify attribute centering works correctly
- Test batch updates on multiple sheets

### Technical Implementation Details

#### Drawing State Detection Pattern
```csharp
enum DrawingState { Active, Inactive, Closed, NotFound, Error }
- Active: doc == Application.DocumentManager.MdiActiveDocument
- Inactive: Found in DocumentManager but not active
- Closed: File exists but not in DocumentManager
```

#### External Drawing Pattern
```csharp
using (var db = new Database(false, true))
{
    db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, null);
    using (var tr = db.TransactionManager.StartTransaction())
    {
        // Operations on external drawing
    }
    db.SaveAs(filePath, DwgVersion.Current);
}
```

#### Attribute Positioning Solution ✅ **COMPLETE**
**Critical Discovery**: `AdjustAlignment()` requires `HostApplicationServices.WorkingDatabase` to match the target database for proper alignment calculations.

**Working Solution**: Minimal-scope WorkingDatabase switching
```csharp
// CRITICAL: Minimal WorkingDatabase switch ONLY for AdjustAlignment
var originalWdb = HostApplicationServices.WorkingDatabase;
try
{
    HostApplicationServices.WorkingDatabase = db;
    attRef.AdjustAlignment(db);
}
finally
{
    HostApplicationServices.WorkingDatabase = originalWdb;
}
```

**Key Features**:
- **Precise Timing**: WorkingDatabase switched only during `AdjustAlignment()` call
- **Safety Guaranteed**: `try/finally` pattern ensures database restoration 
- **Viewport Protection**: Minimal scope prevents the database-wide corruption
- **Perfect Alignment**: Attributes center automatically without manual ATTSYNC
- **Production Tested**: Confirmed working with TESTCLOSEDUPDATE ✅

### Implementation Status & Next Steps

#### ✅ Task 1 Complete: DrawingAccessService
**File**: `DrawingAccessService.cs` - **COMPLETED & TESTED**
- **Drawing State Detection**: Successfully detects Active/Inactive/Closed states ✅
- **Document Management**: Get Document objects, make drawings active ✅
- **File Path Resolution**: Maps sheet names to DWG paths using Excel SHEET_INDEX ✅
- **Test Commands**: `TESTDRAWINGLIST`, `TESTDRAWINGSTATE` working perfectly ✅

#### ✅ Task 2 Complete: External Drawing Manager  
**File**: `ExternalDrawingManager.cs` - **COMPLETED & PRODUCTION READY** ✅
- **Closed Drawing Operations**: Successfully handles external DWG files ✅
- **Viewport Protection**: Multi-layer protection prevents viewport corruption ✅
- **Attribute Alignment**: Working minimal-scope WorkingDatabase switching ✅
- **Comprehensive Logging**: Debug messages for troubleshooting ✅
- **Test Commands**: `TESTCLOSEDUPDATE` working perfectly ✅
- **Commit**: `992f75a` - Fix attribute alignment with minimal-scope WorkingDatabase switching

**Production Features**:
- **Safe File Operations**: Temporary file strategy prevents corruption
- **Robust Error Handling**: Graceful handling of all edge cases
- **NT Block Targeting**: Precise filtering to only affect construction note blocks
- **Dynamic Block Support**: Proper handling of dynamic block properties
- **Transaction Safety**: Proper commit/rollback for all database operations

#### 🚧 Next Implementation Phase: Multi-Drawing Operations

##### Task 3: Enhanced CurrentDrawingBlockManager
**Update**: `CurrentDrawingBlockManager.cs` - **NEXT PRIORITY**
- Add constructor overload: `CurrentDrawingBlockManager(Database externalDatabase, ILogger logger)`
- Support operations on external databases (not just current drawing)  
- Maintain same interface, different data source

##### Task 4: Multi-Drawing Construction Notes Service
**New File**: `MultiDrawingConstructionNotesService.cs`
```csharp
public class MultiDrawingConstructionNotesService
{
    public async Task UpdateConstructionNotesAcrossDrawings(
        Dictionary<string, List<int>> sheetToNotes, // Sheet -> Note numbers
        ProjectConfiguration config)
    {
        foreach (var (sheetName, noteNumbers) in sheetToNotes)
        {
            var dwgPath = drawingAccessService.GetDrawingFilePath(sheetName, config, sheetInfos);
            var state = drawingAccessService.GetDrawingState(dwgPath);
            
            switch (state)
            {
                case DrawingState.Active:
                    // Use current DrawingOperations
                    await drawingOperations.UpdateConstructionNoteBlocksAsync(...);
                    break;
                    
                case DrawingState.Inactive:
                    // Make active, then update
                    drawingAccessService.TryMakeDrawingActive(dwgPath);
                    await drawingOperations.UpdateConstructionNoteBlocksAsync(...);
                    break;
                    
                case DrawingState.Closed:
                    // Use ExternalDrawingManager
                    externalDrawingManager.UpdateClosedDrawing(dwgPath, (db, tr) => {
                        var blockManager = new CurrentDrawingBlockManager(db, logger);
                        // Update blocks using external database
                    });
                    break;
            }
        }
    }
}
```

##### Task 5: UI Integration
**Update**: Construction Notes UI
- **Batch Mode Selection**: Single sheet vs. Multiple sheets
- **Drawing State Display**: Show which drawings are Open/Closed
- **Progress Tracking**: Real-time progress for multi-drawing operations
- **Error Reporting**: Clear feedback for failed operations

### Technical Implementation Details

#### Drawing Access Patterns (IMPLEMENTED ✅)
```csharp
enum DrawingState { Active, Inactive, Closed, NotFound, Error }
- Active: Successfully tested with TESTDRAWINGLIST ✅
- Inactive: Successfully tested with TESTDRAWINGSTATE, can make active ✅  
- Closed: Successfully detected, ready for external processing ✅
```

#### External Drawing Processing Pattern (NEXT)
```csharp
using (var db = new Database(false, true))
{
    db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, null);
    using (var tr = db.TransactionManager.StartTransaction())
    {
        // 1. Reset/update construction note blocks
        // 2. Apply Justify + AdjustAlignment for proper attribute positioning
        tr.Commit();
    }
    db.SaveAs(filePath, DwgVersion.Current);
}
```

#### Attribute Alignment Implementation ✅
- **Working Approach**: Uses Justify + AdjustAlignment pattern
- **Validation**: TESTCLOSEDUPDATE confirms proper attribute centering  
- **Benefits**: Reliable positioning without external dependencies

### Benefits  
- **Seamless Multi-Drawing Updates**: No manual opening required ✅ (Framework ready)
- **Batch Processing Capability**: Update entire projects efficiently (Next phase)  
- **Maintains Attribute Formatting**: ATTSYNC-like functionality implemented ✅
- **Works with Existing UI**: Integrates with current construction notes workflow ✅
- **Handles Edge Cases**: Robust error handling and state detection ✅

### Current Test Commands Status
- `TESTDRAWINGLIST` ✅ **Working** - Lists all open drawings with states
- `TESTDRAWINGSTATE` ✅ **Working** - Interactive file testing, can make drawings active  
- `TESTCLOSEDUPDATE` ⚠️ **Broken after refactoring** - External drawing operations are sensitive to implementation details
- `TESTDRAWINGACCESS` ❌ **Crashes AutoCAD** - Needs investigation (likely Excel/async issue)

### Key Learning: External Drawing Operations
- **Complexity**: External drawing updates involve WorkingDatabase, file I/O, transaction management, and subtle AutoCAD behaviors
- **Fragility**: Small changes can break working functionality in unpredictable ways  
- **Strategy**: Preserve working external drawing code, extend carefully with minimal changes
- **Testing**: TESTCLOSEDUPDATE is the critical validation - must work before any changes are considered complete

## Current System Status ✅ **PRODUCTION READY**

### ✅ **Completed Systems**
- **Excel Notes**: Complete and operational with TESTCLOSEDUPDATE ✅
- **Auto Notes Backend**: Complete with TESTAUTONOTES command ✅
- **External Drawing Operations**: Production-ready with proper attribute alignment ✅
- **Viewport Protection**: Multi-layer protection prevents all corruption ✅
- **Drawing Access Framework**: Supports Active/Inactive/Closed drawing states ✅

### **Current Test Commands Status**
- `TESTCLOSEDUPDATE` ✅ **Working** - External drawing updates with perfect alignment
- `TESTDRAWINGLIST` ✅ **Working** - Lists all open drawings with states  
- `TESTDRAWINGSTATE` ✅ **Working** - Interactive file testing, can make drawings active
- `TESTAUTONOTES` ✅ **Working** - Auto Notes backend validation 
- `TESTBLOCKDISCOVERY` ✅ **Working** - Block analysis and viewport protection diagnostics

### **Critical Success: Attribute Alignment Solution**
**Problem Solved**: External drawing operations now produce perfectly centered attributes without manual ATTSYNC
**Root Cause**: `AdjustAlignment()` requires `WorkingDatabase` to match the target database
**Solution**: Minimal-scope WorkingDatabase switching during alignment operations only
**Result**: Production-ready external drawing updates with zero viewport corruption

### **Ready for Next Phase**
- **Foundation Complete**: All core external drawing operations working ✅
- **Safety Proven**: Extensive testing confirms viewport protection ✅ 
- **Architecture Sound**: Clean separation between drawing states and operations ✅
- **Next**: Multi-drawing batch operations for enhanced productivity