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

### AutoCAD .NET API Reference
The `docs/api-reference/` folder contains extracted AutoCAD API documentation organized by namespace. Key classes include BlockReference, Layout, MLeader, Viewport, and Transaction management.

### LISP Reference Implementation  
The `docs/lisp/` folder contains the original LISP implementation with working algorithms for ray-casting point-in-polygon detection and construction note block updates.

### Completed Implementation Summary
- **Phase 1-3**: Block discovery, updates, and Excel Notes integration complete ✅
- **Current Architecture**: NT01-NT24 blocks with dynamic attributes and visibility
- **Excel Integration**: ClosedXML-based reading with comprehensive error handling  
- **Testing**: All test commands operational (TESTPHASE1, TESTPHASE2, CLEARNOTES, etc.)

### Development Architecture
- **Block Names**: `NT01`, `NT02`, ... `NT24` (simplified pattern)
- **Safety Measures**: Deferred initialization, safe object access, transaction handling
- **Excel Processing**: Async background processing with Task.Run() pattern

### Current System Status ✅
**Production Commands**: DRAFTINGASSISTANT (UI), Excel Notes workflow complete, TESTAUTONOTES validation  
**System Capabilities**: Configuration management, block operations, Auto Notes backend, comprehensive error handling  
**Completed Features**: Phase 5A (viewport boundaries), Phase 5B (multileader analysis), Excel Notes integration  
**Next Priority**: Phase 5C (Auto Notes UI integration), enhanced sheet selection UI

## Auto Notes Implementation Plan with GeometryExtensions

### Overview
Auto Notes automatically detects construction notes from drawing viewports by analyzing bubble multileaders within viewport boundaries. This eliminates the manual Excel-based approach and provides real-time note detection from the actual drawing geometry.

### Core Technology: GeometryExtensions Library
We're using Gilles Chanteau's **GeometryExtensions** library (https://github.com/gileCAD/GeometryExtensions) which provides critical viewport transformation methods:
- **`viewport.DCS2WCS(Point3d)`**: Transform points from Display Coordinate System to World Coordinate System
- **Proven, tested code**: Used by many AutoCAD developers
- **Handles all complexity**: Rotation, scale, view direction, coordinate systems

### Requirements & Specifications

#### Viewport Support
- **All Viewport Types**: Rectangular, rotated, and polygonal viewports
- **Boundary Calculation**: Use GeometryExtensions for accurate model space boundaries
- **No Manual Setup**: Eliminate LISP requirement for manual COGO point placement

#### Multileader Analysis
- **Style Filtering**: Filter multileaders by style specified in `config.ConstructionNotes.MultileaderStyleName`
- **Block Content**: Extract note numbers from block attributes within multileaders
- **Attribute Extraction**: Get `TAGNUMBER` attribute from multileader blocks (e.g., `_TagCircle`)
- **Integer Validation**: Only accept integer note numbers, skip non-integer content
- **Duplicate Handling**: Consolidate duplicate note numbers (e.g., two instances of note 1 → single note 1)

#### Point-in-Polygon Detection
- **Ray Casting Algorithm**: Port the proven ray casting algorithm from LISP implementation
- **Geometric Containment**: Determine which multileaders fall within viewport boundaries
- **Robust Handling**: Handle edge cases and boundary conditions

#### Integration Requirements
- **Block Reset**: Clear all construction note blocks before updating (set Visibility=OFF)
- **Existing Pipeline**: Use existing construction note block update infrastructure
- **Error Handling**: Empty viewports are normal - just clear blocks and continue
- **Logging**: Comprehensive logging for debugging and audit trails

### Test Data & Setup

#### Test Drawing: PROJ-ABC-100.dwg
- **Multiple Layouts**: ABC-101, ABC-102, ABC-103 with viewports containing multileaders
- **Multileader Style**: `ML-STYLE-01` with `Circle` source block
- **Block Structure**: Block name `_TagCircle` with `TAGNUMBER` attribute
- **Duplicate Test Case**: One layout has two multileaders both containing note number 1
- **Expected Behavior**: System should detect duplicates and only place note 1 once

#### Multileader Properties (from test data)
- **Style**: ML-STYLE-01
- **Source Block**: Circle
- **Block Name**: _TagCircle (when exploded)
- **Attribute**: TAGNUMBER with integer values
- **Content Access**: Available via Edit Attributes dialog or programmatically

### Implementation Plan

#### Step 1: Setup GeometryExtensions
- Add GeometryExtensions NuGet package to Core project
- Create `TESTVIEWPORT` command to validate transformations work correctly

#### Step 2: Core Utilities
- **ViewportBoundaryCalculator** (`Core/Utilities/`)
  - `GetViewportFootprint(viewport)` using `viewport.DCS2WCS()`
  - Handle rectangular and polygonal viewports
- **PointInPolygonDetector** (`Core/Utilities/`)
  - Ray casting algorithm for containment testing

#### Step 3: Multileader Analysis
- **MultileaderAnalyzer** (`Core/Services/`)
  - Find all multileaders in model space
  - Filter by style (ML-STYLE-01 from config)
  - Extract TAGNUMBER attribute values

#### Step 4: Auto Notes Service
- **AutoNotesService** (`Core/Services/`)
  - Orchestrate viewport analysis and multileader detection
  - Map multileaders to viewports using point-in-polygon
  - Return consolidated note numbers

#### Step 5: Integration
- Update `ConstructionNotesService.GetAutoNotesForSheetAsync()`
- Wire into existing UI and block update pipeline
- Register in DI container

#### Step 6: Testing & Validation
- Test with PROJ-ABC-100.dwg layouts
- Verify all viewport types work correctly
- Confirm duplicate handling and note extraction

### Sample Implementation

#### ViewportBoundaryCalculator.cs
```csharp
public static Point3dCollection GetViewportFootprint(Viewport vp)
{
    var result = new Point3dCollection();
    
    if (!vp.NonRectClipOn)
    {
        // Rectangular boundary
        double hw = vp.Width / 2.0;
        double hh = vp.Height / 2.0;
        
        var dcsCorners = new[]
        {
            new Point3d(-hw, -hh, 0),
            new Point3d(-hw,  hh, 0),
            new Point3d( hw,  hh, 0),
            new Point3d( hw, -hh, 0)
        };
        
        foreach (var pt in dcsCorners)
            result.Add(vp.DCS2WCS(pt)); // GeometryExtensions!
    }
    else
    {
        // Polygonal boundary
        var clipPts = vp.GetNonRectClipBoundary();
        if (clipPts != null && clipPts.Count > 0)
        {
            foreach (Point2d p in clipPts)
            {
                var dcsPt = new Point3d(p.X - vp.CenterPoint.X, p.Y - vp.CenterPoint.Y, 0);
                result.Add(vp.DCS2WCS(dcsPt));
            }
        }
    }
    
    return result;
}
```

### Benefits Over LISP Implementation
- **No Manual Setup**: Eliminates manual COGO point placement
- **Proven Transformations**: Uses GeometryExtensions' tested coordinate math
- **Better Error Handling**: Robust exception handling and logging  
- **Configurable**: Uses project configuration for multileader style
- **Integrated**: Seamlessly works with existing construction notes system
- **Maintainable**: Clean C# code with proper separation of concerns
- **Testable**: Each component can be unit tested independently

### Benefits Over Manual Implementation
- **No Coordinate Math Bugs**: GeometryExtensions handles all transformation complexity
- **Faster Development**: Proven library eliminates trial-and-error coordinate calculations
- **All Viewport Types**: Support for rectangular, rotated, and polygonal viewports
- **Production Ready**: Library used by many AutoCAD developers in production

## Auto Notes Phase 5A Implementation ✅ **COMPLETE**

### ViewportBoundaryCalculator Success
The viewport boundary calculation has been successfully implemented and tested with the `TESTVIEWPORT` command:

#### Key Implementation Details ✅
- **GeometryExtensions Integration**: Successfully integrated `Gile.AutoCAD.R25.Geometry` namespace
- **Manual Scale Calculation**: Uses `ViewCenter` property for accurate model space center detection
- **Robust Error Handling**: Graceful fallback to empty results on transformation errors
- **Direct Coordinate Math**: Applied `1.0 / viewport.CustomScale` for scale factor calculation

#### Test Results Validation ✅
**ABC-101 Layout**: 
- ViewCenter: (80, 50), Scale: 1:20 (0.05), Size: 160×100 paper units
- **Calculated Area**: (0,0) to (160,100) model space ✅
- Formula: Center ± (PaperSize × ScaleFactor) / 2

**ABC-102 Layout**:
- ViewCenter: (240, 50), Scale: 1:20 (0.05), Size: 160×100 paper units  
- **Calculated Area**: (160,0) to (320,100) model space ✅
- Formula: Center ± (PaperSize × ScaleFactor) / 2

#### Core Implementation (ViewportBoundaryCalculator.cs)
```csharp
double scaleFactor = 1.0 / viewport.CustomScale; // 0.05 -> 20
double halfWidth = (viewport.Width * scaleFactor) / 2.0;
double halfHeight = (viewport.Height * scaleFactor) / 2.0;

// ViewCenter is the center of displayed model space area
Point3d modelSpaceCenter = new Point3d(viewport.ViewCenter.X, viewport.ViewCenter.Y, 0);

var modelCorners = new[]
{
    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y - halfHeight, 0),
    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y + halfHeight, 0),
    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y + halfHeight, 0),
    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y - halfHeight, 0)
};
```

#### Project Configuration Updates ✅
- **Core Project**: Added GeometryExtensions DLL reference to `libs/GeometryExtensionsR25.dll`
- **PlatformTarget**: Set to x64 to match DLL architecture
- **BlockTableRecord Access**: Uses `layout.BlockTableRecordId` to access viewports without layout switching

#### Development Dependencies
- **GeometryExtensions R25**: Viewport transformation library (Gilles Chanteau)
- **Reference Path**: `C:\Users\trevorp\Dev\KPFF.AutoCAD.DraftingAssistant\libs\GeometryExtensionsR25.dll`
- **Namespace**: `Gile.AutoCAD.R25.Geometry` (R25 for AutoCAD 2025 compatibility)

### Enhanced Viewport Transformation Implementation ✅ **COMPLETE**

#### Refactored ViewportBoundaryCalculator with Gile's ViewportExtension
The implementation was completely refactored to use Gile's ViewportExtension methods instead of manual coordinate calculations:

#### Key Transformation Chain
```csharp
// Get the transformation chain: PSDCS → DCS → WCS using Gile's ViewportExtension
Matrix3d psdcs2dcs = viewport.PSDCS2DCS();
Matrix3d dcs2wcs = viewport.DCS2WCS();

// Order matters: matrices apply right-to-left.
// Start in PSDCS, transform → DCS, then → WCS
Matrix3d fullTransform = dcs2wcs * psdcs2dcs;
```

#### Coordinate System Understanding
- **PSDCS (Paper Space Display Coordinate System)**: Layout coordinates as they appear on the sheet
- **DCS (Display Coordinate System)**: Each viewport's camera coordinate system 
- **WCS (World Coordinate System)**: Global fixed coordinate system for the entire DWG
- **ViewCenter Property**: Exists in DCS coordinates, critical for rotated viewports

#### Matrix Multiplication Order Fix ✅
**Critical Bug Fix**: Initial implementation had backwards matrix multiplication order:
- **Wrong**: `psdcs2dcs * dcs2wcs` (applies DCS→WCS transformation first)
- **Correct**: `dcs2wcs * psdcs2dcs` (applies PSDCS→DCS first, then DCS→WCS)
- **Impact**: Fixed twisted/rotated viewport calculations that were rotating around wrong origin

#### Supported Viewport Types
- **Rectangular Viewports**: Standard paper space viewports
- **Rotated Viewports**: Viewports with twist angle transformations
- **Polygonal Viewports**: Custom-clipped viewports with complex boundaries

#### Enhanced Features
- **Transaction Parameter**: Optional transaction for efficient polygonal viewport processing
- **Comprehensive Error Handling**: Specific exceptions for different failure modes
- **Transformation Diagnostics**: `GetTransformationDiagnostics()` method shows step-by-step coordinate transformations
- **Robust Polygonal Support**: Handles Polyline, Polyline2d, and Polyline3d clip entities

#### Validation Results ✅
All test cases confirmed working:
- **Non-rotated viewports**: Accurate boundary detection
- **Twisted/rotated viewports**: Correct transformation with proper matrix order
- **Polygonal viewports**: Proper clip entity processing
- **Mixed scenarios**: Robust handling across all viewport configurations

### ✅ Auto Notes Phase 5B Implementation Complete

#### Implemented Components
- **PointInPolygonDetector** (`Core/Utilities/`): Ray casting algorithm with tolerance support and winding number calculations
- **MultileaderAnalyzer** (`Core/Services/`): Comprehensive multileader discovery, style filtering, and TAGNUMBER extraction
- **AutoNotesService** (`Core/Services/`): Complete viewport-to-notes orchestration pipeline with layout management
- **TESTAUTONOTES Command**: Validation tool with diagnostics, point-in-polygon testing, and service initialization

#### Key Capabilities Delivered ✅
- **Multileader Discovery**: Finds all multileaders in model space with configurable style filtering (ML-STYLE-01)
- **Attribute Extraction**: Extracts TAGNUMBER attributes from multileader blocks (_TagCircle) with integer validation
- **Point-in-Polygon Detection**: Ray casting algorithm handles viewport boundary containment testing
- **Note Consolidation**: Removes duplicates and sorts note numbers for clean output
- **Full Integration**: Seamlessly integrated with existing ConstructionNotesService and Excel Notes workflow

#### Technical Implementation Details
- **Robust Error Handling**: Graceful handling of missing layouts, empty viewports, and invalid multileaders
- **Comprehensive Logging**: Detailed progress tracking and diagnostic information for troubleshooting
- **Service Integration**: Proper dependency injection and service initialization patterns
- **Transaction Management**: Safe AutoCAD object access with proper transaction handling

#### Testing & Validation ✅
- **TESTAUTONOTES Command**: Provides comprehensive testing with layout input, diagnostic output, and algorithm validation
- **Point-in-Polygon Verification**: Built-in tests validate ray casting algorithm with known test cases
- **Service Initialization**: Fixed service startup issues to ensure commands work reliably
- **Production Ready**: Successfully tested with PROJ-ABC-100.dwg test data

#### Commit Status ✅
**Commit `abd18e4`**: "Implement Auto Notes Phase 5B: Multileader Discovery and Analysis"
- 5 files changed, 986 insertions(+), 10 deletions(-)
- 3 new core services created
- Full backend implementation complete

### Next Implementation Phase: UI Integration (Phase 5C)
With the Auto Notes backend complete, the final step is user interface integration to provide seamless Auto/Excel mode selection:

#### 1. Update Construction Notes UI (ConstructionNotesView.xaml)
- **Add Radio Button Group**: "Auto Notes" (default) and "Excel Notes" mode selection
- **Multileader Style Configuration**: Text field for specifying target multileader style name
- **Mode-Specific Controls**: Show/hide relevant options based on selected mode
- **User Feedback Area**: Display detected note counts and processing status

#### 2. Enhance ViewModel Logic (ConstructionNotesViewModel.cs)
- **NotesMode Property**: Enum-based mode selection (Auto/Excel) with property change notification
- **Configuration Binding**: Two-way binding for multileader style name setting
- **Command Routing**: Update `UpdateNotesCommand` to call appropriate service based on mode
- **Progress Feedback**: Show real-time progress and results from Auto Notes detection

#### 3. Configuration Management
- **Project Settings**: Store multileader style preference in ProjectConfiguration
- **Default Values**: Set reasonable defaults (ML-STYLE-01) for new projects
- **Validation**: Ensure style name is valid before processing
- **Persistence**: Save user preferences across sessions

#### 4. User Experience Enhancements
- **Progress Indicators**: Show processing status during Auto Notes detection
- **Result Display**: List detected notes with counts and locations
- **Error Handling**: Clear error messages for common issues (no multileaders, layout not found)
- **Help Integration**: Tooltips and help text explaining Auto Notes requirements

#### 5. Testing & Quality Assurance
- **Mode Switching**: Verify smooth transitions between Auto and Excel modes
- **End-to-End Testing**: Test complete workflow from UI to drawing updates
- **Error Scenarios**: Test handling of missing layouts, invalid styles, empty viewports
- **Performance**: Ensure UI remains responsive during processing

#### Expected UI Layout
```
Construction Notes Tab:
┌─ Mode Selection ─────────────────────────────────┐
│ ○ Auto Notes (detect from viewports)            │
│ ○ Excel Notes (use manual configuration)        │
└──────────────────────────────────────────────────┘
┌─ Auto Notes Settings ────────────────────────────┐
│ Multileader Style: [ML-STYLE-01    ] [?]        │
│ └─ Style used to filter multileaders            │
└──────────────────────────────────────────────────┘
┌─ Processing Status ──────────────────────────────┐
│ ● Found 3 multileaders in ABC-101               │
│ ● Detected notes: 1, 2, 4                       │
│ ● Updated 3 construction note blocks             │
└──────────────────────────────────────────────────┘
                [Update Notes]
```

#### Implementation Priority
1. **Core UI Changes**: Radio buttons and basic mode switching
2. **Service Integration**: Route commands to Auto/Excel services
3. **Configuration**: Multileader style settings and persistence  
4. **User Feedback**: Progress display and result messaging
5. **Polish & Testing**: Error handling, tooltips, comprehensive testing

#### Benefits of Completion
- **Eliminates Manual Setup**: No more Excel configuration for basic note detection
- **Real-time Detection**: Notes automatically found from actual drawing content
- **Maintains Flexibility**: Excel mode still available for complex scenarios
- **Improved Workflow**: Faster, more accurate construction note management