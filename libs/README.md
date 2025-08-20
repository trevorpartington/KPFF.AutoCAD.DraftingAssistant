# External Dependencies

## GeometryExtensions Library

### Download Instructions
1. Download GeometryExtensions.zip from: https://gilecad.azurewebsites.net/Resources/GeometryExtensions.zip
2. Extract the ZIP file
3. Copy `GeometryExtensionsR25.dll` to this `libs/` directory (for AutoCAD 2025)
4. Add reference to the DLL in the Core project

### Usage
- The library provides viewport transformation methods like `viewport.DCS2WCS(Point3d)`
- Documentation: https://gilecad.azurewebsites.net/Resources/GeometryExtensionsHelp/index.html
- License: MIT
- Author: Gilles Chanteau

### Project Reference
Once the DLL is placed in this directory, add a reference in Core project:
```xml
<Reference Include="GeometryExtensionsR25">
  <HintPath>..\..\libs\GeometryExtensionsR25.dll</HintPath>
  <Private>True</Private>
</Reference>
```