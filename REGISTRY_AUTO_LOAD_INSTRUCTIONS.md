# Registry-Based Auto-Loading Instructions

## Overview
The KPFF Drafting Assistant plugin now uses registry-based auto-loading instead of startup scripts to eliminate plugin loading conflicts with ProjectWise and other AutoCAD integrations.

## How It Works

### 1. **Plugin Loading (Automatic)**
- Plugin DLL registers itself in the Windows registry for auto-loading with AutoCAD
- Loads automatically when AutoCAD starts (no manual NETLOAD required)
- No startup script conflicts with other plugins

### 2. **Service Initialization (On-Demand)**
- Services remain dormant until first command use
- Type `KPFF` to initialize all services and show the palette
- Subsequent commands use already-initialized services

## User Experience

### **Before (Script-Based Loading)**
1. AutoCAD starts → Script runs → Plugin loads → Services initialize
2. ProjectWise conflicts caused tab-switching issues
3. Drawing opens but focus reverts to Drawing1

### **After (Registry-Based Loading)**
1. AutoCAD starts → Plugin loads silently (no conflicts)
2. User types `KPFF` → Services initialize → Palette appears
3. No tab-switching issues, works with any AutoCAD configuration

## Available Commands

- **`KPFF`** - Main entry point command (initializes services and shows palette)
- `KPFFSTART` - Alias for the main command
- `KPFFHELP` - Show help information
- `KPFFTOGGLE` - Toggle palette visibility
- `KPFFHIDE` - Hide palette

## Registry Installation

The plugin automatically registers itself for auto-loading in:
```
HKEY_CURRENT_USER\SOFTWARE\Autodesk\AutoCAD\R25.0\ACAD-5001:409\Applications\KPFFDraftingAssistant
```

**Registry Values:**
- `DESCRIPTION`: "KPFF Drafting Assistant Plugin"
- `LOADCTRLS`: 14 (Load on startup + demand load)
- `LOADER`: Path to plugin DLL
- `MANAGED`: 1 (Managed .NET assembly)

## Troubleshooting

### Plugin Not Loading
1. Check if registry entry exists (auto-created on first run)
2. Verify DLL path in registry is correct
3. Restart AutoCAD to reload registry settings

### Services Not Initializing
1. Type `KPFF` command to manually trigger initialization
2. Check debug output for error messages
3. Verify AutoCAD document is open and accessible

## Development Testing

### Debug Mode
- Use Visual Studio launch profiles (script removed)
- Plugin loads via registry instead of start.scr
- No script execution during AutoCAD startup

### Production Deployment
- Plugin auto-registers on first load
- Works in any AutoCAD environment (standalone, ProjectWise, Vault)
- No IT configuration required

## Benefits

✅ **Universal Compatibility** - Works with any AutoCAD configuration  
✅ **No Startup Conflicts** - Eliminates script execution issues  
✅ **User Controlled** - Services load only when needed  
✅ **Enterprise Ready** - Registry-based loading is the standard approach  
✅ **Future Proof** - Handles unknown plugin conflicts gracefully