# GURI GUI - Startup and System Tray Configuration

## Overview

The GURI GUI application now supports:
- **System Tray Icon** - Runs in the background with a tray icon
- **Startup on Login** - Automatically starts when Windows boots
- **Minimize to Tray** - Closing the window minimizes to tray instead of quitting

## System Tray Features

### Tray Icon Menu
- **Show Window** - Restore and show the main window
- **Hide Window** - Minimize to tray
- **Refresh Records** - Refresh the database records
- **Quit** - Completely exit the application

### Behavior
- **Double-click tray icon** - Toggle window visibility
- **Close window (X button)** - Minimizes to tray (doesn't quit)
- **Quit from menu** - Completely exits the application

## Installation Methods

### Method 1: Startup Folder (Recommended - Simple)

1. **Run the installer:**
   ```
   install_startup_service.bat
   ```

2. This creates a shortcut in Windows Startup folder
3. The application will start automatically when you log in
4. It will appear in the system tray

**To remove:**
```
uninstall_startup_service.bat
```

### Method 2: Windows Service (Advanced)

1. **Install pywin32:**
   ```
   pip install pywin32
   ```

2. **Run the service installer (as Administrator):**
   ```
   install_windows_service.bat
   ```

3. **Start the service (as Administrator):**
   ```
   python guri_gui_service.py start
   ```

**Service Management:**
- Start: `python guri_gui_service.py start`
- Stop: `python guri_gui_service.py stop`
- Remove: `python guri_gui_service.py remove`

Or use Windows Services (services.msc) to manage it.

### Method 3: Manual Startup Folder

1. Press `Win + R`
2. Type: `shell:startup`
3. Create a shortcut to `start_guri_gui.bat`
4. Or create a shortcut directly to `guri_gui.py`

## Files Created

- `start_guri_gui.bat` - Startup script
- `install_startup_service.bat` - Installer for startup folder
- `uninstall_startup_service.bat` - Uninstaller for startup folder
- `guri_gui_service.py` - Windows service wrapper
- `install_windows_service.bat` - Service installer

## Troubleshooting

### System Tray Icon Not Appearing

- Check if system tray is available: The application checks for tray availability on startup
- Some Windows configurations may hide tray icons - check Windows tray settings
- Restart the application

### Application Not Starting on Login

- Check if the shortcut exists in Startup folder: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`
- Verify the batch file path is correct
- Check Windows Event Viewer for errors

### Service Won't Start

- Ensure you're running as Administrator
- Check if pywin32 is installed: `pip list | findstr pywin32`
- Check Windows Event Viewer for service errors
- Verify Python path is correct in the service script

## Manual Start

To start manually without startup:
```
python guri_gui.py
```

Or use the batch file:
```
start_guri_gui.bat
```

