#!/usr/bin/env python3
"""
GURI GUI Windows Service Wrapper
Allows the GURI GUI to run as a Windows service.
"""

import sys
import os
import time
import subprocess
import logging

# Try to import win32serviceutil for Windows service support
try:
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    WINDOWS_SERVICE_AVAILABLE = True
except ImportError:
    WINDOWS_SERVICE_AVAILABLE = False
    print("Windows service modules not available. Install pywin32: pip install pywin32")


class GURIGUIService(win32serviceutil.ServiceFramework if WINDOWS_SERVICE_AVAILABLE else object):
    """Windows Service wrapper for GURI GUI."""
    
    _svc_name_ = "GURIGUIService"
    _svc_display_name_ = "GURI Database Viewer & Manager"
    _svc_description_ = "GURI Database Viewer and Manager - Runs in system tray"
    
    def __init__(self, args):
        if WINDOWS_SERVICE_AVAILABLE:
            win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None) if WINDOWS_SERVICE_AVAILABLE else None
        self.process = None
        self.logger = logging.getLogger(__name__)
        
    def SvcStop(self):
        """Stop the service."""
        if WINDOWS_SERVICE_AVAILABLE:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self.hWaitStop)
            if self.process:
                self.process.terminate()
                self.process.wait()
    
    def SvcDoRun(self):
        """Run the service."""
        if not WINDOWS_SERVICE_AVAILABLE:
            return
        
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        suite_root = os.path.dirname(script_dir)
        gui_script = os.path.join(script_dir, "gui.py")
        
        try:
            self.process = subprocess.Popen(
                [sys.executable, gui_script],
                cwd=suite_root,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Wait for stop event
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error starting GURI GUI: {e}")
        finally:
            if self.process:
                self.process.terminate()
                self.process.wait()
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, '')
            )


def main():
    """Main entry point for service installation/removal."""
    if not WINDOWS_SERVICE_AVAILABLE:
        print("Windows service support requires pywin32.")
        print("Install with: pip install pywin32")
        print("\nAlternatively, use install_startup_service.bat for startup folder installation.")
        return
    
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(GURIGUIService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(GURIGUIService)


if __name__ == '__main__':
    main()

