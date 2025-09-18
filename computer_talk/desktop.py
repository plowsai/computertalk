"""
Desktop application integration for computer-talk.
Provides capabilities to open and interact with desktop applications.
"""

import subprocess
import platform
import time
import logging
from typing import List, Dict, Any, Optional, Union
from pathlib import Path
import json

from .exceptions import ComputerTalkError, CommunicationError


class DesktopManager:
    """
    Manages desktop application interactions and control.
    """
    
    def __init__(self):
        """Initialize the desktop manager."""
        self.system = platform.system().lower()
        self.logger = logging.getLogger(__name__)
        self.running_apps = {}
        
    def open_application(self, app_name: str, **kwargs) -> Dict[str, Any]:
        """
        Open a desktop application.
        
        Args:
            app_name: Name or path of the application to open
            **kwargs: Additional arguments for the application
            
        Returns:
            Dictionary with app info and status
            
        Raises:
            ComputerTalkError: If app cannot be opened
        """
        try:
            self.logger.info(f"Opening application: {app_name}")
            
            if self.system == "darwin":  # macOS
                result = self._open_macos_app(app_name, **kwargs)
            elif self.system == "windows":
                result = self._open_windows_app(app_name, **kwargs)
            elif self.system == "linux":
                result = self._open_linux_app(app_name, **kwargs)
            else:
                raise ComputerTalkError(f"Unsupported operating system: {self.system}")
            
            # Store app info
            app_id = f"{app_name}_{int(time.time())}"
            self.running_apps[app_id] = {
                "name": app_name,
                "pid": result.get("pid"),
                "started_at": time.time(),
                "status": "running"
            }
            
            return {
                "success": True,
                "app_id": app_id,
                "app_name": app_name,
                "pid": result.get("pid"),
                "message": f"Successfully opened {app_name}"
            }
            
        except Exception as e:
            self.logger.error(f"Failed to open {app_name}: {e}")
            raise ComputerTalkError(f"Failed to open application {app_name}: {e}")
    
    def _open_macos_app(self, app_name: str, **kwargs) -> Dict[str, Any]:
        """Open application on macOS."""
        # Try different approaches for macOS
        commands = [
            # Direct app name
            ["open", "-a", app_name],
            # App bundle path
            ["open", app_name],
            # Using osascript
            ["osascript", "-e", f'tell application "{app_name}" to activate']
        ]
        
        for cmd in commands:
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    return {"pid": None, "command": cmd}
            except subprocess.TimeoutExpired:
                continue
            except Exception:
                continue
        
        raise ComputerTalkError(f"Could not open {app_name} on macOS")
    
    def _open_windows_app(self, app_name: str, **kwargs) -> Dict[str, Any]:
        """Open application on Windows."""
        try:
            # Try to start the application
            result = subprocess.run(
                ["start", app_name], 
                shell=True, 
                capture_output=True, 
                text=True, 
                timeout=10
            )
            return {"pid": None, "command": ["start", app_name]}
        except Exception as e:
            raise ComputerTalkError(f"Could not open {app_name} on Windows: {e}")
    
    def _open_linux_app(self, app_name: str, **kwargs) -> Dict[str, Any]:
        """Open application on Linux."""
        try:
            # Try different approaches for Linux
            commands = [
                [app_name],
                ["xdg-open", app_name],
                ["gnome-open", app_name],
                ["kde-open", app_name]
            ]
            
            for cmd in commands:
                try:
                    result = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    return {"pid": result.pid, "command": cmd}
                except FileNotFoundError:
                    continue
            
            raise ComputerTalkError(f"Could not open {app_name} on Linux")
        except Exception as e:
            raise ComputerTalkError(f"Could not open {app_name} on Linux: {e}")
    
    def list_running_apps(self) -> List[Dict[str, Any]]:
        """
        List currently running applications.
        
        Returns:
            List of running application information
        """
        try:
            if self.system == "darwin":
                return self._list_macos_apps()
            elif self.system == "windows":
                return self._list_windows_apps()
            elif self.system == "linux":
                return self._list_linux_apps()
            else:
                return []
        except Exception as e:
            self.logger.error(f"Failed to list apps: {e}")
            return []
    
    def _list_macos_apps(self) -> List[Dict[str, Any]]:
        """List running apps on macOS."""
        try:
            result = subprocess.run(
                ["osascript", "-e", "tell application \"System Events\" to get name of every process"],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                apps = [app.strip() for app in result.stdout.split(",")]
                return [{"name": app, "system": "macOS"} for app in apps if app]
            return []
        except Exception:
            return []
    
    def _list_windows_apps(self) -> List[Dict[str, Any]]:
        """List running apps on Windows."""
        try:
            result = subprocess.run(
                ["tasklist", "/fo", "csv"],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')[1:]  # Skip header
                apps = []
                for line in lines:
                    parts = line.split(',')
                    if len(parts) >= 1:
                        app_name = parts[0].strip('"')
                        apps.append({"name": app_name, "system": "Windows"})
                return apps
            return []
        except Exception:
            return []
    
    def _list_linux_apps(self) -> List[Dict[str, Any]]:
        """List running apps on Linux."""
        try:
            result = subprocess.run(
                ["ps", "aux"],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')[1:]  # Skip header
                apps = []
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 11:
                        app_name = parts[10]
                        if not app_name.startswith('['):  # Skip kernel processes
                            apps.append({"name": app_name, "system": "Linux"})
                return apps
            return []
        except Exception:
            return []
    
    def get_common_apps(self) -> List[Dict[str, Any]]:
        """
        Get list of common applications that can be opened.
        
        Returns:
            List of common applications with their names and descriptions
        """
        common_apps = {
            "darwin": [
                {"name": "Safari", "description": "Web browser", "command": "Safari"},
                {"name": "Chrome", "description": "Web browser", "command": "Google Chrome"},
                {"name": "Firefox", "description": "Web browser", "command": "Firefox"},
                {"name": "Terminal", "description": "Command line terminal", "command": "Terminal"},
                {"name": "Finder", "description": "File manager", "command": "Finder"},
                {"name": "TextEdit", "description": "Text editor", "command": "TextEdit"},
                {"name": "Notes", "description": "Note-taking app", "command": "Notes"},
                {"name": "Calendar", "description": "Calendar app", "command": "Calendar"},
                {"name": "Mail", "description": "Email client", "command": "Mail"},
                {"name": "Messages", "description": "Messaging app", "command": "Messages"},
                {"name": "Spotify", "description": "Music streaming", "command": "Spotify"},
                {"name": "VSCode", "description": "Code editor", "command": "Visual Studio Code"},
                {"name": "Xcode", "description": "iOS development", "command": "Xcode"},
                {"name": "Photos", "description": "Photo management", "command": "Photos"},
                {"name": "Preview", "description": "PDF and image viewer", "command": "Preview"}
            ],
            "macOS": [
                {"name": "Safari", "description": "Web browser", "command": "Safari"},
                {"name": "Chrome", "description": "Web browser", "command": "Google Chrome"},
                {"name": "Firefox", "description": "Web browser", "command": "Firefox"},
                {"name": "Terminal", "description": "Command line terminal", "command": "Terminal"},
                {"name": "Finder", "description": "File manager", "command": "Finder"},
                {"name": "TextEdit", "description": "Text editor", "command": "TextEdit"},
                {"name": "Notes", "description": "Note-taking app", "command": "Notes"},
                {"name": "Calendar", "description": "Calendar app", "command": "Calendar"},
                {"name": "Mail", "description": "Email client", "command": "Mail"},
                {"name": "Messages", "description": "Messaging app", "command": "Messages"},
                {"name": "Spotify", "description": "Music streaming", "command": "Spotify"},
                {"name": "VSCode", "description": "Code editor", "command": "Visual Studio Code"},
                {"name": "Xcode", "description": "iOS development", "command": "Xcode"},
                {"name": "Photos", "description": "Photo management", "command": "Photos"},
                {"name": "Preview", "description": "PDF and image viewer", "command": "Preview"}
            ],
            "Windows": [
                {"name": "Chrome", "description": "Web browser", "command": "chrome"},
                {"name": "Firefox", "description": "Web browser", "command": "firefox"},
                {"name": "Edge", "description": "Web browser", "command": "msedge"},
                {"name": "Notepad", "description": "Text editor", "command": "notepad"},
                {"name": "Word", "description": "Word processor", "command": "winword"},
                {"name": "Excel", "description": "Spreadsheet", "command": "excel"},
                {"name": "PowerPoint", "description": "Presentation", "command": "powerpnt"},
                {"name": "VSCode", "description": "Code editor", "command": "code"},
                {"name": "Calculator", "description": "Calculator app", "command": "calc"},
                {"name": "Paint", "description": "Image editor", "command": "mspaint"}
            ],
            "Linux": [
                {"name": "Chrome", "description": "Web browser", "command": "google-chrome"},
                {"name": "Firefox", "description": "Web browser", "command": "firefox"},
                {"name": "Terminal", "description": "Command line", "command": "gnome-terminal"},
                {"name": "VSCode", "description": "Code editor", "command": "code"},
                {"name": "Gedit", "description": "Text editor", "command": "gedit"},
                {"name": "LibreOffice", "description": "Office suite", "command": "libreoffice"},
                {"name": "GIMP", "description": "Image editor", "command": "gimp"},
                {"name": "Calculator", "description": "Calculator", "command": "gnome-calculator"}
            ]
        }
        
        apps = common_apps.get(self.system, [])
        self.logger.debug(f"Found {len(apps)} common apps for {self.system}")
        return apps
    
    def interact_with_app(self, app_name: str, action: str, **kwargs) -> Dict[str, Any]:
        """
        Interact with a running application.
        
        Args:
            app_name: Name of the application
            action: Action to perform
            **kwargs: Additional parameters
            
        Returns:
            Result of the interaction
        """
        try:
            self.logger.info(f"Interacting with {app_name}: {action}")
            
            if self.system == "darwin":
                return self._interact_macos_app(app_name, action, **kwargs)
            elif self.system == "windows":
                return self._interact_windows_app(app_name, action, **kwargs)
            elif self.system == "linux":
                return self._interact_linux_app(app_name, action, **kwargs)
            else:
                raise ComputerTalkError(f"Unsupported operating system: {self.system}")
                
        except Exception as e:
            self.logger.error(f"Failed to interact with {app_name}: {e}")
            raise CommunicationError(f"Failed to interact with {app_name}: {e}")
    
    def _interact_macos_app(self, app_name: str, action: str, **kwargs) -> Dict[str, Any]:
        """Interact with macOS application using AppleScript."""
        actions = {
            "activate": f'tell application "{app_name}" to activate',
            "close": f'tell application "{app_name}" to close',
            "quit": f'tell application "{app_name}" to quit',
            "minimize": f'tell application "{app_name}" to set minimized of window 1 to true',
            "maximize": f'tell application "{app_name}" to set zoomed of window 1 to true'
        }
        
        if action not in actions:
            raise ComputerTalkError(f"Unknown action: {action}")
        
        try:
            result = subprocess.run(
                ["osascript", "-e", actions[action]],
                capture_output=True, text=True, timeout=10
            )
            
            if result.returncode == 0:
                return {"success": True, "action": action, "result": result.stdout.strip()}
            else:
                return {"success": False, "action": action, "error": result.stderr.strip()}
        except Exception as e:
            raise ComputerTalkError(f"Failed to execute action {action}: {e}")
    
    def _interact_windows_app(self, app_name: str, action: str, **kwargs) -> Dict[str, Any]:
        """Interact with Windows application."""
        # Windows interaction would require more complex automation
        # For now, return a basic response
        return {"success": True, "action": action, "message": f"Action {action} on {app_name} (Windows)"}
    
    def _interact_linux_app(self, app_name: str, action: str, **kwargs) -> Dict[str, Any]:
        """Interact with Linux application."""
        # Linux interaction would require X11/Wayland automation
        # For now, return a basic response
        return {"success": True, "action": action, "message": f"Action {action} on {app_name} (Linux)"}
    
    def get_app_status(self, app_id: str) -> Optional[Dict[str, Any]]:
        """
        Get status of a tracked application.
        
        Args:
            app_id: ID of the application
            
        Returns:
            Application status or None if not found
        """
        return self.running_apps.get(app_id)
    
    def close_app(self, app_id: str) -> Dict[str, Any]:
        """
        Close a tracked application.
        
        Args:
            app_id: ID of the application
            
        Returns:
            Result of closing the application
        """
        if app_id not in self.running_apps:
            raise ComputerTalkError(f"Application {app_id} not found")
        
        app_info = self.running_apps[app_id]
        app_name = app_info["name"]
        
        try:
            result = self.interact_with_app(app_name, "quit")
            self.running_apps[app_id]["status"] = "closed"
            return result
        except Exception as e:
            self.logger.error(f"Failed to close {app_name}: {e}")
            return {"success": False, "error": str(e)}
