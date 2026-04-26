"""
Integration with MSAccess VCS Add-in via COM automation.

This module provides a lightweight wrapper around the MSAccess VCS add-in,
delegating all export/import/build operations to the battle-tested add-in
rather than reimplementing them in Python.
"""

import os
from pathlib import Path
from typing import Any, Optional

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def get_access_info(app) -> dict[str, Any]:
    """
    Get Access application version and bitness.
    
    Args:
        app: Access Application COM object
        
    Returns:
        Dictionary with access_version and bitness
    """
    try:
        # Get Access version (e.g., "16.0", "15.0", "14.0")
        access_version = app.Version
        
        # Detect bitness - Access itself reports if it's 64-bit
        # We can check the build number or system architecture
        try:
            # Try to check if we're running in 64-bit mode
            # In 64-bit Access, VBE version is typically 7.1, in 32-bit it's 7.0
            import platform
            import sys
            
            # Check Python's architecture (which should match Access if they're compatible)
            is_64bit = sys.maxsize > 2**32
            bitness = "64-bit" if is_64bit else "32-bit"
        except Exception:
            bitness = "unknown"
        
        return {
            "access_version": access_version,
            "bitness": bitness
        }
    except Exception as e:
        return {
            "access_version": "unknown",
            "bitness": "unknown",
            "error": str(e)
        }


class VCSAddinIntegration:
    """
    Integration layer for MSAccess VCS add-in.
    
    This class handles:
    - Loading the VCS add-in into Access
    - Calling add-in API functions via Application.Run
    - Translating between MCP and add-in formats
    - Parsing results from add-in operations
    """
    
    def __init__(self, addin_path: Optional[str] = None):
        """
        Initialize add-in integration.
        
        Args:
            addin_path: Path to VCS add-in file (.accda). If None, uses default location.
        """
        if not COM_AVAILABLE:
            raise ImportError(
                "pywin32 is required for COM automation. "
                "Install it with: pip install pywin32"
            )
        
        self.addin_path = addin_path or self._get_default_addin_path()
        self._app = None
        self._addin_loaded = False
    
    def _get_default_addin_path(self) -> str:
        """
        Get default VCS add-in installation path.
        
        Returns:
            Path to add-in file (may not exist)
        """
        # Default installation location: %AppData%\MSAccessVCS\Version Control.accda
        appdata = os.environ.get("APPDATA", "")
        return os.path.join(appdata, "MSAccessVCS", "Version Control.accda")
    
    def verify_addin_exists(self) -> bool:
        """
        Check if add-in file exists at configured path.
        
        Returns:
            True if add-in file exists
        """
        return os.path.isfile(self.addin_path)
    
    def load_addin(self, app) -> bool:
        """
        Verify add-in is accessible via the new API method.
        
        Note: The add-in doesn't need to be explicitly "loaded" - Access will
        load it automatically when we call Application.Run. However, the target
        database must be open first.
        
        Args:
            app: Access Application COM object (with a database open)
            
        Returns:
            True if add-in can be called successfully
            
        Raises:
            RuntimeError: If add-in cannot be accessed
        """
        if not self.verify_addin_exists():
            raise RuntimeError(
                f"VCS add-in not found at: {self.addin_path}\n"
                f"Please install the MSAccess VCS add-in or set ACCESS_VCS_ADDIN_PATH.\n"
                f"Download from: https://github.com/joyfullservice/msaccess-vcs-integration/releases"
            )
        
        try:
            # Store app reference
            self._app = app
            
            # Verify add-in is accessible using the new API method
            # Format: app.Run("C:\Path\Version Control.API", "GetVCSVersion")
            addin_lib_name = os.path.splitext(self.addin_path)[0]
            api_function_name = f'{addin_lib_name}.API'
            
            # Try calling a simple function to verify the add-in works
            # This will also load the add-in if it's not already loaded
            app.Run(api_function_name, "GetVCSVersion")
            
            self._addin_loaded = True
            return True
            
        except Exception as e:
            raise RuntimeError(
                f"Failed to load VCS add-in: {e}\n"
                f"Ensure a database is open and the add-in is trusted by Access."
            )
    
    def _call_addin_function(self, function_name: str, *args) -> Any:
        """
        Call a function in the VCS add-in using the new API method.
        
        Note: A database must be open in the Access application before calling
        add-in functions. The add-in loads automatically when called.
        
        Args:
            function_name: Name of function to call (e.g., "GetVCSVersion", "HandleRibbonCommand")
            *args: Arguments to pass to the function
            
        Returns:
            Result from add-in function
            
        Raises:
            RuntimeError: If call fails
        """
        # Gate on add-in load state so callers get a clear lifecycle error
        # instead of a downstream COM message.
        if not self._addin_loaded or not self._app:
            raise RuntimeError("VCS add-in not loaded. Call load_addin() first.")

        try:
            # New API format: Path without extension + ".API", then function name as first argument
            # Example: app.Run("C:\Path\Version Control.API", "GetVCSVersion")
            addin_path_abs = os.path.abspath(self.addin_path)
            addin_lib_name = os.path.splitext(addin_path_abs)[0]
            api_function_name = f'{addin_lib_name}.API'

            # Verify database is open (required for add-in to work)
            try:
                current_db = self._app.CurrentDb()
                if not current_db:
                    raise RuntimeError("No database is currently open in Access. The add-in requires a database to be open.")
                # Force Access to recognize the current database by accessing a property
                # This ensures the database context is fully established
                _ = current_db.Name
            except Exception as db_error:
                raise RuntimeError(f"Cannot access current database: {db_error}. Ensure a database is open before calling add-in functions.")

            # With early binding (gencache.EnsureDispatch), Run returns a tuple
            # where the first element is the actual result.
            try:
                if args:
                    result = self._app.Run(api_function_name, function_name, *args)
                else:
                    result = self._app.Run(api_function_name, function_name)
                if isinstance(result, tuple) and len(result) > 0:
                    return result[0]
                return result
            except Exception as run_error:
                # First call may fail while Access is loading/initializing the
                # add-in. Retry once -- the add-in should now be resident.
                try:
                    if args:
                        result = self._app.Run(api_function_name, function_name, *args)
                    else:
                        result = self._app.Run(api_function_name, function_name)
                    if isinstance(result, tuple) and len(result) > 0:
                        return result[0]
                    return result
                except Exception as second_run_error:
                    raise RuntimeError(
                        f"Failed to call add-in function '{function_name}': {second_run_error}\n"
                        f"First attempt error: {run_error}\n"
                        f"API path used: {api_function_name}\n"
                        f"Add-in path: {self.addin_path}\n"
                        f"Ensure a database is open and the add-in path is correct."
                    )

        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(
                f"Failed to call add-in function '{function_name}': {e}\n"
                f"Ensure a database is open and the add-in path is correct."
            )
    
    def _get_export_folder(self, db_path: str, source_folder: Optional[str] = None) -> str:
        """
        Determine export folder path.
        
        Args:
            db_path: Path to database file
            source_folder: Optional explicit source folder path
            
        Returns:
            Path to export folder
        """
        if source_folder:
            return source_folder
        
        # Default: database_name.src folder next to database
        db_file = Path(db_path)
        return str(db_file.parent / f"{db_file.stem}.src")
    
    def export_source(
        self,
        db_path: str,
        source_folder: Optional[str] = None,
        full_export: bool = False
    ) -> dict[str, Any]:
        """
        Export database to source files using VCS add-in.
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "Export" command, which spawns via timer and allows
        the UI to show while posting progress callbacks.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom export folder (default: db_name.src)
            full_export: If True, force full export; if False, use fast save
            
        Returns:
            Dictionary with export results:
            - success: Boolean
            - export_path: Path where files were exported
            - log_path: Path to Export.log file
            - message: Status message
        """
        export_path = self._get_export_folder(db_path, source_folder)
        
        try:
            # Call VCS API directly (Export or FullExport based on flag)
            command = "FullExport" if full_export else "Export"
            self._call_addin_function(command)
            
            # Check for log file to confirm export completed
            log_path = os.path.join(export_path, "Export.log")
            
            return {
                "success": True,
                "export_path": export_path,
                "log_path": log_path if os.path.exists(log_path) else None,
                "message": "Export completed successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "export_path": export_path,
                "log_path": None,
                "message": f"Export failed: {e}"
            }
    
    def export_vba(self, db_path: str, source_folder: Optional[str] = None) -> dict[str, Any]:
        """
        Export only VBA components (modules, class modules).
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "ExportVBA" command.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom export folder
            
        Returns:
            Dictionary with export results
        """
        export_path = self._get_export_folder(db_path, source_folder)
        
        try:
            self._call_addin_function("ExportVBA")
            
            log_path = os.path.join(export_path, "Export.log")
            
            return {
                "success": True,
                "export_path": export_path,
                "log_path": log_path if os.path.exists(log_path) else None,
                "message": "VBA export completed successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "export_path": export_path,
                "log_path": None,
                "message": f"VBA export failed: {e}"
            }
    
    def merge_build(self, db_path: str, source_folder: Optional[str] = None) -> dict[str, Any]:
        """
        Merge source files into existing database.
        
        This updates modified objects without rebuilding the entire database.
        
        Note: This is the synchronous fallback. The preferred approach is to use
        call_async() with the "MergeBuild" command.
        
        Args:
            db_path: Path to Access database
            source_folder: Optional custom source folder
            
        Returns:
            Dictionary with build results:
            - success: Boolean
            - database_path: Path to database
            - log_path: Path to Build.log file
            - message: Status message
        """
        source_path = self._get_export_folder(db_path, source_folder)
        
        try:
            self._call_addin_function("MergeBuild")
            
            log_path = os.path.join(source_path, "Build.log")
            
            return {
                "success": True,
                "database_path": db_path,
                "log_path": log_path if os.path.exists(log_path) else None,
                "message": "Merge build completed successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "database_path": db_path,
                "log_path": None,
                "message": f"Merge build failed: {e}"
            }
    
    def build_from_source(
        self,
        source_folder: str,
        output_path: Optional[str] = None
    ) -> dict[str, Any]:
        """
        Build database from source files.
        
        Creates a fresh database from source files.
        
        Args:
            source_folder: Path to source files folder
            output_path: Optional path for new database (default: build in place)
            
        Returns:
            Dictionary with build results:
            - success: Boolean
            - output_path: Path to built database
            - log_path: Path to Build.log file
            - message: Status message
        """
        try:
            # Build always takes an optional source folder argument.
            # BuildAs is interactive-only (shows file dialogs), so we use
            # Build for both cases in headless mode.
            self._call_addin_function("Build", source_folder)
            
            log_path = os.path.join(source_folder, "Build.log")
            
            return {
                "success": True,
                "output_path": output_path,
                "log_path": log_path if os.path.exists(log_path) else None,
                "message": "Build from source completed successfully"
            }
            
        except Exception as e:
            return {
                "success": False,
                "output_path": None,
                "log_path": None,
                "message": f"Build from source failed: {e}"
            }
    
    def parse_log_file(self, log_path: str) -> dict[str, Any]:
        """
        Parse add-in log file for detailed results.
        
        Args:
            log_path: Path to Export.log or Build.log
            
        Returns:
            Dictionary with parsed log information
        """
        if not os.path.exists(log_path):
            return {"found": False}
        
        try:
            with open(log_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            return {
                "found": True,
                "content": content,
                "path": log_path
            }
            
        except Exception as e:
            return {
                "found": False,
                "error": str(e)
            }
    
    # =========================================================================
    # Sync/Async API Methods (for callback-enabled operations)
    # =========================================================================
    
    def call_sync(self, command: str, *args) -> Any:
        """
        Call a VBA function synchronously using the API entry point.
        
        This calls the existing API function which returns results immediately.
        Use for quick operations that don't need progress reporting.
        
        Args:
            command: Command name (e.g., "GetVCSVersion", "GetOptions")
            *args: Additional arguments to pass
            
        Returns:
            Result from the VBA function
            
        Raises:
            RuntimeError: If call fails
        """
        return self._call_addin_function(command, *args)
    
    def call_async(self, callback_info: str, command: str, *args) -> dict[str, Any]:
        """
        Call a VBA function asynchronously using the APIAsync entry point.
        
        This calls the new APIAsync function which:
        - Spawns a detached process for long-running operations
        - Returns immediately with async marker and timeout hint
        - Sends progress updates via HTTP callbacks
        - Sends completion or error when done
        
        Args:
            callback_info: JSON string with callback_url and operation_id
            command: Command name (e.g., "Export", "Build", "MergeBuild")
            *args: Additional arguments to pass
            
        Returns:
            Dict with either:
            - {"sync": true, "result": ...} for quick operations
            - {"async": true, "timeout_ms": ...} for async operations
            
        Raises:
            RuntimeError: If call fails
        """
        import json
        
        try:
            # Call APIAsync entry point
            # Format: APIAsync(strCallbackInfo, strCommand, [args...])
            addin_path_abs = os.path.abspath(self.addin_path)
            addin_lib_name = os.path.splitext(addin_path_abs)[0]
            api_async_name = f'{addin_lib_name}.APIAsync'
            
            # Ensure app reference is set
            if not self._app:
                raise RuntimeError("Access Application object not set.")
            
            # Call APIAsync with callback info as first arg
            if args:
                result = self._app.Run(api_async_name, callback_info, command, *args)
            else:
                result = self._app.Run(api_async_name, callback_info, command)
            
            # Handle tuple return from early-bound Run method
            if isinstance(result, tuple) and len(result) > 0:
                result = result[0]
            
            # Parse JSON response from VBA
            if isinstance(result, str):
                return json.loads(result)
            else:
                # Unexpected return type
                return {"sync": True, "result": result}
                
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON response from APIAsync: {e}")
        except Exception as e:
            raise RuntimeError(f"Failed to call APIAsync '{command}': {e}")
    
    def get_version_info(self, app) -> dict[str, Any]:
        """
        Get comprehensive version information for VCS add-in and Access.
        
        Note: A database must be open in the Access application before calling this.
        
        Args:
            app: Access Application COM object (with database open)
            
        Returns:
            Dictionary with vcs_version, access_version, bitness, and paths
        """
        result = {
            "success": True,
            "addin_path": self.addin_path,
        }
        
        # Store app reference temporarily for this call
        self._app = app
        
        # Get Access application info
        access_info = get_access_info(app)
        result.update(access_info)
        
        # Get VCS add-in version
        try:
            vcs_version = self._call_addin_function("GetVCSVersion")
            result["vcs_version"] = vcs_version
        except Exception as e:
            result["vcs_version"] = None
            result["vcs_error"] = f"Failed to get VCS version: {e}"
            result["success"] = False
        
        return result
