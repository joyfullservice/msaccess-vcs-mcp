"""
Integration with MSAccess VCS Add-in via COM automation.

This module provides a lightweight wrapper around the MSAccess VCS add-in,
delegating all export/import/build operations to the battle-tested add-in
rather than reimplementing them in Python.
"""

import os
import threading
import time
from pathlib import Path
from typing import Any, Optional

try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

from .usage_logging import log_addin_probe


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
    
    # Class-level guard against zombie probe threads piling up if Access
    # stops responding (VBA break mode, modal dialog, true hang).  Per-
    # instance state would be useless: every tool call constructs a fresh
    # VCSAddinIntegration, so a hung probe on one instance is invisible to
    # the next instance unless we track it on the class.
    _active_probe_thread: Optional[threading.Thread] = None
    
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
    
    def load_addin(self, app, db_path: Optional[str] = None) -> bool:
        """
        Verify add-in is accessible via the new API method.
        
        Probes the add-in by calling ``GetVCSVersion`` through
        ``Application.Run`` with a hard timeout, surfacing dialog-blocked,
        VBA-break-mode, or hung Access instances as a clear lifecycle error
        before any real work is dispatched.  Idempotent: a second call on
        the same instance is a no-op once the probe has succeeded.
        
        Args:
            app: Access Application COM object (with a database open).
            db_path: Optional path to the target database.  When provided,
                two things change:
                  * a fast ``os.path.isfile`` pre-flight catches stale or
                    typo paths in ~1ms instead of burning the full timeout;
                  * the worker thread re-acquires Access via the Running
                    Object Table for proper cross-apartment timeout
                    enforcement.  When ``None``, the worker falls back to
                    sharing the main thread's ``app`` proxy (best-effort
                    timeout).
        
        Returns:
            True if add-in can be called successfully.
        
        Raises:
            RuntimeError: If add-in cannot be accessed (missing file,
                untrusted, missing database, dialog-blocked, etc.).
            TimeoutError: If the probe exceeds
                ``ACCESS_VCS_PROBE_TIMEOUT_SEC`` (default 10s) -- typically
                because Access is in VBA break mode or has a modal dialog
                open.
        """
        # Idempotent: skip the probe if we've already loaded the add-in
        # against this instance.  validate_components() and get_version_info
        # both call load_addin and we don't want to double-pay.
        if self._addin_loaded:
            return True
        
        if not self.verify_addin_exists():
            raise RuntimeError(
                f"VCS add-in not found at: {self.addin_path}\n"
                f"Please install the MSAccess VCS add-in or set ACCESS_VCS_ADDIN_PATH.\n"
                f"Download from: https://github.com/joyfullservice/msaccess-vcs-integration/releases"
            )
        
        # Pre-flight: fast file-existence check.  When db_path is provided,
        # catching a missing/moved/typo path here in ~1ms avoids burning the
        # full probe timeout in ROT lookup for a database that isn't there.
        if db_path is not None and not os.path.isfile(db_path):
            raise RuntimeError(
                f"Database file not found at: {db_path}.  The Access "
                f"instance may have closed it, the file was moved/deleted, "
                f"or the path is incorrect."
            )
        
        try:
            timeout_sec = float(os.environ.get("ACCESS_VCS_PROBE_TIMEOUT_SEC", "10"))
        except ValueError:
            timeout_sec = 10.0
        
        probe_start = time.perf_counter()
        probe_error: Optional[str] = None
        timed_out = False
        try:
            self._probe_with_timeout(app, db_path, timeout_sec)
        except TimeoutError as e:
            timed_out = True
            probe_error = str(e)
            raise
        except RuntimeError as e:
            probe_error = str(e)
            raise
        except Exception as e:
            probe_error = str(e)
            raise RuntimeError(
                f"Failed to load VCS add-in: {e}\n"
                f"Ensure a database is open and the add-in is trusted by Access."
            )
        finally:
            duration_ms = round((time.perf_counter() - probe_start) * 1000, 2)
            log_addin_probe(
                addin_path=self.addin_path,
                duration_ms=duration_ms,
                success=probe_error is None,
                timed_out=timed_out,
                error=probe_error,
            )
        
        self._app = app
        self._addin_loaded = True
        return True
    
    def _probe_with_timeout(self, app, db_path: Optional[str], timeout_sec: float) -> None:
        """
        Run ``GetVCSVersion`` in a daemon worker thread with a hard timeout.
        
        Adapted from db-inspector-mcp's ``_run_dao_with_timeout`` (see that
        project's DECISIONS.md for the full rationale).  The short version:
        Access COM has no native timeout knob, ``CoCancelCall`` requires
        server-side cooperation Jet/ACE doesn't implement, and killing
        ``MSACCESS.EXE`` would lose the user's unsaved work.  A daemon
        thread + ``thread.join(timeout)`` is the only practical way to
        recover responsiveness when Access is stuck.
        
        Raises:
            TimeoutError: If the worker doesn't finish within ``timeout_sec``.
            RuntimeError: If a previous probe is still pending, or if the
                worker can't acquire an Access instance.
        """
        cls = type(self)
        if cls._active_probe_thread is not None and cls._active_probe_thread.is_alive():
            raise RuntimeError(
                "A previous VCS add-in probe is still pending against Access. "
                "This usually means Access is in VBA break mode or has an "
                "open modal dialog.  Resume execution in the VBE, dismiss "
                "any dialog, or restart Access, then retry."
            )
        
        addin_path_abs = os.path.abspath(self.addin_path)
        addin_lib_name = os.path.splitext(addin_path_abs)[0]
        api_function_name = f'{addin_lib_name}.API'
        
        result_box: dict[str, Any] = {}
        
        def worker() -> None:
            try:
                pythoncom.CoInitialize()
                try:
                    # When db_path is known, re-acquire Access via the ROT
                    # so the worker has its own apartment-local proxy --
                    # sharing the main thread's STA proxy across apartments
                    # either fails to marshal or serializes back to the main
                    # thread (which defeats the timeout).  When db_path is
                    # None we fall back to the main proxy: best-effort, the
                    # timeout may not fire reliably but behavior is no
                    # worse than before.
                    worker_app = (
                        self._find_access_in_rot(db_path) if db_path else app
                    )
                    if worker_app is None:
                        raise RuntimeError(
                            f"Cannot find Access instance for {db_path} "
                            f"from worker thread.  The Access application "
                            f"may have been closed."
                        )
                    worker_app.Run(api_function_name, "GetVCSVersion")
                    result_box["ok"] = True
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                result_box["error"] = exc
        
        thread = threading.Thread(
            target=worker, daemon=True, name="vcs-addin-probe"
        )
        cls._active_probe_thread = thread
        thread.start()
        thread.join(timeout=timeout_sec)
        
        if thread.is_alive():
            # Leave _active_probe_thread set so the next call's guard
            # detects the lingering worker and short-circuits with a clear
            # message.  The thread is daemon, so it will be reaped on
            # process exit if Access never responds.
            raise TimeoutError(
                f"VCS add-in probe timed out after {timeout_sec}s "
                f"(no response from Access).  Access is likely in VBA "
                f"break mode (check the VBE), blocked on a modal dialog, "
                f"or hung.  The probe thread will complete naturally once "
                f"Access responds -- no data was lost.  To recover, "
                f"resume execution in the VBE, dismiss any dialog, or "
                f"close and reopen the database."
            )
        
        cls._active_probe_thread = None
        
        if "error" in result_box:
            raise result_box["error"]
    
    @staticmethod
    def _find_access_in_rot(db_path: str):
        """
        Find an Access instance in the Running Object Table that has
        ``db_path`` open.
        
        Two-tier strategy lifted with light edits from db-inspector-mcp's
        ``_find_existing_instance``:
          * Tier 1 -- direct file moniker lookup (~1ms).  ``GetObject`` only
            inspects the ROT; it does NOT fall through to moniker binding,
            so there's no risk of triggering a file-open or password
            dialog.
          * Tier 2 -- enumerate all ROT entries, call ``CurrentDb`` on each
            (~10-50ms).  Catches Access instances that opened the database
            via ``OpenCurrentDatabase`` from a COM client and therefore
            don't appear under a file moniker.  Non-Access entries fail
            on ``CurrentDb`` and are silently skipped.
        
        Returns:
            An Access Application COM object, or None if no match found.
        """
        try:
            pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable(0)
        except Exception:
            return None
        
        # Tier 1: direct file moniker lookup
        try:
            moniker = pythoncom.CreateFileMoniker(os.path.abspath(db_path))
            obj = rot.GetObject(moniker)
            return win32com.client.Dispatch(
                obj.QueryInterface(pythoncom.IID_IDispatch)
            )
        except Exception:
            pass
        
        # Tier 2: enumerate ROT entries, check CurrentDb on each
        try:
            enum = rot.EnumRunning()
            target = os.path.normpath(os.path.abspath(db_path)).lower()
            while True:
                monikers = enum.Next(1)
                if not monikers:
                    break
                try:
                    obj = rot.GetObject(monikers[0])
                    dispatch = win32com.client.Dispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch)
                    )
                    cdb = dispatch.CurrentDb()
                    if cdb is not None:
                        cdb_path = os.path.normpath(
                            os.path.abspath(cdb.Name)
                        ).lower()
                        if cdb_path == target:
                            return dispatch
                except Exception:
                    continue
        except Exception:
            pass
        
        return None
    
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
        
        # Get Access application info first -- this works even if the add-in
        # itself is unhealthy, so callers always get at least the Access
        # version/bitness fields.
        access_info = get_access_info(app)
        result.update(access_info)
        
        # Load the add-in via the lifecycle gate.  Idempotent: if a caller
        # (e.g. validate_components) already loaded it, this is an early
        # return.  db_path=None: get_version_info doesn't have the DB path
        # in scope, so the probe falls back to the main thread's app proxy.
        try:
            self.load_addin(app)
        except Exception as load_error:
            result["vcs_version"] = None
            result["vcs_error"] = f"Failed to load VCS add-in: {load_error}"
            result["success"] = False
            return result
        
        # Get VCS add-in version
        try:
            vcs_version = self._call_addin_function("GetVCSVersion")
            result["vcs_version"] = vcs_version
        except Exception as e:
            result["vcs_version"] = None
            result["vcs_error"] = f"Failed to get VCS version: {e}"
            result["success"] = False
        
        return result
