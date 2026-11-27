"""
Access COM connection management.

Adapted from db-inspector-mcp with ownership tracking and cleanup patterns.
Manages COM connections to Access databases with proper resource cleanup.

Key Complexity Areas:
- Multiple connection strategies based on whether Access is already running
- Ownership tracking to avoid interfering with user's Access instances
- COM cleanup challenges and garbage collection issues
"""

import os
import re
import shutil
import sys
import threading
from typing import Any

try:
    import pythoncom
    import win32com.client
    from win32com.client import gencache
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

_GEN_PY_MODULE_RE = re.compile(r"win32com\.gen_py\.([^\s'\"]+)")


def _extract_gen_py_folder_from_error(message: str) -> str | None:
    """Return the gen_py type-library folder name embedded in an error message."""
    match = _GEN_PY_MODULE_RE.search(message)
    return match.group(1) if match else None


def _gen_py_folder_incomplete(folder_name: str) -> bool:
    """True when a gen_py folder exists but lacks the wrapper package files."""
    if not COM_AVAILABLE:
        return False
    folder_path = os.path.join(gencache.GetGeneratePath(), folder_name)
    if not os.path.isdir(folder_path):
        return False
    return not os.path.isfile(os.path.join(folder_path, "__init__.py"))


def _should_heal_gen_py_cache(exc: BaseException) -> bool:
    """Detect a stale pywin32 gen_py wrapper that can be rebuilt safely."""
    if not isinstance(exc, AttributeError):
        return False
    message = str(exc)
    if "win32com.gen_py." not in message:
        return False
    if "CLSIDToClassMap" in message:
        return True
    folder = _extract_gen_py_folder_from_error(message)
    return bool(folder and _gen_py_folder_incomplete(folder))


def _purge_gen_py_cache_folder(folder_name: str, *, prog_id: str) -> None:
    """Delete one corrupted gen_py type-library folder and drop cached imports."""
    from ..usage_logging import log_diagnostic_event

    folder_path = os.path.join(gencache.GetGeneratePath(), folder_name)
    removed = False
    if os.path.isdir(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)
        removed = True

    module_prefix = f"win32com.gen_py.{folder_name}"
    for name in list(sys.modules):
        if name == module_prefix or name.startswith(f"{module_prefix}."):
            del sys.modules[name]

    try:
        gencache.Rebuild()
    except Exception:
        pass

    log_diagnostic_event(
        "gen_py_cache_rebuilt",
        prog_id=prog_id,
        folder=folder_name,
        path=folder_path,
        removed=removed,
    )
    print(
        f"Rebuilt pywin32 gen_py cache for {prog_id} "
        f"(removed stale folder {folder_name})",
        file=sys.stderr,
    )


def ensure_dispatch(prog_id: str):
    """Early-bound COM dispatch with one-shot gen_py cache self-heal.

    ``gencache.EnsureDispatch`` can fail when a type-library folder under
    ``%TEMP%\\gen_py`` is left in a half-built state (for example only
    ``__pycache__`` remains).  Delete that folder, rebuild the cache, and
    retry once so MCP startup does not die on a recoverable local error.
    """
    if not COM_AVAILABLE:
        raise ImportError(
            "pywin32 is required for COM automation. "
            "Install it with: pip install pywin32"
        )

    for attempt in range(2):
        try:
            return gencache.EnsureDispatch(prog_id)
        except AttributeError as exc:
            if attempt == 0 and _should_heal_gen_py_cache(exc):
                folder = _extract_gen_py_folder_from_error(str(exc))
                if folder:
                    _purge_gen_py_cache_folder(folder, prog_id=prog_id)
                    continue
            raise


def _paths_match(a: str, b: str) -> bool:
    """Case-insensitive, normalised path comparison."""
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except (OSError, ValueError):
        return False


def leave_access_open() -> bool:
    """True unless ACCESS_VCS_LEAVE_ACCESS_OPEN is an explicit false.

    Default is to leave a COM-created Access window running so later
    calls reattach instead of spawning a new process. Set the env var
    to ``false`` / ``0`` / ``no`` to restore quit-per-call.
    """
    raw = os.environ.get("ACCESS_VCS_LEAVE_ACCESS_OPEN", "true").strip().lower()
    return raw not in ("0", "false", "no")


def ensure_access_visible(app) -> bool:
    """Show the Access window for any instance we drive a database through.

    A hidden instance strands the user: an error dialog, a VBA break, or a
    trust prompt blocks every later call with nothing on screen to explain
    why, and nobody can dismiss what they cannot see.  Visibility is
    therefore not optional for database work -- it applies to instances we
    created and to ones we attached to, since either can raise a dialog.

    Call this *after* the database is open.  Making the window visible can
    set ``UserControl``, and startup code reads that flag to decide whether
    a person is watching (see ``open_current_database``).

    Best-effort by design: failing to show the window is never worth failing
    an operation over.
    """
    try:
        app.Visible = True
        return True
    except Exception as e:
        print(f"Could not make Access visible: {e}", file=sys.stderr)
        return False


def open_current_database(app, db_path: str) -> None:
    """Open ``db_path`` as ``app``'s current database, as automation.

    ``OpenCurrentDatabase`` runs the target's AutoExec.  Startup code
    commonly branches on ``Application.UserControl`` to decide whether a
    person is watching -- the VCS add-in's own ``AutoRun`` opens its
    installer form when it believes one is, which strands the instance we
    are about to automate.  The flag gets set both deliberately (see
    ``AccessConnection._create_isolated_instance``, which needs the process
    to outlive a client teardown) and as a side effect of making the window
    visible, so lower it across the open and restore it afterwards.

    Use this in place of a bare ``OpenCurrentDatabase`` call: it also leaves
    the instance visible, which is the invariant every database-holding
    instance owes the user.
    """
    try:
        was_set = bool(app.UserControl)
    except Exception:
        was_set = False

    if was_set:
        try:
            app.UserControl = False
        except Exception:
            was_set = False
    try:
        app.OpenCurrentDatabase(db_path)
    finally:
        if was_set:
            try:
                app.UserControl = True
            except Exception:
                pass
    ensure_access_visible(app)


def access_instance_is_live(db_path: str) -> bool:
    """True when an Access process already has ``db_path`` open.

    Uses the Running Object Table only -- it does not bind a file moniker
    and therefore cannot spawn a new instance.
    """
    from ..addin_integration import VCSAddinIntegration

    return VCSAddinIntegration._find_access_in_rot(db_path) is not None


# Gate lanes for different databases connect at the same time. EnsureDispatch
# returns whichever Access registered itself as active, so a lane that has just
# started an instance but not yet opened its database could see that empty
# instance adopted by a second lane. Finding or starting an instance, through
# opening its database, is therefore one step at a time across lanes.
_INSTANCE_CREATION_LOCK = threading.Lock()
INSTANCE_CREATION_TIMEOUT_SEC = 60.0


DEFAULT_CLOSE_TIMEOUT_SEC = 10.0


def get_close_timeout() -> float:
    """Seconds to wait for a graceful Quit before force-terminating."""
    try:
        value = float(
            os.environ.get(
                "ACCESS_VCS_CLOSE_TIMEOUT_SEC", str(DEFAULT_CLOSE_TIMEOUT_SEC)
            )
        )
    except ValueError:
        return DEFAULT_CLOSE_TIMEOUT_SEC
    return value if value > 0 else DEFAULT_CLOSE_TIMEOUT_SEC


def _resolve_instance_for_record(record) -> Any:
    """Return the COM object for ``record``, or None if it is not that process.

    Both lookups resolve by path, which can land on a *different* Access
    instance than the one recorded -- two processes can hold the same file,
    and a moniker bind can even start a new one. The PID and creation stamp
    are therefore re-checked before the caller is allowed to close anything.
    """
    from .instance_registry import process_create_time
    from .process_qos import pid_from_access_app

    if not COM_AVAILABLE:
        return None

    app = None
    try:
        app = win32com.client.GetObject(record.database_path)
    except Exception:
        from ..addin_integration import VCSAddinIntegration

        app = VCSAddinIntegration._find_access_in_rot(record.database_path)
    if app is None:
        return None

    pid = pid_from_access_app(app)
    if pid != record.pid:
        return None
    if process_create_time(pid) != record.create_time:
        return None
    return app


def _quit_with_timeout(app, timeout_sec: float) -> bool:
    """Run CloseCurrentDatabase + Quit in a daemon thread with a hard timeout.

    This runs against an instance that has usually just failed a health
    probe, so the COM call may never return. A blocking call here would
    hang the Access gate thread and take the whole server down with it.

    ``app`` belongs to the calling thread's apartment; used directly from
    the worker every call fails with RPC_E_WRONG_THREAD, so the interface
    is marshalled across. An object that cannot be marshalled (a test
    double) is handed over as is.
    """
    done: dict[str, bool] = {}
    stream = None
    if COM_AVAILABLE:
        try:
            stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, app._oleobj_
            )
        except Exception:
            stream = None

    def worker() -> None:
        if COM_AVAILABLE:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        try:
            target = app
            if stream is not None:
                target = win32com.client.Dispatch(
                    pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch)
                )
            try:
                target.CloseCurrentDatabase()
            except Exception:
                pass
            target.Quit()
            target = None
            done["ok"] = True
        except Exception:
            done["ok"] = False
        finally:
            if COM_AVAILABLE:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    thread = threading.Thread(target=worker, daemon=True, name="vcs-access-quit")
    thread.start()
    thread.join(timeout=timeout_sec)
    return bool(done.get("ok"))


def _close_owned_record(record) -> dict[str, Any] | None:
    """Close one verified server-created instance. Returns a result or None.

    Graceful Quit first; a hung instance is force-terminated, which is only
    reachable for a PID whose identity was confirmed against the registry.
    """
    from .instance_registry import unregister_owned
    from .process_qos import process_is_alive, terminate_pid

    app = _resolve_instance_for_record(record)
    if app is None:
        # Either the process is gone, or the path now resolves to somebody
        # else's instance. Drop a stale claim; never close what we cannot
        # prove is ours.
        if process_is_alive(record.pid) is False:
            unregister_owned(record.pid, record.create_time)
        else:
            _log_instance_event(
                "owned_instance_close_skipped",
                pid=record.pid,
                database=record.database_path,
            )
        return None

    quit_ok = _quit_with_timeout(app, get_close_timeout())
    app = None
    terminated = False

    if process_is_alive(record.pid) is not False:
        terminated = terminate_pid(record.pid)
        _log_instance_event(
            "owned_instance_terminated",
            pid=record.pid,
            database=record.database_path,
            succeeded=terminated,
        )

    if process_is_alive(record.pid) is not False and not terminated:
        return None

    unregister_owned(record.pid, record.create_time)
    _log_instance_event(
        "owned_instance_closed",
        pid=record.pid,
        database=record.database_path,
        quit=quit_ok,
        terminated=terminated,
    )
    return {
        "pid": record.pid,
        "database_path": record.database_path,
        "quit": quit_ok,
        "terminated": terminated,
    }


def _log_instance_event(event: str, **fields: Any) -> None:
    try:
        from ..usage_logging import log_diagnostic_event

        log_diagnostic_event(event, **fields)
    except Exception:
        pass


def close_owned_instances_holding(
    paths: list[str] | tuple[str, ...],
    addin_paths: list[str] | tuple[str, ...] = (),
) -> list[dict[str, Any]]:
    """Quit server-created Access instances holding any of the given files.

    ``paths`` matches an instance's open database; ``addin_paths`` also
    matches a loaded add-in library, which locks its file even when the
    instance has an unrelated database open. User-owned windows are never
    touched. Used before rebuilds that must replace a file, and when
    recycling a stuck owned instance.
    """
    from .instance_registry import owned_records_for_paths

    closed: list[dict[str, Any]] = []
    for record in owned_records_for_paths(paths, addin_paths):
        result = _close_owned_record(record)
        if result is not None:
            closed.append(result)
    return closed


def recycle_owned_instance(database_path: str) -> bool:
    """Close a server-owned instance for ``database_path`` and open a fresh one.

    Returns True only when an owned instance was closed and a replacement
    was started. User-owned windows are never touched.
    """
    from .instance_registry import owned_records_for_paths

    if not owned_records_for_paths([database_path]):
        return False
    if not close_owned_instances_holding([database_path]):
        return False
    with AccessConnection(database_path) as conn:
        conn.connect()
    return True


class AccessConnection:
    """
    Manages COM connection to Access database.
    
    Uses ownership tracking to determine cleanup responsibility:
    - If connecting to user's existing Access instance, we don't close it
    - If we create our own Access instance, we're responsible for cleanup
      (quit only when ACCESS_VCS_LEAVE_ACCESS_OPEN is explicitly false, or
      when a rebuild / recycle must replace the file)
    """
    
    def __init__(self, db_path: str):
        """
        Initialize Access COM connection manager.
        
        Args:
            db_path: Path to Access database file (.accdb, .accda, .mdb)
        """
        if not COM_AVAILABLE:
            raise ImportError(
                "pywin32 is required for COM automation. "
                "Install it with: pip install pywin32"
            )
        
        self._db_path = db_path
        self._app = None
        self._db = None
        
        # Ownership tracking flags - determine cleanup responsibility
        self._owns_app = False  # True only if we created Access via Dispatch
        self._owns_db = False   # True only if we opened db via DBEngine
        self._db_opened_via_getobject = False  # True if connected to existing instance
        self._db_opened_as_current = False  # True if we called OpenCurrentDatabase
    
    def _get_access_app(self):
        """
        Get or create Access COM application.
        
        Uses a multi-step approach:
        1. Try GetObject(path) to connect to our database if already open
        2. Try EnsureDispatch -- but check whether it returned an existing
           user instance (has a different database open) vs a fresh one
        3. If EnsureDispatch attached to a user instance with a different
           DB, fall back to DispatchEx to create a guaranteed-isolated
           Access process
        
        IMPORTANT: Access can only have ONE database open at a time.
        - If user has OUR database open -> connect to their instance
        - If user has DIFFERENT database open -> create our OWN isolated instance
        - If no Access running -> create new instance
        
        The ownership check after EnsureDispatch prevents the dangerous
        scenario where close() calls Quit() on the user's Access window
        (ported from db-inspector-mcp's DispatchEx fallback pattern).

        Whichever route wins, the instance ends up visible: a dialog or a
        VBA break in a hidden window is unresolvable by the person who has
        to resolve it.  That happens last, once the database is open, for
        the ``UserControl`` reason in ``open_current_database``.  Instances
        we created also get ``UserControl = True`` afterwards so the window
        is a normal interactive Access app, not an automation ghost.
        """
        if self._app is None:
            created_now = False
            if not _INSTANCE_CREATION_LOCK.acquire(timeout=INSTANCE_CREATION_TIMEOUT_SEC):
                raise RuntimeError(
                    "Another Access instance is still starting, and its database may "
                    "be showing a dialog. Retry shortly, or check it with "
                    "vcs_capture_screenshot."
                )
            try:
                try:
                    # A moniker bind *launches* Access when nothing has the file
                    # open, so whether an instance was already running decides
                    # whether the process below is ours. The ROT lookup only
                    # inspects the table -- it cannot start anything itself.
                    was_running = access_instance_is_live(self._db_path)
                    self._app = win32com.client.GetObject(self._db_path)
                    self._db_opened_via_getobject = True
                    created_now = not was_running
                except Exception:
                    self._app, created_now = self._create_or_reuse_instance()
                    self._db_opened_via_getobject = False
                    # Never displace a user's database: only open on instances
                    # we just created (empty or isolated).
                    self._owns_app = created_now
                    self._open_as_current_database(self._app)
                self._resolve_ownership(created_now)
            finally:
                _INSTANCE_CREATION_LOCK.release()
            if self._owns_app:
                from .process_qos import prefer_full_power_app

                prefer_full_power_app(self._app)
            ensure_access_visible(self._app)
            if self._owns_app:
                self._ensure_owned_instance_interactive()
        return self._app

    def _resolve_ownership(self, created_now: bool) -> None:
        """Set ``_owns_app`` from creation or the durable PID registry.

        GetObject on a left-open window would otherwise reclassify a
        server-created instance as user-owned. The registry survives
        process restarts; create_time rejects PID reuse.
        """
        from .instance_registry import is_owned, process_create_time, register_owned
        from .process_qos import pid_from_access_app

        pid = pid_from_access_app(self._app)
        create_time = process_create_time(pid) if pid else None
        if created_now:
            self._owns_app = True
            if pid:
                register_owned(pid, self._db_path, create_time=create_time)
            return
        self._owns_app = bool(pid and is_owned(pid, create_time))

    def _ensure_owned_instance_interactive(self) -> None:
        """Give a COM-created Access instance a normal interactive window.

        ``Visible = True`` alone often leaves an automation-created process
        out of the desktop. ``UserControl`` after the database is open does
        not retrigger AutoExec -- that already ran with the flag down.
        Best-effort: never fail the operation over this.
        """
        try:
            self._app.UserControl = True
        except Exception as e:
            print(f"Could not set Access UserControl: {e}", file=sys.stderr)

    def _open_as_current_database(self, app):
        """Make ``self._db_path`` the instance's current database.

        ``GetObject(path)`` binds a file moniker, which resolves through the
        COM registration for the file's extension.  That registration opens
        .accdb and .mdb as the current database, but not .accda -- Access
        treats it as an add-in and moniker binding fails.  Without this
        fallback the instance has no current database, so add-in calls and
        the Running Object Table lookup in ``_find_access_in_rot`` (which
        matches on ``CurrentDb().Name``) both come up empty.

        A failed open is not fatal: leaving ``_db_opened_as_current`` False
        lets ``_get_current_db`` fall through to its DAO strategies, which
        still serve read-only callers when Access itself cannot take the
        file as the current database.
        """
        try:
            existing = app.CurrentDb()
        except Exception:
            existing = None

        if existing is not None:
            if _paths_match(existing.Name, self._db_path):
                self._db_opened_as_current = True
                return
            # A different database is open.  _create_or_reuse_instance
            # routes that case to an isolated process, so this only happens
            # on an instance we own; never displace a user's database.
            if not self._owns_app:
                return

        try:
            open_current_database(app, self._db_path)
            self._db_opened_as_current = True
        except Exception as e:
            print(
                f"[{os.path.basename(self._db_path)}] OpenCurrentDatabase "
                f"failed ({e}) -- falling back to DAO",
                file=sys.stderr,
            )

    def _create_or_reuse_instance(self):
        """Create or attach to an Access instance with correct ownership.

        EnsureDispatch("Access.Application") can silently return an
        existing user-owned instance instead of creating a new one.
        Returns ``(app, created_now)`` so the caller can register a new
        process or look the PID up in the durable ownership registry.

        Creation is decided by comparing the resulting PID against a
        snapshot taken beforehand, not by whether the instance has a
        database open. A user can leave Access sitting at an empty shell
        with no current database, and that window is still theirs.
        """
        from .process_qos import list_access_pids_or_none, pid_from_access_app

        before = list_access_pids_or_none()
        app = ensure_dispatch("Access.Application")

        try:
            existing_db = app.CurrentDb()
        except Exception:
            existing_db = None

        if existing_db is not None:
            if _paths_match(existing_db.Name, self._db_path):
                return app, False
            db_name = os.path.basename(self._db_path)
            other_name = os.path.basename(existing_db.Name)
            print(
                f"[{db_name}] EnsureDispatch returned instance with "
                f"'{other_name}' open -- launching isolated instance via DispatchEx",
                file=sys.stderr,
            )
            return self._create_isolated_instance(), True

        if before is None:
            # Cannot prove we created it; treat as the user's.
            return app, False
        pid = pid_from_access_app(app)
        return app, bool(pid and pid not in before)

    def _create_isolated_instance(self):
        """Create a guaranteed-isolated Access process via DispatchEx.

        Unlike EnsureDispatch, DispatchEx always spawns a new COM server
        process independent of any existing Access instance.  ``UserControl``
        keeps that process alive if the client goes away mid-operation; see
        ``open_current_database`` for why it has to be lowered again while a
        database opens.
        """
        app = win32com.client.DispatchEx("Access.Application")
        app.UserControl = True
        return app
    
    def _get_current_db(self):
        """
        Get database object for DAO operations.
        
        Strategy:
        1. If database was opened via GetObject (Access already had it open), use CurrentDb()
        2. Otherwise, use DBEngine.OpenDatabase() which is more reliable
        
        IMPORTANT: This method tracks whether we OWN the database connection:
        - If we use CurrentDb() (user's database), we do NOT own it
        - If we open via DBEngine, we OWN it and are responsible for closing
        """
        if self._db is None:
            app = self._get_access_app()
            
            # If Access already had the database open, CurrentDb() should work
            if self._db_opened_via_getobject or self._db_opened_as_current:
                try:
                    db = app.CurrentDb()
                    if db is not None:
                        self._db = db
                        self._owns_db = False  # User's database - do NOT close it
                        return self._db
                except Exception:
                    pass
            
            # Use DBEngine.OpenDatabase() - more reliable
            try:
                dbe = app.DBEngine
                # Open database in shared mode (Exclusive=False, ReadOnly=False)
                self._db = dbe.OpenDatabase(self._db_path, False, False)
                self._owns_db = True  # We opened this - we're responsible for closing
            except Exception:
                # Try read-only mode
                try:
                    dbe = app.DBEngine
                    self._db = dbe.OpenDatabase(self._db_path, False, True)
                    self._owns_db = True
                except Exception:
                    # Last resort: try direct DAO without Access
                    try:
                        dbe = win32com.client.Dispatch("DAO.DBEngine.120")
                        self._db = dbe.OpenDatabase(self._db_path, False, True)
                        self._owns_db = True
                    except Exception as e:
                        raise RuntimeError(
                            f"Failed to open database '{self._db_path}' via COM. "
                            f"Ensure the database file exists and is not corrupted. Error: {e}"
                        )
        
        return self._db
    
    def connect(self):
        """
        Connect to Access database via COM.
        
        Returns:
            Tuple of (Application, Database) objects
        """
        app = self._get_access_app()
        db = self._get_current_db()
        return app, db
    
    def get_app(self):
        """Get Access Application object."""
        return self._get_access_app()
    
    def get_db(self):
        """Get DAO Database object."""
        return self._get_current_db()
    
    def close(self):
        """
        Close the database connection and cleanup COM resources.
        
        IMPORTANT: This method respects ownership:
        - If we connected to an existing Access instance (via GetObject), we do NOT
          close Access or the database - the user is still using them!
        - If we created our own Access instance, we leave it running by default
          so later calls reattach. Set ACCESS_VCS_LEAVE_ACCESS_OPEN=false to
          restore quit-per-call. Rebuild and recycle close owned instances
          when they must replace the file.
        """
        # Only close the database if we opened it ourselves
        if self._db is not None and self._owns_db:
            try:
                self._db.Close()
            except Exception:
                pass
        self._db = None
        
        # Only quit Access if we created it ourselves AND the operator
        # opted back into quit-per-call.
        if self._app is not None and self._owns_app and not leave_access_open():
            from .instance_registry import process_create_time, unregister_owned
            from .process_qos import pid_from_access_app, process_is_alive

            pid = pid_from_access_app(self._app)
            create_time = process_create_time(pid) if pid else None
            try:
                self._app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                self._app.Quit()
            except Exception:
                pass
            # Only drop the claim once the process is confirmed gone. A
            # failed Quit that also forgot the PID would leave a live
            # instance nobody can close or recycle later.
            if pid and process_is_alive(pid) is False:
                unregister_owned(pid, create_time)
        self._app = None
        
        # Reset ownership flags
        self._owns_app = False
        self._owns_db = False
        self._db_opened_via_getobject = False
        self._db_opened_as_current = False
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
        return False
