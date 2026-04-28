"""
MCP tool definitions for msaccess-vcs-mcp.

This module provides database version control tools for AI assistants working with
Microsoft Access databases. All tools use the ``vcs_`` prefix to indicate they
control the VCS add-in, not the Access application itself.

**Getting Started Workflow:**
1. Use vcs_list_objects() to see what's in a database
2. Use vcs_export_database() to export all objects to source directory
3. Edit source files using your preferred tools
4. Use vcs_diff_database() to see what changed
5. Use vcs_import_objects() or vcs_rebuild_database() to apply changes

**Key Features:**
- Export Access objects to git-friendly text files
- Import objects from source files back into Access
- Rebuild entire databases from source
- Track changes between database and source
- Export/import individual objects by name and type
- Execute read-only SQL queries via the add-in's DAO connection
- Call existing VBA functions or run agent-generated VBA code
- Read/write add-in options for session-level configuration
- Read operations always available, write operations require permission
- Long-running operations support progress reporting via callbacks
"""

import asyncio
import functools
import inspect
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

from mcp.server.fastmcp import FastMCP, Context

from .access_com.connection import AccessConnection
from .access_com.dao_helpers import list_query_defs, list_table_defs
from .config import (
    get_config,
    get_callback_url,
    get_session_id,
    initialize_from_workspace,
    load_config,
)
from .addin_integration import VCSAddinIntegration
from .security import (
    validate_database_path,
    validate_export_directory,
    validate_source_directory,
    check_write_permission,
)
from .usage_logging import log_code_execution, log_diagnostic_event, with_logging


def _get_operation_manager():
    """Get the operation manager instance if available."""
    try:
        from .operation_manager import OperationManager
        return OperationManager.get_instance()
    except Exception:
        return None


def _is_async_available() -> bool:
    """Check if async callbacks are available."""
    return get_callback_url() is not None


def _check_database_busy(database_path: str) -> dict[str, Any] | None:
    """
    Check if a database has an operation in progress.
    
    Args:
        database_path: Path to the database
        
    Returns:
        Error dict if busy, None if available
    """
    op_manager = _get_operation_manager()
    if not op_manager:
        return None
    
    busy_status = op_manager.get_busy_status(database_path)
    if busy_status:
        return {
            "success": False,
            "error": busy_status["message"],
            "busy": True,
            "active_operation_id": busy_status["operation_id"],
            "active_command": busy_status["command"],
            "elapsed_seconds": busy_status["elapsed_seconds"],
            "hint": "Wait for the current operation to complete, or cancel it with vcs_cancel_operation()"
        }
    return None


# Create FastMCP server instance with proper metadata
mcp = FastMCP(
    name="msaccess-vcs-mcp",
    instructions=(
        "Microsoft Access version control MCP server. "
        "Export Access database objects to source files, import them back, "
        "rebuild databases from source, and track changes. "
        "All tools use the vcs_ prefix.\n\n"
        "**Recommended workflow:**\n"
        "1. Use vcs_export_database() to export all objects to source directory\n"
        "2. Edit source files using your preferred tools\n"
        "3. Use vcs_import_objects() to merge changes back into database\n"
        "4. Use vcs_rebuild_database() to create fresh database from source\n"
        "5. Use vcs_diff_database() to see what changed\n\n"
        "**Configuration:**\n"
        "Set ACCESS_VCS_DATABASE to your target database path.\n"
        "Set ACCESS_VCS_DISABLE_WRITES=true to prevent database modifications.\n\n"
        "**Rebuilding the VCS add-in:**\n"
        "The VCS add-in (`Version Control.accda`) itself cannot be rebuilt via MCP tools. "
        "Rebuilding it requires all Access instances to be closed, which would also close "
        "any database files the user currently has open. If you change add-in source files "
        "(e.g. `clsQueryComposer.cls`), ask the user to rebuild the add-in manually -- they "
        "must close every open Access window first, then run the add-in's own build "
        "process. After the user confirms the rebuild is complete, you can re-run "
        "verification steps (export/import/rebuild of target databases) through the MCP.\n\n"
        "**Logs:**\n"
        "Two JSON Lines streams (both prefixed `vcs-mcp-` so they don't collide with "
        "other tools' logs in a shared directory).\n"
        "1. `vcs-mcp-diagnostic.jsonl` -- always-on lifecycle log "
        "(server_start, startup_env_load, lazy_init_*). Lives at "
        "`~/.msaccess-vcs-mcp/logs/`. Opt out with "
        "ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true.\n"
        "2. `vcs-mcp-usage.jsonl` -- tool-call audit + code-execution events. "
        "Default-on; opt out with ACCESS_VCS_ENABLE_LOGGING=false. "
        "SQL/VBA bodies are recorded as `code_length` only (full `code` "
        "field requires ACCESS_VCS_LOG_CODE_CONTENT=true). Param keys "
        "matching password/secret/token/api_key/connection_string are "
        "auto-masked to `<redacted>`. Override location with "
        "ACCESS_VCS_LOG_DIR. Call `vcs_get_version_info()` to discover "
        "both active log paths."
    )
)


# ---------------------------------------------------------------------------
# Lazy .env discovery via MCP workspace roots
# ---------------------------------------------------------------------------
# When the server is configured at the *user* level (e.g. in
# ~/.cursor/mcp.json) the working directory is typically the user's home
# folder, not the project root. In that case the upward search in
# ``config._find_project_root`` won't find the project's ``.env``. We use the
# MCP ``roots/list`` call (available after the protocol handshake) to discover
# the workspace and load its ``.env`` lazily on the first tool call.
#
# The lazy-init handshake fires from inside ``vcs_tool``'s wrapper for
# *every* registered tool -- sync or async, with or without a declared
# ``ctx: Context`` parameter -- by retrieving the request context via
# ``mcp.get_context()``. FastMCP sets the request contextvar before
# dispatching any tool handler, so this works uniformly across the entire
# tool surface.

_lazy_init_attempted = False
# Set after the first post-init "already_attempted" diagnostic is emitted.
# Subsequent tool calls take the same fast path but skip the disk write so
# we don't spam the diagnostic log once per tool invocation. The first
# repeated call still emits an event (kept for test coverage and so an
# operator can confirm the cache is working).
_lazy_init_skip_logged = False


def _file_uri_to_path(uri: str) -> Path | None:
    """Convert a ``file://`` URI to a local ``Path``, or return ``None``."""
    parsed = urlparse(uri)
    if parsed.scheme != "file":
        return None
    raw_path = unquote(parsed.path)
    # On Windows, file:///C:/path -> /C:/path -- strip leading slash.
    if len(raw_path) >= 3 and raw_path[0] == "/" and raw_path[2] == ":":
        raw_path = raw_path[1:]
    return Path(raw_path)


def _resolve_session(ctx: Context | None):
    """Return ``ctx.session`` or ``None`` if ctx isn't request-scoped.

    ``Context.request_context`` raises ``ValueError`` ("Context is not
    available outside of a request") when no request is active. That can
    happen if ``mcp.get_context()`` is called outside an MCP request --
    for example, during a unit test that calls a wrapper directly. We
    treat that as "no session" and skip lazy init rather than crashing
    the tool call.
    """
    if ctx is None:
        return None
    try:
        return ctx.session
    except (ValueError, LookupError, AttributeError):
        return None


async def _ensure_env_loaded(ctx: Context | None) -> None:
    """Lazily load the project's ``.env`` from MCP workspace roots.

    Only fires once per process. Every branch emits a diagnostic event to
    the always-on diagnostic stream so an operator can answer "why didn't
    my .env get loaded?" without having to enable usage logging first
    (which is exactly the case where ``.env`` may have been missed).
    """
    global _lazy_init_attempted, _lazy_init_skip_logged
    if _lazy_init_attempted:
        # Emit the skip once so tests / operators can verify the cache is
        # functioning, then go silent for the lifetime of the process.
        if not _lazy_init_skip_logged:
            log_diagnostic_event("lazy_init_skipped", reason="already_attempted")
            _lazy_init_skip_logged = True
        return
    session = _resolve_session(ctx)
    if session is None:
        # ctx may be present but lack a request-scoped session (e.g. unit
        # tests calling the wrapper directly). Don't flip the flag -- a
        # later real request can still retry.
        if ctx is None:
            log_diagnostic_event("lazy_init_skipped", reason="no_ctx")
        else:
            log_diagnostic_event("lazy_init_skipped", reason="no_session")
        return
    _lazy_init_attempted = True

    from .config import _get_project_root
    startup_root = _get_project_root()
    startup_has_env = (startup_root / ".env").exists()
    log_diagnostic_event(
        "lazy_init_started",
        startup_root=str(startup_root),
        startup_root_has_env=startup_has_env,
    )
    if startup_has_env:
        log_diagnostic_event("lazy_init_skipped", reason="startup_env_present")
        return

    try:
        roots_result = await session.list_roots()
    except Exception as exc:
        log_diagnostic_event(
            "list_roots_failed",
            error=str(exc),
            error_type=type(exc).__name__,
        )
        return

    root_uris = [str(r.uri) for r in roots_result.roots]
    log_diagnostic_event("list_roots_response", roots=root_uris)

    for root in roots_result.roots:
        workspace = _file_uri_to_path(str(root.uri))
        if workspace is None:
            continue
        env_path = workspace / ".env"
        if not env_path.exists():
            continue
        try:
            log_diagnostic_event(
                "lazy_init_loaded",
                workspace=str(workspace),
                env_path=str(env_path),
            )
            initialize_from_workspace(workspace)
            return
        except Exception as exc:
            log_diagnostic_event(
                "lazy_init_load_failed",
                workspace=str(workspace),
                error=str(exc),
                error_type=type(exc).__name__,
            )

    log_diagnostic_event("lazy_init_no_env_in_roots", roots=root_uris)


def vcs_tool(name: str):
    """Register an MCP tool with lazy .env discovery, config reload, and logging.

    Composes these concerns in the correct order so that every tool call:
    1. Lazily discovers the project's ``.env`` via MCP workspace roots
       (only on the first call). The discovery uses ``mcp.get_context()``
       to access the active request session, so it works for *every*
       registered tool -- sync or async, with or without an explicit
       ``ctx: Context`` parameter. FastMCP sets the request contextvar
       before dispatching the handler, so the session is always available
       during tool execution.
    2. Refreshes configuration from ``.env`` (picks up edits made while
       the server is running).
    3. Initializes or re-initializes usage logging with the current env vars.
    4. Executes the tool body and logs the outcome.

    The wrapper is *always* an async coroutine. FastMCP detects this via
    ``inspect.iscoroutinefunction`` and awaits it correctly. Sync tool
    bodies are still invoked synchronously inside the wrapper -- there is
    no concurrency change relative to FastMCP's default sync handling.
    """
    def decorator(func):
        logged = with_logging(name)(func)
        is_async_body = inspect.iscoroutinefunction(func)

        @functools.wraps(func)
        async def with_refresh(*args, **kwargs):
            # ``mcp.get_context()`` returns a Context bound to the active
            # request even when the tool itself doesn't declare a ctx
            # parameter -- the lowlevel server sets the contextvar before
            # dispatching. Outside of a request this returns a Context
            # whose ``session`` access raises; ``_ensure_env_loaded``
            # handles that case via ``_resolve_session``.
            try:
                ctx = mcp.get_context()
            except Exception:
                ctx = None
            await _ensure_env_loaded(ctx)
            load_config()
            if is_async_body:
                return await logged(*args, **kwargs)
            return logged(*args, **kwargs)

        return mcp.tool()(with_refresh)
    return decorator


@vcs_tool("vcs_export_database")
async def vcs_export_database(
    database_path: str,
    output_dir: str,
    object_types: list[str] | None = None,
    full_export: bool = False,
    ctx: Context = None
) -> dict[str, Any]:
    """
    Export Access database objects to source files.
    
    Exports tables, queries, forms, reports, macros, and modules to 
    text-based files suitable for version control.
    
    This operation supports progress reporting - you'll receive updates
    as objects are exported.
    
    Examples:
        # Export entire database (quick/fast save - only changed objects)
        vcs_export_database("C:\\\\db.accdb", "C:\\\\src\\\\mydb")
        
        # Full export (all objects, regardless of changes)
        vcs_export_database("C:\\\\db.accdb", "C:\\\\src\\\\mydb", full_export=True)
        
        # Export only queries and modules
        vcs_export_database(
            "C:\\\\db.accdb", 
            "C:\\\\src\\\\mydb",
            object_types=["queries", "modules"]
        )
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        output_dir: Directory to export source files to
        object_types: Optional list of types to export: 
            ["tables", "queries", "forms", "reports", "modules", "macros"]
            If None, exports all types
        full_export: If True, export all objects; if False (default), only export changed objects
    
    Returns:
        Dictionary with:
        - exported_count: Number of objects exported
        - export_path: Path where files were written
        - objects_by_type: Breakdown of what was exported
        - errors: List of any errors encountered
    """
    try:
        # Validate paths
        db_path = validate_database_path(database_path)
        export_path = validate_export_directory(output_dir, allow_create=True)
        
        # Get configuration
        config = get_config()
        callback_url = get_callback_url()
        op_manager = _get_operation_manager()
        
        # Check if database is already busy
        busy_error = _check_database_busy(str(db_path))
        if busy_error:
            return busy_error
        
        # Determine export command
        vba_only = object_types and set(t.lower().strip() for t in object_types) <= {"module", "modules"}
        if vba_only:
            command = "ExportVBA"
        elif full_export:
            command = "FullExport"
        else:
            command = "Export"
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            # Initialize add-in integration; load_addin probes the add-in
            # with a hard timeout, surfacing dialog-blocked / VBA-break /
            # hung Access instances as a clear lifecycle error before any
            # real work is dispatched.
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            try:
                addin.load_addin(app, db_path=str(db_path))
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Add-in not responsive (may have a dialog open): {e}",
                    "exported_count": 0,
                    "export_path": str(export_path),
                    "objects_by_type": {},
                    "hint": "Check if Access has any open dialogs or message boxes"
                }
            
            # Check if async export is available
            if callback_url and op_manager:
                # Ensure operation manager uses the correct event loop (FastMCP's loop)
                op_manager.set_event_loop(asyncio.get_running_loop())
                # Use async path with progress callbacks
                operation_id, queue = op_manager.register_operation(
                    database_path=str(db_path),
                    command=command
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                
                try:
                    # Call async API (Export/ExportVBA don't take arguments - they use VCS options)
                    async_result = addin.call_async(callback_info, command)
                    
                    completion = None
                    if async_result.get("sync"):
                        # VBA returned sync result
                        pass  # Fall through to count objects
                    elif async_result.get("async"):
                        # Wait for completion with progress reporting
                        timeout_ms = async_result.get("timeout_ms", 300000)
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=ctx,  # Pass context for progress reporting
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if not completion.get("success"):
                            return {
                                "success": False,
                                "error": completion.get("error", "Export failed"),
                                "exported_count": 0,
                                "export_path": str(export_path),
                                "objects_by_type": {},
                            }
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    # Perform sync export
                    if vba_only:
                        result = addin.export_vba(str(db_path), str(export_path))
                    else:
                        result = addin.export_source(str(db_path), str(export_path))
                    
                    if not result["success"]:
                        return {
                            "success": False,
                            "error": result["message"],
                            "exported_count": 0,
                            "export_path": str(export_path),
                            "objects_by_type": {},
                        }
            else:
                # Use sync path (no callbacks available)
                completion = None
                if vba_only:
                    result = addin.export_vba(str(db_path), str(export_path))
                else:
                    result = addin.export_source(str(db_path), str(export_path))
                
                if not result["success"]:
                    return {
                        "success": False,
                        "error": result["message"],
                        "exported_count": 0,
                        "export_path": str(export_path),
                        "objects_by_type": {},
                    }
            
            # Get log path and messages from callback result
            log_path = None
            log_messages = None
            if completion:
                log_path = completion.get("log_path")
                log_messages = completion.get("log_messages")
            
            if not log_path:
                # Fallback to legacy location
                legacy_path = os.path.join(str(export_path), "Export.log")
                if os.path.exists(legacy_path):
                    log_path = legacy_path
            
            return {
                "success": True,
                "export_path": str(export_path),
                "log_path": log_path if (log_path and os.path.exists(log_path)) else None,
                "messages": log_messages,
            }
    
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "exported_count": 0,
            "export_path": None,
            "objects_by_type": {},
        }


@vcs_tool("vcs_list_objects")
async def vcs_list_objects(
    database_path: str,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    List all objects in an Access database.
    
    Provides an inventory of tables, queries, forms, reports, 
    modules, and macros.
    
    Examples:
        # List all objects
        vcs_list_objects("C:\\\\db.accdb")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
    
    Returns:
        Dictionary with object lists by type:
        - tables: List of table names
        - queries: List of query names with types
        - modules: List of module names
        - forms: List of form names (future)
        - reports: List of report names (future)
        - macros: List of macro names (future)
    """
    try:
        # Validate path
        db_path = validate_database_path(database_path)
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            # List tables
            tables = list_table_defs(db)
            table_names = [t["name"] for t in tables]
            
            # List queries
            queries = list_query_defs(db)
            
            # List modules
            module_names = []
            try:
                vbe = app.VBE
                vb_project = vbe.ActiveVBProject
                for component in vb_project.VBComponents:
                    if component.Type in (1, 2):  # Standard and class modules
                        module_names.append(component.Name)
            except Exception as e:
                print(f"Warning: Could not list modules: {e}")
            
            return {
                "success": True,
                "database": str(db_path),
                "tables": table_names,
                "queries": queries,
                "modules": module_names,
                "forms": [],  # Future implementation
                "reports": [],  # Future implementation
                "macros": [],  # Future implementation
            }
    
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "database": database_path,
            "tables": [],
            "queries": [],
            "modules": [],
        }


@vcs_tool("vcs_diff_database")
async def vcs_diff_database(
    database_path: str,
    source_dir: str,
    show_details: bool = False,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Compare database objects against source files.
    
    Shows which objects have changed, been added, or deleted
    compared to the source directory.
    
    Examples:
        # Basic diff
        vcs_diff_database("C:\\\\db.accdb", "C:\\\\src\\\\mydb")
        
        # Detailed diff with line-by-line comparison
        vcs_diff_database(
            "C:\\\\db.accdb",
            "C:\\\\src\\\\mydb",
            show_details=True
        )
    
    Args:
        database_path: Path to Access database
        source_dir: Directory containing source files
        show_details: If True, show detailed diff of changes
    
    Returns:
        Dictionary with:
        - modified_objects: List of changed objects
        - new_in_db: Objects in database but not in source
        - new_in_source: Objects in source but not in database
        - unchanged_objects: Objects that match
        - details: (if show_details=True) Detailed differences
    """
    try:
        # Validate paths
        db_path = validate_database_path(database_path)
        src_path = validate_source_directory(source_dir)
        
        # Get objects from database
        db_objects = vcs_list_objects(str(db_path))
        if not db_objects.get("success"):
            return db_objects
        
        # Get objects from source directory
        source_queries = set()
        query_dir = src_path / "queries"
        if query_dir.exists():
            source_queries = {f.stem for f in query_dir.glob("*.sql")}
        
        source_modules = set()
        module_dir = src_path / "modules"
        if module_dir.exists():
            source_modules = {f.stem for f in module_dir.glob("*.bas")}
        
        # Compare
        db_queries = {q["name"] for q in db_objects["queries"]}
        db_modules = set(db_objects["modules"])
        
        result = {
            "success": True,
            "queries": {
                "new_in_db": list(db_queries - source_queries),
                "new_in_source": list(source_queries - db_queries),
                "in_both": list(db_queries & source_queries),
            },
            "modules": {
                "new_in_db": list(db_modules - source_modules),
                "new_in_source": list(source_modules - db_modules),
                "in_both": list(db_modules & source_modules),
            },
        }
        
        if show_details:
            result["note"] = "Detailed line-by-line diff not yet implemented"
        
        return result
    
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
        }


@vcs_tool("vcs_import_objects")
async def vcs_import_objects(
    database_path: str,
    source_dir: str,
    object_types: list[str] | None = None,
    overwrite: bool = False,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Import objects from source files into Access database.
    
    Merges source files back into database. Can update existing objects
    or add new ones.
    
    This operation supports progress reporting - you'll receive updates
    as objects are imported.
    
    Examples:
        # Import all objects
        vcs_import_objects("C:\\\\db.accdb", "C:\\\\src\\\\mydb", overwrite=True)
        
        # Import only queries
        vcs_import_objects(
            "C:\\\\db.accdb",
            "C:\\\\src\\\\mydb",
            object_types=["queries"],
            overwrite=True
        )
    
    Args:
        database_path: Path to Access database
        source_dir: Directory containing source files
        object_types: Optional types to import (default: all)
        overwrite: If True, replace existing objects; if False, skip
    
    Returns:
        Dictionary with import results and any errors
    """
    config = get_config()
    callback_url = get_callback_url()
    op_manager = _get_operation_manager()
    
    try:
        # Check if writes are disabled
        check_write_permission(config)
        
        # Validate paths
        db_path = validate_database_path(database_path)
        src_path = validate_source_directory(source_dir)
        
        # Check if database is already busy
        busy_error = _check_database_busy(str(db_path))
        if busy_error:
            return busy_error
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            # Initialize add-in integration; load_addin probes the add-in
            # with a hard timeout, surfacing dialog-blocked / VBA-break /
            # hung Access instances as a clear lifecycle error before any
            # real work is dispatched.
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            try:
                addin.load_addin(app, db_path=str(db_path))
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Add-in not responsive (may have a dialog open): {e}",
                    "imported_count": 0,
                    "hint": "Check if Access has any open dialogs or message boxes"
                }
            
            # Check if async import is available
            if callback_url and op_manager:
                # Ensure operation manager uses the correct event loop (FastMCP's loop)
                op_manager.set_event_loop(asyncio.get_running_loop())
                
                # Use async path with progress callbacks
                operation_id, queue = op_manager.register_operation(
                    database_path=str(db_path),
                    command="MergeBuild"
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                
                try:
                    # Call async API for MergeBuild
                    completion = None
                    async_result = addin.call_async(callback_info, "MergeBuild")
                    
                    if async_result.get("async"):
                        # Wait for completion with progress reporting
                        timeout_ms = async_result.get("timeout_ms", 300000)
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=None,
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if not completion.get("success"):
                            return {
                                "success": False,
                                "error": completion.get("error", "Import failed"),
                                "imported_count": 0,
                            }
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.merge_build(str(db_path), str(src_path))
                    if not result["success"]:
                        return {
                            "success": False,
                            "error": result["message"],
                            "imported_count": 0,
                        }
            else:
                # Use sync path
                completion = None
                result = addin.merge_build(str(db_path), str(src_path))
                if not result["success"]:
                    return {
                        "success": False,
                        "error": result["message"],
                        "imported_count": 0,
                    }
            
            # Get log path from callback result, or use legacy path as fallback
            log_path = None
            if completion:
                log_path = completion.get("log_path")
            if not log_path:
                legacy_path = os.path.join(str(src_path), "Build.log")
                if os.path.exists(legacy_path):
                    log_path = legacy_path
            
            return {
                "success": True,
                "imported_count": "See log for details",
                "database_path": str(db_path),
                "source_dir": str(src_path),
                "log_path": log_path if (log_path and os.path.exists(log_path)) else None,
            }
    
    except PermissionError as e:
        return {
            "success": False,
            "error": str(e),
            "imported_count": 0,
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "imported_count": 0,
        }


@vcs_tool("vcs_rebuild_database")
async def vcs_rebuild_database(
    source_dir: str,
    output_path: str,
    template_path: str | None = None,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Build a complete Access database from source files.
    
    Creates a fresh database and imports all objects from source.
    Useful for clean builds and distribution.
    
    This operation supports progress reporting - you'll receive updates
    as the database is built.
    
    Examples:
        # Rebuild from source
        vcs_rebuild_database("C:\\\\src\\\\mydb", "C:\\\\output\\\\rebuilt.accdb")
        
        # Rebuild using template
        vcs_rebuild_database(
            "C:\\\\src\\\\mydb",
            "C:\\\\output\\\\rebuilt.accdb",
            template_path="C:\\\\templates\\\\blank.accdb"
        )
    
    Args:
        source_dir: Directory containing source files
        output_path: Path for new database file
        template_path: Optional template database to start from
    
    Returns:
        Dictionary with build results
    """
    config = get_config()
    callback_url = get_callback_url()
    op_manager = _get_operation_manager()
    
    try:
        # Check if writes are disabled
        check_write_permission(config)
        
        # Validate source directory
        src_path = validate_source_directory(source_dir)
        
        # Check if target database is already busy (if it exists)
        if output_path:
            busy_error = _check_database_busy(output_path)
            if busy_error:
                return busy_error
        
        # We need Access running to call the add-in, but no database is
        # open yet -- the add-in's build process creates it.  Create a
        # bare Access instance (no AccessConnection, which requires a
        # database path) and manage its lifecycle with try/finally.
        from win32com.client import gencache
        app = gencache.EnsureDispatch("Access.Application")
        
        try:
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=None)
            
            # Determine command
            command = "BuildAs" if output_path else "Build"
            
            # Check if async build is available
            if callback_url and op_manager:
                # Ensure operation manager uses the correct event loop (FastMCP's loop)
                op_manager.set_event_loop(asyncio.get_running_loop())
                
                # Use async path with progress callbacks
                operation_id, queue = op_manager.register_operation(
                    database_path=output_path or str(src_path),
                    command=command
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                
                try:
                    completion = None
                    if command == "Build":
                        async_result = addin.call_async(callback_info, command, str(src_path))
                    else:
                        async_result = addin.call_async(callback_info, command)
                    
                    if async_result.get("async"):
                        timeout_ms = async_result.get("timeout_ms", 600000)  # 10 min for builds
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=None,
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if not completion.get("success"):
                            return {
                                "success": False,
                                "error": completion.get("error", "Build failed"),
                                "output_path": None,
                            }
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.build_from_source(str(src_path), output_path)
                    if not result["success"]:
                        return {
                            "success": False,
                            "error": result["message"],
                            "output_path": None,
                        }
            else:
                # Use sync path
                completion = None
                result = addin.build_from_source(str(src_path), output_path)
                if not result["success"]:
                    return {
                        "success": False,
                        "error": result["message"],
                        "output_path": None,
                    }
            
            # Get log path from callback result, or use legacy path as fallback
            log_path = None
            if completion:
                log_path = completion.get("log_path")
            if not log_path:
                legacy_path = os.path.join(str(src_path), "Build.log")
                if os.path.exists(legacy_path):
                    log_path = legacy_path
            
            return {
                "success": True,
                "output_path": output_path,
                "source_dir": str(src_path),
                "log_path": log_path if (log_path and os.path.exists(log_path)) else None,
            }
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    
    except PermissionError as e:
        return {
            "success": False,
            "error": str(e),
            "output_path": None,
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "output_path": None,
        }


@vcs_tool("vcs_get_version_info")
async def vcs_get_version_info(
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Get version information for MCP server, MSAccess VCS add-in, and Access application.
    
    Returns comprehensive version information useful for troubleshooting 
    compatibility issues, including:
    - MCP server version
    - VCS add-in version
    - Access application version
    - Access bitness (32-bit or 64-bit)
    - Configured target database path
    - Add-in file path
    - Callback server status (for async operations)
    
    Examples:
        # Get version information
        vcs_get_version_info()
    
    Returns:
        Dictionary with:
        - success: Boolean indicating if info was retrieved
        - mcp_version: Version of the MCP server (e.g., "0.1.0")
        - vcs_version: Version of the VCS add-in (e.g., "4.1.4")
        - access_version: Access application version (e.g., "16.0")
        - bitness: "32-bit" or "64-bit"
        - target_database: Configured database path from ACCESS_VCS_DATABASE
        - addin_path: Path to the VCS add-in file
        - callback_url: URL for async callbacks (None if not available)
        - async_available: Boolean indicating if async operations are supported
        - usage_log_path: Path to ``vcs-mcp-usage.jsonl`` (None if usage
          logging is disabled)
        - diagnostic_log_path: Path to ``vcs-mcp-diagnostic.jsonl`` (None
          if the always-on diagnostic stream has been opted out)
        - log_code_content: Boolean -- whether ``code_execution`` events
          record the full SQL/VBA body or only ``code_length``
        - errors: List of validation errors
        - warnings: List of validation warnings
    """
    from .usage_logging import (
        get_diagnostic_log_path,
        get_log_file_path,
        is_diagnostic_logging_enabled,
        is_logging_enabled,
    )
    from .validation import get_version_info_safe

    result = get_version_info_safe()

    callback_url = get_callback_url()
    op_manager = _get_operation_manager()

    result["callback_url"] = callback_url
    result["async_available"] = bool(callback_url and op_manager)

    usage_path = get_log_file_path() if is_logging_enabled() else None
    diag_path = get_diagnostic_log_path() if is_diagnostic_logging_enabled() else None
    result["usage_log_path"] = str(usage_path) if usage_path else None
    result["diagnostic_log_path"] = str(diag_path) if diag_path else None
    result["log_code_content"] = (
        os.getenv("ACCESS_VCS_LOG_CODE_CONTENT", "false").lower() == "true"
    )

    if not callback_url:
        result["warnings"] = result.get("warnings", []) + [
            "Callback server not running - async operations will fall back to sync mode"
        ]

    return result


@vcs_tool("vcs_cancel_operation")
def vcs_cancel_operation(operation_id: str) -> dict[str, Any]:
    """
    Cancel a running async operation.
    
    Requests cancellation of a long-running operation (export, build, etc.).
    The VBA add-in will detect the cancellation request during its next
    DoEvents cycle and abort the operation.
    
    Note: Cancellation is cooperative - the operation will stop at the next
    safe point, not immediately. The operation may take a few seconds to
    respond depending on what it's doing.
    
    Examples:
        # Cancel an export operation
        vcs_cancel_operation("a1b2c3d4-5678-90ab-cdef-1234567890ab")
    
    Args:
        operation_id: The UUID of the operation to cancel
    
    Returns:
        Dictionary with:
        - success: Boolean indicating if cancellation was requested
        - operation_id: The operation ID that was cancelled
        - message: Status message
    """
    op_manager = _get_operation_manager()
    
    if not op_manager:
        return {
            "success": False,
            "error": "Callback system not available",
            "operation_id": operation_id,
        }
    
    # Request cancellation
    cancelled = op_manager.request_cancel(operation_id)
    
    if cancelled:
        # Also try to notify VBA immediately via COM (best effort)
        # This is non-blocking - VBA will also poll /cancel-status
        try:
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            # Attempt COM call to Cancel API - may block if Access is busy
            # Using a short timeout would be ideal but COM doesn't support that
            # So we just do best-effort here
            # addin.call_sync("Cancel", operation_id)  # Uncomment when VBA side is ready
        except Exception:
            # COM call failed - that's OK, VBA will poll
            pass
        
        return {
            "success": True,
            "operation_id": operation_id,
            "message": "Cancellation requested. Operation will stop at next safe point.",
        }
    else:
        return {
            "success": False,
            "operation_id": operation_id,
            "error": "Operation not found or already completed",
        }


@vcs_tool("vcs_check_vba_compiled")
def vcs_check_vba_compiled(database_path: str) -> dict[str, Any]:
    """
    Check if VBA code in an Access database is compiled.
    
    Returns the compilation state without attempting to compile.
    Useful for establishing a baseline before making code changes.
    
    Examples:
        # Check compilation state
        result = vcs_check_vba_compiled("C:\\\\db.accdb")
        if result["compiled"]:
            print("Code is compiled")
        else:
            print("Code is not compiled (may need compilation)")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
    
    Returns:
        Dictionary with:
        - success: Boolean indicating if the check completed successfully
        - compiled: Boolean - True if project is compiled, False otherwise
        - error: Error message if check failed
    """
    try:
        # Validate path
        db_path = validate_database_path(database_path)
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            # Call IsVBACompiled API
            is_compiled = addin.call_sync("IsVBACompiled")
            
            return {
                "success": True,
                "compiled": bool(is_compiled),
            }
    
    except Exception as e:
        return {
            "success": False,
            "compiled": False,
            "error": str(e),
        }


@vcs_tool("vcs_compile_vba")
def vcs_compile_vba(
    database_path: str,
    suppress_warnings: bool = False
) -> dict[str, Any]:
    """
    Compile all VBA modules in an Access database and return success status.
    
    Attempts to compile all VBA code in the database. Returns True if compilation
    succeeded (project is compiled), False if compilation failed.
    
    **Important:** If compilation fails, do not proceed with code edits as there
    are existing compilation errors that must be fixed first.
    
    Examples:
        # Compile VBA code
        result = vcs_compile_vba("C:\\\\db.accdb", suppress_warnings=True)
        if result["success"]:
            print("Compilation successful!")
        else:
            print("Compilation failed - do not proceed with edits")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        suppress_warnings: If True, suppress message boxes during compilation.
                         Warning: If code crashes, warnings may remain disabled.
    
    Returns:
        Dictionary with:
        - success: Boolean - True if compilation succeeded (project is compiled),
                  False if compilation failed
        - error: Error message if compilation check failed
    """
    try:
        # Validate path
        db_path = validate_database_path(database_path)
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            # Call CompileVBA API with suppress_warnings parameter
            compile_result = addin.call_sync("CompileVBA", suppress_warnings)
            
            return {
                "success": bool(compile_result),
            }
    
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
        }


@vcs_tool("vcs_export_object")
def vcs_export_object(
    database_path: str,
    object_type: str,
    object_name: str = ""
) -> dict[str, Any]:
    """
    Export a single database object or component type to source files.
    
    Exports one object to its source file representation. Much faster than a
    full database export when you only need to refresh one object.
    
    Accepts singular or plural type names. For single-file component types
    (like vbe_project or db_property), the object_name is ignored.
    
    Examples:
        vcs_export_object("C:\\\\db.accdb", "query", "qryCustomers")
        vcs_export_object("C:\\\\db.accdb", "form", "frmMain")
        vcs_export_object("C:\\\\db.accdb", "module", "modUtils")
        vcs_export_object("C:\\\\db.accdb", "imex_spec", "MyImportSpec")
        vcs_export_object("C:\\\\db.accdb", "vbe_project")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        object_type: Type of object. Core types: "query", "form", "report",
            "module", "table", "macro". Extended types: "table_data",
            "table_data_macro", "relation", "saved_spec", "imex_spec",
            "theme", "shared_image", "vbe_form", "command_bar".
            Single-file types (no name needed): "vbe_project", "vbe_reference",
            "project", "connection", "db_property", "project_property",
            "document", "hidden_attribute", "nav_pane_group".
            Plural forms and common aliases are also accepted.
        object_name: Name of the object to export. Required for multi-file
            types, ignored for single-file types.
    
    Returns:
        Dictionary with success status, file path, and any errors
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("ExportObject", object_type, object_name)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_import_object")
def vcs_import_object(
    database_path: str,
    object_type: str,
    object_name: str = ""
) -> dict[str, Any]:
    """
    Import a single object or component type from source files into the database.
    
    Loads one object from its source file back into the Access database.
    The source file must exist in the project's export folder.
    
    Accepts singular or plural type names. For single-file component types
    (like vbe_project or db_property), the object_name is ignored.
    
    Examples:
        vcs_import_object("C:\\\\db.accdb", "query", "qryCustomers")
        vcs_import_object("C:\\\\db.accdb", "module", "modUtils")
        vcs_import_object("C:\\\\db.accdb", "imex_spec", "MyImportSpec")
        vcs_import_object("C:\\\\db.accdb", "vbe_project")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        object_type: Type of object. Core types: "query", "form", "report",
            "module", "table", "macro". Extended types: "table_data",
            "table_data_macro", "relation", "saved_spec", "imex_spec",
            "theme", "shared_image", "vbe_form", "command_bar".
            Single-file types (no name needed): "vbe_project", "vbe_reference",
            "project", "connection", "db_property", "project_property",
            "document", "hidden_attribute", "nav_pane_group".
            Plural forms and common aliases are also accepted.
        object_name: Name of the object to import. Required for multi-file
            types, ignored for single-file types.
    
    Returns:
        Dictionary with success status and any errors
    """
    try:
        config = get_config()
        check_write_permission(config)
        
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("ImportObject", object_type, object_name)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except PermissionError as e:
        return {"success": False, "error": str(e)}
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_execute_sql")
def vcs_execute_sql(
    database_path: str,
    sql: str,
    max_rows: int = 100
) -> dict[str, Any]:
    """
    Execute a read-only SELECT query against the database via the add-in's DAO connection.
    
    Runs a SELECT statement and returns the results as JSON rows. Only SELECT
    statements are allowed -- INSERT, UPDATE, DELETE, and DDL are rejected.
    
    Useful for inspecting MSysObjects, MSysQueries, table data, and query results
    without needing a separate database connection.
    
    Examples:
        vcs_execute_sql("C:\\\\db.accdb", "SELECT Name, Type FROM MSysObjects WHERE Type=5")
        vcs_execute_sql("C:\\\\db.accdb", "SELECT * FROM Customers", max_rows=50)
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        sql: SELECT statement to execute
        max_rows: Maximum number of rows to return (default: 100)
    
    Returns:
        Dictionary with rows, rowCount, and truncated flag
    """
    try:
        db_path = validate_database_path(database_path)
        log_code_execution("vcs_execute_sql", str(db_path), sql, code_type="sql")
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("ExecuteSQL", sql, max_rows)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_call_vba")
def vcs_call_vba(
    database_path: str,
    function_name: str,
    args: list[str] | None = None
) -> dict[str, Any]:
    """
    Call an existing public VBA function by name.
    
    Invokes a function that already exists in the database or a loaded library
    via Application.Run. Lighter weight than vcs_run_vba since there is no
    temp module creation or compilation step.
    
    Examples:
        vcs_call_vba("C:\\\\db.accdb", "MyModule.GetQuerySQL", ["qryCustomers"])
        vcs_call_vba("C:\\\\db.accdb", "Version Control.API", ["GetVCSVersion"])
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        function_name: Fully qualified function name (e.g., "ModuleName.FunctionName")
        args: Optional list of string arguments to pass to the function
    
    Returns:
        Dictionary with the function's return value or error
    """
    try:
        db_path = validate_database_path(database_path)
        call_args = args or []
        call_description = function_name
        if call_args:
            call_description += f"({', '.join(repr(a) for a in call_args)})"
        log_code_execution("vcs_call_vba", str(db_path), call_description, code_type="vba_call")
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            try:
                if len(call_args) == 0:
                    result = app.Run(function_name)
                elif len(call_args) == 1:
                    result = app.Run(function_name, call_args[0])
                elif len(call_args) == 2:
                    result = app.Run(function_name, call_args[0], call_args[1])
                elif len(call_args) == 3:
                    result = app.Run(function_name, call_args[0], call_args[1], call_args[2])
                else:
                    return {
                        "success": False,
                        "error": "Maximum 3 arguments supported for vcs_call_vba"
                    }
            except Exception as e:
                return {
                    "success": False,
                    "error": f"VBA function call failed: {e}",
                    "function": function_name,
                }
            
            return {
                "success": True,
                "result": str(result) if result is not None else None,
                "function": function_name,
            }
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_run_vba")
def vcs_run_vba(
    database_path: str,
    code: str
) -> dict[str, Any]:
    """
    Execute agent-generated VBA code in a temporary module.
    
    Sends a block of VBA code to the add-in's RunVBA method, which handles the
    full lifecycle: creates a temp module, wraps the code in a function with error
    handling, compiles the project to validate, executes, captures the result,
    removes the temp module, and returns structured JSON.
    
    **Requires McpAllowRunVBA option to be enabled** (default: off).
    The user must enable this manually in the VCS Options form.
    
    The agent's code should set the function return value via the
    MCP_TempFunction identifier. Example:
        Dim result As String
        result = CurrentDb.QueryDefs("qryCustomers").SQL
        MCP_TempFunction = result
    
    **Line-number debugging:**
    The add-in auto-prepends sequential 1-based VBA line numbers to every
    executable line in `code` before running it. When a runtime error fires
    inside the wrapper, the response includes an `errorLine` field whose
    value equals the 1-based line number within the `code` string you
    submitted. The counter advances on every physical input line (blanks,
    comments, and `_` continuations included) even though only executable
    lines actually carry a number, so `errorLine: 7` means "line 7 of what
    I sent" -- you can index into your own `code` directly.
    
    Default behavior is to capture the LAST runtime error (the wrapper uses
    `On Error Resume Next` so all statements run). For a richer pattern that
    collects every failing line in one round-trip, use an explicit handler:
    
        Dim col As New Collection
        On Error GoTo H
        CurrentDb.Execute "DELETE * FROM tblA"
        CurrentDb.Execute "INSERT INTO tblB SELECT * FROM nope"
        CurrentDb.Execute "UPDATE tblC SET x = 1"
        MCP_TempFunction = "errors=" & col.Count
        Exit Function
        H: col.Add Erl & ": " & Err.Number & " " & Err.Description
        Resume Next
    
    Each `Erl` value inside the handler is meaningful (and matches an
    `errorLine` you would have seen) because the wrapper auto-numbered
    every line for you.
    
    Examples:
        vcs_run_vba("C:\\\\db.accdb", "MCP_TempFunction = CurrentDb.TableDefs.Count")
        vcs_run_vba("C:\\\\db.accdb", "Dim qd As DAO.QueryDef\\nSet qd = CurrentDb.QueryDefs(\\"qryTest\\")\\nMCP_TempFunction = qd.SQL")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        code: VBA code to execute (statements, not just an expression)
    
    Returns:
        Dictionary with `success`, `result`, and on failure `error`,
        `errorNumber`, and `errorLine` (the 1-based line in `code` that
        raised the captured error; omitted when not available).
    """
    try:
        db_path = validate_database_path(database_path)
        log_code_execution("vcs_run_vba", str(db_path), code, code_type="vba")
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("RunVBA", code)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_set_option")
def vcs_set_option(
    database_path: str,
    option_name: str,
    value: str | bool | int
) -> dict[str, Any]:
    """
    Set a VCS add-in option for the current MCP session.
    
    Changes take effect immediately and persist across operations within
    this session. The user's vcs-options.json is never modified -- overrides
    are stored in a session-scoped file under the mcp/ subfolder of the
    export directory. Stale override files are auto-cleaned after 30 days.
    
    Examples:
        vcs_set_option("C:\\\\db.accdb", "ShowDebug", True)
        vcs_set_option("C:\\\\db.accdb", "BreakOnError", True)
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        option_name: Name of the VCS option property
        value: Value to set
    
    Returns:
        Dictionary with success status and the option that was set
    """
    PROTECTED_OPTIONS = {"mcpallowrunvba"}
    if option_name.lower() in PROTECTED_OPTIONS:
        return {
            "success": False,
            "error": (
                f"The '{option_name}' option cannot be changed by agents. "
                "It controls arbitrary VBA code execution and requires explicit "
                "user consent. Enable it manually in the VCS Options form."
            ),
        }

    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            # Register session so the add-in scopes the override file correctly
            session_id = get_session_id()
            if session_id:
                addin.call_sync("RegisterSession", session_id)
            
            result_json = addin.call_sync("SetOption", option_name, value)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_get_option")
def vcs_get_option(
    database_path: str,
    option_name: str
) -> dict[str, Any]:
    """
    Read a VCS add-in option value.
    
    Returns the current in-memory value of any VCS add-in option property.
    If session overrides have been applied via vcs_set_option, those
    overridden values are reflected here.
    
    Examples:
        vcs_get_option("C:\\\\db.accdb", "ShowDebug")
        vcs_get_option("C:\\\\db.accdb", "McpAllowRunVBA")
        vcs_get_option("C:\\\\db.accdb", "ExportFormatVersion")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        option_name: Name of the VCS option property to read
    
    Returns:
        Dictionary with success status and the option value
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result = addin.call_sync("GetOption", option_name)
            
            # GetOption returns the raw value, not JSON
            if isinstance(result, str) and result.startswith("{"):
                try:
                    parsed = json.loads(result)
                    if isinstance(parsed, dict) and "success" in parsed:
                        return parsed
                except json.JSONDecodeError:
                    pass
            
            return {
                "success": True,
                "option": option_name,
                "value": result,
            }
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_get_log")
def vcs_get_log(
    database_path: str,
    log_type: str = "Export"
) -> dict[str, Any]:
    """
    Read the most recent operation log file.
    
    Finds and returns the content of the most recent log file matching the
    specified type (Export, Build, etc.) from the source folder's logs directory.
    
    Examples:
        vcs_get_log("C:\\\\db.accdb")
        vcs_get_log("C:\\\\db.accdb", log_type="Build")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        log_type: Type of log to read: "Export" (default) or "Build"
    
    Returns:
        Dictionary with log content, path, and success status
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("GetLogContent", log_type)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "content": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_end_session")
def vcs_end_session(
    database_path: str,
) -> dict[str, Any]:
    """
    End the current MCP session and remove all option overrides.
    
    Deletes the session-scoped override file and reloads the project options
    to their original state. Called automatically on MCP server shutdown,
    but can be called explicitly to clear overrides mid-conversation.
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
    
    Returns:
        Dictionary with success status
    """
    try:
        db_path = validate_database_path(database_path)
        session_id = get_session_id() or "default"
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("EndSession", session_id)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}
