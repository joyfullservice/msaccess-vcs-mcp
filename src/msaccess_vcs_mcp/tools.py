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
import glob
import inspect
import json
import os
import shutil
import sys
import tempfile
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal
from urllib.parse import unquote, urlparse

from mcp.server.fastmcp import FastMCP, Context
from mcp.types import CallToolResult, ToolAnnotations

from .access_com.connection import (
    AccessConnection,
    access_instance_is_live,
    close_owned_instances_holding,
    ensure_access_visible,
)
from .access_com.dao_helpers import list_query_defs, list_table_defs
from .access_com.process_qos import list_access_pids, prefer_full_power_app
from .access_com.vba_project import current_vb_project
from .access_gate import EXEMPT_TOOLS, get_access_gate
from .access_state import attach_access_state
from .guidance import with_guidance
from .shared_open import shared_open_warnings, with_warnings_first
from .tips import get_tips
from .config import (
    get_config,
    get_callback_url,
    get_session_id,
    initialize_from_workspace,
    load_config,
)
from .addin_integration import (
    API_REFUSED_PREFIX as _API_REFUSED_PREFIX,
    MERGE_NOT_AVAILABLE_MESSAGE,
    VCSAddinIntegration,
)
from .security import (
    validate_database_path,
    validate_export_directory,
    validate_source_directory,
    check_write_permission,
)
from .usage_logging import (
    log_code_execution,
    log_diagnostic_event,
    read_recent_tool_calls,
    with_logging,
)
from .rebuild_watcher import (
    TERMINAL_STATUSES as REBUILD_TERMINAL_STATUSES,
    get_rebuild_timeout,
    wait_for_rebuild_status,
)
from .operation_manager import MonotonicProgressReporter
from . import screenshot as _screenshot
from .vba_worker_manager import get_build_timeout, get_call_vba_timeout, run_vba_resilient

_COMPILE_FAILURE_AGENT_GUIDANCE = (
    "Compilation failed and the failing line could not be read. "
    "Call vcs_capture_screenshot(database_path, target=\"vbe\"): the editor "
    "highlights the error. If that does not show it, stop editing source "
    "files and ask the user to compile in the VBE (Debug → Compile) and paste "
    "the code around the error line. Do not guess fixes or iterate blindly. "
    "For more, call vcs_tips(\"compile\")."
)

_COMPILE_ERROR_LOCATED_GUIDANCE = (
    "Compilation stopped at the error in `compile_error` (module, line, and "
    "code). Fix that line in the source file, import it, and run "
    "vcs_compile_vba again. The compiler stops at the first error, so others "
    "may follow. For more, call vcs_tips(\"compile\")."
)

_ADDIN_NOT_RESPONSIVE_GUIDANCE = (
    "Call vcs_capture_screenshot(database_path, include_image=False) to see whether "
    "Access is showing a dialog or is in VBA break mode, then retry. For more, call "
    "vcs_tips(\"stuck\")."
)

_NOT_COMPILED_AGENT_GUIDANCE = (
    "The VBA project is not compiled. Run vcs_compile_vba before editing code: "
    "it compiles and, on an error, returns the failing module, line, and code "
    "in `compile_error`."
)


def _addin_json_result(result_json: Any, raw_key: str = "result") -> dict[str, Any]:
    """
    Parse a sync add-in API response into an MCP tool result.

    The add-in's sync API uses camelCase keys (``logPath``, ``errorNumber``)
    while MCP results use snake_case. Translating here keeps each side
    idiomatic and lets a newer server work against an older add-in build that
    only emits one of the two spellings.

    Args:
        result_json: Raw return value from ``VCSAddinIntegration.call_sync``
        raw_key: Key to file the value under when it is not a JSON object
    """
    result = result_json
    if isinstance(result, str):
        try:
            result = json.loads(result)
        except json.JSONDecodeError:
            return {"success": True, raw_key: result_json}

    if not isinstance(result, dict):
        return {"success": True, raw_key: result_json}

    # Keep the original camelCase key for one release so existing callers
    # that already read logPath keep working.
    log_path = result.get("log_path") or result.get("logPath")
    if log_path:
        result["log_path"] = log_path
        result["logPath"] = log_path

        if result.get("success") is False and "log_excerpt" not in result:
            excerpt = _read_log_excerpt(log_path)
            if excerpt:
                result["log_excerpt"] = excerpt

    return result


def _newest_log(source_dir: str | os.PathLike[str], base_name: str) -> str | None:
    """
    Return the newest ``{source_dir}/logs/{base_name}_*.log``, or None.

    Mirrors the add-in's own GetLogContent lookup. Log names embed a sortable
    ``yyyymmdd_hhnnss_fff`` stamp, so lexical max is newest.
    """
    pattern = os.path.join(str(source_dir), "logs", f"{base_name}_*.log")
    matches = glob.glob(pattern)
    return max(matches) if matches else None


def _read_log_excerpt(log_path: str, tail_lines: int = 50, max_chars: int = 4000) -> str | None:
    """
    Return the last ``tail_lines`` of a log file, truncated to ``max_chars``.

    Callers surface this on failure so an agent gets the error inline. The
    add-in gitignores its ``logs/`` folder, so Glob/Grep silently skip it and
    an agent that has only a path still burns calls trying to read around it.
    """
    if not log_path or not os.path.exists(log_path):
        return None

    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
    except OSError:
        return None

    excerpt = "".join(lines[-tail_lines:]).strip()
    if len(excerpt) > max_chars:
        excerpt = "...(truncated)...\n" + excerpt[-max_chars:]
    return excerpt or None


_TEST_NO_RESULTS_ERROR = (
    "Test runner returned no results. Ensure modTestAssert is "
    "installed in the target database and that test modules "
    "contain TestAssert calls."
)


def _apply_test_run_success(parsed: dict[str, Any]) -> dict[str, Any]:
    """Set ``success`` from the runner summary, not from Operation.Result."""
    summary = parsed.get("summary", {})
    parsed["success"] = (
        summary.get("failed", 1) == 0
        and summary.get("errored", 1) == 0
        and summary.get("subs", 0) > 0
    )
    return parsed


def _parse_test_runner_json(result_json: Any) -> dict[str, Any]:
    """Parse the sync ``RunFilteredTests`` return into a tool result."""
    if not result_json or (isinstance(result_json, str) and not str(result_json).strip()):
        return {"success": False, "error": _TEST_NO_RESULTS_ERROR}

    if isinstance(result_json, str):
        try:
            parsed = json.loads(result_json)
        except json.JSONDecodeError:
            return {
                "success": False,
                "error": f"Failed to parse test results JSON: {result_json[:200]}",
            }
    else:
        parsed = result_json if isinstance(result_json, dict) else {"result": result_json}

    if not isinstance(parsed, dict):
        return {"success": True, "result": parsed}
    if "summary" not in parsed and "tests" not in parsed:
        return parsed if "success" in parsed else {"success": True, "result": parsed}
    return _apply_test_run_success(parsed)


def _load_test_results_file(path: str | None) -> dict[str, Any] | None:
    """Read ``TestResults_*.json`` written by the add-in, or None if unreadable."""
    if not path:
        return None
    try:
        with open(path, encoding="utf-8-sig") as handle:
            parsed = json.load(handle)
    except (OSError, json.JSONDecodeError, TypeError):
        return None
    return parsed if isinstance(parsed, dict) else None


def _test_results_from_completion(completion: dict[str, Any]) -> dict[str, Any]:
    """Build the tool result from an MCP complete/error/cancelled callback.

    Failed tests finish as ``eorFailed`` (callback type ``error``) but still
    write ``results_path``. Prefer that file over the callback's error string.
    """
    loaded = _load_test_results_file(completion.get("results_path"))
    if loaded is not None:
        parsed = _apply_test_run_success(loaded)
        if completion.get("cancelled"):
            parsed["cancelled"] = True
            parsed["success"] = False
            parsed["error_pattern"] = "operation_cancelled"
        log_path = completion.get("log_path")
        if log_path:
            parsed.setdefault("log_path", log_path)
            parsed.setdefault("logPath", log_path)
        return parsed

    raw = completion.get("result")
    if raw not in (None, ""):
        parsed = _parse_test_runner_json(raw)
        return _mark_cancelled(parsed, completion) if completion.get("cancelled") else parsed

    if completion.get("cancelled"):
        return {
            "success": False,
            "cancelled": True,
            "error_pattern": "operation_cancelled",
            "error": completion.get("message") or "Test run was cancelled",
        }

    if completion.get("success") and not completion.get("error"):
        return {"success": False, "error": _TEST_NO_RESULTS_ERROR}

    result: dict[str, Any] = {
        "success": False,
        "error": completion.get("error")
        or completion.get("message")
        or "Test run failed",
    }
    for key in ("error_pattern", "operation_id", "log_path"):
        if completion.get(key):
            result[key] = completion[key]
    return result


def _attach_log_context(
    result: dict[str, Any],
    source_dir: str | os.PathLike[str],
    base_name: str,
    completion: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Add ``log_path`` to a tool result, plus ``log_excerpt`` when it failed.

    The add-in gitignores its ``logs/`` folder, so Glob/Grep skip it entirely
    and an agent told only that "the build failed" cannot search its way to the
    reason. Returning the tail inline on failure removes that dead end.
    """
    # Prefer an explicit path from the completion callback or an already-
    # normalized tool result (e.g. sync ImportByType / ExportByType JSON).
    log_path = (completion or {}).get("log_path") or result.get("log_path")
    if not log_path or not os.path.exists(log_path):
        # The operation just ran, so the newest log for this family is its own.
        log_path = _newest_log(source_dir, base_name)

    result["log_path"] = log_path
    if not result.get("success") and log_path:
        excerpt = _read_log_excerpt(log_path)
        if excerpt:
            result["log_excerpt"] = excerpt
    return result


def _scoped_types_arg(object_types: list[str]) -> str | list[str]:
    """
    Shape ``object_types`` for ``ExportByType`` / ``ImportByType``.

    The add-in accepts a single alias string or an array. A one-element list
    is sent as a bare string so COM marshalling does not depend on
    ``VT_ARRAY|VT_VARIANT`` for the common single-category case.
    """
    cleaned = [t.strip() for t in object_types if t and t.strip()]
    if len(cleaned) == 1:
        return cleaned[0]
    return cleaned


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
            "agent_guidance": (
                "Wait for the current operation to complete, or cancel it with "
                "vcs_cancel_operation()."
            ),
        }
    return None


# Create FastMCP server instance with proper metadata
# Claude Code cuts server instructions at 4,096 characters and some clients never
# show them at all, so the rules every call needs come first, and failure-time
# guidance travels in results (guidance.py) and vcs_tips.
SERVER_INSTRUCTIONS = (
    "Version control for Microsoft Access databases (VBA modules, forms, reports, "
    "queries, macros, tables), driven through the MSAccess VCS add-in. Export a "
    "database to text source files, edit them, import or merge changes back, rebuild "
    "a database from source, compile and run VBA, run tests, and capture screenshots "
    "of Access windows. All tools use the vcs_ prefix.\n\n"
    "Rules for every call:\n"
    "- The installed add-in (under %APPDATA%\\MSAccessVCS) is never a target. Tools "
    "refuse it as database_path, output_path, or template_path. Work on the "
    "development copy in the add-in's repository and rebuild from there.\n"
    "- Access runs one call at a time. Issue Access tool calls one after another, "
    "not in parallel; a call that waits too long (15 seconds by default) for another "
    "returns error_pattern \"server_busy\" and can be retried.\n"
    "- If Access seems stuck or a call timed out, call "
    "vcs_capture_screenshot(database_path, include_image=False). It lists every "
    "dialog Access is showing and whether VBA is in break mode, and it works while "
    "Access is blocked.\n"
    "- A failed result may carry agent_guidance with the next step; follow it.\n\n"
    "Call vcs_tips(topic) for depth:\n"
    "- If you are working on the msaccess-vcs-addin repository itself, call "
    "vcs_tips(\"addin-development\") first.\n"
    "- Access stuck, a dialog, break mode, or a timeout: vcs_tips(\"stuck\").\n"
    "- VBA compile errors: vcs_tips(\"compile\").\n"
    "- Writing code for vcs_run_vba: vcs_tips(\"run-vba\").\n"
    "- Finding operation logs: vcs_tips(\"logs\").\n\n"
    "Workflow: vcs_export_database writes source files; edit them; "
    "vcs_import_objects merges changes back; vcs_diff_database compares the database "
    "with source; vcs_rebuild_database builds a fresh database from source. Most "
    "tools take database_path, the full path to the database file.\n\n"
    "Configuration comes from environment variables, usually in a .env file. "
    "ACCESS_VCS_DATABASE names the default database. ACCESS_VCS_DISABLE_WRITES=true "
    "refuses import and rebuild.\n\n"
    "Leave Access windows open. Never stop or kill MSACCESS.EXE: another agent or "
    "the user may own the window. vcs_dismiss_dialog and vcs_reset_vba act only on "
    "windows this server created; for the user's windows, ask the user.\n\n"
    "vcs_rebuild_database builds a user project from source. vcs_rebuild_addin "
    "rebuilds the VCS add-in itself from its Version Control.accda.src folder.\n\n"
    "If the client times out (-32001), no result reaches you. Call "
    "vcs_get_recent_calls() to see what actually ran before retrying."
)

mcp = FastMCP(
    name="msaccess-vcs-mcp",
    instructions=SERVER_INSTRUCTIONS,
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


READ_ONLY_TOOLS = frozenset({
    "vcs_list_objects",
    "vcs_diff_database",
    "vcs_get_option",
    "vcs_get_log",
    "vcs_get_version_info",
    "vcs_get_recent_calls",
    "vcs_check_vba_compiled",
    "vcs_execute_sql",
    "vcs_capture_screenshot",
    "vcs_tips",
})

DESTRUCTIVE_TOOLS = frozenset({
    "vcs_import_objects",
    "vcs_import_object",
    "vcs_rebuild_database",
    "vcs_rebuild_addin",
    "vcs_set_option",
})


def _default_annotations(name: str) -> ToolAnnotations | None:
    if name in READ_ONLY_TOOLS:
        return ToolAnnotations(readOnlyHint=True)
    if name in DESTRUCTIVE_TOOLS:
        return ToolAnnotations(readOnlyHint=False, destructiveHint=True)
    return None


def vcs_tool(
    name: str,
    *,
    annotations: ToolAnnotations | None = None,
    meta: dict[str, Any] | None = None,
):
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
    5. On a timeout or unresponsive Access, attaches ``access_state``: the
       dialogs and break mode Access shows, read with window calls only
       (:mod:`access_state`).
    6. Adds ``agent_guidance`` to a failed result that has none
       (:mod:`guidance`).

    The wrapper is *always* an async coroutine. FastMCP detects this via
    ``inspect.iscoroutinefunction`` and awaits it correctly. Sync tool
    bodies run in a single COM apartment thread via :mod:`access_gate` so
    they no longer block the asyncio event loop. One Access operation runs
    at a time across all Cursor windows sharing this server process.
    """
    def decorator(func):
        logged = with_logging(name)(func)
        is_async_body = inspect.iscoroutinefunction(func)

        @functools.wraps(func)
        async def with_refresh(*args, **kwargs):
            result = await _run_tool(*args, **kwargs)
            result = await attach_access_state(name, func, args, kwargs, result)
            return with_guidance(name, result)

        async def _run_tool(*args, **kwargs):
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

            refusal = _refuse_installed_addin_target(name, func, args, kwargs)
            if refusal is not None:
                return refusal

            if name in EXEMPT_TOOLS:
                if is_async_body:
                    return await logged(*args, **kwargs)
                return logged(*args, **kwargs)

            database = kwargs.get("database_path")
            if database is None and args:
                sig = inspect.signature(func)
                param_names = list(sig.parameters.keys())
                if param_names and param_names[0] == "database_path":
                    database = args[0]

            gate = get_access_gate()
            return await gate.run_exclusive(
                name,
                str(database) if database is not None else None,
                logged,
                is_async_body,
                *args,
                **kwargs,
            )

        return mcp.tool(
            annotations=annotations or _default_annotations(name), meta=meta
        )(with_refresh)
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
    
    With no ``object_types``, runs a full-project export (``Export`` /
    ``FullExport``) with progress callbacks. With ``object_types``, runs a
    category-scoped export via ``ExportByType`` — only those categories are
    written, deletions within them are reconciled, and the call is
    synchronous (no progress reporting). Prefer ``vcs_export_object`` for a
    single named object.
    
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
        object_types: Optional categories to export (e.g. ``["queries"]``,
            ``["modules", "forms"]``). If None, exports the entire project.
        full_export: If True, export all objects in scope; if False (default),
            only export changed objects (per the VCS index)
    
    Returns:
        Dictionary with:
        - exported_count: Number of objects exported
        - export_path: Path where files were written
        - objects_by_type: Breakdown of what was exported
        - errors: List of any errors encountered
        - log_path: Full path to this run's log file
        - log_excerpt: Tail of the log, included only on failure
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Export").
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
        
        # Determine full-project export command (scoped path uses ExportByType)
        command = "FullExport" if full_export else "Export"
        
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
                    "agent_guidance": _ADDIN_NOT_RESPONSIVE_GUIDANCE,
                }

            # Category-scoped export: sync ExportByType (no progress callbacks).
            if object_types:
                types_arg = _scoped_types_arg(object_types)
                if not types_arg:
                    return {
                        "success": False,
                        "error": "object_types was empty after stripping blanks",
                        "exported_count": 0,
                        "export_path": str(export_path),
                        "objects_by_type": {},
                    }
                result = await _run_scoped_api(
                    app, addin, "ExportByType", [types_arg, full_export],
                    str(db_path), ctx, "Exporting",
                )
                result.setdefault("export_path", str(export_path))
                return _attach_log_context(result, export_path, "Export")
            
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
                    # Call async API (Export/FullExport use VCS options)
                    async_result = addin.call_async(callback_info, command)
                    
                    completion = None
                    if async_result.get("sync"):
                        # VBA returned sync result
                        op_manager.unregister_operation(operation_id)
                    elif async_result.get("async"):
                        # Wait for completion with progress reporting
                        timeout_ms = async_result.get("timeout_ms", 300000)
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=ctx,  # Pass context for progress reporting
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if not completion.get("success"):
                            return _attach_log_context(_mark_cancelled({
                                "success": False,
                                "error": completion.get("error", "Export failed"),
                                "error_pattern": completion.get("error_pattern"),
                                "operation_id": completion.get("operation_id") or operation_id,
                                "exported_count": 0,
                                "export_path": str(export_path),
                                "objects_by_type": {},
                            }, completion), export_path, "Export", completion)
                    else:
                        # Neither marker: the add-in never started the operation,
                        # so run it synchronously rather than reporting success
                        # for work that never happened.
                        op_manager.unregister_operation(operation_id)
                        result = addin.export_source(str(db_path), str(export_path))
                        
                        if not result["success"]:
                            return _attach_log_context({
                                "success": False,
                                "error": result["message"],
                                "exported_count": 0,
                                "export_path": str(export_path),
                                "objects_by_type": {},
                            }, export_path, "Export")
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.export_source(str(db_path), str(export_path))
                    
                    if not result["success"]:
                        return _attach_log_context({
                            "success": False,
                            "error": result["message"],
                            "exported_count": 0,
                            "export_path": str(export_path),
                            "objects_by_type": {},
                        }, export_path, "Export")
            else:
                # Use sync path (no callbacks available)
                completion = None
                result = addin.export_source(str(db_path), str(export_path))
                
                if not result["success"]:
                    return _attach_log_context({
                        "success": False,
                        "error": result["message"],
                        "exported_count": 0,
                        "export_path": str(export_path),
                        "objects_by_type": {},
                    }, export_path, "Export")
            
            return _attach_log_context({
                "success": True,
                "export_path": str(export_path),
                "messages": (completion or {}).get("log_messages"),
            }, export_path, "Export", completion)
    
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
                vb_project = current_vb_project(app)
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


def _merge_fallback_failure(result: dict[str, Any], src_path: Path) -> dict[str, Any]:
    """Tool result for a failed synchronous ``merge_build`` fallback."""
    failure: dict[str, Any] = {
        "success": False,
        "error": result["message"],
        "imported_count": 0,
    }
    if result["message"] == MERGE_NOT_AVAILABLE_MESSAGE:
        # No log was started, so the newest Merge log on disk belongs to an
        # earlier run.
        failure["error_pattern"] = "merge_not_available"
        return failure
    return _attach_log_context(failure, src_path, "Merge")


@vcs_tool("vcs_import_objects")
async def vcs_import_objects(
    database_path: str,
    source_dir: str,
    object_types: list[str] | None = None,
    full_import: bool = False,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Import objects from source files into Access database.
    
    Merges source files back into the database. With no ``object_types``, runs
    a full project merge (``MergeBuild``) with progress callbacks. With
    ``object_types``, runs a category-scoped merge via ``ImportByType`` —
    only those categories are touched.
    
    A full project merge needs a prior full build or full export of the
    database. Without one it returns ``error_pattern: merge_not_available``
    and imports nothing; run ``vcs_rebuild_database`` (or a full export)
    first. (Add-in builds before 5.1.1 report nothing in that case, and the
    call waits out its timeout.)
    
    **Scoped import (when object_types is set):**
    
    - Database objects in the named categories with no corresponding source
      file are **deleted** (orphan reconciliation).
    - No database backup is taken (unlike a full ``MergeBuild``). Own your
      backup before a destructive scoped merge.
    - ``full_import=False`` (default) merges only files the VCS index marks
      as changed. A stale or reset index can under-report; pass
      ``full_import=True`` to reload every source file in those categories
      regardless of the index. That path also skips conflict detection
      (source wins for the named categories).
    - Passing every category with ``full_import=True`` approximates a full
      build without a backup — prefer ``vcs_rebuild_database`` there.
    - Scoped calls are synchronous (no progress reporting). Prefer a single
      object via ``vcs_import_object`` when you only need one name.
    
    ``source_dir`` must exist and is used for log resolution; the merge itself
    reads from the project's configured export folder.
    
    Examples:
        # Full project merge
        vcs_import_objects("C:\\\\db.accdb", "C:\\\\src\\\\mydb")
        
        # Merge only queries (changed source files)
        vcs_import_objects(
            "C:\\\\db.accdb",
            "C:\\\\src\\\\mydb",
            object_types=["queries"],
        )
        
        # Reload every module from source, ignoring the change index
        vcs_import_objects(
            "C:\\\\db.accdb",
            "C:\\\\src\\\\mydb",
            object_types=["modules"],
            full_import=True,
        )
    
    Args:
        database_path: Path to Access database
        source_dir: Directory containing source files (must exist; merge uses
            the project's export folder)
        object_types: Optional categories to merge (e.g. ``["queries"]``,
            ``["modules", "forms"]``). If None, merges the entire project.
        full_import: When ``object_types`` is set: if False (default), merge
            only changed source files; if True, merge all source files in
            those categories (ignores the change index, skips conflict
            prompts). Ignored for a full project merge.
    
    Returns:
        Dictionary with import results and any errors, plus ``log_path`` for
        this run's log and ``log_excerpt`` (tail of the log) on failure.
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Merge").
    A ``warnings`` entry says when another session has the database open.
    """
    advisory: list[str] = []
    result = await _import_objects(
        database_path, source_dir, object_types, full_import, ctx, advisory
    )
    return with_warnings_first(result, advisory)


async def _import_objects(
    database_path: str,
    source_dir: str,
    object_types: list[str] | None,
    full_import: bool,
    ctx: Context | None,
    advisory: list[str],
) -> dict[str, Any]:
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
            advisory.extend(shared_open_warnings(str(db_path), app))
            
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
                    "agent_guidance": _ADDIN_NOT_RESPONSIVE_GUIDANCE,
                }

            # Category-scoped merge: ImportByType runs synchronously in Access.
            if object_types:
                types_arg = _scoped_types_arg(object_types)
                if not types_arg:
                    return {
                        "success": False,
                        "error": "object_types was empty after stripping blanks",
                        "imported_count": 0,
                    }
                result = await _run_scoped_api(
                    app, addin, "ImportByType", [types_arg, full_import],
                    str(db_path), ctx, "Merging",
                )
                result.setdefault("database_path", str(db_path))
                result.setdefault("source_dir", str(src_path))
                if result.get("success") and "imported_count" not in result:
                    result["imported_count"] = "See log for details"
                return _attach_log_context(result, src_path, "Merge")
            
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
                            ctx=ctx,
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if completion.get("code") == "merge_not_available":
                            # No log was started, so the newest Merge log on
                            # disk belongs to an earlier run.
                            return {
                                "success": False,
                                "error": MERGE_NOT_AVAILABLE_MESSAGE,
                                "error_pattern": "merge_not_available",
                                "imported_count": 0,
                            }
                        if not completion.get("success"):
                            return _attach_log_context(_mark_cancelled({
                                "success": False,
                                "error": completion.get("error", "Import failed"),
                                "error_pattern": completion.get("error_pattern"),
                                "operation_id": completion.get("operation_id") or operation_id,
                                "imported_count": 0,
                            }, completion), src_path, "Merge", completion)
                    elif async_result.get("sync"):
                        # The add-in already ran the merge inline; re-running it
                        # here would merge twice. Fall through and resolve the
                        # log from disk.
                        op_manager.unregister_operation(operation_id)
                    else:
                        # Neither marker: the add-in never started the operation,
                        # so run it synchronously rather than reporting success
                        # for work that never happened.
                        op_manager.unregister_operation(operation_id)
                        result = addin.merge_build(str(db_path), str(src_path))
                        if not result["success"]:
                            return _merge_fallback_failure(result, src_path)
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.merge_build(str(db_path), str(src_path))
                    if not result["success"]:
                        return _merge_fallback_failure(result, src_path)
            else:
                # Use sync path
                completion = None
                result = addin.merge_build(str(db_path), str(src_path))
                if not result["success"]:
                    return _merge_fallback_failure(result, src_path)
            
            return _attach_log_context({
                "success": True,
                "imported_count": "See log for details",
                "database_path": str(db_path),
                "source_dir": str(src_path),
            }, src_path, "Merge", completion)
    
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


# BuildHeadless arrived in 5.1.0 and took an output path from 5.1.1. An extra
# argument raises inside the library, which Access shows as a modal dialog rather
# than returning, so the version has to be checked before the call is made.
_BUILD_HEADLESS_MIN_VERSION = (5, 1, 0)
_BUILD_HEADLESS_OUTPUT_MIN_VERSION = (5, 1, 1)
_BUILD_HEARTBEAT_SEC = 15.0
_BUILD_HOST_RELEASE_SEC = 5.0


def _parse_addin_version(value: Any) -> tuple[int, ...] | None:
    """Return ``(5, 1, 1)`` for ``"5.1.1"``, or None when it is not a version."""
    try:
        return tuple(int(part) for part in str(value).strip().split("."))
    except (TypeError, ValueError):
        return None


# RunTestsMatching takes the filter as an argument. Older add-ins only read it from
# the DefaultTestFilter option, and setting that through SetOption writes
# mcp/options-default.json into the user's export folder, where nothing deletes it.
# Calling a method the add-in lacks raises a modal dialog, so this is a version check.
_RUN_TESTS_MATCHING_MIN_VERSION = (5, 2, 0)


def _test_run_command(version_text: Any, filter: str | None) -> tuple[str, tuple[str, ...]]:
    """Return the add-in method and arguments that run the tests matching ``filter``."""
    version = _parse_addin_version(version_text)
    if version is not None and version >= _RUN_TESTS_MATCHING_MIN_VERSION:
        return "RunTestsMatching", (filter or "",)
    return "RunFilteredTests", ()


def _build_headless_refusal(version_text: Any, output_path: str | None) -> dict[str, Any] | None:
    """Refuse a build the installed add-in cannot run unattended, or return None."""
    needed = _BUILD_HEADLESS_OUTPUT_MIN_VERSION if output_path else _BUILD_HEADLESS_MIN_VERSION
    version = _parse_addin_version(version_text)
    if version is not None and version >= needed:
        return None

    needed_text = ".".join(str(part) for part in needed)
    if output_path:
        reason = (
            f"Building to output_path needs VCS add-in {needed_text} or later "
            f"(installed: {version_text}). Omit output_path to build to the file "
            "the source files name, or update the add-in."
        )
    else:
        reason = (
            f"vcs_rebuild_database needs VCS add-in {needed_text} or later "
            f"(installed: {version_text})."
        )
    return {
        "success": False,
        "error": reason,
        "error_pattern": "addin_version",
        "output_path": None,
    }


def _create_build_instance() -> Any:
    """Start a new Access process for one build.

    ``ensure_dispatch`` can return a window the user already has open; the
    build would open its scratch database there and quit the window when it
    finished.
    """
    import win32com.client

    app = win32com.client.DispatchEx("Access.Application")
    prefer_full_power_app(app)
    return app


def _open_build_host(app: Any) -> str:
    """Open a scratch database in ``app`` and return its path.

    ``Application.Run`` into the add-in library with no database open crashes
    Access, so a full build needs something open to call it from. The build
    closes this database before it creates the new one.
    """
    folder = tempfile.mkdtemp(prefix="vcs-build-host-")
    path = os.path.join(folder, "BuildHost.accdb")
    try:
        app.NewCurrentDatabase(path)
    except Exception:
        shutil.rmtree(folder, ignore_errors=True)
        raise
    return path


async def _remove_build_host(path: str) -> None:
    """Delete the scratch folder once Access lets go of it.

    ``Quit`` returns before Access releases the file, so the first attempt
    can fail on a file still in use.
    """
    folder = os.path.dirname(path)
    deadline = time.perf_counter() + _BUILD_HOST_RELEASE_SEC
    while True:
        shutil.rmtree(folder, ignore_errors=True)
        if not os.path.exists(folder) or time.perf_counter() >= deadline:
            return
        await asyncio.sleep(0.25)


async def _run_api_with_callbacks(
    acquire_app: Any,
    addin_path: str,
    command: str,
    args: list[str],
    *,
    database_path: str,
    timeout: float,
    ctx: Any,
    heartbeat: str,
) -> dict[str, Any]:
    """Run a long synchronous API method on a worker thread, cancellable from MCP.

    The method goes through ``APIAsync`` with callback info, so the add-in streams
    its log, reports progress, and polls for cancellation, while the event loop
    stays free to answer ``vcs_cancel_operation``. Without the callback server it
    runs through ``API``. An add-in older than the callback support runs the
    method through ``API`` either way and only streams nothing.

    Returns ``run_result`` (the ``_start_application_call`` box, with the
    ``{"sync":true,"result":...}`` wrapper removed), ``completion`` (the terminal
    callback, if one arrived), ``log_messages``, ``operation_id``, and
    ``timed_out``. On a timeout or client abandonment the operation is asked to
    stop and left tombstoned.
    """
    lib_name = os.path.splitext(os.path.abspath(addin_path))[0]
    op_manager = _get_operation_manager()
    callback_url = get_callback_url()
    operation_id: str | None = None
    queue: asyncio.Queue | None = None

    if op_manager and callback_url:
        op_manager.set_event_loop(asyncio.get_running_loop())
        operation_id, queue = op_manager.register_operation(
            timeout_ms=int(timeout * 1000),
            database_path=database_path,
            command=command,
        )
        callback_info = op_manager.create_callback_info(operation_id, callback_url, "cursor")
        run_name, call_args = f"{lib_name}.APIAsync", [callback_info, command, *args]
    else:
        run_name, call_args = f"{lib_name}.API", [command, *args]

    reporter = MonotonicProgressReporter()
    log_messages: list[str] = []
    completion: dict[str, Any] | None = None

    async def _drain() -> None:
        nonlocal completion
        while queue is not None and not queue.empty():
            callback = queue.get_nowait()
            msg_type = str(callback.get("type") or "")
            message = str(callback.get("message") or "")
            if msg_type == "progress":
                await reporter.emit(
                    ctx,
                    message=message,
                    vba_progress=callback.get("progress"),
                    vba_total=callback.get("total"),
                )
            elif msg_type == "log" and message:
                log_messages.append(message)
                await reporter.emit(ctx, message=message)
            elif msg_type in ("complete", "error", "cancelled"):
                completion = dict(callback)
                if msg_type == "cancelled":
                    completion["cancelled"] = True

    thread, run_result = _start_application_call(acquire_app, run_name, call_args)
    start = time.perf_counter()
    next_beat = start + _BUILD_HEARTBEAT_SEC
    timed_out = False
    try:
        while thread.is_alive():
            now = time.perf_counter()
            if now - start >= timeout:
                timed_out = True
                break
            had_messages = len(log_messages)
            await _drain()
            if len(log_messages) != had_messages:
                next_beat = now + _BUILD_HEARTBEAT_SEC
            elif now >= next_beat:
                await reporter.emit(ctx, message=f"{heartbeat} ({now - start:.0f}s)")
                next_beat = now + _BUILD_HEARTBEAT_SEC
            await asyncio.sleep(0.25)
        # Callbacks posted just before the call returned may still be in flight.
        await asyncio.sleep(0)
        await _drain()
    except asyncio.CancelledError:
        if op_manager and operation_id:
            op_manager.abandon_operation(operation_id)
        raise

    if op_manager and operation_id:
        if timed_out:
            op_manager.abandon_operation(operation_id)
        else:
            op_manager.unregister_operation(operation_id)

    raw = run_result.get("result")
    if isinstance(raw, tuple) and raw:
        raw = raw[0]
    if isinstance(raw, str) and run_name.endswith(".APIAsync"):
        try:
            parsed = json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            parsed = None
        if isinstance(parsed, dict) and parsed.get("sync") and "result" in parsed:
            raw = parsed["result"]
        run_result["result"] = raw

    return {
        "run_result": run_result,
        "completion": completion,
        "log_messages": log_messages,
        "operation_id": operation_id,
        "timed_out": timed_out,
    }


async def _run_scoped_api(
    app: Any,
    addin: VCSAddinIntegration,
    command: str,
    args: list[Any],
    db_path: str,
    ctx: Any,
    heartbeat: str,
) -> dict[str, Any]:
    """Run a category-scoped ``ExportByType`` / ``ImportByType`` and parse its JSON."""
    timeout = get_build_timeout()
    run = await _run_api_with_callbacks(
        VCSAddinIntegration.worker_app_source(app, None),
        addin.addin_path,
        command,
        args,
        database_path=db_path,
        timeout=timeout,
        ctx=ctx,
        heartbeat=heartbeat,
    )
    if run["timed_out"]:
        return {
            "success": False,
            "error": (
                f"{command} did not finish within {timeout:g} seconds "
                "(ACCESS_VCS_BUILD_TIMEOUT_SEC). The add-in was asked to stop it."
            ),
            "error_pattern": "timeout",
            "timed_out": True,
            "operation_id": run["operation_id"],
        }
    run_result = run["run_result"]
    if not run_result.get("success"):
        return {
            "success": False,
            "error": _describe_run_failure(run_result.get("error"), command, command),
        }
    raw = run_result.get("result")
    if isinstance(raw, str) and raw.startswith(_API_REFUSED_PREFIX):
        return {"success": False, "error": raw[len(_API_REFUSED_PREFIX):]}
    result = _addin_json_result(raw)
    if run["log_messages"]:
        result.setdefault("log_messages", run["log_messages"])
    return _mark_cancelled(result, run["completion"])


def _mark_cancelled(result: dict[str, Any], completion: dict[str, Any] | None) -> dict[str, Any]:
    """Flag a tool result as canceled when the add-in or its callback says so."""
    if result.get("cancelled") or (completion or {}).get("cancelled"):
        result["success"] = False
        result["cancelled"] = True
        result["error_pattern"] = "operation_cancelled"
        result.setdefault("error", "Operation was canceled.")
    return result


def _build_headless_result(
    run_result: dict[str, Any],
    api_name: str,
    src_path: Path,
    output_path: str | None,
) -> dict[str, Any]:
    """Turn a finished ``BuildHeadless`` call into a tool result."""
    if not run_result.get("success"):
        return {
            "success": False,
            "error": _describe_run_failure(run_result.get("error"), api_name, api_name),
            "output_path": None,
        }

    raw = run_result.get("result")
    if isinstance(raw, tuple) and raw:
        raw = raw[0]
    if isinstance(raw, str) and raw.startswith(_API_REFUSED_PREFIX):
        return {
            "success": False,
            "error": raw[len(_API_REFUSED_PREFIX):],
            "output_path": None,
        }

    result = _addin_json_result(raw)
    if "databasePath" in result:
        result["output_path"] = result["databasePath"]
    else:
        result.setdefault("output_path", output_path if result.get("success") else None)
    result["source_dir"] = str(src_path)
    return result


@vcs_tool("vcs_rebuild_database")
async def vcs_rebuild_database(
    source_dir: str,
    output_path: str | None = None,
    template_path: str | None = None,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Build a complete Access database from source files.
    
    Creates a fresh database and imports all objects from source.
    Useful for clean builds and distribution.
    
    Runs the add-in's ``BuildHeadless``, which never prompts and returns when
    the build is over. Progress notifications carry the build log, with a
    heartbeat when it is quiet. vcs_cancel_operation stops the build, which then
    returns ``cancelled: true``. The call is bounded by
    ``ACCESS_VCS_BUILD_TIMEOUT_SEC`` (default 600 seconds).
    
    Examples:
        # Build to the file the source files name
        vcs_rebuild_database("C:\\\\src\\\\mydb")
        
        # Build to a different file (add-in 5.1.1 or later)
        vcs_rebuild_database("C:\\\\src\\\\mydb", "C:\\\\output\\\\rebuilt.accdb")
    
    Args:
        source_dir: Directory containing source files
        output_path: Optional path for the new database file. Omit it to build
            to the file the source files name, which works with any add-in that
            has ``BuildHeadless`` (5.1.0). A different path needs add-in 5.1.1
            or later; against an older add-in the call is refused before the
            build starts.
        template_path: Accepted for compatibility and ignored; the add-in
            builds into a new blank database.
    
    Returns:
        Dictionary with ``success``, ``output_path`` (the file that was built),
        ``errorCount``, ``warningCount``, ``durationMs``, ``log_path`` for this
        run's log, and ``log_excerpt`` (tail of the log) on failure. A build
        refused before it started carries only ``success`` and ``error``.
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Build").
    """
    config = get_config()
    
    try:
        # Check if writes are disabled
        check_write_permission(config)
        
        # Validate source directory
        src_path = validate_source_directory(source_dir)
        
        # Check if target database is already busy (if it exists)
        if output_path:
            close_owned_instances_holding([output_path])
            busy_error = _check_database_busy(output_path)
            if busy_error:
                return busy_error
        
        # No database exists yet -- the build creates it -- so there is no
        # AccessConnection; the instance's lifecycle is managed here.
        app = _create_build_instance()
        host_path: str | None = None
        build_running = False

        try:
            host_path = _open_build_host(app)
            # The build creates and populates a database in this instance, so
            # whatever it shows has to be visible to be dismissed.
            ensure_access_visible(app)

            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=None)

            refusal = _build_headless_refusal(
                addin.call_sync("GetVCSVersion"), output_path
            )
            if refusal:
                return refusal

            build_args = [str(src_path)]
            if output_path:
                build_args.append(str(output_path))
            api_name = f"{os.path.splitext(os.path.abspath(addin.addin_path))[0]}.API"
            timeout = get_build_timeout()

            build_running = True
            run = await _run_api_with_callbacks(
                VCSAddinIntegration.worker_app_source(app, None),
                addin.addin_path,
                "BuildHeadless",
                build_args,
                database_path=str(output_path or src_path),
                timeout=timeout,
                ctx=ctx,
                heartbeat="Building from source",
            )

            if run["timed_out"]:
                return _attach_log_context({
                    "success": False,
                    "error": (
                        f"Build did not finish within {timeout:g} seconds "
                        "(ACCESS_VCS_BUILD_TIMEOUT_SEC). The add-in was asked to "
                        "stop it; it may still be running in the Access window, "
                        "which is left open."
                    ),
                    "error_pattern": "timeout",
                    "timed_out": True,
                    "operation_id": run["operation_id"],
                    "output_path": output_path,
                    "source_dir": str(src_path),
                }, src_path, "Build")
            build_running = False

            result = _build_headless_result(
                run["run_result"], api_name, src_path, output_path
            )
            if run["log_messages"]:
                result.setdefault("log_messages", run["log_messages"])
            return _mark_cancelled(result, run["completion"])
        finally:
            # A build still running owns the instance; quitting it from here
            # would block until the build ends, and the host may still be open.
            if not build_running:
                try:
                    app.Quit()
                except Exception:
                    pass
                if host_path:
                    await _remove_build_host(host_path)

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


@vcs_tool("vcs_rebuild_addin")
async def vcs_rebuild_addin(
    source_dir: str,
    timeout_seconds: float | None = None,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Rebuild the VCS add-in from source and wait until it is installed.

    Derives the development copy (the ``Version Control.accda`` beside
    ``source_dir``), launches ``RebuildAddIn``, and watches
    ``<source>/logs/rebuild-status.json`` until this attempt reaches
    ``complete`` or a terminal failure. The Access gate is held only for
    the launch; other tools can run while the worker builds.

    Before launch, the server closes Access windows it created that hold
    the development copy or the installed add-in. User-owned windows are
    left alone and still produce ``refused`` / ``otherInstances``.

    This is not ``vcs_rebuild_database``, which rebuilds a user project.
    ``vcs_call_vba(..., ["RebuildAddIn", source])`` remains a launch-only
    escape hatch and does not wait for the rebuild to finish.

    Args:
        source_dir: Path to ``Version Control.accda.src``
        timeout_seconds: How long to wait after launch. Defaults to
            ACCESS_VCS_REBUILD_TIMEOUT_SEC (20 minutes).

    Returns:
        Terminal rebuild status, including ``status``, ``status_file``,
        ``phaseStarted``, and ``buildLog`` when present.
    """
    operation_id: str | None = None
    op_manager = None
    callback_task: asyncio.Task | None = None
    # Set while a launched rebuild may still be running unobserved: on a watch
    # that ends without a terminal status, or when the client stops waiting.
    abandoned = False

    try:
        check_write_permission(get_config())
        src_path = validate_source_directory(source_dir)
        host_path = _development_addin_from_source(src_path)
        if _is_installed_addin_path(str(host_path)):
            return _installed_addin_target_refusal(
                "vcs_rebuild_addin", "source_dir", str(host_path)
            )

        status_file, status_before = _snapshot_rebuild_status(str(src_path))
        timeout = get_rebuild_timeout(timeout_seconds)
        callback_url = get_callback_url()
        callback_info: str | None = None
        callback_queue: asyncio.Queue | None = None
        callback_state: dict[str, Any] = {"log_messages": []}
        reporter = MonotonicProgressReporter()

        if callback_url:
            op_manager = _get_operation_manager()
            if op_manager:
                op_manager.set_event_loop(asyncio.get_running_loop())
                operation_id, callback_queue = op_manager.register_operation(
                    timeout_ms=int(timeout * 1000),
                    database_path=str(host_path),
                    command="RebuildAddIn",
                )
                callback_info = op_manager.create_callback_info(
                    operation_id,
                    callback_url,
                    "cursor",
                )

        async def _launch() -> dict[str, Any]:
            # Inside the gate: another window's tool call could otherwise
            # open a fresh instance between the close and the launch, and
            # the rebuild would refuse on a file we had just freed.
            # Every tool call loads the add-in as a library, which locks it
            # regardless of which database that instance has open, so the
            # installed path is matched against loaded libraries too.
            installed = get_config().get("ACCESS_VCS_ADDIN_PATH")
            close_owned_instances_holding(
                [str(host_path)],
                [str(installed)] if installed else [],
            )

            call_args = ["RebuildAddIn", str(src_path)]
            if callback_info:
                call_args.append(callback_info)
            return _execute_call_vba(
                str(host_path),
                "VCS.API",
                call_args,
            )

        # COM launch blocks this task until RebuildAddIn returns, so emit
        # once beforehand rather than fake steps we cannot observe.
        await reporter.emit(ctx, message="Starting Access...")

        preexisting_access_pids = list_access_pids()

        gate = get_access_gate()
        launch = await gate.run_exclusive(
            "vcs_rebuild_addin",
            str(host_path),
            _launch,
            True,
        )
        if isinstance(launch, dict) and launch.get("error_pattern") == "server_busy":
            return launch

        launch = _with_rebuild_fields(launch, status_file, status_before)
        if not launch.get("success"):
            return launch

        parsed = _parse_rebuild_launch(launch.get("result"))
        status = parsed.get("status")
        if status in ("refused", "launch-failed") or not status:
            result = dict(launch)
            result.update(parsed)
            result["status_file"] = status_file
            if status != "launched":
                result["success"] = status == "complete"
                if status and status != "complete":
                    result.setdefault(
                        "error",
                        parsed.get("error") or f"Rebuild ended with status {status}",
                    )
            return result

        if status != "launched":
            result = dict(launch)
            result.update(parsed)
            result["status_file"] = status_file
            return result

        phase_started = (
            parsed.get("phaseStarted")
            or launch.get("rebuild_phase_started")
        )
        if not phase_started:
            return {
                "success": False,
                "error": (
                    "Rebuild launched but did not return phaseStarted; "
                    "cannot correlate the status file."
                ),
                "status_file": status_file,
                "result": launch.get("result"),
            }

        await reporter.emit(
            ctx,
            message="Rebuild launched; waiting for build callbacks",
        )

        if callback_queue is not None:
            callback_task = asyncio.create_task(
                _forward_rebuild_callbacks(
                    callback_queue,
                    ctx,
                    reporter,
                    callback_state,
                )
            )

        try:
            watched = await wait_for_rebuild_status(
                status_file,
                str(phase_started),
                timeout_sec=timeout,
                ctx=ctx,
                reporter=reporter,
                preexisting_access_pids=preexisting_access_pids,
            )
        except asyncio.CancelledError:
            abandoned = True
            raise
        abandoned = watched.get("status") not in REBUILD_TERMINAL_STATUSES
        if callback_task is not None and callback_task.done():
            await callback_task
            callback_task = None

        watched["rebuild_phase_started"] = phase_started
        watched["rebuild_status_file"] = status_file
        if callback_state["log_messages"]:
            watched["log_messages"] = callback_state["log_messages"]
        if callback_state.get("log_path"):
            watched.setdefault("log_path", callback_state["log_path"])
        if status_before is not None:
            watched["rebuild_status_before"] = status_before
        if launch.get("rebuild_status_superseded") is not None:
            watched["rebuild_status_superseded"] = launch["rebuild_status_superseded"]
        return watched

    except PermissionError as e:
        return {"success": False, "error": str(e)}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        if callback_task is not None:
            callback_task.cancel()
            try:
                await callback_task
            except asyncio.CancelledError:
                pass
        if op_manager is not None and operation_id is not None:
            if abandoned:
                op_manager.abandon_operation(operation_id)
            else:
                op_manager.unregister_operation(operation_id)


async def _forward_rebuild_callbacks(
    queue: asyncio.Queue,
    ctx: Context | None,
    reporter: MonotonicProgressReporter,
    state: dict[str, Any],
) -> None:
    """Forward the builder Access instance's existing HTTP callback stream.

    The build phase emits its own terminal callback before the disconnected
    worker compiles and installs the add-in. That callback ends this pump, not
    ``vcs_rebuild_addin``; the status watcher remains authoritative for the
    whole rebuild.
    """
    while True:
        callback = await queue.get()
        msg_type = str(callback.get("type") or "")
        message = str(callback.get("message") or "")

        if msg_type == "progress":
            await reporter.emit(
                ctx,
                message=message,
                vba_progress=callback.get("progress"),
                vba_total=callback.get("total"),
            )
        elif msg_type == "log":
            if message:
                state["log_messages"].append(message)
                await reporter.emit(ctx, message=message)
        elif msg_type == "complete":
            state["log_path"] = callback.get("log_path")
            state["result"] = callback.get("result")
            await reporter.emit(
                ctx,
                message=f"Build phase complete: {message}".rstrip(": "),
            )
            return
        elif msg_type == "error":
            state["log_path"] = callback.get("log_path")
            state["build_error"] = message or "Build phase failed"
            await reporter.emit(
                ctx,
                message=f"Build phase error: {state['build_error']}",
            )
            return
        elif msg_type == "cancelled":
            state["build_cancelled"] = True
            await reporter.emit(
                ctx,
                message=f"Build phase cancelled: {message}".rstrip(": "),
            )
            return


def _development_addin_from_source(source_dir: Path) -> Path:
    """Return the development ``Version Control.accda`` beside a source folder."""
    name = source_dir.name
    if name.lower().endswith(".src"):
        host = source_dir.parent / name[:-4]
    else:
        host = source_dir.parent / "Version Control.accda"
    if not host.is_file():
        raise ValueError(
            f"Development add-in not found beside source folder: {host}. "
            "vcs_rebuild_addin expects Version Control.accda.src next to "
            "Version Control.accda in the add-in repository."
        )
    return host


def _parse_rebuild_launch(result: Any) -> dict[str, Any]:
    if isinstance(result, dict):
        return result
    if not isinstance(result, str):
        return {}
    try:
        parsed = json.loads(result)
    except (json.JSONDecodeError, TypeError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _with_rebuild_fields(
    result: dict[str, Any],
    status_file: str | None,
    status_before: dict[str, Any] | None,
) -> dict[str, Any]:
    if status_file:
        result["rebuild_status_file"] = status_file
        result["status_file"] = status_file
    if status_before is not None:
        result["rebuild_status_before"] = status_before
    result.update(_describe_rebuild_attempt(result.get("result"), status_before))
    return result


@vcs_tool("vcs_tips", meta={"anthropic/alwaysLoad": True})
def vcs_tips(topic: str | None = None) -> dict[str, Any]:
    """
    Guidance for situations the server instructions only mention: troubleshooting
    stuck or hung Access, dialogs and VBA break mode, timeouts, server_busy, VBA
    compile errors, finding logs, writing code for vcs_run_vba, and developing or
    testing the VCS add-in itself (its rebuild and self-tests).

    Call with no topic for the list. Topics: addin-development, stuck, compile,
    logs, run-vba. An unknown topic is matched against topic text. Works while
    Access is blocked or another call is running.

    Args:
        topic: Topic name, or a word to search for

    Returns:
        ``topic`` and its markdown ``content``, or ``topics`` (name and summary
        of each) when no topic is given or none matches.
    """
    return get_tips(topic)


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


@vcs_tool("vcs_get_recent_calls")
def vcs_get_recent_calls(limit: int = 10) -> dict[str, Any]:
    """
    Return recent tool-call entries from the usage JSONL log.

    After an MCP client timeout (``-32001``), the server may still finish the
    request and write a usage-log entry even though the response was discarded.
    Use this tool to see what actually ran instead of inferring from side
    effects such as ``rebuild-status.json``.

    Args:
        limit: Maximum number of recent ``tool_call`` entries (default 10)

    Returns:
        Dictionary with ``success``, ``entries``, and ``log_path``
    """
    from .usage_logging import get_log_file_path

    entries = read_recent_tool_calls(limit=limit)
    log_path = get_log_file_path()
    return {
        "success": True,
        "entries": entries,
        "log_path": str(log_path) if log_path else None,
        "count": len(entries),
    }


@vcs_tool("vcs_cancel_operation")
def vcs_cancel_operation(
    operation_id: str | None = None,
    database_path: str | None = None,
) -> dict[str, Any]:
    """
    Cancel a running long operation: export, import, rebuild, or test run.

    The target is, in order: the given operation_id; the operation running on
    database_path; or the only operation running. With several running and
    neither argument, nothing is cancelled and the running operations are listed.

    Cancellation is cooperative. The add-in polls for it about twice a second and
    stops at its next safe point; a test run stops after the current test. The
    tool that started the operation then returns cancelled: true. A canceled
    export or merge leaves the index as it was, so the next run redoes the work.

    This tool is exempt from the Access gate, so it answers while the operation
    it cancels holds Access.

    Examples:
        vcs_cancel_operation()
        vcs_cancel_operation(database_path=r"C:\\path\\to\\db.accdb")
        vcs_cancel_operation("a1b2c3d4-5678-90ab-cdef-1234567890ab")

    Args:
        operation_id: Operation to cancel (busy errors report it as active_operation_id)
        database_path: Database whose running operation to cancel

    Returns:
        Dictionary with success, the operation_id and command cancelled, and a
        message; or an error, with "operations" listing what is running when the
        target was ambiguous or not found.
    """
    op_manager = _get_operation_manager()
    if not op_manager:
        return {"success": False, "error": "Callback system not available"}

    def _cancelled(op_id: str, command: str | None) -> dict[str, Any]:
        return {
            "success": True,
            "operation_id": op_id,
            "command": command,
            "message": "Cancellation requested. The operation stops at its next safe point.",
        }

    if operation_id:
        operation = op_manager.get_operation(operation_id)
        if op_manager.request_cancel(operation_id):
            return _cancelled(operation_id, operation.command if operation else None)
        return {
            "success": False,
            "operation_id": operation_id,
            "error": "Operation not found or already completed",
            "operations": op_manager.list_active_operations(),
        }

    if database_path:
        operation = op_manager.get_active_operation_for_database(database_path)
        if operation and op_manager.request_cancel(operation.operation_id):
            return _cancelled(operation.operation_id, operation.command)
        return {
            "success": False,
            "database_path": database_path,
            "error": "No operation is running on that database",
            "operations": op_manager.list_active_operations(),
        }

    active = op_manager.list_active_operations()
    if not active:
        return {"success": False, "error": "No operation is running"}
    if len(active) > 1:
        return {
            "success": False,
            "error": "Several operations are running. Pass operation_id or database_path.",
            "operations": active,
        }
    only = active[0]
    op_manager.request_cancel(only["operation_id"])
    return _cancelled(only["operation_id"], only["command"])


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
        - agent_guidance: Present when compiled is False — hand off to the user
          via VBE compile before proceeding with edits
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
            
            result: dict[str, Any] = {
                "success": True,
                "compiled": bool(is_compiled),
            }
            if not is_compiled:
                result["agent_guidance"] = _NOT_COMPILED_AGENT_GUIDANCE
            return result
    
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
    
    Compiles from the VBA editor, the same as Debug → Compile. On an error
    the editor moves to the failing line; the tool reads it, dismisses the
    "Compile error" dialog, and returns the location in ``compile_error``.
    The editor is left open on the error.

    When compilation fails, fix the reported line in source, import it, and
    compile again. If ``compile_error`` has no ``line`` (only the module from
    the editor's window title), take a screenshot with target="vbe".

    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        suppress_warnings: If True, suppress message boxes during compilation.
                         Warning: If code crashes, warnings may remain disabled.

    Returns:
        Dictionary with:
        - success: True if the project is compiled
        - compile_error: On failure, ``message`` ("Compile error: Variable not
          defined"), ``module``, ``line``, ``start_column``, ``end_line``,
          ``end_column``, ``code``, ``context`` (nearby lines), and
          ``location_source`` ("editor" from the code pane, "window" from
          the window title only)
        - agent_guidance: Present when success is False
        - error / error_pattern: when the compile could not run or finish
          (``dialog_shown``: another dialog appeared; see ``dialogs``)
    """
    try:
        # Validate path
        db_path = validate_database_path(database_path)
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            from . import compile_check
            from .access_com.process_qos import pid_from_access_app

            access_pid = pid_from_access_app(app)
            located = None
            if access_pid:
                if suppress_warnings:
                    app.DoCmd.SetWarnings(False)
                try:
                    located = compile_check.compile_with_location(app, access_pid)
                finally:
                    if suppress_warnings:
                        app.DoCmd.SetWarnings(True)
            if located is not None:
                if not located.get("success"):
                    error = located.get("compile_error") or {}
                    located["agent_guidance"] = (
                        _COMPILE_ERROR_LOCATED_GUIDANCE if error.get("line")
                        else _COMPILE_FAILURE_AGENT_GUIDANCE
                    )
                return located

            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            # Call CompileVBA API with suppress_warnings parameter
            compile_result = addin.call_sync("CompileVBA", suppress_warnings)
            
            if compile_result:
                return {"success": True}

            return {
                "success": False,
                "agent_guidance": _COMPILE_FAILURE_AGENT_GUIDANCE,
            }
    
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "agent_guidance": _COMPILE_FAILURE_AGENT_GUIDANCE,
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
        Dictionary with success status, file path, and any errors, plus
        ``log_path`` for this run's log and ``log_excerpt`` on failure.
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Export").
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("ExportObject", object_type, object_name)
            
            return _addin_json_result(result_json)
    
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
        Dictionary with success status and any errors, plus ``log_path`` for
        this run's log and ``log_excerpt`` (tail of the log) on failure.
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Merge").
    A ``warnings`` entry says when another session has the database open.
    """
    try:
        config = get_config()
        check_write_permission(config)
        
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            advisory = shared_open_warnings(str(db_path), app)
            
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("ImportObject", object_type, object_name)
            
            return with_warnings_first(_addin_json_result(result_json), advisory)
    
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
    args: list[str] | None = None,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """
    Call an existing public VBA function by name.
    
    Invokes a function that already exists in the database or a loaded library
    via Application.Run. Lighter weight than vcs_run_vba since there is no
    temp module creation or compilation step.

    **This is the correct tool for reaching the VCS add-in's own API.** Pass "VCS.API"
    (or "Version Control.API" / "MSAccessVCS.API") as function_name and it is rewritten
    to the configured add-in's full path, which also loads the add-in on demand. Do NOT
    try to reach the API from inside vcs_run_vba: that code is itself delivered through
    modAPI.API, so calling back into the API is a re-entrant call and will be refused.

    **Prefer ``vcs_rebuild_addin(source_dir)`` to rebuild the add-in.** This tool
    remains a launch-only escape hatch for ``RebuildAddIn``: it returns when the
    worker is confirmed and does not wait for install. Host it on the development
    copy (the ``Version Control.accda`` beside the source folder). The installed
    add-in is refused as ``database_path`` here as it is everywhere, with
    ``installed_addin_refused``.

    Examples:
        vcs_call_vba("C:\\\\db.accdb", "MyModule.GetQuerySQL", ["qryCustomers"])
        vcs_call_vba("C:\\\\db.accdb", "VCS.API", ["GetVCSVersion"])
        vcs_call_vba("C:\\\\db.accdb", "VCS.API", ["RunRoundtripTests", "C:\\\\fixtures\\\\"])
        vcs_call_vba("C:\\\\db.accdb", "VCS.API", ["RebuildAddIn", "C:\\\\Repos\\\\msaccess-vcs-addin\\\\Version Control.accda.src\\\\"])

    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        function_name: Fully qualified function name (e.g., "ModuleName.FunctionName"),
            or a VCS add-in alias such as "VCS.API"
        args: Optional list of string arguments to pass to the function
        timeout_seconds: Optional server-side timeout. Defaults to
            ACCESS_VCS_CALL_VBA_TIMEOUT_SEC (45 seconds). Keep below the MCP
            client's request timeout or the client may return -32001 with no JSON.

    Returns:
        Dictionary with the function's return value or error. RebuildAddIn calls
        also include ``rebuild_status_file`` and ``rebuild_status_before``.
    """
    return _execute_call_vba(database_path, function_name, args, timeout_seconds)


def _execute_call_vba(
    database_path: str,
    function_name: str,
    args: list[str] | None = None,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """Invoke Application.Run and normalize the result. Used by vcs_call_vba
    and by vcs_rebuild_addin for the launch phase only.
    """
    rebuild_status_file: str | None = None
    rebuild_status_before: dict[str, Any] | None = None

    def _with_rebuild_context(result: dict[str, Any]) -> dict[str, Any]:
        if rebuild_status_file:
            result["rebuild_status_file"] = rebuild_status_file
        if rebuild_status_before is not None:
            result["rebuild_status_before"] = rebuild_status_before
        return result

    try:
        db_path = validate_database_path(database_path)
        call_args = args or []
        resolved_name = _resolve_addin_function_name(function_name)
        call_description = resolved_name
        if call_args:
            call_description += f"({', '.join(repr(a) for a in call_args)})"
        log_code_execution("vcs_call_vba", str(db_path), call_description, code_type="vba_call")

        if len(call_args) > 3:
            return _with_rebuild_context({
                "success": False,
                "error": "Maximum 3 arguments supported for vcs_call_vba",
            })

        if (
            _is_addin_api_resolved_name(resolved_name)
            and len(call_args) >= 2
            and str(call_args[0]).lower() == "setoption"
            and _is_user_controlled_option(call_args[1])
        ):
            return _user_controlled_option_refusal(call_args[1])

        is_rebuild = (
            _is_addin_api_resolved_name(resolved_name)
            and len(call_args) >= 2
            and str(call_args[0]) == "RebuildAddIn"
        )
        if is_rebuild:
            rebuild_status_file, rebuild_status_before = _snapshot_rebuild_status(str(call_args[1]))

        timeout = get_call_vba_timeout(timeout_seconds)

        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()

            run_result = _run_application_call_with_timeout(
                app, resolved_name, call_args, timeout, str(db_path)
            )

            if run_result.get("timed_out"):
                return _with_rebuild_context({
                    "success": False,
                    "error": run_result["error"],
                    "error_pattern": "timeout",
                    "recoverable": True,
                    "timed_out": True,
                    "function": resolved_name,
                    "duration_ms": run_result.get("duration_ms"),
                })

            if not run_result.get("success"):
                exc = run_result.get("error")
                return _with_rebuild_context({
                    "success": False,
                    "error": _describe_run_failure(exc, resolved_name, function_name),
                    "function": resolved_name,
                })

            result = run_result["result"]

            # Early-bound Run returns a tuple: the function's return value followed by
            # Run's own 30 Arg slots (unused ones show DISP_E_PARAMNOTFOUND). Only the
            # first element is ours. Matches VCSAddinIntegration.call_api_function.
            if isinstance(result, tuple) and len(result) > 0:
                result = result[0]

            # A refused call comes back as a marked string rather than an exception --
            # the add-in cannot raise across the library boundary without producing a
            # modal dialog. Report it as a failure so it is not mistaken for data.
            if isinstance(result, str) and result.startswith(_API_REFUSED_PREFIX):
                return _with_rebuild_context({
                    "success": False,
                    "error": result[len(_API_REFUSED_PREFIX):],
                    "function": resolved_name,
                })

            response: dict[str, Any] = {
                "success": True,
                "result": str(result) if result is not None else None,
                "function": resolved_name,
            }

            if is_rebuild:
                _log_rebuild_call_result(result)
                response.update(
                    _describe_rebuild_attempt(result, rebuild_status_before)
                )

            return _with_rebuild_context(response)

    except Exception as e:
        return _with_rebuild_context({"success": False, "error": str(e)})


# Qualifiers that mean "the VCS add-in library". Application.Run resolves a bare
# qualifier against the VBA *project* name (MSAccessVCS), not the file name
# (Version Control), and only once the add-in is already loaded -- so the file-name
# form never works and the project-name form works only sometimes. Rewriting to the
# configured full path is correct from a cold start and loads the add-in on demand.
_ADDIN_QUALIFIER_ALIASES = frozenset({"vcs", "version control", "msaccessvcs"})

def _resolve_addin_function_name(function_name: str) -> str:
    """Rewrite a VCS add-in alias qualifier to the configured add-in's full path."""
    qualifier, sep, member = function_name.rpartition(".")
    if not sep or qualifier.lower() not in _ADDIN_QUALIFIER_ALIASES:
        return function_name

    addin_path = get_config().get("ACCESS_VCS_ADDIN_PATH")
    if not addin_path:
        return function_name

    return f"{os.path.splitext(os.path.abspath(addin_path))[0]}.{member}"


def _describe_run_failure(error: Exception, resolved_name: str, requested_name: str) -> str:
    """Add context to an Application.Run failure where the raw COM error is unhelpful.

    "Cannot find the procedure" is the one worth explaining: it usually means the
    qualifier was a file name, and Application.Run matches loaded VBA project names.
    """
    text = f"VBA function call failed: {error}"

    if "cannot find the procedure" not in str(error).lower():
        return text

    if resolved_name != requested_name:
        text += f"\nResolved '{requested_name}' to '{resolved_name}'."

    return text + (
        "\nApplication.Run resolves the qualifier against a loaded VBA project name, "
        'not a file name. To reach the VCS add-in, pass "VCS.API" as function_name; '
        "it is rewritten to the add-in's full path, which also loads it on demand."
    )


def _configured_addin_lib_prefix() -> str | None:
    addin_path = get_config().get("ACCESS_VCS_ADDIN_PATH")
    if not addin_path:
        return None
    return os.path.splitext(os.path.abspath(addin_path))[0]


def _is_addin_api_resolved_name(resolved_name: str) -> bool:
    prefix = _configured_addin_lib_prefix()
    if not prefix:
        return False
    return resolved_name.lower().startswith(prefix.lower() + ".")


def _is_installed_addin_path(database_path: str) -> bool:
    """True when a path names the installed add-in file.

    Compared without the extension, mirroring the add-in's own
    ``modInstall.PathsMatchIgnoringExtension``: an install configured for the compiled
    add-in is a ``.accde`` built from the same ``.accda``, and the configured path names
    only one of the two. Matching on the full name lets the other variant slip past.
    """
    addin_path = get_config().get("ACCESS_VCS_ADDIN_PATH")
    if not addin_path:
        return False
    return os.path.normcase(os.path.splitext(database_path)[0]) == os.path.normcase(
        os.path.splitext(os.path.abspath(addin_path))[0]
    )


# Parameters that name a file a tool opens, writes, or replaces. Checked against the
# installed add-in on every call. `source_dir` and `output_dir` are folders, and an
# export folder beside the install is not the install.
_ADDIN_TARGET_PARAMETERS = ("database_path", "output_path", "template_path")


def _installed_addin_target_refusal(
    tool: str, parameter: str, target: str
) -> dict[str, Any]:
    return {
        "success": False,
        "error": (
            f"Refusing {tool}: {parameter} names the installed add-in ({target}). That "
            "file exists to be loaded as a library -- it is never opened as a database, "
            "and never modified in place, which would reset a VBA project while it is "
            "executing. Work on the development copy of the add-in in its repository, "
            "the 'Version Control.accda' beside its source folder, and rebuild from "
            "there; the rebuild is what replaces the installed file. Do not substitute "
            "a user database, anything in the repository's Testing folder, or a scratch "
            ".accdb. vcs_get_version_info() reports the installed add-in's version "
            "without opening it."
        ),
        "error_pattern": "installed_addin_refused",
        "recoverable": True,
    }


def _refuse_installed_addin_target(
    tool: str, func: Any, args: tuple[Any, ...], kwargs: dict[str, Any]
) -> dict[str, Any] | None:
    """Refuse any call that names the installed add-in as a file to act on.

    One check for every tool, ahead of the gate and of any COM work, because the harm
    is in opening or replacing the file at all rather than in what a particular tool
    goes on to do with it. Per-tool guards had already been written twice and still
    left `vcs_export_database`, `vcs_run_vba`, and the rebuild's `output_path`
    uncovered.
    """
    try:
        bound = inspect.signature(func).bind_partial(*args, **kwargs)
    except TypeError:
        # A malformed call: let the real signature error surface from the body.
        return None

    for parameter in _ADDIN_TARGET_PARAMETERS:
        value = bound.arguments.get(parameter)
        if isinstance(value, str) and value and _is_installed_addin_path(value):
            log_diagnostic_event(
                "installed_addin_refused", tool=tool, parameter=parameter
            )
            return _installed_addin_target_refusal(tool, parameter, value)
    return None


def _describe_rebuild_attempt(
    result: Any, status_before: dict[str, Any] | None
) -> dict[str, Any]:
    """Report which attempt the status file now describes.

    RebuildAddIn returns the ``phaseStarted`` it stamped, and holds that value for the
    rest of the run, so it identifies the attempt across every record the run writes.
    Comparing it against the snapshot taken before the call is what distinguishes a file
    this attempt wrote from one left behind by an earlier run -- the mistake that used to
    let a refusal be read as a prior ``complete``.
    """
    if not isinstance(result, str):
        return {}

    try:
        parsed = json.loads(result)
    except (json.JSONDecodeError, TypeError):
        return {}
    if not isinstance(parsed, dict):
        return {}

    phase_started = parsed.get("phaseStarted")
    if not phase_started:
        return {}

    described: dict[str, Any] = {"rebuild_phase_started": phase_started}
    if status_before is not None:
        described["rebuild_status_superseded"] = (
            status_before.get("phaseStarted") != phase_started
        )
    return described


def _rebuild_status_path(source_dir: str) -> str:
    return os.path.join(source_dir.rstrip("\\/"), "logs", "rebuild-status.json")


def _snapshot_rebuild_status(source_dir: str) -> tuple[str, dict[str, Any] | None]:
    status_file = _rebuild_status_path(source_dir)
    if not os.path.isfile(status_file):
        return status_file, None

    try:
        mtime = os.path.getmtime(status_file)
        # The add-in writes this file as UTF-8 with a BOM, which plain "utf-8" keeps
        # in the string and json.load then rejects. Reading it as "utf-8" reported
        # read_error for every snapshot and left the phaseStarted comparison inert.
        with open(status_file, encoding="utf-8-sig") as handle:
            data = json.load(handle)
        return status_file, {
            "status": data.get("status"),
            "updated": data.get("updated"),
            "phaseStarted": data.get("phaseStarted"),
            "mtime": datetime.fromtimestamp(mtime, tz=timezone.utc).isoformat(),
        }
    except (OSError, json.JSONDecodeError):
        return status_file, {"status": None, "updated": None, "read_error": True}


def _log_rebuild_call_result(result: Any) -> None:
    status_value: str | None = None
    if isinstance(result, str):
        try:
            parsed = json.loads(result)
            if isinstance(parsed, dict):
                status_value = parsed.get("status")
        except json.JSONDecodeError:
            pass
    elif isinstance(result, dict):
        status_value = result.get("status")

    log_diagnostic_event(
        "rebuild_addin_call_result",
        status=status_value,
        result_type=type(result).__name__,
    )


def _run_application_call_with_timeout(
    app: Any,
    resolved_name: str,
    call_args: list[str],
    timeout_seconds: float,
    db_path: str | None = None,
) -> dict[str, Any]:
    """Run ``Application.Run`` in a daemon thread with a hard join timeout.

    The thread needs its own COM apartment and its own proxy. Handing it the caller's
    STA proxy fails before it ever reaches Access -- "CoInitialize has not been called"
    while the thread has no apartment, RPC_E_WRONG_THREAD once it does -- which is what
    made RebuildAddIn unreachable through this tool.
    ``VCSAddinIntegration.worker_app_source`` supplies one: from the Running Object
    Table with ``db_path``, marshalled from the caller without it.
    """
    start = time.perf_counter()
    thread, result_box = _start_application_call(
        VCSAddinIntegration.worker_app_source(app, db_path), resolved_name, call_args
    )
    thread.join(timeout=timeout_seconds)
    duration_ms = round((time.perf_counter() - start) * 1000, 2)

    if thread.is_alive():
        return {
            "success": False,
            "timed_out": True,
            "error": f"VBA function call timed out after {timeout_seconds} seconds",
            "duration_ms": duration_ms,
        }

    result_box["duration_ms"] = duration_ms
    return result_box


def _start_application_call(
    acquire_app: Any,
    resolved_name: str,
    call_args: list[str],
) -> tuple[threading.Thread, dict[str, Any]]:
    """Start ``Application.Run`` on a daemon thread; return it and its result box.

    ``acquire_app`` comes from ``VCSAddinIntegration.worker_app_source``, called on
    the thread that owns the proxy. The box gets ``success`` and ``result`` or
    ``error`` once the thread ends.
    """
    result_box: dict[str, Any] = {}

    def worker() -> None:
        import pythoncom

        try:
            pythoncom.CoInitialize()
            try:
                worker_app = acquire_app()
                if worker_app is None:
                    raise RuntimeError(
                        "Cannot find the Access instance from the worker thread. "
                        "The Access application may have been closed."
                    )
                result_box["result"] = worker_app.Run(resolved_name, *call_args)
                result_box["success"] = True
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        except Exception as exc:
            result_box["success"] = False
            result_box["error"] = exc

    thread = threading.Thread(target=worker, daemon=True, name="vcs-call-vba")
    thread.start()
    return thread, result_box


@vcs_tool("vcs_run_vba")
def vcs_run_vba(
    database_path: str,
    code: str,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """
    Execute agent-generated VBA code in a temporary module.

    Your code becomes the body of a function named MCP_TempFunction; assign to
    it to return a value. Send statements, not a bare expression:
        Dim qd As DAO.QueryDef
        Set qd = CurrentDb.QueryDefs("qryExample")
        MCP_TempFunction = qd.SQL

    Returns `success` and `result`; on failure `error`, `error_pattern`,
    `errorNumber`, and `errorLine` (the 1-based line of your `code` that raised
    the captured error). `compileError`, `sweptModules`, and the temp-module
    fields below appear when they apply.

    **Requires McpAllowRunVBA** (default: off). The user sets it in the add-in's
    MCP options form; no tool can change it. To call an existing function,
    including the add-in's API, use vcs_call_vba instead. Do not call the API
    from inside this code; the re-entrant call is refused.

    **Errors and line numbers:** the wrapper runs with `On Error Resume Next`
    and reports the last runtime error. Lines are numbered for you, counting
    blank lines, comments, and `_` continuations, so `errorLine: 7` is line 7
    of what you sent. To collect every failing line in one call, use your own
    handler; `Erl` matches `errorLine`:

        Dim col As New Collection
        On Error GoTo H
        CurrentDb.Execute "DELETE * FROM tblA"
        CurrentDb.Execute "UPDATE tblC SET x = 1"
        MCP_TempFunction = "errors=" & col.Count
        Exit Function
        H: col.Add Erl & ": " & Err.Number & " " & Err.Description
        Resume Next

    **Host-project compile failures:** `compileError` saying the host project
    does not compile on its own means the error is in the database's existing
    VBA, not in your `code`. Run vcs_compile_vba, which returns the failing
    module and line, rather than editing your snippet.

    **Reset and temporary modules:** before each run the worker resets the host
    project's VBA and sweeps stale `MCP_Temp_*` modules. A failed reset is
    fail-closed: your code does not run.
      - `sweptModules`: stale modules removed before the payload
      - `temp_module_sweep_failed` + `orphanModules`: stale modules remain;
        payload not started
      - `temp_module_unresolvable` + `tempModule`: post-compile canary failed;
        payload not started
      - `temp_module_cleanup_failed` + `cleanupFailed` + `orphanModule`: the
        wrapper survived cleanup; a completed return is kept as `payloadResult`
    When cleanup fails, stop and tell the user rather than calling again.

    More patterns: vcs_tips("run-vba").

    Examples:
        vcs_run_vba("C:\\\\db.accdb", "MCP_TempFunction = CurrentDb.TableDefs.Count")

    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        code: VBA code to execute (statements, not just an expression)
        timeout_seconds: Optional parent-side timeout. Defaults to
            ACCESS_VCS_RUN_VBA_TIMEOUT_SEC (45 seconds). Keep it below the
            client's request timeout.
    """
    try:
        db_path = validate_database_path(database_path)
        log_code_execution("vcs_run_vba", str(db_path), code, code_type="vba")

        config = get_config()
        worker_result = run_vba_resilient(
            database_path=str(db_path),
            code=code,
            addin_path=config.get("ACCESS_VCS_ADDIN_PATH"),
            timeout_seconds=timeout_seconds,
        )
        if not worker_result.get("success"):
            return worker_result

        result_json = worker_result.get("result")
        if isinstance(result_json, str):
            try:
                return json.loads(result_json)
            except json.JSONDecodeError:
                return {"success": True, "result": result_json}

        return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


# The MCP permissions. The add-in refuses them in SetOption and in session override
# files too; refusing here as well covers add-in builds older than that check.
USER_CONTROLLED_OPTIONS = frozenset({
    "mcpallowrunvba",
    "mcpallowcallvba",
    "mcpallowexecutesql",
    "mcpallowimport",
})


def _is_user_controlled_option(option_name: Any) -> bool:
    return str(option_name).strip().lower() in USER_CONTROLLED_OPTIONS


def _user_controlled_option_refusal(option_name: Any) -> dict[str, Any]:
    return {
        "success": False,
        "error": (
            f"{option_name} is the user's decision. It is set in the add-in's MCP "
            "options form and cannot be changed through any tool, including "
            "vcs_call_vba. Ask the user to change it."
        ),
        "error_pattern": "user_controlled_option",
    }


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
    if _is_user_controlled_option(option_name):
        return _user_controlled_option_refusal(option_name)

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
    specified type, from the source folder's ``logs`` directory.
    
    **Pick the type that matches the operation you ran.** Each operation writes
    its own log family, so asking for the wrong one silently returns a stale
    log from a different run:
    
    - ``"Export"``  -- vcs_export_database, vcs_export_object
    - ``"Merge"``   -- vcs_import_objects, vcs_import_object
    - ``"Build"``   -- vcs_rebuild_database
    - ``"TestRun"`` -- vcs_run_tests
    - ``"Other"``   -- anything else
    
    Prefer the ``log_path`` returned by the operation itself; use this tool
    when you no longer have it. These logs are gitignored, so Glob/Grep will
    not find them.
    
    Examples:
        vcs_get_log("C:\\\\db.accdb")
        vcs_get_log("C:\\\\db.accdb", log_type="Merge")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        log_type: Log family to read: "Export" (default), "Merge", "Build",
            "TestRun", or "Other"
    
    Returns:
        Dictionary with log content, log_path, and success status
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))
            
            result_json = addin.call_sync("GetLogContent", log_type)

            return _addin_json_result(result_json, raw_key="content")
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_run_tests")
async def vcs_run_tests(
    database_path: str,
    filter: str | None = None,
    timeout_seconds: float | None = None,
    ctx: Context = None,
) -> dict[str, Any]:
    """
    Run VBA tests in the database using the VCS add-in's built-in test runner.

    Discovers test modules (standard modules and classes containing TestAssert
    calls), executes their test procedures, and returns structured JSON results
    with per-test status, assertion details, timing, and tags.

    Headless here means no add-in UI (no web runner, no console form, silent
    dialogs) -- not a hidden Access window. The host instance stays visible.

    **Live output:** MCP progress is best-effort in Cursor. For a live stream
    (dots for fast passes, names for tests ≥ 1s, FAIL lines, then a human
    completion line), run ``msaccess-vcs run-tests <database>`` from a
    terminal and keep that command in the foreground. This tool still returns
    the full ``tests`` map for programmatic reruns.

    **Filter syntax** (comma-separated, applied as a single string):

    Each comma-separated element is resolved in priority order:
    1. Module name -- exact match (e.g. ``modTestEncoding``)
    2. Suite/folder -- match against ``@Folder`` annotations (e.g. ``SQL``)
    3. Procedure name -- match on procedure or ``Module.Procedure`` key
    4. Tag -- match against ``@Tag`` annotations (e.g. ``unit``)

    Prefix any element with ``-`` to exclude. Inclusions combine with OR;
    exclusions combine with AND.

    **Iterative workflow:** The response includes per-test entries keyed by
    ``Module.Procedure``. To rerun only failures, pass those keys back as
    the filter (e.g. ``"modTestFoo.TestBar,clsTestBaz.TestQux"``).

    **Prerequisite:** The target database must have ``modTestAssert`` installed
    (via the VCS ribbon or ``VCS.InstallTestAssertModule``). In unattended mode
    the install prompt is suppressed, so pre-install before calling this tool.

    **To run the add-in's own suite, pass the development copy in its repository** --
    the ``Version Control.accda`` beside ``Version Control.accda.src``. The runner scans
    the current VBA project, so the host database is the code under test, while the
    installed add-in loads as a library and supplies the runner and TestAssert. Passing
    the installed add-in is refused with ``installed_addin_refused`` before it is opened.

    Examples:
        vcs_run_tests("C:\\\\db.accdb")
        vcs_run_tests("C:\\\\db.accdb", filter="modTestEncoding")
        vcs_run_tests("C:\\\\db.accdb", filter="SQL,-slow")
        vcs_run_tests("C:\\\\db.accdb", filter="modTestFoo.TestSpecificProc")

    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        filter: Optional comma-separated filter string. When omitted, runs
            all tests.
        timeout_seconds: How long to wait for the async run (default from the
            add-in's timeout_ms, 10 minutes). Unused on the sync fallback.

    Returns:
        Dictionary with ``success`` (True when all tests pass, none errored,
        and at least one test ran), ``summary``, ``tests``, ``durationMs``,
        and other fields from the test runner JSON output.
    """
    try:
        db_path = validate_database_path(database_path)

        busy_error = _check_database_busy(str(db_path))
        if busy_error:
            return busy_error

        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()

            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=str(db_path))

            # Silent mode: suppress MsgBox dialogs during test run
            addin_lib = os.path.splitext(os.path.abspath(addin.addin_path))[0]
            app.Run(f"{addin_lib}.SetInteractionMode", 1)

            command, command_args = _test_run_command(
                addin.call_sync("GetVCSVersion"), filter
            )
            if command == "RunFilteredTests":
                addin.call_sync("SetOption", "DefaultTestFilter", filter or "")

            callback_url = get_callback_url()
            op_manager = _get_operation_manager()

            if callback_url and op_manager:
                op_manager.set_event_loop(asyncio.get_running_loop())
                operation_id, _queue = op_manager.register_operation(
                    database_path=str(db_path),
                    command=command,
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                try:
                    async_result = addin.call_async(callback_info, command, *command_args)
                    if async_result.get("sync"):
                        op_manager.unregister_operation(operation_id)
                        return _parse_test_runner_json(async_result.get("result"))
                    if async_result.get("async"):
                        timeout_ms = async_result.get("timeout_ms", 600000)
                        wait_timeout = (
                            timeout_seconds
                            if timeout_seconds is not None
                            else timeout_ms / 1000
                        )
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=ctx,
                            timeout_seconds=wait_timeout,
                        )
                        return _test_results_from_completion(completion)
                    op_manager.unregister_operation(operation_id)
                except Exception:
                    op_manager.unregister_operation(operation_id)

            return _parse_test_runner_json(addin.call_sync(command, *command_args))

    except Exception as e:
        return {"success": False, "error": str(e)}


@vcs_tool("vcs_capture_screenshot")
async def vcs_capture_screenshot(
    database_path: str | None = None,
    target: Literal["application", "vbe", "active", "object"] = "application",
    object_type: Literal["form", "report", "table", "query"] | None = None,
    object_name: str | None = None,
    open_if_closed: bool = False,
    pid: int | None = None,
    max_dimension: int = _screenshot.DEFAULT_MAX_DIMENSION,
    include_image: bool = True,
    view: Literal["normal", "design"] = "normal",
    settle_ms: int = _screenshot.DEFAULT_SETTLE_MS,
) -> CallToolResult:
    """
    Capture a PNG of a running Access window, plus the text of its dialogs.

    Use it to see what Access is doing: when a call hung or timed out, when
    Access may be waiting on a MsgBox or prompt, after a compile failure
    (target="vbe" shows the module and highlighted line), or to check a
    form's or report's layout. Read ``dialogs`` in the metadata before the
    image: it carries each dialog's text and button captions.

    Targets:
    - "application": the Access main window, with its dialogs and popups.
    - "vbe": the VBA editor window.
    - "active": the active form, report, or datasheet (needs database_path).
    - "object": a named form, report, table, or query (needs database_path,
      object_type, and object_name).

    "application" and "vbe" use Windows APIs only. They never call into
    Access or wait for other tool calls, so they work while Access sits on a
    dialog, in break mode, or inside a long call. "active" and "object" ask
    Access for the object's window; they refuse (``access_blocked``) when
    Access is showing a dialog, and return the application capture instead.

    Each dialog carries a ``dialog_id``. When ``owned_by_server`` is true,
    pass it to vcs_dismiss_dialog to answer the dialog; ``vbe_state: "break"``
    means vcs_reset_vba can end a run stopped in break mode.

    The tool never starts Access, opens a database, restores, focuses, or
    closes a window, or clicks a dialog button. With
    ``open_if_closed=True`` it opens a closed object for the capture and
    closes it without saving. Opening runs the object's event code. Reports
    open in print preview; tables and queries open read-only; only select,
    crosstab, and union queries are opened. ``view="design"`` opens a form or
    report in design view instead, which runs no event code: use it to check
    a layout changed in source. An object already open is captured as it is,
    and its view is reported as ``object.current_view``.

    Args:
        database_path: Database whose Access window to capture. Optional for
            "application" and "vbe" when only one Access is running or pid
            is given.
        target: "application" (default), "vbe", "active", or "object".
        object_type: "form", "report", "table", or "query" (target="object").
        object_name: Object name (target="object").
        open_if_closed: Open a closed object for the capture, then close it.
        pid: MSACCESS.EXE process id, to pick one of several instances.
        max_dimension: Longest image edge in pixels (320-3840). Images are
            scaled down, never up.
        include_image: False returns the metadata and dialog text only.
        view: "normal" (default) or "design", for forms and reports opened
            with open_if_closed.
        settle_ms: Wait after opening an object before capturing (0-5000,
            default 150). Raise it when an opened form shows blank controls.

    Returns:
        A JSON metadata block (``success``, ``pid``, ``owned_by_server``,
        ``window``, ``vbe_state``, ``dialogs``, ``modal_blocked``, ``image``,
        ``warnings``; ``error`` and ``error_pattern`` on failure) followed by
        the PNG when one was taken.

    Limits: minimized, hidden, or cloaked windows cannot be captured. A window
    that does not answer messages is copied from the screen instead
    (``capture_method: "screen"``), with other applications' windows grayed out.
    With ``Application.Echo False`` the capture shows a stale frame. Captures
    can show client data; they are returned inline and never saved to disk.
    """
    if database_path:
        try:
            database_path = str(validate_database_path(database_path))
        except ValueError as exc:
            return _screenshot.error_result("invalid_argument", str(exc), target=target)
    return await _screenshot.capture_screenshot(
        database_path=database_path,
        target=target,
        object_type=object_type,
        object_name=object_name,
        open_if_closed=open_if_closed,
        pid=pid,
        max_dimension=max_dimension,
        include_image=include_image,
        get_gate=get_access_gate,
        view=view,
        settle_ms=settle_ms,
    )


async def _run_blocking(fn: Any, timeout: float) -> dict[str, Any]:
    try:
        return await asyncio.wait_for(asyncio.to_thread(fn), timeout)
    except asyncio.TimeoutError:
        return {
            "success": False,
            "error": f"The call did not finish within {timeout:.0f} s.",
            "error_pattern": "access_unresponsive",
            "recoverable": True,
        }


def _optional_database(database_path: str | None) -> str | None:
    return str(validate_database_path(database_path)) if database_path else None


@vcs_tool("vcs_dismiss_dialog")
async def vcs_dismiss_dialog(
    dialog_id: str,
    button: str,
    database_path: str | None = None,
    pid: int | None = None,
) -> dict[str, Any]:
    """
    Click a button on a dialog that a server-created Access window is showing.

    Call vcs_capture_screenshot first (include_image=False is enough): pass
    a ``dialog_id`` and one of the ``buttons`` it lists. The click goes only
    to that dialog. If it was answered or replaced since the screenshot,
    the call is refused with ``dialog_changed`` and returns the current
    ``dialogs``.

    Works while another tool call is blocked behind the dialog; it does not
    wait for the Access gate. Refused (``user_owned_instance``) for Access
    windows this server did not open: ask the user to answer those.
    Refused when ACCESS_VCS_DISABLE_WRITES is set.

    Args:
        dialog_id: The dialog's ``dialog_id`` from vcs_capture_screenshot.
        button: A button caption as listed, for example "OK" or "No".
        database_path: Database whose Access window shows the dialog.
        pid: MSACCESS.EXE process id, instead of or with database_path.

    Returns:
        ``success``, ``pid``, ``dismissed`` (dialog_id, title, button),
        ``closed``, the ``dialogs`` still showing, and ``vbe_state`` (a
        dismissed error can leave VBA in break mode).
    """
    from . import vba_control

    try:
        check_write_permission(get_config())
        database_path = _optional_database(database_path)
    except (PermissionError, ValueError) as exc:
        pattern = "write_disabled" if isinstance(exc, PermissionError) else "invalid_argument"
        return {"success": False, "error": str(exc), "error_pattern": pattern}
    return await _run_blocking(
        lambda: vba_control.dismiss_dialog(database_path, pid, dialog_id, button), 20.0
    )


@vcs_tool("vcs_reset_vba")
async def vcs_reset_vba(
    database_path: str | None = None,
    pid: int | None = None,
) -> dict[str, Any]:
    """
    End VBA break mode in a server-created Access window, reporting where it stopped.

    Use when vcs_capture_screenshot reports ``vbe_state: "break"``: code
    stopped on a run-time error or a Stop statement, and every call into
    Access waits behind it. Answer any error dialog with vcs_dismiss_dialog
    first. The reset reads the module, line, and code where execution
    stopped, then runs the VBA editor's Reset command, which ends the run
    and clears module-level variables. The blocked tool call then returns.

    Does not wait for the Access gate. Refused (``user_owned_instance``) for
    Access windows this server did not open, and when
    ACCESS_VCS_DISABLE_WRITES is set.

    Args:
        database_path: Database open in the Access window in break mode.
        pid: MSACCESS.EXE process id, instead of or with database_path.

    Returns:
        ``success``, ``pid``, ``stopped_at`` (module, line, start_column,
        end_line, end_column, code, context), and ``vbe_state``.
    """
    from . import vba_control

    try:
        check_write_permission(get_config())
        database_path = _optional_database(database_path)
    except (PermissionError, ValueError) as exc:
        pattern = "write_disabled" if isinstance(exc, PermissionError) else "invalid_argument"
        return {"success": False, "error": str(exc), "error_pattern": pattern}
    return await _run_blocking(
        lambda: vba_control.reset_vba(database_path, pid),
        vba_control.RESET_TIMEOUT_SEC + vba_control.RESET_SETTLE_SEC + 10.0,
    )


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

        # The add-in clears session overrides on the running instance. With
        # no instance running there is nothing to clear, and launching one
        # just to end a session is the opposite of what this tool is for --
        # shutdown calls it unconditionally.
        if not access_instance_is_live(str(db_path)):
            return {
                "success": True,
                "session_id": session_id,
                "message": "No live Access instance; session overrides left on disk",
            }

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
