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
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

from mcp.server.fastmcp import FastMCP, Context

from .access_com.connection import (
    AccessConnection,
    access_instance_is_live,
    close_owned_instances_holding,
    ensure_access_visible,
    ensure_dispatch,
)
from .access_com.dao_helpers import list_query_defs, list_table_defs
from .access_com.process_qos import list_access_pids, prefer_full_power_if_created
from .access_gate import EXEMPT_TOOLS, get_access_gate
from .config import (
    get_config,
    get_callback_url,
    get_session_id,
    initialize_from_workspace,
    load_config,
)
from .addin_integration import API_REFUSED_PREFIX as _API_REFUSED_PREFIX, VCSAddinIntegration
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
from .rebuild_watcher import get_rebuild_timeout, wait_for_rebuild_status
from .operation_manager import MonotonicProgressReporter
from .vba_worker_manager import get_call_vba_timeout, run_vba_resilient

_COMPILE_FAILURE_AGENT_GUIDANCE = (
    "Compilation failed. MCP cannot report the failing module or line. "
    "Stop editing source files. Ask the user to compile in the VBE "
    "(Debug → Compile) — Access will jump to the error line. "
    "Ask the user to paste the code snippet around that line (a few lines "
    "above and below). Do not guess fixes or iterate blindly."
)

_NOT_COMPILED_AGENT_GUIDANCE = (
    "The VBA project is not compiled. MCP cannot report the failing module or line. "
    "Before proceeding with code edits, ask the user to compile in the VBE "
    "(Debug → Compile) and paste the code snippet around any error line, "
    "or confirm the project compiles cleanly."
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
        log_path = completion.get("log_path")
        if log_path:
            parsed.setdefault("log_path", log_path)
            parsed.setdefault("logPath", log_path)
        return parsed

    raw = completion.get("result")
    if raw not in (None, ""):
        return _parse_test_runner_json(raw)

    if completion.get("cancelled"):
        return {
            "success": False,
            "cancelled": True,
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
    if completion.get("log_path"):
        result["log_path"] = completion["log_path"]
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
        "To rebuild `Version Control.accda` from source after editing add-in files, "
        "call vcs_rebuild_addin(\"<source folder>\"). It derives the development copy "
        "of the add-in (the `Version Control.accda` beside the source folder), launches "
        "RebuildAddIn, and waits on `<source>/logs/rebuild-status.json` with filesystem "
        "notifications until status is `complete` or `*-failed`. Do NOT open a user "
        "database, anything in the repo's Testing folder, or a scratch database to host "
        "this -- rebuilding the add-in is a repository operation and belongs to the "
        "repository's own copy. Prefer this over polling the status file yourself.\n"
        "`refused` and `launch-failed` are returned immediately and mean nothing was "
        "rebuilt. Before launch the server closes Access windows it created that hold "
        "a file the rebuild replaces. The rebuild still refuses when a user-owned "
        "MSACCESS.EXE holds one of those files, and never closes another process. "
        "On refusal, `otherInstances` names each process and which file it holds; close "
        "those and call again. `launch-failed` means the helper script never started: "
        "Access is left open and the call is safe to retry.\n"
        "This is not vcs_rebuild_database, which rebuilds a user project. "
        "vcs_call_vba(db, \"VCS.API\", [\"RebuildAddIn\", source]) remains a launch-only "
        "escape hatch; it does not wait for the rebuild to finish.\n"
        "MCP progress notifications are best-effort in Cursor: the server emits them, "
        "but some Cursor builds show only \"Running...\" until the tool returns. For "
        "live terminal output, run `msaccess-vcs rebuild-addin <source>` (or export / "
        "merge / rebuild-database / run-tests).\n"
        "After any MCP client timeout (`-32001`), call vcs_get_recent_calls() to learn "
        "what actually executed — the server may have finished after the client gave up.\n"
        "One server process is shared across Cursor windows. When another window holds the "
        "Access gate, tools return `error_pattern: server_busy` with `busy_with` naming the "
        "in-flight tool — retry instead of waiting for a client timeout.\n"
        "vcs_call_vba has a server-side timeout (default 45s via "
        "ACCESS_VCS_CALL_VBA_TIMEOUT_SEC); raising it above the client's request timeout "
        "brings `-32001` back.\n\n"
        "**Running the add-in's own tests:**\n"
        "Pass the **development copy** in the add-in's repository -- the `Version "
        "Control.accda` beside `Version Control.accda.src` -- as database_path. The "
        "runner scans the current VBA project, so that copy is the code under test, "
        "while the installed add-in loads as a library and supplies the runner and "
        "TestAssert. Both roles are required and they are different files. A user "
        "database or anything in the repo's Testing folder finds that database's tests "
        "instead. This server opens the development copy for you, so no manual pre-open "
        "step is needed. Always run these tests through this server rather than from the "
        "add-in's own window -- an all-EMPTY result (zero assertions) means the harness "
        "was bypassed, not that the tests passed. For live per-test output, run "
        "`msaccess-vcs run-tests <database>` from a terminal (same pattern as "
        "rebuild-addin). Headless means no add-in UI, not a hidden Access window.\n\n"
        "**The installed add-in is never a target:**\n"
        "No tool accepts the installed add-in under %APPDATA%\\MSAccessVCS as "
        "database_path, output_path, or template_path. That file exists to be loaded as "
        "a library: opening it as a database, or writing into it, resets a VBA project "
        "while it is executing. Any such call is refused with "
        "`error_pattern: installed_addin_refused` before Access is touched, whichever "
        "tool it was -- export, import, rebuild, run_vba, run_tests, call_vba, or "
        "anything else. Work on the development copy in the add-in's repository and "
        "rebuild from there; the rebuild is what replaces the installed file. "
        "vcs_get_version_info() reports the installed add-in's version without opening "
        "it. The comparison ignores the file extension, because a compiled install is a "
        "`.accde` built from the same `.accda`.\n\n"
        "**VBA compile failures:**\n"
        "MCP compile tools return success/failure only — not the failing module or line. "
        "When vcs_compile_vba returns success=false (or vcs_check_vba_compiled shows "
        "compiled=false), stop editing source files. Ask the user to compile in the VBE "
        "(Debug → Compile); Access navigates to the error line. Wait for the user to "
        "paste the code snippet around that line before proposing a fix. Do not guess "
        "or iterate through speculative edits.\n\n"
        "**Tool quick-reference (required parameters marked with *, optional with ?):**\n"
        "- vcs_get_version_info() — server, add-in, and Access version info\n"
        "- vcs_list_objects(database_path*) — list all objects by type\n"
        "- vcs_export_database(database_path*, output_dir*, object_types?, full_export?) "
        "— export project (or scoped categories via object_types) to source files\n"
        "- vcs_export_object(database_path*, object_type*, object_name?) "
        "— export a single object/type to source\n"
        "- vcs_import_objects(database_path*, source_dir*, object_types?, full_import?) "
        "— merge project (or scoped categories via object_types) from source; "
        "scoped merges reconcile deletions and take no backup\n"
        "- vcs_import_object(database_path*, object_type*, object_name?) "
        "— import a single object/type from source\n"
        "- vcs_rebuild_database(source_dir*, output_path*, template_path?) "
        "— build fresh database from source\n"
        "- vcs_rebuild_addin(source_dir*, timeout_seconds?) "
        "— rebuild the VCS add-in from source and wait for install\n"
        "- vcs_diff_database(database_path*, source_dir*, show_details?) "
        "— compare database against source files\n"
        "- vcs_run_vba(database_path*, code*, timeout_seconds?) "
        "— execute agent-generated VBA in a temporary module\n"
        "- vcs_call_vba(database_path*, function_name*, args?, timeout_seconds?) "
        "— call an existing public VBA function\n"
        "- vcs_execute_sql(database_path*, sql*, max_rows?) "
        "— run a read-only SELECT query via DAO\n"
        "- vcs_run_tests(database_path*, filter?, timeout_seconds?) — run VBA tests "
        "(prefer `msaccess-vcs run-tests` for live per-test output)\n"
        "- vcs_compile_vba(database_path*, suppress_warnings?) — compile all VBA modules\n"
        "- vcs_check_vba_compiled(database_path*) — check VBA compilation status\n"
        "- vcs_get_option(database_path*, option_name*) — read a VCS option value\n"
        "- vcs_set_option(database_path*, option_name*, value*) "
        "— set a VCS option for this session\n"
        "- vcs_get_log(database_path*, log_type?) — read Export or Build log\n"
        "- vcs_get_recent_calls(limit?) — recent tool_call entries from the usage log\n"
        "- vcs_end_session(database_path*) — end session, remove option overrides\n"
        "- vcs_cancel_operation(operation_id*) — cancel a running async operation\n\n"
        "**Common mistakes to avoid:**\n"
        "- Almost every tool requires database_path* — pass the full .accdb path.\n"
        "- vcs_run_vba() executes arbitrary VBA code you provide in the code* parameter. "
        "Do NOT guess tool names like vcs_eval, vcs_execute_code, or vcs_run_code — "
        "they do not exist.\n"
        "- vcs_run_vba() returns values via a MCP_TempFunction pattern — "
        "read the full tool description for details.\n"
        "- To call a VCS add-in API method (Export, Build, RunTests, "
        "RunRoundtripTests, ...), use vcs_call_vba(db, \"VCS.API\", [\"<Method>\", ...]). "
        "Do NOT call the API from inside vcs_run_vba: that code is itself delivered "
        "through modAPI.API, so calling back into the API is re-entrant and is refused.\n"
        "- Application.Run qualifiers match a loaded VBA *project* name, not a file "
        "name, so \"Version Control.API\" does not resolve. Pass \"VCS.API\" and the "
        "server rewrites it to the add-in's full path.\n\n"
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
    bodies run in a single COM apartment thread via :mod:`access_gate` so
    they no longer block the asyncio event loop. One Access operation runs
    at a time across all Cursor windows sharing this server process.
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
                    "hint": "Check if Access has any open dialogs or message boxes"
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
                result = _addin_json_result(
                    addin.call_sync("ExportByType", types_arg, full_export)
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
                            return _attach_log_context({
                                "success": False,
                                "error": completion.get("error", "Export failed"),
                                "exported_count": 0,
                                "export_path": str(export_path),
                                "objects_by_type": {},
                            }, export_path, "Export", completion)
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
    full_import: bool = False,
    ctx: Context | None = None,
) -> dict[str, Any]:
    """
    Import objects from source files into Access database.
    
    Merges source files back into the database. With no ``object_types``, runs
    a full project merge (``MergeBuild``) with progress callbacks. With
    ``object_types``, runs a category-scoped merge via ``ImportByType`` —
    only those categories are touched.
    
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

            # Category-scoped merge: sync ImportByType (no progress callbacks).
            if object_types:
                types_arg = _scoped_types_arg(object_types)
                if not types_arg:
                    return {
                        "success": False,
                        "error": "object_types was empty after stripping blanks",
                        "imported_count": 0,
                    }
                result = _addin_json_result(
                    addin.call_sync("ImportByType", types_arg, full_import)
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
                        
                        if not completion.get("success"):
                            return _attach_log_context({
                                "success": False,
                                "error": completion.get("error", "Import failed"),
                                "imported_count": 0,
                            }, src_path, "Merge", completion)
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
                            return _attach_log_context({
                                "success": False,
                                "error": result["message"],
                                "imported_count": 0,
                            }, src_path, "Merge")
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.merge_build(str(db_path), str(src_path))
                    if not result["success"]:
                        return _attach_log_context({
                            "success": False,
                            "error": result["message"],
                            "imported_count": 0,
                        }, src_path, "Merge")
            else:
                # Use sync path
                completion = None
                result = addin.merge_build(str(db_path), str(src_path))
                if not result["success"]:
                    return _attach_log_context({
                        "success": False,
                        "error": result["message"],
                        "imported_count": 0,
                    }, src_path, "Merge")
            
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
        Dictionary with build results, plus ``log_path`` for this run's log
        and ``log_excerpt`` (tail of the log) on failure.
    
    The add-in gitignores its ``logs`` folder, so Glob/Grep will not find
    these files. Open ``log_path`` directly, or call vcs_get_log("Build").
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
            close_owned_instances_holding([output_path])
            busy_error = _check_database_busy(output_path)
            if busy_error:
                return busy_error
        
        # We need Access running to call the add-in, but no database is
        # open yet -- the add-in's build process creates it.  Create a
        # bare Access instance (no AccessConnection, which requires a
        # database path) and manage its lifecycle with try/finally.
        app = ensure_dispatch("Access.Application")
        prefer_full_power_if_created(app)
        # The build creates and populates a database in this instance, so any
        # prompt it raises has to be visible to be answerable.
        ensure_access_visible(app)
        
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
                            ctx=ctx,
                            timeout_seconds=timeout_ms / 1000
                        )
                        
                        if not completion.get("success"):
                            return _attach_log_context({
                                "success": False,
                                "error": completion.get("error", "Build failed"),
                                "output_path": None,
                            }, src_path, "Build", completion)
                    elif async_result.get("sync"):
                        # The add-in already ran the build inline; re-running it
                        # here would build twice. Fall through and resolve the
                        # log from disk.
                        op_manager.unregister_operation(operation_id)
                    else:
                        # Neither marker: the add-in never started the operation,
                        # so run it synchronously rather than reporting success
                        # for work that never happened.
                        op_manager.unregister_operation(operation_id)
                        result = addin.build_from_source(str(src_path), output_path)
                        if not result["success"]:
                            return _attach_log_context({
                                "success": False,
                                "error": result["message"],
                                "output_path": None,
                            }, src_path, "Build")
                except Exception as e:
                    # Async call failed - fall back to sync
                    completion = None
                    op_manager.unregister_operation(operation_id)
                    result = addin.build_from_source(str(src_path), output_path)
                    if not result["success"]:
                        return _attach_log_context({
                            "success": False,
                            "error": result["message"],
                            "output_path": None,
                        }, src_path, "Build")
            else:
                # Use sync path
                completion = None
                result = addin.build_from_source(str(src_path), output_path)
                if not result["success"]:
                    return _attach_log_context({
                        "success": False,
                        "error": result["message"],
                        "output_path": None,
                    }, src_path, "Build")
            
            return _attach_log_context({
                "success": True,
                "output_path": output_path,
                "source_dir": str(src_path),
            }, src_path, "Build", completion)
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

        watched = await wait_for_rebuild_status(
            status_file,
            str(phase_started),
            timeout_sec=timeout,
            ctx=ctx,
            reporter=reporter,
            preexisting_access_pids=preexisting_access_pids,
        )
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
    
    Attempts to compile all VBA code in the database. Returns True if compilation
    succeeded (project is compiled), False if compilation failed.
    
    MCP cannot report the failing module or line. When compilation fails:
    
    1. **Stop** — do not import more changes, edit `.bas`/`.cls` files
       speculatively, or run trial-and-error fixes.
    2. **Ask the user** to open the database in Access, open the VBE, and run
       **Debug → Compile** (toolbar compile button). Access highlights the
       first error line.
    3. **Ask the user** to paste the code snippet around that line (±5 lines
       is enough; full error text is optional).
    4. **Then** fix the identified issue in source and re-import / re-compile.
    
    Examples:
        # Compile VBA code
        result = vcs_compile_vba("C:\\\\db.accdb", suppress_warnings=True)
        if result["success"]:
            print("Compilation successful!")
        else:
            print(result["agent_guidance"])
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        suppress_warnings: If True, suppress message boxes during compilation.
                         Warning: If code crashes, warnings may remain disabled.
    
    Returns:
        Dictionary with:
        - success: Boolean - True if compilation succeeded (project is compiled),
                  False if compilation failed
        - agent_guidance: Present when success is False — stop and hand off to
          the user via VBE compile before editing source
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
            
            return _addin_json_result(result_json)
    
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
    made RebuildAddIn unreachable through this tool. Re-acquiring the instance from the
    Running Object Table yields an apartment-local proxy, the approach
    ``VCSAddinIntegration`` already uses for its probe. Marshalling the caller's pointer
    across instead would serialize the call back onto the calling thread and defeat the
    timeout this function exists to impose.

    Without ``db_path`` there is nothing to look up, so the caller's proxy is used as
    before: no worse than it was, and the timeout may simply not fire.
    """
    result_box: dict[str, Any] = {}
    start = time.perf_counter()

    def worker() -> None:
        import pythoncom

        try:
            pythoncom.CoInitialize()
            try:
                worker_app = (
                    VCSAddinIntegration._find_access_in_rot(db_path) if db_path else app
                )
                if worker_app is None:
                    raise RuntimeError(
                        f"Cannot find Access instance for {db_path} from the worker "
                        "thread. The Access application may have been closed."
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


@vcs_tool("vcs_run_vba")
def vcs_run_vba(
    database_path: str,
    code: str,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """
    Execute agent-generated VBA code in a temporary module.
    
    Sends a block of VBA code to the add-in's RunVBA method, which handles the
    full lifecycle: creates a temp module, wraps the code in a function with error
    handling, compiles the project to validate, executes, captures the result,
    removes the temp module, and returns structured JSON.
    
    **Requires McpAllowRunVBA option to be enabled** (default: off).
    The user must enable this manually in the VCS Options form.

    **Reset and temporary-module recovery:**
    Before creating the wrapper, the worker queues a host-project VBA reset
    in a separate API call, runs a harmless COM message-pump barrier, and
    reacquires its Access references. A reset refusal or failure is fail-closed:
    the submitted code does not run.

    The add-in sweeps stale `MCP_Temp_*` standard modules before creating the
    wrapper. Recovery and cleanup fields are actionable:
      - `sweptModules`: stale modules removed before the payload
      - `temp_module_sweep_failed` + `orphanModules`: stale modules remain;
        payload not started
      - `temp_module_unresolvable` + `tempModule`: post-compile canary failed;
        payload not started
      - `temp_module_cleanup_failed` + `cleanupFailed` + `orphanModule`:
        wrapper survived cleanup; a completed return is preserved as
        `payloadResult`

    Stop and tell the user when cleanup fails rather than repeatedly issuing
    calls against a project with an orphaned wrapper.
    
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
    
    **Host-project compile failures:**
    If the response includes `compileError` stating the host project does not
    compile on its own, the failure is in existing database VBA — not your
    submitted `code`. Stop and ask the user to compile in the VBE (Debug →
    Compile) and paste the code snippet around the highlighted line. Do not
    treat `generatedSource` as the fix target unless the response says the
    failure is in the agent code.
    
    Examples:
        vcs_run_vba("C:\\\\db.accdb", "MCP_TempFunction = CurrentDb.TableDefs.Count")
        vcs_run_vba("C:\\\\db.accdb", "Dim qd As DAO.QueryDef\\nSet qd = CurrentDb.QueryDefs(\\"qryTest\\")\\nMCP_TempFunction = qd.SQL")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        code: VBA code to execute (statements, not just an expression)
        timeout_seconds: Optional parent-side timeout. Defaults to
            ACCESS_VCS_RUN_VBA_TIMEOUT_SEC (45 seconds).
    
    Returns:
        Dictionary with `success`, `result`, and on failure `error`,
        `error_pattern`, `errorNumber`, and `errorLine` (the 1-based line
        in `code` that raised the captured error; omitted when not
        available). Reset, sweep, canary, and cleanup failures include the
        actionable fields described above.
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

            # Set the filter option (session-scoped, does not modify user's vcs-options.json)
            addin.call_sync("SetOption", "DefaultTestFilter", filter or "")

            callback_url = get_callback_url()
            op_manager = _get_operation_manager()

            if callback_url and op_manager:
                op_manager.set_event_loop(asyncio.get_running_loop())
                operation_id, _queue = op_manager.register_operation(
                    database_path=str(db_path),
                    command="RunFilteredTests",
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                try:
                    async_result = addin.call_async(callback_info, "RunFilteredTests")
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

            return _parse_test_runner_json(addin.call_sync("RunFilteredTests"))

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
