"""
Usage logging for msaccess-vcs-mcp.

Provides structured JSON logging for tool usage analytics and debugging.
Logs are stored with automatic rotation.

**Two parallel streams:**

1. **Diagnostic stream** (``vcs-mcp-diagnostic.jsonl``) -- always-on lifecycle
   log: server start, project-root resolution, lazy MCP-roots ``.env``
   discovery, and reset transitions. Independent of
   ``ACCESS_VCS_ENABLE_LOGGING`` so it can be used to debug exactly the
   case where ``.env`` was not found. Lives at
   ``ACCESS_VCS_DIAGNOSTIC_LOG_DIR`` or ``~/.msaccess-vcs-mcp/logs/``.
   Opt out with ``ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true``.

2. **Usage stream** (``vcs-mcp-usage.jsonl``) -- tool calls and
   code-execution events. Gated on ``ACCESS_VCS_ENABLE_LOGGING`` (default
   ``true``).

**Log Location Strategy (usage stream):**
- Development mode: Logs to ``{project_root}/logs/`` within the source dir
- Package mode: Logs to ``~/.msaccess-vcs-mcp/logs/``

Development mode is detected when the source files are in a directory with
pyproject.toml and src/ structure, indicating an editable install.

This allows developers to:
1. Collect usage logs from any client project using the tool
2. Open the MCP source project in an editor
3. Have agents analyze logs alongside source code to suggest improvements

Configuration via environment variables:
- ACCESS_VCS_ENABLE_LOGGING: Set to "true" to enable usage logging (default: true)
- ACCESS_VCS_LOG_DIR: Custom log directory (overrides auto-detection)
- ACCESS_VCS_LOG_MAX_SIZE_MB: Max size before rotation (default: 10)
- ACCESS_VCS_LOG_BACKUP_COUNT: Number of rotated files to keep (default: 5)
- ACCESS_VCS_LOG_CODE_CONTENT: Set to "true" to record full SQL/VBA bodies in
  ``code_execution`` events (default: false -- only ``code_length`` is logged).
- ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG: Set to "true" to disable the always-on
  diagnostic stream (default: false).
- ACCESS_VCS_DIAGNOSTIC_LOG_DIR: Custom diagnostic log directory (default:
  ``~/.msaccess-vcs-mcp/logs/``).
"""

import functools
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable

def _strip_path_quotes(value: str) -> str:
    """Strip matching outer quotes from an env-var value (see config._strip_quotes)."""
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
        return value[1:-1]
    return value


_logging_enabled: bool | None = None
_log_handler: RotatingFileHandler | None = None
_log_file: Path | None = None

_diagnostic_handler: RotatingFileHandler | None = None
_diagnostic_file: Path | None = None
_diagnostic_initialized: bool = False
_diagnostic_enabled: bool = False
_diagnostic_disabled_reason: str | None = None


def _is_development_install() -> bool:
    """
    Check if we're running from a development (editable) install.

    Returns True if the source files are in a directory structure that
    indicates a development environment (has pyproject.toml, src/, etc.).
    """
    try:
        this_file = Path(__file__).resolve()
        package_dir = this_file.parent       # msaccess_vcs_mcp/
        src_dir = package_dir.parent          # src/
        project_root = src_dir.parent         # project root

        has_pyproject = (project_root / "pyproject.toml").exists()
        has_src_structure = src_dir.name == "src" and package_dir.name == "msaccess_vcs_mcp"
        has_tests = (project_root / "tests").exists()

        return has_pyproject and has_src_structure and has_tests
    except Exception:
        return False


def _get_project_root() -> Path | None:
    """Get the project root directory if running from development install."""
    try:
        this_file = Path(__file__).resolve()
        package_dir = this_file.parent
        src_dir = package_dir.parent
        project_root = src_dir.parent

        if _is_development_install():
            return project_root
        return None
    except Exception:
        return None


def _get_workspace_project_root() -> Path | None:
    """Return the user's workspace project root if discovered by config.

    The config module resolves this via ``ACCESS_VCS_PROJECT_DIR``,
    upward CWD walk, or the lazy MCP ``roots/list`` handshake (see
    ``tools._ensure_env_loaded``). Imported lazily to avoid a circular
    import: ``config`` imports ``usage_logging`` for diagnostic events.

    Returns ``None`` when:
      * Discovery hasn't happened yet.
      * Discovery hit the no-marker CWD fallback (``RESOLUTION_CWD_FALLBACK``).
      * Discovery resolved to the user's *home* directory -- this
        typically happens when Cursor launches the server from
        ``~`` and the upward walk finds ``~/.cursor/mcp.json``
        (the user-level Cursor config, not a project-level one).
        Writing logs into ``~/logs/`` would litter the user's home
        directory, so we treat that as "no workspace discovered" and
        let the caller fall back to ``~/.msaccess-vcs-mcp/logs/``.

    Callers fall back to the development-install root or the user-home
    log dir in those cases.
    """
    try:
        from . import config as _config
    except Exception:
        return None
    root = getattr(_config, "_project_root", None)
    method = getattr(_config, "_project_root_method", None)
    if root is None:
        return None
    if method == _config.RESOLUTION_CWD_FALLBACK:
        return None
    try:
        if Path(root).resolve() == Path.home().resolve():
            return None
    except OSError:
        return None
    return root


def _get_default_log_dir() -> Path:
    """
    Get the default log directory.

    Resolution order:
      1. **User workspace** -- if config has discovered a real project
         root (via ``ACCESS_VCS_PROJECT_DIR``, CWD walk, or MCP
         ``roots/list``), log into ``{workspace}/logs/``. This puts
         tool-call audit data alongside the project the user is
         actually working on, not the MCP server's own repo.
      2. **MCP server development install** -- if running from a
         cloned ``msaccess-vcs-mcp`` checkout, log into that repo's
         ``logs/`` folder. Useful for MCP development.
      3. **User home fallback** -- ``~/.msaccess-vcs-mcp/logs/``.

    The workspace path may not be available on the very first
    ``startup_env_load`` (before the lazy MCP roots handshake fires).
    The first lazy-init triggers ``reset_logging()`` so the next call
    picks up the workspace-aware path.
    """
    workspace_root = _get_workspace_project_root()
    if workspace_root is not None:
        return workspace_root / "logs"

    dev_root = _get_project_root()
    if dev_root is not None:
        return dev_root / "logs"

    return Path.home() / ".msaccess-vcs-mcp" / "logs"


def _get_logging_config() -> dict[str, Any]:
    """Load logging configuration from environment variables."""
    return {
        "enabled": os.getenv("ACCESS_VCS_ENABLE_LOGGING", "true").lower() == "true",
        "log_dir": _strip_path_quotes(os.getenv("ACCESS_VCS_LOG_DIR", "")),
        "max_size_mb": int(os.getenv("ACCESS_VCS_LOG_MAX_SIZE_MB", "10")),
        "backup_count": int(os.getenv("ACCESS_VCS_LOG_BACKUP_COUNT", "5")),
        "log_code_content": os.getenv("ACCESS_VCS_LOG_CODE_CONTENT", "false").lower() == "true",
    }


def _get_diagnostic_log_dir() -> Path:
    """
    Get the diagnostic log directory.

    Diagnostics always default to the per-user package location so that a
    diagnostic record exists even when the project ``.env`` cannot be
    discovered (which is the exact scenario the diagnostic stream is
    designed to debug). Override with ``ACCESS_VCS_DIAGNOSTIC_LOG_DIR``.
    """
    override = _strip_path_quotes(os.getenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", "")).strip()
    if override:
        return Path(override)
    return Path.home() / ".msaccess-vcs-mcp" / "logs"


def _ensure_log_dir(log_dir: Path) -> bool:
    """
    Ensure the log directory exists.

    Returns True if directory exists or was created, False if creation failed.
    """
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        print(f"Warning: Could not create log directory {log_dir}: {e}", file=sys.stderr)
        return False


def _initialize_logging() -> bool:
    """
    Initialize the logging system.

    Returns True if logging is enabled and initialized, False otherwise.
    """
    global _logging_enabled, _log_handler, _log_file

    if _logging_enabled is True:
        return True

    config = _get_logging_config()

    if not config["enabled"]:
        # Don't cache False: the .env file may not have been loaded yet.
        # Leaving _logging_enabled as None lets us re-check on the next call
        # once load_config() has populated the environment.
        return False

    if config["log_dir"]:
        log_dir = Path(config["log_dir"])
    else:
        log_dir = _get_default_log_dir()

    if not _ensure_log_dir(log_dir):
        _logging_enabled = False
        return False

    _log_file = log_dir / "vcs-mcp-usage.jsonl"
    max_bytes = config["max_size_mb"] * 1024 * 1024

    try:
        _log_handler = RotatingFileHandler(
            filename=str(_log_file),
            maxBytes=max_bytes,
            backupCount=config["backup_count"],
            encoding="utf-8",
        )
        _logging_enabled = True

        # Include project-root resolution provenance so users can audit
        # *how* the .env was discovered for this session (matters most
        # when the server is launched at the user level and several
        # discovery paths are possible).
        try:
            from .config import get_project_root_info
            root_info = get_project_root_info()
        except Exception:
            root_info = {"project_root": None, "resolution_method": None}

        _write_log_entry({
            "event": "logging_initialized",
            "log_file": str(_log_file),
            "max_size_mb": config["max_size_mb"],
            "backup_count": config["backup_count"],
            "project_root": root_info.get("project_root"),
            "project_root_resolution": root_info.get("resolution_method"),
        })

        return True

    except Exception as e:
        print(f"Warning: Could not initialize logging: {e}", file=sys.stderr)
        _logging_enabled = False
        return False


def reset_logging() -> None:
    """Clear logging state so it re-initializes from fresh env vars.

    Called by :func:`config.load_config` when configuration changes,
    ensuring that logging configuration changes (enable/disable,
    log directory, rotation settings) take effect without a server restart.

    Resets *both* the usage and diagnostic streams so changes to
    ``ACCESS_VCS_DIAGNOSTIC_LOG_DIR`` / ``ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG``
    also take effect after a config reload.
    """
    global _logging_enabled, _log_handler, _log_file
    global _diagnostic_handler, _diagnostic_file
    global _diagnostic_initialized, _diagnostic_enabled, _diagnostic_disabled_reason

    if _log_handler is not None:
        try:
            _log_handler.close()
        except Exception:
            pass
    _logging_enabled = None
    _log_handler = None
    _log_file = None

    if _diagnostic_handler is not None:
        try:
            _diagnostic_handler.close()
        except Exception:
            pass
    _diagnostic_handler = None
    _diagnostic_file = None
    _diagnostic_initialized = False
    _diagnostic_enabled = False
    _diagnostic_disabled_reason = None


def _write_log_entry(
    entry: dict[str, Any],
    handler: RotatingFileHandler | None = None,
) -> None:
    """Write a single JSONL entry to a rotating log handler.

    Defaults to the usage-log handler. Pass ``handler`` to target the
    diagnostic stream (or any other rotating handler).
    """
    if handler is None:
        handler = _log_handler

    if handler is None:
        return

    if "timestamp" not in entry:
        entry["timestamp"] = datetime.now(timezone.utc).isoformat()
    if "version" not in entry:
        from . import __version__
        entry["version"] = __version__

    try:
        log_line = json.dumps(entry, default=str) + "\n"

        handler.stream.write(log_line)
        handler.stream.flush()

        # Check if rotation is needed by comparing file size directly.
        # We can't use shouldRollover() because it expects a LogRecord.
        if handler.maxBytes > 0:
            handler.stream.seek(0, 2)
            if handler.stream.tell() >= handler.maxBytes:
                handler.doRollover()

    except Exception as e:
        print(f"Warning: Failed to write log entry: {e}", file=sys.stderr)


def _initialize_diagnostic_logging() -> bool:
    """
    Initialize the always-on diagnostic logging stream.

    The diagnostic stream captures server lifecycle events (startup,
    project-root resolution, ``.env`` loading, lazy MCP-roots init) so
    that an operator can answer "why didn't logging work?" without
    relying on the very ``.env`` discovery that may have failed.

    Returns ``True`` if the handler is open and ready, ``False`` if the
    stream has been opted out or the handler could not be created.
    Idempotent: subsequent calls return the cached result.
    """
    global _diagnostic_handler, _diagnostic_file
    global _diagnostic_initialized, _diagnostic_enabled
    global _diagnostic_disabled_reason

    if _diagnostic_initialized:
        return _diagnostic_enabled

    if os.getenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "false").lower() == "true":
        _diagnostic_initialized = True
        _diagnostic_enabled = False
        _diagnostic_disabled_reason = "opt_out_env_var"
        return False

    log_dir = _get_diagnostic_log_dir()

    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(
            f"Warning: Could not create diagnostic log directory {log_dir}: {e}",
            file=sys.stderr,
        )
        _diagnostic_initialized = True
        _diagnostic_enabled = False
        _diagnostic_disabled_reason = f"mkdir_failed: {e}"
        return False

    _diagnostic_file = log_dir / "vcs-mcp-diagnostic.jsonl"
    # Diagnostic log: smaller cap, fewer backups -- it's lifecycle-only.
    max_bytes = 1 * 1024 * 1024  # 1 MB
    backup_count = 3

    try:
        _diagnostic_handler = RotatingFileHandler(
            filename=str(_diagnostic_file),
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding="utf-8",
        )
    except Exception as e:
        print(
            f"Warning: Could not initialize diagnostic logging at {_diagnostic_file}: {e}",
            file=sys.stderr,
        )
        _diagnostic_initialized = True
        _diagnostic_enabled = False
        _diagnostic_disabled_reason = f"handler_failed: {e}"
        return False

    _diagnostic_initialized = True
    _diagnostic_enabled = True
    _diagnostic_disabled_reason = None
    return True


def log_diagnostic_event(event: str, **fields: Any) -> None:
    """
    Append a single lifecycle event to ``vcs-mcp-diagnostic.jsonl``.

    Always-on by design: this is the stream of last resort for
    diagnosing why the *usage* stream might be silent. No-ops cleanly
    if diagnostic logging has been opted out or could not be opened.

    Args:
        event: Short event identifier (e.g. ``"server_start"``,
            ``"lazy_init_loaded"``).
        **fields: Arbitrary structured fields to record alongside the
            event. Values must be JSON-serializable (``json.dumps``
            falls back to ``default=str``).
    """
    if not _initialize_diagnostic_logging():
        return

    entry: dict[str, Any] = {"event": event, **fields}
    _write_log_entry(entry, handler=_diagnostic_handler)


def get_diagnostic_log_path() -> Path | None:
    """Return the diagnostic log file path, or ``None`` if disabled."""
    if _initialize_diagnostic_logging():
        return _diagnostic_file
    return None


def is_diagnostic_logging_enabled() -> bool:
    """Return whether the diagnostic stream is open and writing."""
    return _initialize_diagnostic_logging()


def log_tool_call(
    tool_name: str,
    parameters: dict[str, Any],
    result: dict[str, Any] | None = None,
    error: str | None = None,
    execution_time_ms: float | None = None,
) -> None:
    """
    Log a tool call with its parameters and result.

    Args:
        tool_name: Name of the tool called
        parameters: Input parameters (will be truncated if too large)
        result: Result dictionary (success case)
        error: Error message (failure case)
        execution_time_ms: Execution time in milliseconds
    """
    if not _initialize_logging():
        return

    sanitized_params = _sanitize_parameters(parameters)

    entry: dict[str, Any] = {
        "event": "tool_call",
        "tool": tool_name,
        "parameters": sanitized_params,
        "success": error is None,
        "execution_time_ms": execution_time_ms,
    }

    if error:
        entry["error"] = _truncate_string(error, max_length=500)
        entry["error_pattern"] = _extract_error_pattern(error)

    if result and isinstance(result, dict) and "error" in result:
        entry["success"] = False
        entry["error"] = _truncate_string(str(result.get("error", "")), max_length=500)
        entry["error_pattern"] = _extract_error_pattern(str(result.get("error", "")))

    _write_log_entry(entry)


# Defense-in-depth: parameter keys that *look* like credentials are
# replaced with ``"<redacted>"`` before they are written to disk,
# regardless of any other logging switch. This catches the common case
# of an agent or downstream tool inventing a parameter name like
# ``password=`` or ``api_key=``.
_SECRET_KEY_PATTERN = re.compile(
    r"(password|secret|token|api[_-]?key|connection[_-]?string)",
    re.IGNORECASE,
)

# Parameter keys that carry executable code (SQL, VBA).  When
# ``ACCESS_VCS_LOG_CODE_CONTENT`` is false (the default), these are
# replaced with ``"<code_length:N>"`` so the audit trail records
# *that* code was passed and how large it was, without persisting
# the body.  This closes the gap where ``code_execution`` events
# correctly honour the switch but ``tool_call`` events would leak
# the code through the parameter dict.
_CODE_KEY_PATTERN = re.compile(r"^(code|sql)$", re.IGNORECASE)


def _sanitize_parameters(params: dict[str, Any], max_string_length: int = 500) -> dict[str, Any]:
    """Sanitize parameters for logging.

    - Truncates oversized strings.
    - Recursively sanitizes nested dicts.
    - Caps lists at 10 items.
    - Replaces values whose keys match :data:`_SECRET_KEY_PATTERN` with
      ``"<redacted>"`` so credential-shaped fields never reach disk.
    - Replaces values whose keys match :data:`_CODE_KEY_PATTERN` with
      a length stub when ``ACCESS_VCS_LOG_CODE_CONTENT`` is disabled,
      matching the behaviour of ``code_execution`` events.
    """
    log_code = _get_logging_config()["log_code_content"]
    sanitized: dict[str, Any] = {}
    for key, value in params.items():
        if isinstance(key, str) and _SECRET_KEY_PATTERN.search(key):
            sanitized[key] = "<redacted>"
            continue
        if (
            not log_code
            and isinstance(key, str)
            and _CODE_KEY_PATTERN.search(key)
            and isinstance(value, str)
        ):
            sanitized[key] = f"<code_length:{len(value)}>"
            continue
        if isinstance(value, str):
            sanitized[key] = _truncate_string(value, max_string_length)
        elif isinstance(value, dict):
            sanitized[key] = _sanitize_parameters(value, max_string_length)
        elif isinstance(value, list):
            sanitized[key] = [
                _truncate_string(v, max_string_length) if isinstance(v, str) else v
                for v in value[:10]
            ]
            if len(value) > 10:
                sanitized[key].append(f"... ({len(value) - 10} more items)")
        else:
            sanitized[key] = value
    return sanitized


def _truncate_string(s: str, max_length: int = 500) -> str:
    """Truncate a string if it exceeds max_length."""
    if len(s) <= max_length:
        return s
    return s[:max_length] + f"... (truncated, {len(s)} chars total)"


def _extract_error_pattern(error: str) -> str:
    """
    Extract a normalized error pattern for categorization.

    Helps identify common error types for analysis.
    """
    error_lower = error.lower()

    # COM / Access errors
    if "com_error" in error_lower or "pywintypes.com_error" in error_lower:
        return "com_error"

    if "prevents it from being opened or locked" in error_lower:
        return "database_exclusive_lock"

    if "file already in use" in error_lower:
        return "file_in_use"

    if "not found" in error_lower:
        if "database" in error_lower or "file" in error_lower:
            return "file_not_found"
        if "object" in error_lower or "module" in error_lower:
            return "object_not_found"
        return "not_found_other"

    if "permission" in error_lower or "access denied" in error_lower:
        return "permission_denied"

    if "write" in error_lower and "disabled" in error_lower:
        return "write_disabled"

    if "timeout" in error_lower or "timed out" in error_lower:
        return "timeout"

    if "callback" in error_lower:
        return "callback_error"

    if "compile" in error_lower or "syntax" in error_lower:
        return "vba_compile_error"

    if "addin" in error_lower or "add-in" in error_lower:
        return "addin_error"

    if "cancelled" in error_lower or "canceled" in error_lower:
        return "operation_cancelled"

    if "busy" in error_lower:
        return "database_busy"

    if "serializ" in error_lower or "not json serializable" in error_lower:
        return "serialization_error"

    if "encoding" in error_lower or "utf" in error_lower or "unicode" in error_lower:
        return "encoding_error"

    if "error" in error_lower:
        return "generic_error"

    return "unknown"


def with_logging(tool_name: str):
    """
    Decorator to add usage logging to a tool function.

    Supports both sync and async tool functions.

    Usage::

        @with_logging("vcs_list_objects")
        def vcs_list_objects(database_path: str) -> dict:
            ...

        @with_logging("vcs_export_database")
        async def vcs_export_database(database_path: str, ...) -> dict:
            ...

    Args:
        tool_name: Name of the tool for logging purposes
    """
    def _log_call(func, args, kwargs, result, error_msg, serialization_warning, start_time):
        """Shared logging logic for sync and async wrappers."""
        execution_time_ms = (time.time() - start_time) * 1000

        import inspect
        parameters = dict(kwargs)
        sig = inspect.signature(func)
        param_names = list(sig.parameters.keys())
        for i, arg in enumerate(args):
            if i < len(param_names):
                parameters[param_names[i]] = arg

        # Filter out Context objects (not serializable)
        parameters = {
            k: v for k, v in parameters.items()
            if not (hasattr(v, '__class__') and v.__class__.__name__ == 'Context')
        }

        logged_error = error_msg or serialization_warning
        log_tool_call(
            tool_name=tool_name,
            parameters=parameters,
            result=result,
            error=logged_error,
            execution_time_ms=round(execution_time_ms, 2),
        )

    def _check_serialization(result):
        """Validate result is JSON-serializable, return (result, warning)."""
        if result is None:
            return result, None
        try:
            json.dumps(result, default=str)
            return result, None
        except (TypeError, ValueError, UnicodeEncodeError) as ser_err:
            warning = (
                f"Result passed tool execution but failed JSON serialization: "
                f"{type(ser_err).__name__}: {ser_err}"
            )
            error_result = {
                "error": (
                    f"Tool succeeded but the result contains values that "
                    f"cannot be serialized to JSON ({type(ser_err).__name__}: {ser_err})."
                )
            }
            return error_result, warning

    def decorator(func: Callable) -> Callable:
        import inspect as _inspect

        if _inspect.iscoroutinefunction(func):
            @functools.wraps(func)
            async def async_wrapper(*args, **kwargs) -> Any:
                if not _initialize_logging():
                    return await func(*args, **kwargs)

                start_time = time.time()
                error_msg = None
                result = None
                serialization_warning = None
                try:
                    result = await func(*args, **kwargs)
                    result, serialization_warning = _check_serialization(result)
                    return result
                except Exception as e:
                    error_msg = str(e)
                    raise
                finally:
                    _log_call(func, args, kwargs, result, error_msg, serialization_warning, start_time)

            return async_wrapper
        else:
            @functools.wraps(func)
            def wrapper(*args, **kwargs) -> Any:
                if not _initialize_logging():
                    return func(*args, **kwargs)

                start_time = time.time()
                error_msg = None
                result = None
                serialization_warning = None
                try:
                    result = func(*args, **kwargs)
                    result, serialization_warning = _check_serialization(result)
                    return result
                except Exception as e:
                    error_msg = str(e)
                    raise
                finally:
                    _log_call(func, args, kwargs, result, error_msg, serialization_warning, start_time)

            return wrapper
    return decorator


def log_code_execution(
    tool_name: str,
    database_path: str,
    code: str,
    code_type: str = "sql",
) -> None:
    """
    Log a code execution attempt *before* it runs.

    Written as a separate ``"code_execution"`` event so it can be found
    quickly in the audit trail. ``code_length`` is always recorded so an
    analyst can spot anomalies (e.g. "an agent ran a 4 KB VBA block")
    without seeing the body itself.

    The full ``code`` body is recorded **only** when
    ``ACCESS_VCS_LOG_CODE_CONTENT=true``. By default it is omitted so
    SQL/VBA fragments that may contain business data, table names, or
    other sensitive context are not persisted to disk. Operators who
    need a complete forensic record can opt in.

    Args:
        tool_name: MCP tool name (e.g. ``"vcs_execute_sql"``).
        database_path: Target database path.
        code: The SQL statement, VBA code block, or function call string.
        code_type: ``"sql"``, ``"vba"``, or ``"vba_call"``.
    """
    if not _initialize_logging():
        return

    config = _get_logging_config()
    entry: dict[str, Any] = {
        "event": "code_execution",
        "tool": tool_name,
        "database": database_path,
        "code_type": code_type,
        "code_length": len(code),
    }
    if config["log_code_content"]:
        entry["code"] = code
    _write_log_entry(entry)


def log_addin_probe(
    addin_path: str,
    duration_ms: float,
    success: bool,
    timed_out: bool = False,
    error: str | None = None,
) -> None:
    """
    Log a single VCS add-in ``GetVCSVersion`` probe.

    Each call to :meth:`VCSAddinIntegration.load_addin` performs a probe to
    verify the add-in is healthy before any real work is dispatched.  This
    helper records the probe's outcome -- duration, success, and whether
    the hard timeout fired -- so the cumulative cost of probing every tool
    call can be audited from the same ``vcs-mcp-usage.jsonl`` stream as
    ``tool_call`` and ``code_execution`` events.

    Args:
        addin_path: Path to the ``.accda`` add-in being probed.
        duration_ms: Wall-clock duration of the probe in milliseconds.
        success: True if the probe completed without raising.
        timed_out: True iff the hard timeout (``ACCESS_VCS_PROBE_TIMEOUT_SEC``)
            fired -- distinguishes true hangs from generic COM errors.
        error: Error message when ``success`` is False; otherwise ``None``.
    """
    if not _initialize_logging():
        return

    entry: dict[str, Any] = {
        "event": "addin_probe",
        "addin_path": addin_path,
        "duration_ms": duration_ms,
        "success": success,
        "timed_out": timed_out,
    }

    if error:
        entry["error"] = _truncate_string(error, max_length=500)
        entry["error_pattern"] = _extract_error_pattern(error)

    _write_log_entry(entry)


def get_log_file_path() -> Path | None:
    """
    Get the current log file path.

    Returns:
        Path to log file if logging is enabled, None otherwise.
    """
    if _initialize_logging():
        return _log_file
    return None


def is_logging_enabled() -> bool:
    """Check if logging is currently enabled."""
    return _initialize_logging()
