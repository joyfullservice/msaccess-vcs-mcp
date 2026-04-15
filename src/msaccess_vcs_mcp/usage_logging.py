"""
Usage logging for msaccess-vcs-mcp.

Provides structured JSON logging for tool usage analytics and debugging.
Logs are stored with automatic rotation.

**Log Location Strategy:**
- Development mode: Logs to {project_root}/logs/ within the source directory
- Package mode: Logs to ~/.msaccess-vcs-mcp/logs/

Development mode is detected when the source files are in a directory with
pyproject.toml and src/ structure, indicating an editable install.

This allows developers to:
1. Collect usage logs from any client project using the tool
2. Open the MCP source project in an editor
3. Have agents analyze logs alongside source code to suggest improvements

Configuration via environment variables:
- ACCESS_VCS_ENABLE_LOGGING: Set to "true" to enable logging (default: false)
- ACCESS_VCS_LOG_DIR: Custom log directory (overrides auto-detection)
- ACCESS_VCS_LOG_MAX_SIZE_MB: Max size before rotation (default: 10)
- ACCESS_VCS_LOG_BACKUP_COUNT: Number of rotated files to keep (default: 5)
"""

import functools
import json
import os
import sys
import time
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable

_logging_enabled: bool | None = None
_log_handler: RotatingFileHandler | None = None
_log_file: Path | None = None


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


def _get_default_log_dir() -> Path:
    """
    Get the default log directory.

    - Development mode: {project_root}/logs/
    - Package mode: ~/.msaccess-vcs-mcp/logs/
    """
    project_root = _get_project_root()

    if project_root is not None:
        return project_root / "logs"
    else:
        return Path.home() / ".msaccess-vcs-mcp" / "logs"


def _get_logging_config() -> dict[str, Any]:
    """Load logging configuration from environment variables."""
    return {
        "enabled": os.getenv("ACCESS_VCS_ENABLE_LOGGING", "false").lower() == "true",
        "log_dir": os.getenv("ACCESS_VCS_LOG_DIR", ""),
        "max_size_mb": int(os.getenv("ACCESS_VCS_LOG_MAX_SIZE_MB", "10")),
        "backup_count": int(os.getenv("ACCESS_VCS_LOG_BACKUP_COUNT", "5")),
    }


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

    _log_file = log_dir / "usage.jsonl"
    max_bytes = config["max_size_mb"] * 1024 * 1024

    try:
        _log_handler = RotatingFileHandler(
            filename=str(_log_file),
            maxBytes=max_bytes,
            backupCount=config["backup_count"],
            encoding="utf-8",
        )
        _logging_enabled = True

        _write_log_entry({
            "event": "logging_initialized",
            "log_file": str(_log_file),
            "max_size_mb": config["max_size_mb"],
            "backup_count": config["backup_count"],
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
    """
    global _logging_enabled, _log_handler, _log_file
    if _log_handler is not None:
        try:
            _log_handler.close()
        except Exception:
            pass
    _logging_enabled = None
    _log_handler = None
    _log_file = None


def _write_log_entry(entry: dict[str, Any]) -> None:
    """Write a single log entry to the log file."""
    global _log_handler

    if _log_handler is None:
        return

    if "timestamp" not in entry:
        entry["timestamp"] = datetime.now(timezone.utc).isoformat()
    if "version" not in entry:
        from . import __version__
        entry["version"] = __version__

    try:
        log_line = json.dumps(entry, default=str) + "\n"

        _log_handler.stream.write(log_line)
        _log_handler.stream.flush()

        # Check if rotation is needed by comparing file size directly.
        # We can't use shouldRollover() because it expects a LogRecord.
        if _log_handler.maxBytes > 0:
            _log_handler.stream.seek(0, 2)
            if _log_handler.stream.tell() >= _log_handler.maxBytes:
                _log_handler.doRollover()

    except Exception as e:
        print(f"Warning: Failed to write log entry: {e}", file=sys.stderr)


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


def _sanitize_parameters(params: dict[str, Any], max_string_length: int = 500) -> dict[str, Any]:
    """Sanitize parameters for logging by truncating large values."""
    sanitized = {}
    for key, value in params.items():
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
