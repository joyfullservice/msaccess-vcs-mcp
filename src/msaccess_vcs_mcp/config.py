"""Configuration management for msaccess-vcs-mcp."""

import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


# Resolution-method labels emitted by ``_find_project_root`` and
# ``initialize_from_workspace``. Surfaced via ``get_project_root_info()``
# so usage logging and stderr diagnostics can show *how* the project root
# was discovered.
RESOLUTION_PROJECT_DIR_ENV = "ACCESS_VCS_PROJECT_DIR"
RESOLUTION_WORKSPACE_ROOTS = "mcp_workspace_roots"
RESOLUTION_CWD_ENV = "cwd_walk_env"
RESOLUTION_CWD_MARKER = "cwd_walk_project_marker"
RESOLUTION_PACKAGE_ENV = "package_walk_env"
RESOLUTION_PACKAGE_MARKER = "package_walk_project_marker"
RESOLUTION_CWD_FALLBACK = "cwd_fallback_no_env_found"


def _find_project_root() -> Path:
    """
    Find the project root by searching for .env file or common project markers.

    Resolution order:
    1. ``ACCESS_VCS_PROJECT_DIR`` env-var (explicit override)
    2. Upward search from the current working directory
    3. Upward search from the installed package location
    4. Falls back to the current working directory

    When Cursor (or another IDE) launches the MCP server, it normally sets
    the working directory to the open workspace root -- so the automatic
    search finds the project's ``.env`` file even when the server is
    configured at the *user* level. ``ACCESS_VCS_PROJECT_DIR`` is available
    as a fallback for environments where the working directory is not set
    to the project root. ``initialize_from_workspace()`` provides a third
    path: lazy discovery via the MCP ``roots/list`` protocol call.

    For each starting point the function walks up the directory tree
    looking for:
    * ``.env`` -- primary indicator
    * ``.cursor/mcp.json`` -- MCP configuration
    * ``pyproject.toml`` -- Python project marker

    Side effect: sets module-level ``_project_root_method`` so callers
    (and usage logging) can see *how* the project root was resolved.

    Returns:
        Path to project root, or current working directory if not found
    """
    global _project_root_method
    import sys

    # 1. Explicit override via ACCESS_VCS_PROJECT_DIR
    explicit_dir = os.getenv("ACCESS_VCS_PROJECT_DIR")
    if explicit_dir:
        explicit_path = Path(explicit_dir).resolve()
        if explicit_path.is_dir():
            _project_root_method = RESOLUTION_PROJECT_DIR_ENV
            return explicit_path
        else:
            print(
                f"Warning: ACCESS_VCS_PROJECT_DIR points to a non-existent directory: {explicit_dir}",
                file=sys.stderr,
            )

    # 2-3. Automatic search from CWD and package location
    cwd_root = Path.cwd().resolve()
    search_roots: list[tuple[Path, str, str]] = [
        (cwd_root, RESOLUTION_CWD_ENV, RESOLUTION_CWD_MARKER),
    ]

    try:
        # Walk up to find project root (should be parent of src/)
        package_dir = Path(__file__).parent.parent.parent.resolve()
        if (package_dir.parent / "pyproject.toml").exists():
            search_roots.append(
                (package_dir.parent, RESOLUTION_PACKAGE_ENV, RESOLUTION_PACKAGE_MARKER),
            )
        search_roots.append(
            (package_dir.parent, RESOLUTION_PACKAGE_ENV, RESOLUTION_PACKAGE_MARKER),
        )
    except Exception:
        pass

    for root, env_method, marker_method in search_roots:
        current = root
        while current != current.parent:
            if (current / ".env").exists():
                _project_root_method = env_method
                return current
            if (current / ".cursor" / "mcp.json").exists():
                _project_root_method = marker_method
                return current
            if (current / "pyproject.toml").exists():
                _project_root_method = marker_method
                return current
            current = current.parent

    # 4. Fall back to current working directory
    _project_root_method = RESOLUTION_CWD_FALLBACK
    return Path.cwd().resolve()


_env_loaded = False
_is_reload = False
_project_root: Path | None = None
_project_root_method: str | None = None
_env_mtimes: dict[str, float] = {}  # path -> mtime at last load


def _get_project_root() -> Path:
    """Return the stored project root, falling back to discovery."""
    if _project_root is not None:
        return _project_root
    return _find_project_root()


def _record_env_mtimes() -> None:
    """Store the modification times of loaded .env files for hot-reload detection."""
    global _env_mtimes
    _env_mtimes.clear()
    root = _get_project_root()
    for name in (".env", ".env.local"):
        path = root / name
        if path.exists():
            try:
                _env_mtimes[str(path)] = path.stat().st_mtime
            except OSError:
                pass


def _check_env_reload() -> bool:
    """Compare current .env mtimes against stored values, trigger reload if changed.

    Returns True if a reload was triggered (``_env_loaded`` reset to False).
    """
    global _env_loaded, _is_reload
    if not _env_loaded or not _env_mtimes:
        return False

    for path_str, old_mtime in _env_mtimes.items():
        try:
            current_mtime = Path(path_str).stat().st_mtime
            if current_mtime != old_mtime:
                _env_loaded = False
                _is_reload = True
                _env_mtimes.clear()
                return True
        except OSError:
            pass

    return False


def _load_env_files() -> None:
    """
    Load .env files from the project root.

    Searches for project root by looking for .env file or project markers,
    then loads:
    1. .env (base configuration)
    2. .env.local (local overrides, if exists)

    On first load, environment variables already set (e.g., from MCP
    server env section) take precedence. On reload (triggered by
    ``_check_env_reload`` detecting an mtime change), ``override=True``
    is used so edited values replace old ones.

    Prints diagnostic messages to stderr so users can verify which files
    were loaded (visible in Cursor's MCP server output pane).
    """
    global _env_loaded, _is_reload, _project_root
    if _env_loaded:
        return
    _env_loaded = True

    import sys

    reloading = _is_reload
    _is_reload = False

    project_root = _find_project_root()
    _project_root = project_root
    cwd = Path.cwd().resolve()

    if not reloading:
        print(f"Working directory: {cwd}", file=sys.stderr)
        if os.getenv("ACCESS_VCS_PROJECT_DIR"):
            print(
                f"ACCESS_VCS_PROJECT_DIR: {os.getenv('ACCESS_VCS_PROJECT_DIR')}",
                file=sys.stderr,
            )
        print(
            f"Resolved project root: {project_root} "
            f"(via {_project_root_method or 'unknown'})",
            file=sys.stderr,
        )

    # On reload, use override=True so edited values replace old ones.
    # On first load, override=False preserves MCP env-section precedence.
    env_override = reloading

    env_path = project_root / ".env"
    if env_path.exists():
        result = load_dotenv(str(env_path), override=env_override)
        if result:
            label = "Reloaded" if reloading else "Loaded"
            print(f"{label} .env from {env_path}", file=sys.stderr)
        elif not reloading:
            print(
                f"Warning: .env file exists at {env_path} but no new variables were loaded "
                f"(they may already be set via mcp.json env section)",
                file=sys.stderr,
            )
    elif not reloading:
        print(
            f"No .env file found at {env_path} -- "
            f"if this is unexpected, set ACCESS_VCS_PROJECT_DIR to your project path",
            file=sys.stderr,
        )

    env_local_path = project_root / ".env.local"
    if env_local_path.exists():
        load_dotenv(str(env_local_path), override=True)
        label = "Reloaded" if reloading else "Loaded"
        print(f"{label} .env.local from {env_local_path}", file=sys.stderr)

    _record_env_mtimes()


def _load_env_from_directory(directory: Path) -> bool:
    """Load .env and .env.local from an explicit directory.

    Used by the MCP workspace-roots lazy-init path when the working
    directory is not the project root.

    Returns True if at least one file was loaded.
    """
    global _project_root
    import sys

    _project_root = directory.resolve()
    loaded = False
    env_path = directory / ".env"
    if env_path.exists():
        result = load_dotenv(str(env_path), override=False)
        if result:
            print(f"Loaded .env from {env_path}", file=sys.stderr)
            loaded = True
        else:
            print(
                f"Warning: .env file exists at {env_path} but no new variables were loaded "
                f"(they may already be set via mcp.json env section)",
                file=sys.stderr,
            )

    env_local_path = directory / ".env.local"
    if env_local_path.exists():
        load_dotenv(str(env_local_path), override=True)
        print(f"Loaded .env.local from {env_local_path}", file=sys.stderr)
        loaded = True

    _record_env_mtimes()
    return loaded


def initialize_from_workspace(workspace_path: Path) -> dict[str, Any]:
    """Initialize configuration from a workspace directory discovered via MCP roots.

    This is the *lazy-init* path used when the server starts without a
    discoverable ``.env`` file (e.g. user-level MCP configuration where
    the working directory is not the project root). The MCP client
    reports its workspace roots via ``roots/list``; the first root that
    contains an ``.env`` file is loaded here.

    Args:
        workspace_path: Absolute path to the workspace root provided by
            the MCP client (Cursor / VS Code).

    Returns:
        The freshly loaded configuration dictionary.
    """
    global _env_loaded, _project_root, _project_root_method
    import sys

    workspace_path = workspace_path.resolve()
    _project_root = workspace_path
    _project_root_method = RESOLUTION_WORKSPACE_ROOTS
    print(
        f"Lazy init: loading .env from workspace root {workspace_path} "
        f"(via {RESOLUTION_WORKSPACE_ROOTS})",
        file=sys.stderr,
    )

    _load_env_from_directory(workspace_path)
    # Mark env as loaded so subsequent load_config() calls don't re-run
    # discovery (which would walk back up from CWD and overwrite our
    # workspace_path resolution).
    _env_loaded = True
    _record_env_mtimes()

    # Reset logging so changes to ACCESS_VCS_ENABLE_LOGGING etc. take
    # effect using the workspace's settings.
    from .usage_logging import reset_logging
    reset_logging()

    return load_config()


def get_project_root_info() -> dict[str, Any]:
    """Return diagnostic info about how the project root was resolved.

    Used by usage logging to record resolution provenance in the
    ``logging_initialized`` event so users can audit which discovery
    mechanism was actually used in a given session.

    Returns:
        Dict with keys:
        - ``project_root``: resolved project-root path as a string, or
          ``None`` if discovery hasn't happened yet.
        - ``resolution_method``: one of the ``RESOLUTION_*`` constants
          in this module, or ``None`` if discovery hasn't happened yet.
    """
    return {
        "project_root": str(_project_root) if _project_root is not None else None,
        "resolution_method": _project_root_method,
    }


def get_default_addin_path() -> str:
    """
    Get default VCS add-in installation path.
    
    Returns:
        Path to add-in file at default installation location
    """
    # Default installation location: %AppData%\MSAccessVCS\Version Control.accda
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, "MSAccessVCS", "Version Control.accda")


def load_config() -> dict[str, Any]:
    """
    Load configuration from environment variables.

    Automatically loads .env and .env.local files from the project root.
    Environment variables passed via MCP server env section take precedence.

    On subsequent calls the .env file modification time is checked; if
    the file was edited since last load the environment is refreshed
    automatically (hot-reload), and logging is re-initialized so the
    new settings take effect without restarting the server.

    Returns:
        Dictionary with configuration values
    """
    # Check for .env file changes and trigger reload if needed.
    reloaded = _check_env_reload()

    # Load .env files (first load or reload).
    _load_env_files()

    # Parse callback enabled setting (default: true)
    callback_enabled_str = os.getenv("ACCESS_VCS_CALLBACK_ENABLED", "true").lower()
    callback_enabled = callback_enabled_str not in ("false", "0", "no", "off")

    config = {
        # Database settings
        "ACCESS_VCS_DATABASE": os.getenv("ACCESS_VCS_DATABASE", ""),
        "ACCESS_VCS_ADDIN_PATH": os.getenv("ACCESS_VCS_ADDIN_PATH", get_default_addin_path()),
        "ACCESS_VCS_DISABLE_WRITES": os.getenv("ACCESS_VCS_DISABLE_WRITES", "false").lower() == "true",

        # Callback server settings
        "ACCESS_VCS_CALLBACK_ENABLED": callback_enabled,
        "ACCESS_VCS_CALLBACK_HOST": os.getenv("ACCESS_VCS_CALLBACK_HOST", "127.0.0.1"),

        # Logging configuration
        "ACCESS_VCS_ENABLE_LOGGING": os.getenv("ACCESS_VCS_ENABLE_LOGGING", "false").lower() == "true",
        "ACCESS_VCS_LOG_DIR": os.getenv("ACCESS_VCS_LOG_DIR", ""),
        "ACCESS_VCS_LOG_MAX_SIZE_MB": int(os.getenv("ACCESS_VCS_LOG_MAX_SIZE_MB", "10")),
        "ACCESS_VCS_LOG_BACKUP_COUNT": int(os.getenv("ACCESS_VCS_LOG_BACKUP_COUNT", "5")),

        # Runtime values (set by main.py after callback server starts)
        # ACCESS_VCS_CALLBACK_URL - set in environment after server starts
    }

    if reloaded:
        # Reset logging so changes to ACCESS_VCS_ENABLE_LOGGING,
        # ACCESS_VCS_LOG_DIR, etc. take effect without a server restart.
        from .usage_logging import reset_logging
        reset_logging()
        import sys
        print("Configuration reloaded from .env", file=sys.stderr)

    return config


def get_config() -> dict[str, Any]:
    """
    Get the current configuration.
    
    Returns:
        Configuration dictionary
    """
    return load_config()


def get_callback_url() -> str | None:
    """
    Get the callback URL if callback server is running.
    
    Returns:
        Callback URL or None if not available
    """
    return os.environ.get("ACCESS_VCS_CALLBACK_URL")


def get_session_id() -> str | None:
    """
    Get the MCP server session ID for option override scoping.
    
    Generated at startup and stored in the environment. Used to create
    session-specific override files (mcp/options-{session_id}.json).
    
    Returns:
        Session ID string or None if not set
    """
    return os.environ.get("ACCESS_VCS_SESSION_ID")


def validate_access_installation() -> None:
    """
    Verify that Microsoft Access COM automation is available.
    
    Raises:
        ImportError: If pywin32 is not installed
        RuntimeError: If Access COM automation is not available
    """
    try:
        import win32com.client
        from win32com.client import gencache
    except ImportError:
        raise ImportError(
            "pywin32 is required for Access COM automation. "
            "Install it with: pip install pywin32"
        )
    
    # Try to create Access application object
    try:
        # This will fail if Access is not installed
        # Use EnsureDispatch for early binding (fixes Application.Run issues)
        app = gencache.EnsureDispatch("Access.Application")
        
        # Check if this is the user's instance (has a database open)
        # If so, do NOT quit - we'd close their work!
        has_user_db = False
        try:
            current_db = app.CurrentDb()
            if current_db is not None:
                has_user_db = True
        except Exception:
            pass
        
        # Only quit if we created a new empty instance
        if not has_user_db:
            try:
                app.Quit()
            except Exception:
                pass
    except Exception as e:
        raise RuntimeError(
            f"Microsoft Access COM automation not available. "
            f"Ensure Microsoft Access is installed. Error: {e}"
        )
