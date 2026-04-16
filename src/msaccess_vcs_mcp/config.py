"""Configuration management for msaccess-vcs-mcp."""

import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


def _find_project_root() -> Path:
    """
    Find the project root by searching for .env file or common project markers.
    
    Searches upward from multiple starting points:
    1. Current working directory
    2. Location of this script (if available)
    
    For each starting point, searches upward for:
    1. .env file (primary indicator)
    2. .cursor/mcp.json (MCP configuration)
    3. pyproject.toml (Python project marker)
    
    Returns:
        Path to project root, or current working directory if not found
    """
    import sys
    
    # Try multiple starting points
    search_roots = [Path.cwd().resolve()]
    
    # Also try to find project root relative to this package's location
    try:
        # Get the directory containing this config.py file
        package_dir = Path(__file__).parent.parent.parent.resolve()
        # Walk up to find project root (should be parent of src/)
        if (package_dir.parent / "pyproject.toml").exists():
            search_roots.append(package_dir.parent)
        search_roots.append(package_dir.parent)
    except Exception:
        pass
    
    # Search from each starting point
    for root in search_roots:
        current = root
        # Walk up the directory tree
        while current != current.parent:
            # Check for .env file (primary indicator)
            if (current / ".env").exists():
                return current
            # Check for MCP config (secondary indicator)
            if (current / ".cursor" / "mcp.json").exists():
                return current
            # Check for Python project marker (tertiary indicator)
            if (current / "pyproject.toml").exists():
                return current
            
            current = current.parent
    
    # Fall back to current working directory
    return Path.cwd().resolve()


def _load_env_files() -> None:
    """
    Load .env files from the project root.
    
    Searches for project root by looking for .env file or project markers,
    then loads:
    1. .env (base configuration)
    2. .env.local (local overrides, if exists)
    
    Environment variables already set (e.g., from MCP server env section) take precedence.
    """
    import sys
    
    # Find project root (searches upward from current working directory)
    project_root = _find_project_root()
    
    # Load .env file if it exists
    env_path = project_root / ".env"
    if env_path.exists():
        # Convert Path to string for load_dotenv compatibility
        result = load_dotenv(str(env_path), override=False)
        if not result:
            print(f"Warning: .env file exists at {env_path} but no variables were loaded", file=sys.stderr)
    
    # Load .env.local if it exists (takes precedence over .env)
    env_local_path = project_root / ".env.local"
    if env_local_path.exists():
        # Convert Path to string for load_dotenv compatibility
        load_dotenv(str(env_local_path), override=True)


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
    
    Returns:
        Dictionary with configuration values
    """
    # Load .env files first (if not already loaded)
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

    # Reset logging state so changes to ACCESS_VCS_ENABLE_LOGGING,
    # ACCESS_VCS_LOG_DIR, etc. take effect without a server restart.
    from .usage_logging import reset_logging
    reset_logging()

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
