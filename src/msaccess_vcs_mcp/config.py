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
    
    return {
        "ACCESS_VCS_DATABASE": os.getenv("ACCESS_VCS_DATABASE", ""),
        "ACCESS_VCS_ADDIN_PATH": os.getenv("ACCESS_VCS_ADDIN_PATH", get_default_addin_path()),
        "ACCESS_VCS_DISABLE_WRITES": os.getenv("ACCESS_VCS_DISABLE_WRITES", "false").lower() == "true",
    }


def get_config() -> dict[str, Any]:
    """
    Get the current configuration.
    
    Returns:
        Configuration dictionary
    """
    return load_config()


def validate_access_installation() -> None:
    """
    Verify that Microsoft Access COM automation is available.
    
    Raises:
        ImportError: If pywin32 is not installed
        RuntimeError: If Access COM automation is not available
    """
    try:
        import win32com.client
    except ImportError:
        raise ImportError(
            "pywin32 is required for Access COM automation. "
            "Install it with: pip install pywin32"
        )
    
    # Try to create Access application object
    try:
        # This will fail if Access is not installed
        app = win32com.client.Dispatch("Access.Application")
        # Clean up immediately
        try:
            app.Quit()
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(
            f"Microsoft Access COM automation not available. "
            f"Ensure Microsoft Access is installed. Error: {e}"
        )
