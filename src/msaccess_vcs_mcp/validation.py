"""
Validation utilities for checking component availability and version information.

This module provides shared validation logic used for:
- Startup validation checks
- Troubleshooting via access_get_version_info() tool
- Pre-operation validation in other tools
"""

import os
from typing import Any

from . import __version__
from .addin_integration import VCSAddinIntegration, get_access_info
from .config import get_config

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def validate_components(load_addin: bool = True) -> dict[str, Any]:
    """
    Validate that all required components are available.
    
    This function checks:
    - Access application is installed and accessible via COM
    - VCS add-in exists at configured path
    - Target database path from configuration (if set)
    - Can successfully retrieve version information
    
    Args:
        load_addin: If True, attempt to load add-in and get versions (slower).
                   If False, only check file existence (faster).
    
    Returns:
        Dictionary with validation results, version information, and configured paths:
        - success: Boolean indicating if all checks passed
        - mcp_version: MCP server version
        - vcs_version: VCS add-in version (if load_addin=True)
        - access_version: Access application version (if COM available)
        - bitness: Access bitness (if COM available)
        - target_database: Configured database path from ACCESS_VCS_DATABASE
        - addin_path: Path to VCS add-in file
        - errors: List of validation errors
        - warnings: List of validation warnings
    """
    result = {
        "success": True,
        "mcp_version": __version__,
        "errors": [],
        "warnings": [],
    }
    
    # Get configuration
    config = get_config()
    
    # Add target database if configured
    target_db = config.get("ACCESS_VCS_DATABASE")
    if target_db:
        result["target_database"] = target_db
        # Check if it exists
        if not os.path.exists(target_db):
            result["warnings"].append(
                f"Target database does not exist: {target_db}"
            )
    else:
        result["target_database"] = None
        result["warnings"].append(
            "No target database configured (ACCESS_VCS_DATABASE not set)"
        )
    
    # Check COM availability
    if not COM_AVAILABLE:
        result["success"] = False
        result["errors"].append(
            "pywin32 is required for COM automation. Install it with: pip install pywin32"
        )
        return result
    
    # Try to create Access instance
    app = None
    try:
        app = win32com.client.Dispatch("Access.Application")
        
        # Get Access info
        access_info = get_access_info(app)
        result.update(access_info)
        
        if access_info.get("error"):
            result["warnings"].append(f"Could not retrieve Access info: {access_info['error']}")
    
    except Exception as e:
        result["success"] = False
        result["errors"].append(
            f"Failed to create Access application instance: {e}. "
            "Is Microsoft Access installed?"
        )
        return result
    
    # Check VCS add-in
    try:
        addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
        result["addin_path"] = addin.addin_path
        
        # Check if add-in file exists
        if not addin.verify_addin_exists():
            result["success"] = False
            result["errors"].append(
                f"VCS add-in not found at: {addin.addin_path}. "
                "Please install the MSAccess VCS add-in or set ACCESS_VCS_ADDIN_PATH. "
                "Download from: https://github.com/joyfullservice/msaccess-vcs-integration/releases"
            )
        elif load_addin and app and target_db and os.path.exists(target_db):
            # To call add-in functions, we need to open the target database first
            # Then call add-in functions using Application.Run syntax
            try:
                app.OpenCurrentDatabase(target_db)
                version_info = addin.get_version_info(app)
                
                if version_info.get("vcs_version"):
                    result["vcs_version"] = version_info["vcs_version"]
                else:
                    result["warnings"].append(
                        version_info.get("vcs_error", "Could not retrieve VCS version")
                    )
                
                app.CloseCurrentDatabase()
            except Exception as e:
                result["warnings"].append(
                    f"Could not retrieve VCS version: {e}"
                )
        elif load_addin and not target_db:
            result["warnings"].append(
                "Cannot retrieve VCS version without a target database configured"
            )
        elif load_addin and target_db and not os.path.exists(target_db):
            result["warnings"].append(
                f"Cannot retrieve VCS version: Target database does not exist: {target_db}"
            )
    
    except Exception as e:
        result["warnings"].append(f"Error checking VCS add-in: {e}")
    
    finally:
        # Clean up Access instance
        if app:
            try:
                app.Quit()
            except Exception:
                pass
    
    return result


def get_version_info_safe() -> dict[str, Any]:
    """
    Safely get version information without raising exceptions.
    
    This is a wrapper around validate_components that's suitable for use
    in MCP tools where we want comprehensive error information in the response
    rather than throwing exceptions.
    
    Returns:
        Dictionary with version information or error details
    """
    try:
        return validate_components(load_addin=True)
    except Exception as e:
        return {
            "success": False,
            "mcp_version": __version__,
            "error": f"Validation failed with exception: {e}",
            "errors": [str(e)],
            "warnings": []
        }
