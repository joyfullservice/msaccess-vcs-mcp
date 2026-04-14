"""
Validation utilities for checking component availability and version information.

This module provides shared validation logic used for:
- Startup validation checks
- Troubleshooting via vcs_get_version_info() tool
- Pre-operation validation in other tools
"""

import os
from typing import Any

from . import __version__
from .addin_integration import VCSAddinIntegration, get_access_info
from .config import get_config

try:
    import win32com.client
    from win32com.client import gencache
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def normalize_path(path: str) -> str:
    """
    Normalize a file path for comparison, handling both local and UNC paths.
    
    Args:
        path: File path (can be local or UNC like \\server\share\file.accdb)
        
    Returns:
        Normalized path suitable for comparison
    """
    # Normalize the path
    normalized = os.path.normpath(path)
    
    # For UNC paths, os.path.abspath() doesn't work correctly
    # For local paths, we can use abspath to resolve relative paths
    if normalized.startswith('\\\\'):
        # UNC path - just normalize, don't use abspath
        return normalized
    else:
        # Local path - use abspath to resolve relative paths
        return os.path.abspath(normalized)


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
    
    # Try to get or create Access instance
    # We need to be careful not to close databases the user has open
    app = None
    owns_app = False  # Track if we created the Access instance
    db_was_already_open = False  # Track if database was already open
    
    try:
        # Try to connect to the specific Access instance that has the target database open
        if target_db and os.path.exists(target_db):
            target_db_normalized = normalize_path(target_db)
            
            # GetObject(target_db) will find the Access instance that has this database open
            # If the database is not open, it will raise an exception
            try:
                # Use GetObject(path) to connect to the running instance with the database open
                # This is more reliable than EnsureDispatch for finding existing instances
                # Note: After gencache.EnsureDispatch has been called once, Run returns tuples
                app = win32com.client.GetObject(target_db)
                owns_app = False  # User's Access instance - don't close it
                db_was_already_open = True
                
                # Verify it's actually the target database (safety check)
                try:
                    current_db = app.CurrentDb()
                    if current_db is not None:
                        current_db_path = normalize_path(current_db.Name)
                        if current_db_path != target_db_normalized:
                            # Path mismatch - this shouldn't happen if GetObject worked correctly
                            result["warnings"].append(
                                f"Database path mismatch: Expected {target_db_normalized}, "
                                f"got {current_db_path}. This may indicate multiple Access instances."
                            )
                except Exception:
                    # Can't verify - assume it's correct since GetObject(target_db) succeeded
                    pass
            except Exception:
                # Database not open in any Access instance - create our own instance
                # Use EnsureDispatch for early binding (fixes Application.Run issues)
                app = gencache.EnsureDispatch("Access.Application")
                owns_app = True
                db_was_already_open = False
        else:
            # No target database configured, just create/get Access instance
            # Use EnsureDispatch for early binding (fixes Application.Run issues)
            try:
                app = gencache.EnsureDispatch("Access.Application")
                owns_app = False
            except Exception:
                app = gencache.EnsureDispatch("Access.Application")
                owns_app = True
        
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
            # To call add-in functions, we need the target database open
            # But we should NOT close it if the user already had it open
            try:
                # Normalize paths for comparison (handles both local and UNC paths)
                target_db_normalized = normalize_path(target_db)
                current_db_path = None
                
                # Check if database is already open and matches target
                if db_was_already_open:
                    try:
                        # Verify the current database matches the target
                        current_db = app.CurrentDb()
                        if current_db:
                            current_db_path = normalize_path(current_db.Name)
                            if current_db_path != target_db_normalized:
                                # Database is open but it's a different one
                                # Only open target if we own the app
                                if owns_app:
                                    app.OpenCurrentDatabase(target_db)
                                    db_was_already_open = False  # We opened it now
                                else:
                                    # Can't change user's database, skip version check
                                    result["warnings"].append(
                                        f"Cannot retrieve VCS version: Access has a different database open. "
                                        f"Current: {current_db_path}, Target: {target_db_normalized}"
                                    )
                                    return result
                    except Exception:
                        # Can't verify current database, try to open target if we own app
                        if owns_app:
                            app.OpenCurrentDatabase(target_db)
                            db_was_already_open = False
                        else:
                            result["warnings"].append(
                                "Cannot retrieve VCS version: Unable to verify current database"
                            )
                            return result
                else:
                    # Database wasn't already open, open it now
                    app.OpenCurrentDatabase(target_db)
                
                # Final verification: ensure we're working with the correct database
                try:
                    current_db = app.CurrentDb()
                    if not current_db:
                        result["warnings"].append(
                            "Cannot retrieve VCS version: Database is not accessible"
                        )
                        return result
                    
                    # Verify the database path matches the target
                    current_db_path = normalize_path(current_db.Name)
                    if current_db_path != target_db_normalized:
                        result["warnings"].append(
                            f"Cannot retrieve VCS version: Database path mismatch. "
                            f"Expected: {target_db_normalized}, Got: {current_db_path}"
                        )
                        return result
                except Exception as db_error:
                    result["warnings"].append(
                        f"Cannot retrieve VCS version: Database not accessible: {db_error}"
                    )
                    return result
                
                # Now that we've verified the correct database is open, we can call the add-in function directly
                # The add-in loads automatically when called via Application.Run
                # No need for a separate "load" step
                # Ensure add-in has app reference set
                addin._app = app
                version_info = addin.get_version_info(app)
                
                if version_info.get("vcs_version"):
                    result["vcs_version"] = version_info["vcs_version"]
                else:
                    result["warnings"].append(
                        version_info.get("vcs_error", "Could not retrieve VCS version")
                    )
                
                # Only close the database if we opened it (not if user had it open)
                if not db_was_already_open and owns_app:
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
        # Only quit Access if we created the instance ourselves
        if app and owns_app:
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
