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
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP, Context

from .access_com.connection import AccessConnection
from .access_com.dao_helpers import list_query_defs, list_table_defs
from .config import get_config, get_callback_url
from .addin_integration import VCSAddinIntegration
from .security import (
    validate_database_path,
    validate_export_directory,
    validate_source_directory,
    check_write_permission,
)


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
        "Set ACCESS_VCS_DISABLE_WRITES=true to prevent database modifications."
    )
)


@mcp.tool()
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
        vba_only = object_types and set(object_types) == {"modules"}
        if vba_only:
            command = "ExportVBA"
        elif full_export:
            command = "FullExport"
        else:
            command = "Export"
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            # Initialize add-in integration
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app  # Set app reference for add-in calls
            
            # Pre-flight check: verify add-in is responsive (not blocked by dialogs)
            try:
                addin.call_sync("GetVCSVersion")
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


@mcp.tool()
def vcs_list_objects(database_path: str) -> dict[str, Any]:
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


@mcp.tool()
def vcs_diff_database(
    database_path: str,
    source_dir: str,
    show_details: bool = False
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


@mcp.tool()
async def vcs_import_objects(
    database_path: str,
    source_dir: str,
    object_types: list[str] | None = None,
    overwrite: bool = False
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
            
            # Initialize add-in integration
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app  # Set app reference for add-in calls
            
            # Pre-flight check: verify add-in is responsive (not blocked by dialogs)
            try:
                addin.call_sync("GetVCSVersion")
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


@mcp.tool()
async def vcs_rebuild_database(
    source_dir: str,
    output_path: str,
    template_path: str | None = None
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
        
        # Note: We need Access to be running to call the add-in,
        # but we don't have a database open yet. The add-in's build
        # process will create the database.
        
        # Create a temporary Access instance just to load the add-in
        with AccessConnection.__new__(AccessConnection) as conn:
            # Create Access app without opening a database
            # Use EnsureDispatch for early binding (fixes Application.Run issues)
            from win32com.client import gencache
            app = gencache.EnsureDispatch("Access.Application")
            
            try:
                # Initialize add-in integration
                addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
                addin._app = app  # Set app reference for add-in calls
                
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
                        # Call async API for Build
                        # Build takes optional source folder, BuildAs does not
                        completion = None
                        if command == "Build":
                            async_result = addin.call_async(callback_info, command, str(src_path))
                        else:
                            async_result = addin.call_async(callback_info, command)
                        
                        if async_result.get("async"):
                            # Wait for completion with progress reporting
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
                # Clean up Access instance
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


@mcp.tool()
def vcs_get_version_info() -> dict[str, Any]:
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
        - errors: List of validation errors
        - warnings: List of validation warnings
    """
    from .validation import get_version_info_safe
    
    result = get_version_info_safe()
    
    # Add callback server status
    callback_url = get_callback_url()
    op_manager = _get_operation_manager()
    
    result["callback_url"] = callback_url
    result["async_available"] = bool(callback_url and op_manager)
    
    if not callback_url:
        result["warnings"] = result.get("warnings", []) + [
            "Callback server not running - async operations will fall back to sync mode"
        ]
    
    return result


@mcp.tool()
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


@mcp.tool()
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
            
            # Initialize add-in integration
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app  # Set app reference for add-in calls
            
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


@mcp.tool()
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
            
            # Initialize add-in integration
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app  # Set app reference for add-in calls
            
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


@mcp.tool()
def vcs_export_object(
    database_path: str,
    object_type: str,
    object_name: str
) -> dict[str, Any]:
    """
    Export a single named database object to source files.
    
    Exports one object (query, form, report, module, table, or macro) to its
    source file representation. Much faster than a full database export when
    you only need to refresh one object.
    
    Examples:
        vcs_export_object("C:\\\\db.accdb", "query", "qryCustomers")
        vcs_export_object("C:\\\\db.accdb", "form", "frmMain")
        vcs_export_object("C:\\\\db.accdb", "module", "modUtils")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        object_type: Type of object: "query", "form", "report", "module", "table", "macro"
        object_name: Name of the object to export
    
    Returns:
        Dictionary with success status, file path, and any errors
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app
            
            result_json = addin.call_sync("ExportObject", object_type, object_name)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def vcs_import_object(
    database_path: str,
    object_type: str,
    object_name: str
) -> dict[str, Any]:
    """
    Import a single named object from source files into the database.
    
    Loads one object from its source file back into the Access database.
    The source file must exist in the project's export folder.
    
    Examples:
        vcs_import_object("C:\\\\db.accdb", "query", "qryCustomers")
        vcs_import_object("C:\\\\db.accdb", "module", "modUtils")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        object_type: Type of object: "query", "form", "report", "module", "table", "macro"
        object_name: Name of the object to import
    
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
            addin._app = app
            
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


@mcp.tool()
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
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app
            
            result_json = addin.call_sync("ExecuteSQL", sql, max_rows)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
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
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            call_args = args or []
            
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


@mcp.tool()
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
    Enable via vcs_set_option or the VCS Options form.
    
    The agent's code should set the function return value. Example:
        Dim result As String
        result = CurrentDb.QueryDefs("qryCustomers").SQL
        _MCP_TempFunction = result
    
    Examples:
        vcs_run_vba("C:\\\\db.accdb", "_MCP_TempFunction = CurrentDb.TableDefs.Count")
        vcs_run_vba("C:\\\\db.accdb", "Dim qd As DAO.QueryDef\\nSet qd = CurrentDb.QueryDefs(\\"qryTest\\")\\n_MCP_TempFunction = qd.SQL")
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        code: VBA code to execute (statements, not just an expression)
    
    Returns:
        Dictionary with success, result, and any compile or runtime errors
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app
            
            result_json = addin.call_sync("RunVBA", code)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def vcs_set_option(
    database_path: str,
    option_name: str,
    value: str | bool | int
) -> dict[str, Any]:
    """
    Set a VCS add-in option for the current session.
    
    Changes take effect immediately but do not persist to vcs-options.json
    until explicitly saved. Useful for controlling add-in behavior during
    a testing session (e.g., enabling debug output).
    
    Examples:
        vcs_set_option("C:\\\\db.accdb", "ShowDebug", True)
        vcs_set_option("C:\\\\db.accdb", "McpAllowRunVBA", True)
        vcs_set_option("C:\\\\db.accdb", "MaxLogFiles", 50)
    
    Args:
        database_path: Path to Access database (.accdb, .accda, .mdb)
        option_name: Name of the VCS option property
        value: Value to set
    
    Returns:
        Dictionary with success status and the option that was set
    """
    try:
        db_path = validate_database_path(database_path)
        
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            config = get_config()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app
            
            result_json = addin.call_sync("SetOption", option_name, value)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "result": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def vcs_get_option(
    database_path: str,
    option_name: str
) -> dict[str, Any]:
    """
    Read a VCS add-in option value.
    
    Returns the current value of any VCS add-in option property.
    
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
            addin._app = app
            
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


@mcp.tool()
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
            addin._app = app
            
            result_json = addin.call_sync("GetLogContent", log_type)
            
            if isinstance(result_json, str):
                try:
                    return json.loads(result_json)
                except json.JSONDecodeError:
                    return {"success": True, "content": result_json}
            
            return {"success": True, "result": result_json}
    
    except Exception as e:
        return {"success": False, "error": str(e)}
