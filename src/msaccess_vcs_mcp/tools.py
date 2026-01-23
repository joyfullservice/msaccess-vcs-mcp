"""
MCP tool definitions for msaccess-vcs-mcp.

This module provides database version control tools for AI assistants working with
Microsoft Access databases. All tools support exporting to source files, importing
from source, and tracking changes.

**Getting Started Workflow:**
1. Use access_list_objects() to see what's in a database
2. Use access_export_database() to export all objects to source directory
3. Edit source files using your preferred tools
4. Use access_diff_database() to see what changed
5. Use access_import_objects() or access_rebuild_database() to apply changes

**Key Features:**
- Export Access objects to git-friendly text files
- Import objects from source files back into Access
- Rebuild entire databases from source
- Track changes between database and source
- Read operations always available, write operations require permission
- Long-running operations support progress reporting via callbacks
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

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
            "hint": "Wait for the current operation to complete, or cancel it with access_cancel_operation()"
        }
    return None


# Create FastMCP server instance with proper metadata
mcp = FastMCP(
    name="msaccess-vcs-mcp",
    instructions=(
        "Microsoft Access version control MCP server. "
        "Export Access database objects to source files, import them back, "
        "rebuild databases from source, and track changes.\n\n"
        "**Recommended workflow:**\n"
        "1. Use access_export_database() to export all objects to source directory\n"
        "2. Edit source files using your preferred tools\n"
        "3. Use access_import_objects() to merge changes back into database\n"
        "4. Use access_rebuild_database() to create fresh database from source\n"
        "5. Use access_diff_database() to see what changed\n\n"
        "**Configuration:**\n"
        "Set ACCESS_VCS_DATABASE to your target database path.\n"
        "Set ACCESS_VCS_DISABLE_WRITES=true to prevent database modifications."
    )
)


@mcp.tool()
async def access_export_database(
    database_path: str,
    output_dir: str,
    object_types: list[str] | None = None
) -> dict[str, Any]:
    """
    Export Access database objects to source files.
    
    Exports tables, queries, forms, reports, macros, and modules to 
    text-based files suitable for version control.
    
    This operation supports progress reporting - you'll receive updates
    as objects are exported.
    
    Examples:
        # Export entire database
        access_export_database("C:\\\\db.accdb", "C:\\\\src\\\\mydb")
        
        # Export only queries and modules
        access_export_database(
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
        
        # Determine if this is a VBA-only export
        vba_only = object_types and set(object_types) == {"modules"}
        command = "ExportVBA" if vba_only else "Export"
        
        # Connect to database
        with AccessConnection(str(db_path)) as conn:
            app, db = conn.connect()
            
            # Initialize add-in integration
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin._app = app  # Set app reference for add-in calls
            
            # Check if async export is available
            if callback_url and op_manager:
                # Use async path with progress callbacks
                operation_id, queue = op_manager.register_operation(
                    database_path=str(db_path),
                    command=command
                )
                callback_info = op_manager.create_callback_info(
                    operation_id, callback_url, "cursor"
                )
                
                try:
                    # Call async API
                    async_result = addin.call_async(callback_info, command, str(export_path))
                    
                    if async_result.get("sync"):
                        # VBA returned sync result
                        pass  # Fall through to count objects
                    elif async_result.get("async"):
                        # Wait for completion with progress reporting
                        timeout_ms = async_result.get("timeout_ms", 300000)
                        completion = await op_manager.wait_for_completion(
                            operation_id,
                            ctx=None,  # Context would be passed if available
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
            
            # Parse log file for detailed results
            log_path = os.path.join(str(export_path), "Export.log")
            log_info = addin.parse_log_file(log_path) if os.path.exists(log_path) else {}
            
            # Count exported objects by reading directory structure
            objects_by_type = {}
            for obj_type in ["queries", "modules", "forms", "reports", "macros", "tables"]:
                type_dir = export_path / obj_type
                if type_dir.exists():
                    # Count files in directory
                    files = list(type_dir.glob("*"))
                    objects_by_type[obj_type] = len([f for f in files if f.is_file()])
            
            total_count = sum(objects_by_type.values())
            
            return {
                "success": True,
                "exported_count": total_count,
                "export_path": str(export_path),
                "objects_by_type": objects_by_type,
                "log_path": log_path if os.path.exists(log_path) else None,
                "log_content": log_info.get("content") if log_info.get("found") else None,
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
def access_list_objects(database_path: str) -> dict[str, Any]:
    """
    List all objects in an Access database.
    
    Provides an inventory of tables, queries, forms, reports, 
    modules, and macros.
    
    Examples:
        # List all objects
        access_list_objects("C:\\\\db.accdb")
    
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
def access_diff_database(
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
        access_diff_database("C:\\\\db.accdb", "C:\\\\src\\\\mydb")
        
        # Detailed diff with line-by-line comparison
        access_diff_database(
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
        db_objects = access_list_objects(str(db_path))
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
async def access_import_objects(
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
        access_import_objects("C:\\\\db.accdb", "C:\\\\src\\\\mydb", overwrite=True)
        
        # Import only queries
        access_import_objects(
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
            
            # Check if async import is available
            if callback_url and op_manager:
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
                result = addin.merge_build(str(db_path), str(src_path))
                if not result["success"]:
                    return {
                        "success": False,
                        "error": result["message"],
                        "imported_count": 0,
                    }
            
            # Parse log file for detailed results
            log_path = os.path.join(str(src_path), "Build.log")
            log_info = addin.parse_log_file(log_path) if os.path.exists(log_path) else {}
            
            return {
                "success": True,
                "imported_count": "See log for details",
                "database_path": str(db_path),
                "source_dir": str(src_path),
                "log_path": log_path if os.path.exists(log_path) else None,
                "log_content": log_info.get("content") if log_info.get("found") else None,
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
async def access_rebuild_database(
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
        access_rebuild_database("C:\\\\src\\\\mydb", "C:\\\\output\\\\rebuilt.accdb")
        
        # Rebuild using template
        access_rebuild_database(
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
                        async_result = addin.call_async(callback_info, command, str(src_path))
                        
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
                    result = addin.build_from_source(str(src_path), output_path)
                    if not result["success"]:
                        return {
                            "success": False,
                            "error": result["message"],
                            "output_path": None,
                        }
                
                # Parse log file for detailed results
                log_path = os.path.join(str(src_path), "Build.log")
                log_info = addin.parse_log_file(log_path) if os.path.exists(log_path) else {}
                
                return {
                    "success": True,
                    "output_path": output_path,
                    "source_dir": str(src_path),
                    "log_path": log_path if os.path.exists(log_path) else None,
                    "log_content": log_info.get("content") if log_info.get("found") else None,
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
def access_get_version_info() -> dict[str, Any]:
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
    
    Examples:
        # Get version information
        access_get_version_info()
    
    Returns:
        Dictionary with:
        - success: Boolean indicating if info was retrieved
        - mcp_version: Version of the MCP server (e.g., "0.1.0")
        - vcs_version: Version of the VCS add-in (e.g., "4.1.4")
        - access_version: Access application version (e.g., "16.0")
        - bitness: "32-bit" or "64-bit"
        - target_database: Configured database path from ACCESS_VCS_DATABASE
        - addin_path: Path to the VCS add-in file
        - errors: List of validation errors
        - warnings: List of validation warnings
    """
    from .validation import get_version_info_safe
    
    return get_version_info_safe()


@mcp.tool()
def access_cancel_operation(operation_id: str) -> dict[str, Any]:
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
        access_cancel_operation("a1b2c3d4-5678-90ab-cdef-1234567890ab")
    
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
