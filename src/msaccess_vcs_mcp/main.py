"""Main entry point for msaccess-vcs-mcp MCP server."""

import asyncio
import atexit
import os
import sys
import uuid
from pathlib import Path

from . import __version__
from .config import get_config, validate_access_installation
from .usage_logging import log_diagnostic_event


# Global reference to callback server for cleanup
_callback_server = None

# Session ID generated at startup for option override scoping
_session_id = uuid.uuid4().hex[:8]


def _start_callback_server(config: dict) -> str | None:
    """
    Start the HTTP callback server for VBA progress updates.
    
    Args:
        config: Configuration dictionary
        
    Returns:
        The callback URL, or None if disabled
    """
    global _callback_server
    
    # Check if callbacks are enabled
    if config.get("ACCESS_VCS_CALLBACK_ENABLED") is False:
        print("Callback server: disabled", file=sys.stderr)
        return None
    
    try:
        from .callback_server import CallbackServer
        from .operation_manager import OperationManager
        
        # Get operation manager (creates singleton)
        op_manager = OperationManager.get_instance()
        
        # Set the event loop for cross-thread queue operations
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            # No running loop yet - it will be set when MCP starts
            loop = asyncio.new_event_loop()
        op_manager.set_event_loop(loop)
        
        # Create and start callback server
        host = config.get("ACCESS_VCS_CALLBACK_HOST", "127.0.0.1")
        _callback_server = CallbackServer(
            callback_router=op_manager.route_callback,
            cancel_checker=op_manager.is_cancelled,
            cancel_requester=op_manager.request_cancel,
            host=host,
            port=0  # OS assigns available port
        )
        _callback_server.start()
        
        # Register cleanup on exit
        atexit.register(_stop_callback_server)
        
        callback_url = _callback_server.callback_url
        print(f"✓ Callback server: {callback_url}", file=sys.stderr)
        return callback_url
        
    except Exception as e:
        print(f"⚠ Callback server failed to start: {e}", file=sys.stderr)
        return None


def _stop_callback_server() -> None:
    """Stop the callback server on exit."""
    global _callback_server
    if _callback_server:
        _callback_server.stop()
        _callback_server = None


def _cleanup_session() -> None:
    """End the MCP session on server shutdown, cleaning up override files."""
    session_id = os.environ.get("ACCESS_VCS_SESSION_ID")
    db_path = os.environ.get("ACCESS_VCS_DATABASE")
    if not session_id or not db_path:
        return
    try:
        from .access_com.connection import AccessConnection
        from .addin_integration import VCSAddinIntegration
        config = get_config()
        with AccessConnection(db_path) as conn:
            app, db = conn.connect()
            addin = VCSAddinIntegration(config.get("ACCESS_VCS_ADDIN_PATH"))
            addin.load_addin(app, db_path=db_path)
            addin.call_sync("EndSession", session_id)
            print(f"Session {session_id}: overrides cleaned up", file=sys.stderr)
    except Exception as e:
        print(f"Session cleanup skipped: {e}", file=sys.stderr)


def main() -> None:
    """Main entry point for the MCP server."""
    # Always-on lifecycle event: emitted before .env discovery so we still
    # have a record even when the project root cannot be resolved.
    log_diagnostic_event(
        "server_start",
        cwd=str(Path.cwd()),
        project_dir_env=os.getenv("ACCESS_VCS_PROJECT_DIR"),
        mcp_version=__version__,
        session_id=_session_id,
    )

    # Load configuration (.env files from project root)
    config = get_config()
    
    # Verify Access COM is available
    try:
        validate_access_installation()
        print("✓ Microsoft Access COM automation available", file=sys.stderr)
    except ImportError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("Install pywin32: pip install pywin32", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Print configuration summary
    write_mode = "disabled" if config.get("ACCESS_VCS_DISABLE_WRITES") else "enabled"
    print(f"Write operations: {write_mode}", file=sys.stderr)
    if config.get("ACCESS_VCS_DATABASE"):
        print(f"Target database: {config['ACCESS_VCS_DATABASE']}", file=sys.stderr)
    
    # Optional startup validation
    if os.environ.get("ACCESS_VCS_VALIDATE_STARTUP", "false").lower() == "true":
        print("\n--- Validating components on startup ---", file=sys.stderr)
        from .validation import validate_components
        
        validation = validate_components(load_addin=False)  # Quick check, don't load add-in
        
        # Display validation results
        print(f"MCP Version: {validation.get('mcp_version', 'unknown')}", file=sys.stderr)
        if validation.get("access_version"):
            print(f"Access Version: {validation.get('access_version')} ({validation.get('bitness', 'unknown')})", file=sys.stderr)
        if validation.get("addin_path"):
            addin_exists = "✓" if os.path.exists(validation.get("addin_path", "")) else "✗"
            print(f"{addin_exists} VCS Add-in: {validation.get('addin_path')}", file=sys.stderr)
        if validation.get("target_database"):
            db_exists = "✓" if os.path.exists(validation.get("target_database", "")) else "✗"
            print(f"{db_exists} Target Database: {validation.get('target_database')}", file=sys.stderr)
        
        # Show errors and warnings
        if validation.get("errors"):
            print("\nErrors:", file=sys.stderr)
            for error in validation["errors"]:
                print(f"  ✗ {error}", file=sys.stderr)
        
        if validation.get("warnings"):
            print("\nWarnings:", file=sys.stderr)
            for warning in validation["warnings"]:
                print(f"  ⚠ {warning}", file=sys.stderr)
        
        if validation.get("success"):
            print("\n✓ Component validation passed", file=sys.stderr)
        else:
            print("\n✗ Component validation failed (server will still start)", file=sys.stderr)
        
        print("--- End validation ---\n", file=sys.stderr)
    
    # Start callback server for VBA progress updates
    callback_url = _start_callback_server(config)
    
    # Store callback URL in environment for tools to access
    if callback_url:
        os.environ["ACCESS_VCS_CALLBACK_URL"] = callback_url
    
    # Store session ID for option override scoping
    os.environ["ACCESS_VCS_SESSION_ID"] = _session_id
    atexit.register(_cleanup_session)
    print(f"Session ID: {_session_id}", file=sys.stderr)
    
    # Show logging status
    from .usage_logging import (
        get_diagnostic_log_path,
        get_log_file_path,
        is_diagnostic_logging_enabled,
        is_logging_enabled,
    )
    usage_enabled = is_logging_enabled()
    usage_path = get_log_file_path() if usage_enabled else None
    if usage_enabled:
        print(f"Usage logging: {usage_path}", file=sys.stderr)
    else:
        print("Usage logging: disabled (set ACCESS_VCS_ENABLE_LOGGING=true to enable)", file=sys.stderr)

    # Diagnostic stream is always-on by default; surface it so the operator
    # knows where to look when usage logging is silent.
    if is_diagnostic_logging_enabled():
        print(f"Diagnostic logging: {get_diagnostic_log_path()}", file=sys.stderr)
    else:
        print("Diagnostic logging: disabled", file=sys.stderr)

    log_diagnostic_event(
        "usage_log_status",
        enabled=usage_enabled,
        log_file=str(usage_path) if usage_path else None,
    )

    # Import and run MCP server
    from .tools import mcp
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
