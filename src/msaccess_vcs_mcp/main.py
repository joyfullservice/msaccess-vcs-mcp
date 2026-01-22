"""Main entry point for msaccess-vcs-mcp MCP server."""

import os
import sys

from .config import get_config, validate_access_installation


def main() -> None:
    """Main entry point for the MCP server."""
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
    
    # Import and run MCP server
    from .tools import mcp
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
