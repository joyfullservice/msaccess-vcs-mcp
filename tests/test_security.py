"""Tests for security and validation."""

import pytest
from pathlib import Path
import tempfile


def test_validate_database_path_invalid_extension():
    """Test that invalid extensions are rejected."""
    from msaccess_vcs_mcp.security import validate_database_path
    
    with pytest.raises(ValueError, match="Invalid database extension"):
        validate_database_path("test.txt")


def test_validate_database_path_not_exists():
    """Test that non-existent paths are rejected."""
    from msaccess_vcs_mcp.security import validate_database_path
    
    with pytest.raises(ValueError, match="Database not found"):
        validate_database_path("C:\\nonexistent\\database.accdb")


def test_validate_export_directory_creates():
    """Test that export directory is created if it doesn't exist."""
    from msaccess_vcs_mcp.security import validate_export_directory
    
    with tempfile.TemporaryDirectory() as tmpdir:
        export_dir = Path(tmpdir) / "exports"
        result = validate_export_directory(str(export_dir), allow_create=True)
        
        assert result.exists()
        assert result.is_dir()


def test_validate_export_directory_system_denied():
    """Test that system directories are rejected."""
    from msaccess_vcs_mcp.security import validate_export_directory
    
    with pytest.raises(ValueError, match="Cannot export to system directory"):
        validate_export_directory("C:\\Windows\\Temp")


def test_check_write_permission_denied():
    """Test write permission check."""
    from msaccess_vcs_mcp.security import check_write_permission

    config = {"ACCESS_VCS_DISABLE_WRITES": True}

    with pytest.raises(PermissionError, match="Write operations are disabled"):
        check_write_permission(config)


def test_check_write_permission_allowed():
    """Test write permission when enabled."""
    from msaccess_vcs_mcp.security import check_write_permission

    config = {"ACCESS_VCS_DISABLE_WRITES": False}

    # Should not raise
    check_write_permission(config)
