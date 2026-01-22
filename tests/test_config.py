"""Tests for configuration management."""

import os
import pytest
from pathlib import Path


def test_load_config():
    """Test configuration loading."""
    from msaccess_vcs_mcp.config import load_config
    
    config = load_config()
    
    assert "ACCESS_VCS_DEFAULT_DB" in config
    assert "ACCESS_VCS_EXPORT_FORMAT" in config
    assert "ACCESS_VCS_ALLOW_WRITES" in config
    assert "ACCESS_VCS_ENCODING" in config


def test_default_values():
    """Test default configuration values."""
    from msaccess_vcs_mcp.config import load_config
    
    # Clear environment to test defaults
    env_vars = [
        "ACCESS_VCS_DEFAULT_DB",
        "ACCESS_VCS_EXPORT_FORMAT",
        "ACCESS_VCS_ALLOW_WRITES",
        "ACCESS_VCS_ENCODING",
    ]
    old_values = {}
    for var in env_vars:
        old_values[var] = os.environ.get(var)
        if var in os.environ:
            del os.environ[var]
    
    try:
        config = load_config()
        
        assert config["ACCESS_VCS_EXPORT_FORMAT"] == "text"
        assert config["ACCESS_VCS_ALLOW_WRITES"] == "false"
        assert config["ACCESS_VCS_ENCODING"] == "utf-8-sig"
    finally:
        # Restore environment
        for var, value in old_values.items():
            if value is not None:
                os.environ[var] = value


def test_validate_access_installation():
    """Test Access COM validation."""
    from msaccess_vcs_mcp.config import validate_access_installation
    
    # This test will only pass if Access is installed
    # On CI/CD without Access, this would be skipped
    try:
        validate_access_installation()
    except (ImportError, RuntimeError) as e:
        pytest.skip(f"Access not available: {e}")
