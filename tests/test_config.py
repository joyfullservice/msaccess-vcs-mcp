"""Tests for configuration management.

Note: load_config schema and defaults are exercised in
``tests/test_config_env_loading.py`` against the current env-var
schema (ACCESS_VCS_DATABASE, ACCESS_VCS_DISABLE_WRITES, etc.).
"""

import pytest


def test_validate_access_installation():
    """Test Access COM validation."""
    from msaccess_vcs_mcp.config import validate_access_installation
    
    # This test will only pass if Access is installed
    # On CI/CD without Access, this would be skipped
    try:
        validate_access_installation()
    except (ImportError, RuntimeError) as e:
        pytest.skip(f"Access not available: {e}")
