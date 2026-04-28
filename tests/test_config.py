"""Tests for configuration management.

Note: load_config schema and defaults are exercised in
``tests/test_config_env_loading.py`` against the current env-var
schema (ACCESS_VCS_DATABASE, ACCESS_VCS_DISABLE_WRITES, etc.).
"""

import pytest

from msaccess_vcs_mcp.config import _strip_quotes


class TestStripQuotes:
    """Tests for the _strip_quotes helper used on path-valued env vars."""

    def test_unquoted_passthrough(self):
        assert _strip_quotes(r"C:\Repos\db.accdb") == r"C:\Repos\db.accdb"

    def test_double_quoted(self):
        assert _strip_quotes('"C:\\Repos\\db.accdb"') == "C:\\Repos\\db.accdb"

    def test_single_quoted(self):
        assert _strip_quotes("'C:\\Repos\\db.accdb'") == "C:\\Repos\\db.accdb"

    def test_path_with_spaces(self):
        assert (
            _strip_quotes('"C:\\Repos\\My Database.accdb"')
            == "C:\\Repos\\My Database.accdb"
        )

    def test_empty_string(self):
        assert _strip_quotes("") == ""

    def test_mismatched_quotes_untouched(self):
        assert _strip_quotes("\"C:\\Repos\\db.accdb'") == "\"C:\\Repos\\db.accdb'"

    def test_single_char_untouched(self):
        assert _strip_quotes('"') == '"'

    def test_empty_quoted_string(self):
        assert _strip_quotes('""') == ""
        assert _strip_quotes("''") == ""


def test_validate_access_installation():
    """Test Access COM validation."""
    from msaccess_vcs_mcp.config import validate_access_installation
    
    # This test will only pass if Access is installed
    # On CI/CD without Access, this would be skipped
    try:
        validate_access_installation()
    except (ImportError, RuntimeError) as e:
        pytest.skip(f"Access not available: {e}")
