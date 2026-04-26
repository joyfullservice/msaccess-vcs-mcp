"""Tests for object exporters.

Note: this project delegates all export/import operations to the
MSAccess VCS add-in via COM (see ``addin_integration.py``); there is
no in-process ``msaccess_vcs_mcp.exporters`` module to unit-test.
The placeholders below mark where future integration tests against a
live Access database would live.
"""

import pytest


@pytest.mark.integration
def test_query_exporter():
    """Test query exporter (requires Access)."""
    pytest.skip("Integration test - requires Access database")


@pytest.mark.integration
def test_module_exporter():
    """Test module exporter (requires Access)."""
    pytest.skip("Integration test - requires Access database")
