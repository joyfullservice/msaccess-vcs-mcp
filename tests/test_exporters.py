"""Tests for object exporters."""

import pytest
from pathlib import Path
import tempfile


@pytest.mark.integration
def test_query_exporter():
    """Test query exporter (requires Access)."""
    pytest.skip("Integration test - requires Access database")


@pytest.mark.integration
def test_module_exporter():
    """Test module exporter (requires Access)."""
    pytest.skip("Integration test - requires Access database")


def test_sanitize_filename():
    """Test filename sanitization."""
    from msaccess_vcs_mcp.exporters.query_exporter import QueryExporter
    
    exporter = QueryExporter()
    
    # Test various invalid characters
    assert exporter._sanitize_filename("test:query") == "test_query"
    assert exporter._sanitize_filename("test/query") == "test_query"
    assert exporter._sanitize_filename("test<>query") == "test__query"
