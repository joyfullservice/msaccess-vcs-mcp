"""Integration tests for full workflows."""

import pytest


@pytest.mark.integration
def test_full_export_import_cycle():
    """
    Test complete export → modify → import workflow.
    
    This test requires:
    - Microsoft Access installed
    - Sample Access database
    """
    pytest.skip("Integration test - requires Access database and setup")


@pytest.mark.integration
def test_rebuild_from_source():
    """
    Test rebuilding database from source files.
    
    This test requires:
    - Microsoft Access installed
    - Sample source files
    """
    pytest.skip("Integration test - requires Access database and setup")
