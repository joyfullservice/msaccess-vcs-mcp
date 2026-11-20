"""Integration tests for full workflows."""

import asyncio

import pytest


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


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
def test_rebuild_from_source(tmp_path, monkeypatch):
    """
    Export a one-table database, then rebuild it to a new path unattended.
    
    A dialog anywhere in the build would hang this test rather than fail it,
    which is the regression it exists to catch.
    
    This test requires:
    - Microsoft Access installed
    - VCS add-in 5.1.1 or later installed
    """
    try:
        import win32com.client
    except ImportError as e:
        pytest.skip(f"pywin32 not available: {e}")

    from msaccess_vcs_mcp.tools import vcs_export_database, vcs_rebuild_database

    monkeypatch.setenv("ACCESS_VCS_LEAVE_ACCESS_OPEN", "false")
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))

    db_path = tmp_path / "Sample.accdb"
    try:
        app = win32com.client.DispatchEx("Access.Application")
    except Exception as e:
        pytest.skip(f"Access not available: {e}")
    try:
        app.NewCurrentDatabase(str(db_path))
        app.CurrentDb().Execute(
            "CREATE TABLE tblExample (ID COUNTER PRIMARY KEY, Label TEXT(50))"
        )
        app.CloseCurrentDatabase()
    finally:
        app.Quit()

    src_path = tmp_path / "Sample.accdb.src"
    export = asyncio.run(
        _unwrap(vcs_export_database)(str(db_path), str(src_path), full_export=True)
    )
    assert export["success"], export

    output = tmp_path / "out" / "Rebuilt.accdb"
    output.parent.mkdir()
    result = asyncio.run(_unwrap(vcs_rebuild_database)(str(src_path), str(output)))
    if result.get("error_pattern") == "addin_version":
        pytest.skip(result["error"])

    assert result["success"], result
    assert result["output_path"] == str(output)
    assert output.exists()
    assert result["errorCount"] == 0
