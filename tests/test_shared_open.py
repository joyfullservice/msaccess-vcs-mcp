"""Imports warn, without refusing, when another session holds the database."""

import os

import pytest

from msaccess_vcs_mcp import shared_open as so


def _lock_bytes(*entries):
    data = b""
    for machine, user in entries:
        data += machine.encode().ljust(32, b"\x00") + user.encode().ljust(32, b"\x00")
    return data


def _warnings(tmp_path, holders=frozenset(), access=frozenset(), driven=10,
              lock=None, local=frozenset({"THISPC"})):
    db = tmp_path / "db.accdb"
    db.write_bytes(b"")
    if lock is not None:
        (tmp_path / "db.laccdb").write_bytes(lock)
    return so.shared_open_warnings(
        str(db), object(),
        holders=lambda path: set(holders),
        access_pids=lambda: set(access),
        driven_pid=lambda app: driven,
        local_names=lambda: set(local),
    )


def test_only_driven_instance_holds_it(tmp_path):
    assert _warnings(tmp_path, holders={10}, access={10}) == []


def test_another_access_session(tmp_path):
    [warning] = _warnings(tmp_path, holders={10, 22}, access={10, 22})
    assert warning.startswith("Another Access session (PID 22) has this database open;")
    assert "may not save" in warning and "vcs_diff_database" in warning


def test_other_process_and_own_process(tmp_path):
    warnings = _warnings(tmp_path, holders={10, 33, os.getpid()}, access={10})
    assert warnings == [
        f"Another process (PID 33) has this database open; {so.CONSEQUENCE}. {so.ACTION}"
    ]


def test_lock_file_other_machine_may_be_stale(tmp_path):
    lock = _lock_bytes(("THISPC", "Admin"), ("OTHERPC", "Admin"), ("OTHERPC", "Admin"))
    [warning] = _warnings(tmp_path, holders={10}, lock=lock)
    assert "computer OTHERPC" in warning and "may be stale" in warning


def test_lock_file_local_machine_only(tmp_path):
    lock = _lock_bytes(("thispc", "Admin"))
    assert _warnings(tmp_path, holders={10}, lock=lock) == []


def test_failed_holder_query_still_reads_lock_file(tmp_path):
    db = tmp_path / "db.accdb"
    (tmp_path / "db.laccdb").write_bytes(_lock_bytes(("OTHERPC", "Admin")))

    def broken(path):
        raise OSError("restart manager unavailable")

    warnings = so.shared_open_warnings(
        str(db), object(), holders=broken, access_pids=set, driven_pid=lambda app: 10,
        local_names=lambda: {"THISPC"},
    )
    assert len(warnings) == 1 and "OTHERPC" in warnings[0]


@pytest.mark.parametrize("name, lock", [
    ("db.accdb", "db.laccdb"), ("db.ACCDA", "db.laccdb"), ("db.mdb", "db.ldb"), ("db.txt", None),
])
def test_lock_file_for(name, lock):
    result = so.lock_file_for(f"C:\\data\\{name}")
    assert (result.name if result else None) == lock


def test_warnings_go_before_detail_and_existing_warnings():
    result = so.with_warnings_first(
        {"success": True, "objects": {"a": 1}, "warnings": ["old"]}, ["new"]
    )
    assert list(result) == ["success", "warnings", "objects"]
    assert result["warnings"] == ["new", "old"]


def test_warnings_after_error_pattern():
    result = so.with_warnings_first(
        {"success": False, "error": "x", "error_pattern": "p", "log_path": "l"}, ["w"]
    )
    assert list(result) == ["success", "error", "error_pattern", "warnings", "log_path"]


def test_no_warnings_leaves_result_alone():
    original = {"success": True}
    assert so.with_warnings_first(original, []) is original


def test_import_object_result_carries_the_warning(monkeypatch, tmp_path):
    import json
    from pathlib import Path
    from unittest.mock import MagicMock

    from msaccess_vcs_mcp import tools

    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path))
    conn = MagicMock()
    conn.__enter__.return_value = conn
    conn.connect.return_value = (MagicMock(), MagicMock())
    addin = MagicMock()
    addin.call_sync.return_value = json.dumps({"success": True, "objectName": "qryExample"})
    monkeypatch.setattr(tools, "validate_database_path", lambda p: Path(p))
    monkeypatch.setattr(tools, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": "x"})
    monkeypatch.setattr(tools, "AccessConnection", lambda path: conn)
    monkeypatch.setattr(tools, "VCSAddinIntegration", lambda path: addin)
    monkeypatch.setattr(tools, "shared_open_warnings", lambda path, app: ["held elsewhere"])

    fn = tools.vcs_import_object
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    result = fn("C:\\data\\db.accdb", "query", "qryExample")

    assert result["success"] is True
    assert result["warnings"] == ["held elsewhere"]
    assert list(result).index("warnings") == 1
