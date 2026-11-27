"""Ownership persistence, rebuild pre-flight, and session-cleanup guards."""

from __future__ import annotations

import asyncio
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod
from msaccess_vcs_mcp.access_com import instance_registry as registry
from msaccess_vcs_mcp.access_com.connection import (
    AccessConnection,
    access_instance_is_live,
    close_owned_instances_holding,
    recycle_owned_instance,
)
import msaccess_vcs_mcp.main as main_module
import msaccess_vcs_mcp.tools as tools_module


DB = r"C:\Projects\other\Some Database.accdb"


@pytest.fixture(autouse=True)
def _isolate_registry(tmp_path, monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(tmp_path / "owned.json"))
    monkeypatch.setattr(registry, "list_access_pids_or_none", lambda: {4242})
    monkeypatch.setattr(
        registry, "process_create_time", lambda pid: 111 if pid == 4242 else None
    )


def _process_table(monkeypatch, *, alive, terminated=None):
    """Script process liveness for the closure path.

    ``alive`` is consulted after each close attempt, so a test can make a
    process survive its Quit and then disappear on terminate.
    """
    monkeypatch.setattr(
        "msaccess_vcs_mcp.access_com.process_qos.process_is_alive",
        lambda pid: alive(pid),
    )
    if terminated is not None:
        monkeypatch.setattr(
            "msaccess_vcs_mcp.access_com.process_qos.terminate_pid", terminated
        )


@pytest.fixture(autouse=True)
def _com_available():
    with patch.object(conn_mod, "COM_AVAILABLE", True):
        yield


def _app_with_current_db(path, hwnd=1):
    app = MagicMock()
    current = MagicMock()
    current.Name = path
    app.CurrentDb.return_value = current
    app.hWndAccessApp.return_value = hwnd
    return app


def test_reattach_to_registered_pid_is_still_owned(monkeypatch):
    registry.register_owned(4242, DB, create_time=111)
    app = _app_with_current_db(DB)
    monkeypatch.setattr(conn_mod, "access_instance_is_live", lambda _p: True)

    with (
        patch.object(conn_mod, "win32com") as mock_win32com,
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            return_value=4242,
        ),
    ):
        mock_win32com.client.GetObject.return_value = app
        first = AccessConnection(DB)
        first._get_access_app()
        assert first._owns_app is True
        first.close()

        second = AccessConnection(DB)
        second._get_access_app()
        assert second._owns_app is True
        second.close()
        app.Quit.assert_not_called()


def test_unregistered_getobject_attach_is_not_owned(monkeypatch):
    app = _app_with_current_db(DB)
    monkeypatch.setattr(conn_mod, "access_instance_is_live", lambda _p: True)

    with (
        patch.object(conn_mod, "win32com") as mock_win32com,
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            return_value=4242,
        ),
    ):
        mock_win32com.client.GetObject.return_value = app
        conn = AccessConnection(DB)
        conn._get_access_app()

    assert conn._owns_app is False


def test_created_instance_is_registered(monkeypatch):
    app = MagicMock()
    app.CurrentDb.return_value = None
    app.hWndAccessApp.return_value = 1

    with (
        patch.object(conn_mod, "win32com") as mock_win32com,
        patch.object(conn_mod, "gencache") as mock_gencache,
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.list_access_pids_or_none",
            return_value=set(),
        ),
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            return_value=4242,
        ),
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.prefer_full_power_app",
        ),
    ):
        mock_win32com.client.GetObject.side_effect = Exception("no moniker")
        mock_gencache.EnsureDispatch.return_value = app
        conn = AccessConnection(DB)
        conn._get_access_app()

    assert conn._owns_app is True
    assert registry.is_owned(4242, 111) is True


def test_getobject_that_launches_access_is_registered(monkeypatch):
    """A moniker bind starts Access when nothing holds the file.

    That process is the server's, and without registering it here the
    persistent-instance feature would never recognize its own windows.
    """
    app = _app_with_current_db(DB)
    monkeypatch.setattr(conn_mod, "access_instance_is_live", lambda _p: False)

    with (
        patch.object(conn_mod, "win32com") as mock_win32com,
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            return_value=4242,
        ),
        patch("msaccess_vcs_mcp.access_com.process_qos.prefer_full_power_app"),
    ):
        mock_win32com.client.GetObject.return_value = app
        conn = AccessConnection(DB)
        conn._get_access_app()

    assert conn._owns_app is True
    assert registry.is_owned(4242, 111) is True


def test_users_empty_access_shell_is_not_claimed(monkeypatch):
    """EnsureDispatch can hand back a user's Access with no database open.

    Having no current database does not make the window ours; only a PID
    that appeared after the dispatch does.
    """
    app = MagicMock()
    app.CurrentDb.return_value = None
    app.hWndAccessApp.return_value = 1

    with (
        patch.object(conn_mod, "win32com") as mock_win32com,
        patch.object(conn_mod, "gencache") as mock_gencache,
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.list_access_pids_or_none",
            return_value={4242},
        ),
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            return_value=4242,
        ),
    ):
        mock_win32com.client.GetObject.side_effect = Exception("no moniker")
        mock_gencache.EnsureDispatch.return_value = app
        conn = AccessConnection(DB)
        conn._get_access_app()

    assert conn._owns_app is False


def _owned_app(monkeypatch, pid=4242, create_time=111):
    """Wire up a resolvable COM instance whose identity matches the record."""
    app = _app_with_current_db(DB)
    monkeypatch.setattr(
        "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
        lambda _a: pid,
    )
    monkeypatch.setattr(
        conn_mod, "_resolve_instance_for_record", lambda record: app
    )
    monkeypatch.setattr(registry, "process_create_time", lambda _p: create_time)
    return app


def test_close_owned_instances_holding_quits_only_owned(monkeypatch):
    registry.register_owned(4242, DB, create_time=111)
    app = _owned_app(monkeypatch)
    _process_table(monkeypatch, alive=lambda _pid: False)

    closed = close_owned_instances_holding([DB])

    assert len(closed) == 1
    assert closed[0]["pid"] == 4242
    assert closed[0]["terminated"] is False
    app.Quit.assert_called_once()
    assert registry.is_owned(4242, 111) is False


def test_close_skips_instance_whose_identity_does_not_match(monkeypatch):
    """The path resolves, but to a different process than the record.

    Two Access processes can hold the same file, and a moniker bind can
    start a third. Closing on a path match alone would kill a user window.
    """
    registry.register_owned(4242, DB, create_time=111)
    other = _app_with_current_db(DB)
    monkeypatch.setattr(
        "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
        lambda _a: 9999,
    )
    monkeypatch.setattr(conn_mod, "win32com", MagicMock())
    conn_mod.win32com.client.GetObject.return_value = other
    _process_table(monkeypatch, alive=lambda _pid: True)

    closed = close_owned_instances_holding([DB])

    assert closed == []
    other.Quit.assert_not_called()
    assert registry.is_owned(4242, 111) is True


def test_hung_instance_is_terminated_after_quit_times_out(monkeypatch):
    """A server-created window that ignores Quit is force-terminated.

    It holds a file a rebuild must replace, and unsaved state in an
    instance the server started is acceptable loss.
    """
    registry.register_owned(4242, DB, create_time=111)
    _owned_app(monkeypatch)
    monkeypatch.setattr(conn_mod, "_quit_with_timeout", lambda _app, _t: False)

    states = iter([True, False, False])
    terminated: list[int] = []

    def _terminate(pid):
        terminated.append(pid)
        return True

    _process_table(
        monkeypatch, alive=lambda _pid: next(states, False), terminated=_terminate
    )

    closed = close_owned_instances_holding([DB])

    assert terminated == [4242]
    assert closed[0]["terminated"] is True
    assert registry.is_owned(4242, 111) is False


def test_surviving_instance_keeps_its_registry_record(monkeypatch):
    """If neither Quit nor terminate worked, the claim has to stay.

    Forgetting a live owned process makes it permanently uncloseable.
    """
    registry.register_owned(4242, DB, create_time=111)
    _owned_app(monkeypatch)
    monkeypatch.setattr(conn_mod, "_quit_with_timeout", lambda _app, _t: False)
    _process_table(
        monkeypatch, alive=lambda _pid: True, terminated=lambda _pid: False
    )

    assert close_owned_instances_holding([DB]) == []
    assert registry.is_owned(4242, 111) is True


def test_close_owned_leaves_unregistered_paths_alone(monkeypatch):
    resolved = MagicMock()
    monkeypatch.setattr(conn_mod, "_resolve_instance_for_record", resolved)

    closed = close_owned_instances_holding([DB])

    assert closed == []
    resolved.assert_not_called()


def test_recycle_owned_instance_respawns(monkeypatch):
    registry.register_owned(4242, DB, create_time=111)
    _owned_app(monkeypatch)
    _process_table(monkeypatch, alive=lambda _pid: False)
    opened = []

    class FakeConn:
        def __init__(self, path):
            opened.append(path)

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def connect(self):
            return None, None

    monkeypatch.setattr(conn_mod, "AccessConnection", FakeConn)

    assert recycle_owned_instance(DB) is True
    assert opened == [DB]


def test_recycle_skips_user_owned():
    assert recycle_owned_instance(DB) is False


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


def test_rebuild_addin_closes_owned_holders(tmp_path, monkeypatch):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    host = tmp_path / "Version Control.accda"
    host.write_bytes(b"x")
    installed = r"C:\Users\me\AppData\Roaming\MSAccessVCS\Version Control.accda"
    captured: list[tuple[list[str], list[str]]] = []
    order: list[str] = []

    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": installed},
    )
    monkeypatch.setattr(tools_module, "get_callback_url", lambda: None)
    monkeypatch.setattr(tools_module, "_is_installed_addin_path", lambda _p: False)
    def _close(paths, addin_paths=()):
        order.append("close")
        captured.append((list(paths), list(addin_paths)))
        return []

    monkeypatch.setattr(tools_module, "close_owned_instances_holding", _close)

    def _call_vba(*_a, **_k):
        order.append("launch")
        return {
            "success": True,
            "result": '{"success": false, "status": "refused", '
            '"phaseStarted": "2026-09-10 12:00:00", "error": "user"}',
        }

    monkeypatch.setattr(tools_module, "_execute_call_vba", _call_vba)

    gate = MagicMock()

    async def _exclusive(_tool, _db, fn, _is_async, /, *args, **kwargs):
        order.append("gate")
        return await fn(*args, **kwargs)

    gate.run_exclusive_all = _exclusive
    monkeypatch.setattr(tools_module, "get_access_gate", lambda: gate)

    result = asyncio.run(_unwrap(tools_module.vcs_rebuild_addin)(str(source)))

    assert result["status"] == "refused"
    # The development copy is matched as an open database; the installed
    # add-in is matched as a loaded library, since every tool call locks
    # it regardless of which database that instance has open.
    assert captured == [([str(host)], [installed])]
    # Closing must happen inside the gate, or another window's tool call
    # can reopen the file between the close and the launch.
    assert order == ["gate", "close", "launch"]


def test_rebuild_database_closes_owned_output(tmp_path, monkeypatch):
    source = tmp_path / "src"
    source.mkdir()
    output = str(tmp_path / "out.accdb")
    captured: list[list[str]] = []

    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(
        tools_module,
        "close_owned_instances_holding",
        lambda paths, *_a: captured.append(list(paths)) or [],
    )
    monkeypatch.setattr(tools_module, "_check_database_busy", lambda _p: None)
    monkeypatch.setattr(
        tools_module,
        "_create_build_instance",
        lambda: (_ for _ in ()).throw(RuntimeError("stop after pre-flight")),
    )

    result = asyncio.run(
        _unwrap(tools_module.vcs_rebuild_database)(str(source), output)
    )

    assert captured == [[output]]
    assert result["success"] is False
    assert "stop after pre-flight" in result["error"]


def test_session_cleanup_skips_when_access_is_not_live(monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_SESSION_ID", "abcd")
    monkeypatch.setenv("ACCESS_VCS_DATABASE", DB)
    monkeypatch.delenv("ACCESS_VCS_SKIP_SESSION_CLEANUP", raising=False)

    opened = []

    class FakeConn:
        def __init__(self, path):
            opened.append(path)

        def __enter__(self):
            raise AssertionError("must not open Access when none is live")

        def __exit__(self, *exc):
            return False

    monkeypatch.setattr(
        "msaccess_vcs_mcp.access_com.connection.access_instance_is_live",
        lambda _p: False,
    )
    monkeypatch.setattr(
        "msaccess_vcs_mcp.access_com.connection.AccessConnection",
        FakeConn,
    )

    main_module._cleanup_session()
    assert opened == []


def test_end_session_does_not_launch_access(monkeypatch):
    """Ending a session must not start the thing it is winding down.

    Shutdown calls this unconditionally, and opening Access here would
    leave a window behind after every server exit.
    """
    def _fail(_path):
        raise AssertionError("must not open Access when none is live")

    monkeypatch.setattr(tools_module, "access_instance_is_live", lambda _p: False)
    monkeypatch.setattr(tools_module, "AccessConnection", _fail)
    monkeypatch.setattr(tools_module, "validate_database_path", lambda p: p)
    monkeypatch.setattr(tools_module, "get_session_id", lambda: "abcd")

    result = _unwrap(tools_module.vcs_end_session)(DB)

    assert result["success"] is True
    assert result["session_id"] == "abcd"


def test_access_instance_is_live_uses_rot(monkeypatch):
    monkeypatch.setattr(
        "msaccess_vcs_mcp.addin_integration.VCSAddinIntegration._find_access_in_rot",
        lambda path: object() if path == DB else None,
    )
    assert access_instance_is_live(DB) is True
    assert access_instance_is_live(r"C:\other.accdb") is False
