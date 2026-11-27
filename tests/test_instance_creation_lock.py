"""Two gate lanes starting Access at once never share an instance."""

import itertools
import threading
import time
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod
from msaccess_vcs_mcp.access_com.connection import AccessConnection

_pids = itertools.count(1000)


class _FakeApp:
    def __init__(self, world):
        self.pid = next(_pids)
        self.opened: list[str] = []
        self.Visible = False
        self.UserControl = False
        world.append(self)

    def CurrentDb(self):
        if not self.opened:
            return None
        db = MagicMock()
        db.Name = self.opened[-1]
        return db

    def OpenCurrentDatabase(self, path):
        # Slow enough that a second lane would reach EnsureDispatch meanwhile.
        time.sleep(0.2)
        self.opened.append(path)


@pytest.fixture
def world():
    apps: list[_FakeApp] = []

    def ensure_dispatch(_prog_id):
        # Like Access: the most recently started instance answers EnsureDispatch.
        return apps[-1] if apps else _FakeApp(apps)

    with (
        patch.object(conn_mod, "COM_AVAILABLE", True),
        patch.object(conn_mod, "win32com") as win32com,
        patch.object(conn_mod, "ensure_dispatch", side_effect=ensure_dispatch),
        patch.object(conn_mod, "access_instance_is_live", return_value=False),
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.list_access_pids_or_none",
            side_effect=lambda: {a.pid for a in apps},
        ),
        patch(
            "msaccess_vcs_mcp.access_com.process_qos.pid_from_access_app",
            side_effect=lambda app: app.pid,
        ),
        patch("msaccess_vcs_mcp.access_com.process_qos.prefer_full_power_app"),
        patch("msaccess_vcs_mcp.access_com.instance_registry.register_owned"),
        patch("msaccess_vcs_mcp.access_com.instance_registry.process_create_time", return_value=1.0),
        patch("msaccess_vcs_mcp.access_com.instance_registry.is_owned", return_value=True),
    ):
        win32com.client.GetObject.side_effect = Exception("no moniker")
        win32com.client.DispatchEx.side_effect = lambda _prog_id: _FakeApp(apps)
        yield apps


def test_simultaneous_first_calls_get_separate_instances(world):
    databases = [r"C:\wt-a\App.accdb", r"C:\wt-b\App.accdb", r"C:\wt-c\App.accdb"]
    got: dict[str, _FakeApp] = {}

    def connect(path):
        got[path] = AccessConnection(path)._get_access_app()

    threads = [threading.Thread(target=connect, args=(db,)) for db in databases]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(5)

    assert len({id(app) for app in got.values()}) == 3
    for path, app in got.items():
        assert app.opened == [path]


def test_creation_waits_are_bounded(world, monkeypatch):
    monkeypatch.setattr(conn_mod, "INSTANCE_CREATION_TIMEOUT_SEC", 0.05)
    assert conn_mod._INSTANCE_CREATION_LOCK.acquire(timeout=1)
    try:
        with pytest.raises(RuntimeError, match="still starting"):
            AccessConnection(r"C:\wt-a\App.accdb")._get_access_app()
    finally:
        conn_mod._INSTANCE_CREATION_LOCK.release()
