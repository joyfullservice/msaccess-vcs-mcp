"""Closing a server-created instance quits it on a worker thread it can reach."""

import threading
from unittest.mock import Mock

import pytest

from msaccess_vcs_mcp.access_com import connection as conn_mod

pytestmark = pytest.mark.skipif(not conn_mod.COM_AVAILABLE, reason="pywin32 required")


def test_quit_uses_a_proxy_marshalled_to_the_worker(monkeypatch):
    caller = threading.get_ident()
    seen: dict[str, int] = {}
    stream = object()
    unmarshalled = object()
    worker_app = Mock()
    worker_app.Quit.side_effect = lambda: seen.setdefault("quit", threading.get_ident())

    def marshal(_iid, _ole):
        seen["marshal"] = threading.get_ident()
        return stream

    def unmarshal(got, _iid):
        assert got is stream
        seen["unmarshal"] = threading.get_ident()
        return unmarshalled

    monkeypatch.setattr(conn_mod.pythoncom, "CoMarshalInterThreadInterfaceInStream", marshal)
    monkeypatch.setattr(conn_mod.pythoncom, "CoGetInterfaceAndReleaseStream", unmarshal)
    monkeypatch.setattr(
        conn_mod.win32com.client, "Dispatch",
        lambda obj: worker_app if obj is unmarshalled else pytest.fail("wrong object"),
    )
    caller_app = Mock()

    assert conn_mod._quit_with_timeout(caller_app, 5) is True

    assert seen["marshal"] == caller
    assert seen["unmarshal"] != caller
    assert seen["quit"] == seen["unmarshal"]
    worker_app.CloseCurrentDatabase.assert_called_once()
    caller_app.Quit.assert_not_called()


def test_unmarshallable_object_is_quit_directly(monkeypatch):
    def refuse(_iid, _ole):
        raise TypeError("not a COM object")

    monkeypatch.setattr(conn_mod.pythoncom, "CoMarshalInterThreadInterfaceInStream", refuse)
    app = Mock()

    assert conn_mod._quit_with_timeout(app, 5) is True
    app.Quit.assert_called_once()


def test_quit_that_raises_reports_failure(monkeypatch):
    monkeypatch.setattr(
        conn_mod.pythoncom, "CoMarshalInterThreadInterfaceInStream",
        Mock(side_effect=TypeError("double")),
    )
    app = Mock()
    app.Quit.side_effect = RuntimeError("disconnected")

    assert conn_mod._quit_with_timeout(app, 5) is False
