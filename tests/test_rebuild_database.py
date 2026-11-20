"""Tests for vcs_rebuild_database on the add-in's BuildHeadless entry point.

The tool must never reach an Access dialog: not BuildAs's file pickers, not an
argument error an older add-in would raise inside the library, and not the crash
Access has when a library is called with no database open.
"""

import asyncio
import json
import os
import time
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp import tools as tools_module
from msaccess_vcs_mcp.tools import _build_headless_refusal, _parse_addin_version


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


class FakeAccess:
    """Stands in for Access.Application; NewCurrentDatabase creates the file."""

    def __init__(self):
        self.opened: list[str] = []
        self.quit_called = False

    def NewCurrentDatabase(self, path):
        with open(path, "wb") as handle:
            handle.write(b"host")
        self.opened.append(path)

    def Quit(self):
        self.quit_called = True


@contextmanager
def _patch_rebuild(tmp_path, *, version="5.1.1", run=None):
    """Patch COM so vcs_rebuild_database runs offline against a fake worker app."""
    app = FakeAccess()
    worker_app = MagicMock()
    worker_app.Run = MagicMock(side_effect=run)

    addin_path = str(tmp_path / "Version Control.accda")
    mock_addin = MagicMock()
    mock_addin.addin_path = addin_path
    mock_addin.call_sync.return_value = version
    addin_class = MagicMock(return_value=mock_addin)
    addin_class.worker_app_source.return_value = lambda: worker_app

    src_path = tmp_path / "src"
    src_path.mkdir(exist_ok=True)

    with (
        patch.object(tools_module.tempfile, "tempdir", str(tmp_path)),
        patch.object(tools_module, "VCSAddinIntegration", addin_class),
        patch.object(tools_module, "_create_build_instance", return_value=app),
        patch.object(tools_module, "ensure_access_visible", return_value=None),
        patch.object(tools_module, "check_write_permission", return_value=None),
        patch.object(tools_module, "close_owned_instances_holding", return_value=[]),
        patch.object(tools_module, "_check_database_busy", return_value=None),
        patch.object(
            tools_module, "get_config", return_value={"ACCESS_VCS_ADDIN_PATH": addin_path}
        ),
    ):
        yield app, worker_app, src_path, os.path.splitext(addin_path)[0] + ".API"


def _build_json(**fields):
    payload = {
        "success": True,
        "errorCount": 0,
        "warningCount": 1,
        "durationMs": 1200,
    }
    payload.update(fields)
    return json.dumps(payload)


def test_builds_to_output_path_with_build_headless(tmp_path):
    output = str(tmp_path / "out" / "Rebuilt.accdb")
    log = str(tmp_path / "src" / "logs" / "Build_1.log")

    with _patch_rebuild(
        tmp_path, run=lambda *a: (_build_json(databasePath=output, logPath=log),)
    ) as (app, worker_app, src_path, api_name):
        result = asyncio.run(
            _unwrap(tools_module.vcs_rebuild_database)(str(src_path), output)
        )

    worker_app.Run.assert_called_once_with(api_name, "BuildHeadless", str(src_path), output)
    assert result["success"] is True
    assert result["output_path"] == output
    assert result["log_path"] == log
    assert result["errorCount"] == 0
    # A database was open for the call, and nothing of it is left behind.
    assert len(app.opened) == 1
    assert not os.path.exists(os.path.dirname(app.opened[0]))
    assert app.quit_called


def test_without_output_path_sends_one_argument(tmp_path):
    """The one-argument form is what every add-in with BuildHeadless accepts."""
    built = str(tmp_path / "MyApp.accdb")

    with _patch_rebuild(
        tmp_path, version="5.1.0", run=lambda *a: _build_json(databasePath=built)
    ) as (_app, worker_app, src_path, api_name):
        result = asyncio.run(_unwrap(tools_module.vcs_rebuild_database)(str(src_path)))

    worker_app.Run.assert_called_once_with(api_name, "BuildHeadless", str(src_path))
    assert result["success"] is True
    assert result["output_path"] == built


def test_output_path_refused_on_older_addin(tmp_path):
    """An extra argument would stop an older add-in on a modal error dialog."""
    with _patch_rebuild(tmp_path, version="5.1.0") as (app, worker_app, src_path, _):
        result = asyncio.run(
            _unwrap(tools_module.vcs_rebuild_database)(
                str(src_path), str(tmp_path / "out.accdb")
            )
        )

    worker_app.Run.assert_not_called()
    assert result["success"] is False
    assert result["error_pattern"] == "addin_version"
    assert "5.1.1" in result["error"]
    assert app.quit_called
    assert not os.path.exists(os.path.dirname(app.opened[0]))


def test_failed_build_reports_add_in_error(tmp_path):
    with _patch_rebuild(
        tmp_path,
        run=lambda *a: json.dumps(
            {"success": False, "error": "Source files not found in folder: x"}
        ),
    ) as (_app, _worker_app, src_path, _):
        result = asyncio.run(_unwrap(tools_module.vcs_rebuild_database)(str(src_path)))

    assert result["success"] is False
    assert "Source files not found" in result["error"]
    assert result["output_path"] is None


def test_timeout_leaves_the_running_build_alone(tmp_path, monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_BUILD_TIMEOUT_SEC", "0.3")

    with _patch_rebuild(tmp_path, run=lambda *a: time.sleep(3)) as (
        app,
        _worker_app,
        src_path,
        _,
    ):
        result = asyncio.run(_unwrap(tools_module.vcs_rebuild_database)(str(src_path)))

    assert result["success"] is False
    assert result["timed_out"] is True
    assert result["error_pattern"] == "timeout"
    # Quitting would block until the build ends; the host may still be open.
    assert not app.quit_called


def test_host_folder_removed_when_host_cannot_open(tmp_path, monkeypatch):
    monkeypatch.setattr(tools_module.tempfile, "tempdir", str(tmp_path))
    app = MagicMock()
    app.NewCurrentDatabase.side_effect = RuntimeError("You already have the database open.")

    with pytest.raises(RuntimeError):
        tools_module._open_build_host(app)

    assert list(tmp_path.iterdir()) == []


def test_host_removal_retries_while_access_holds_the_file(tmp_path):
    folder = tmp_path / "vcs-build-host-x"
    folder.mkdir()
    host = folder / "BuildHost.accdb"
    host.write_bytes(b"host")
    real_rmtree = tools_module.shutil.rmtree
    attempts = []

    def rmtree(path, ignore_errors=False):
        attempts.append(path)
        if len(attempts) > 1:
            real_rmtree(path, ignore_errors=ignore_errors)

    with patch.object(tools_module.shutil, "rmtree", side_effect=rmtree):
        asyncio.run(tools_module._remove_build_host(str(host)))

    assert len(attempts) == 2
    assert not folder.exists()


def test_build_instance_is_a_new_process():
    """The build quits its instance, so it must never attach to the user's window."""
    import win32com.client

    app = MagicMock()
    with (
        patch.object(win32com.client, "DispatchEx", return_value=app) as dispatch_ex,
        patch.object(win32com.client, "Dispatch") as dispatch,
        patch.object(tools_module, "prefer_full_power_app") as full_power,
    ):
        assert tools_module._create_build_instance() is app

    dispatch_ex.assert_called_once_with("Access.Application")
    dispatch.assert_not_called()
    full_power.assert_called_once_with(app)


@pytest.mark.parametrize(
    ("text", "expected"),
    [("5.1.1", (5, 1, 1)), (" 5.1.0 ", (5, 1, 0)), ("", None), (None, None), ("dev", None)],
)
def test_parse_addin_version(text, expected):
    assert _parse_addin_version(text) == expected


@pytest.mark.parametrize(
    ("version", "output_path", "refused"),
    [
        ("5.1.1", "C:\\out.accdb", False),
        ("5.2.0", "C:\\out.accdb", False),
        ("5.1.0", "C:\\out.accdb", True),
        ("5.1.0", None, False),
        ("5.0.2", None, True),
        ("unknown", None, True),
    ],
)
def test_build_headless_refusal(version, output_path, refused):
    assert (_build_headless_refusal(version, output_path) is not None) is refused
