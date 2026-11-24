"""Timeouts say what Access is waiting on, from window calls only."""

import asyncio
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import access_state as ast
from msaccess_vcs_mcp import screenshot


def tool_with_db(database_path, code=""):
    pass


def tool_with_source(source_dir, timeout_seconds=None):
    pass


def tool_without_db(limit=10):
    pass


DIALOG = {"dialog_id": "abc123", "title": "Microsoft Access", "text": "Save changes?",
          "buttons": ["Yes", "No", "Cancel"]}


def _state(**overrides):
    state = {"pid": 42, "owned_by_server": True, "vbe_state": "design",
             "modal_blocked": True, "dialogs": [dict(DIALOG)]}
    state.update(overrides)
    return state


def _attach(tool, func, args, kwargs, result, reader):
    return asyncio.run(ast.attach_access_state(tool, func, args, kwargs, result, reader=reader))


TIMEOUT = {"success": False, "error": "timed out", "error_pattern": "timeout", "recoverable": True}


def test_target_database_from_database_path():
    assert ast.target_database("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}) == "C:\\db.accdb"


@pytest.mark.parametrize("source", [
    "C:\\Repos\\addin\\Version Control.accda.src",
    "C:\\Repos\\addin\\Version Control.accda.src\\",
])
def test_target_database_for_rebuild_addin_is_development_copy(source):
    assert ast.target_database("vcs_rebuild_addin", tool_with_source, (source,), {}) == (
        "C:\\Repos\\addin\\Version Control.accda"
    )


def test_target_database_unknown():
    assert ast.target_database("vcs_get_recent_calls", tool_without_db, (), {}) is None
    assert ast.target_database("vcs_other", tool_with_source, ("C:\\src",), {}) is None


def test_attaches_state_and_owned_dialog_guidance():
    result = _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT),
                     lambda db: _state())
    assert list(result)[:4] == ["success", "error", "error_pattern", "agent_guidance"]
    assert 'vcs_dismiss_dialog(dialog_id="abc123"' in result["agent_guidance"]
    assert '"Yes", "No", "Cancel"' in result["agent_guidance"]
    assert result["access_state"]["dialogs"][0]["text"] == "Save changes?"
    assert result["error"] == "timed out"


def test_user_owned_dialog_says_ask_the_user():
    result = _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT),
                     lambda db: _state(owned_by_server=False))
    assert "Ask the user" in result["agent_guidance"]
    assert "vcs_dismiss_dialog" not in result["agent_guidance"]


def test_break_mode_guidance():
    result = _attach("vcs_call_vba", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT),
                     lambda db: _state(dialogs=[], vbe_state="break"))
    assert "vcs_reset_vba()" in result["agent_guidance"]


def test_no_dialog_keeps_existing_guidance():
    original = {**TIMEOUT, "agent_guidance": "worker guidance"}
    result = _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, original,
                     lambda db: _state(dialogs=[], vbe_state="running"))
    assert result["agent_guidance"] == "worker guidance"
    assert result["access_state"]["vbe_state"] == "running"


@pytest.mark.parametrize("result", [
    {"success": True},
    {"success": False, "error_pattern": "server_busy"},
    {"success": False, "error_pattern": "rebuild_stalled"},
])
def test_other_results_are_not_probed(result):
    def reader(db):
        raise AssertionError("probed")
    assert _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, result, reader) == result


def test_no_database_means_no_probe():
    def reader(db):
        raise AssertionError("probed")
    assert _attach("vcs_x", tool_without_db, (), {}, dict(TIMEOUT), reader) == TIMEOUT


def test_screenshot_tool_is_skipped():
    def reader(db):
        raise AssertionError("probed")
    _attach("vcs_capture_screenshot", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT), reader)


def test_probe_failure_returns_result_unchanged():
    def reader(db):
        raise RuntimeError("boom")
    assert _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT), reader) == TIMEOUT


def test_slow_probe_is_abandoned(monkeypatch):
    import time
    monkeypatch.setattr(ast, "PROBE_TIMEOUT_SEC", 0.05)
    result = _attach("vcs_run_vba", tool_with_db, ("C:\\db.accdb",), {}, dict(TIMEOUT),
                     lambda db: time.sleep(0.5) or _state())
    assert "access_state" not in result


def test_read_access_state_caps_text():
    long_text = "x" * 2000
    meta = {"vbe_state": "design", "modal_blocked": True,
            "dialogs": [{**DIALOG, "text": long_text}]}
    with (
        patch.object(screenshot, "resolve_target_pid", return_value=(42, "restart_manager")),
        patch.object(screenshot, "capture_process", return_value=screenshot.Outcome(meta=meta)),
        patch.object(screenshot, "server_owns", return_value=False),
    ):
        state = ast.read_access_state("C:\\db.accdb")
    assert len(state["dialogs"][0]["text"]) == ast.TEXT_CAP
    assert state == {**state, "pid": 42, "owned_by_server": False, "modal_blocked": True}


def test_read_access_state_uses_error_metadata():
    error = screenshot.ScreenshotError("window_minimized", "minimized",
                                       vbe_state="break", dialogs=[], modal_blocked=False)
    with (
        patch.object(screenshot, "resolve_target_pid", return_value=(7, "window_title")),
        patch.object(screenshot, "capture_process", side_effect=error),
        patch.object(screenshot, "server_owns", return_value=True),
    ):
        state = ast.read_access_state("C:\\db.accdb")
    assert state["vbe_state"] == "break"


def test_read_access_state_none_when_not_open():
    error = screenshot.ScreenshotError("database_not_open", "not open")
    with patch.object(screenshot, "resolve_target_pid", side_effect=error):
        assert ast.read_access_state("C:\\db.accdb") is None
