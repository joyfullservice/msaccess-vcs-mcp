"""Tests for dismissing dialogs and resetting break mode on server-created Access."""

from __future__ import annotations

import time
from types import SimpleNamespace

import pytest

from msaccess_vcs_mcp import vba_control
from msaccess_vcs_mcp import screen_capture as sc
from msaccess_vcs_mcp.screen_capture import Rect, WindowInfo

PID = 4242


def dialog_window(hwnd: int = 7) -> WindowInfo:
    frame = Rect(0, 0, 200, 100)
    return WindowInfo(hwnd=hwnd, pid=PID, cls="#32770", title="Example", rect=frame, frame=frame)


def entry(dialog_id: str = "abc123def456", buttons=("Yes", "No"), source: str = "window") -> dict:
    return {"dialog_id": dialog_id, "title": "Example", "class": "#32770", "text": "Continue?",
            "buttons": list(buttons), "text_source": source}


@pytest.fixture
def process(monkeypatch):
    """One server-created instance showing one dialog, clicked by a recorder."""
    state = {"dialogs": [(dialog_window(), entry())], "pressed": [], "vbe": "design"}
    monkeypatch.setattr(vba_control, "resolve_target_pid", lambda db, pid: (PID, "argument"))

    def current(pid):
        return [w for w, _ in state["dialogs"]], [e for _, e in state["dialogs"]]

    def press(hwnd, item, button):
        state["pressed"].append((hwnd, button))
        state["dialogs"] = []
        return True

    monkeypatch.setattr(vba_control, "current_dialogs", current)
    monkeypatch.setattr(vba_control, "_press", press)
    monkeypatch.setattr(vba_control, "_wait_closed", lambda hwnd, timeout: True)
    monkeypatch.setattr(vba_control, "_process_windows", lambda pid: [])
    monkeypatch.setattr(sc, "vbe_state", lambda windows: state["vbe"])
    return state


def test_dismiss_clicks_the_named_button(process):
    result = vba_control.dismiss_dialog(None, PID, "abc123def456", "&no", owns=lambda p: True)

    assert result["success"] is True
    assert result["closed"] is True
    assert result["dialogs"] == []
    assert result["dismissed"] == {"dialog_id": "abc123def456", "title": "Example", "button": "&no"}
    assert process["pressed"] == [(7, "&no")]


def test_dismiss_refuses_user_owned_instances(process):
    result = vba_control.dismiss_dialog(None, PID, "abc123def456", "No", owns=lambda p: False)

    assert result["success"] is False
    assert result["error_pattern"] == "user_owned_instance"
    assert "Ask the user" in result["agent_guidance"]
    assert process["pressed"] == []


def test_dismiss_refuses_a_dialog_that_changed(process):
    result = vba_control.dismiss_dialog(None, PID, "stale0000000", "No", owns=lambda p: True)

    assert result["error_pattern"] == "dialog_changed"
    assert result["dialogs"][0]["dialog_id"] == "abc123def456"
    assert process["pressed"] == []


def test_dismiss_refuses_an_unlisted_button(process):
    result = vba_control.dismiss_dialog(None, PID, "abc123def456", "Cancel", owns=lambda p: True)

    assert result["error_pattern"] == "invalid_button"
    assert result["buttons"] == ["Yes", "No"]


def test_dismiss_reports_a_failed_press(process, monkeypatch):
    monkeypatch.setattr(vba_control, "_press", lambda hwnd, item, button: False)
    result = vba_control.dismiss_dialog(None, PID, "abc123def456", "No", owns=lambda p: True)
    assert result["error_pattern"] == "dismiss_failed"


def test_dismiss_warns_when_the_dialog_stays(process, monkeypatch):
    monkeypatch.setattr(vba_control, "_wait_closed", lambda hwnd, timeout: False)
    result = vba_control.dismiss_dialog(None, PID, "abc123def456", "Yes", owns=lambda p: True)
    assert result["success"] is True and result["closed"] is False
    assert result["warnings"]


def test_press_uses_window_click_then_accessibility(monkeypatch):
    calls = []
    monkeypatch.setattr(sc, "click_button", lambda hwnd, label: calls.append("click") or False)
    from msaccess_vcs_mcp import accessibility

    monkeypatch.setattr(accessibility, "press_button",
                        lambda hwnd, label: calls.append("msaa") or True)

    assert vba_control._press(7, entry(), "OK") is True
    assert calls == ["click", "msaa"]

    calls.clear()
    assert vba_control._press(7, entry(source="accessibility"), "OK") is True
    assert calls == ["msaa"]


# ---------------------------------------------------------------------------
# Reset
# ---------------------------------------------------------------------------


def test_reset_refuses_user_owned(process):
    process["vbe"] = "break"
    result = vba_control.reset_vba("C:\\db.accdb", None, owns=lambda p: False)
    assert result["error_pattern"] == "user_owned_instance"


def test_reset_needs_break_mode(process):
    result = vba_control.reset_vba("C:\\db.accdb", None, owns=lambda p: True)
    assert result["error_pattern"] == "not_in_break_mode"
    assert result["vbe_state"] == "design"


def test_reset_reports_where_execution_stopped(process):
    process["vbe"] = "break"
    stopped = {"module": "modExample", "line": 8, "code": "    Stop"}

    def session(db, pid, box):
        assert (db, pid) == ("C:\\db.accdb", PID)
        box["stopped_at"] = stopped
        box["executed"] = True
        process["vbe"] = "design"

    result = vba_control.reset_vba("C:\\db.accdb", None, owns=lambda p: True, session=session)

    assert result == {"success": True, "pid": PID, "stopped_at": stopped, "vbe_state": "design"}


def test_reset_times_out_on_an_unresponsive_access(process):
    process["vbe"] = "break"
    result = vba_control.reset_vba(
        "C:\\db.accdb", None, owns=lambda p: True,
        session=lambda db, pid, box: time.sleep(1), timeout=0.05,
    )
    assert result["error_pattern"] == "access_unresponsive"


def test_reset_uses_the_owned_record_when_only_a_pid_is_given(process, monkeypatch):
    process["vbe"] = "break"
    from msaccess_vcs_mcp.access_com import instance_registry

    monkeypatch.setattr(instance_registry, "list_owned",
                        lambda: [SimpleNamespace(pid=PID, database_path="C:\\owned.accdb")])
    seen = []

    def session(db, pid, box):
        seen.append(db)
        process["vbe"] = "design"

    vba_control.reset_vba(None, PID, owns=lambda p: True, session=session)
    assert seen == ["C:\\owned.accdb"]


@pytest.mark.parametrize("tool_name, args", [
    ("vcs_dismiss_dialog", ("abc123def456", "OK")),
    ("vcs_reset_vba", ()),
])
def test_tools_refuse_when_writes_are_disabled(monkeypatch, tool_name, args):
    import asyncio

    import msaccess_vcs_mcp.tools as tools_module

    monkeypatch.setattr(tools_module, "get_config", lambda: {"ACCESS_VCS_DISABLE_WRITES": True})
    func = getattr(tools_module, tool_name)
    while hasattr(func, "__wrapped__"):
        func = func.__wrapped__

    result = asyncio.run(func(*args))

    assert result["success"] is False
    assert result["error_pattern"] == "write_disabled"


class FakeModule:
    def __init__(self, lines):
        self._lines = lines
        self.Parent = SimpleNamespace(Name="modExample")
        self.CountOfLines = len(lines)

    def Lines(self, start, count):
        return "\r\n".join(self._lines[start - 1:start - 1 + count])


def test_code_location_reads_selection_and_context():
    module = FakeModule(["Option Explicit", "", "Sub A()", "    x = 1", "End Sub"])
    pane = SimpleNamespace(CodeModule=module, GetSelection=lambda: (4, 5, 4, 6))

    location = vba_control.code_location(SimpleNamespace(ActiveCodePane=pane), context=1)

    assert location["module"] == "modExample"
    assert (location["line"], location["start_column"], location["end_column"]) == (4, 5, 6)
    assert location["code"] == "    x = 1"
    assert [c["line"] for c in location["context"]] == [3, 4, 5]


def test_code_location_without_a_code_pane():
    assert vba_control.code_location(SimpleNamespace(ActiveCodePane=None)) is None
