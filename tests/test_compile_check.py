"""Tests for the editor-side compile that reports the first error's location."""

from __future__ import annotations

import threading
from types import SimpleNamespace

import pytest

from msaccess_vcs_mcp import compile_check as cc
from msaccess_vcs_mcp.screen_capture import Rect, WindowInfo

PID = 77
VBE_HWND = 0x200


@pytest.mark.parametrize(
    "title, module",
    [
        ("modExample (Code)", "modExample"),
        ("Form_frmExample (Code)", "Form_frmExample"),
        ("Microsoft Visual Basic for Applications - [modExample (Code)]", "modExample"),
        ("Microsoft Visual Basic for Applications - Db1 [break] - [modExample (Code)]",
         "modExample"),
        ("Immediate", None),
        ("Microsoft Visual Basic for Applications - Db1", None),
    ],
)
def test_module_from_title(title, module):
    assert cc.module_from_title(title) == module


def test_normalize_message():
    assert cc.normalize_message("Compile error:\n\nVariable not defined") == (
        "Compile error: Variable not defined"
    )


class FakeModule:
    def __init__(self, name, lines):
        self._lines = lines
        self.Parent = SimpleNamespace(Name=name)
        self.CountOfLines = len(lines)

    def Lines(self, start, count):
        return "\r\n".join(self._lines[start - 1:start - 1 + count])


class FakeControl:
    def __init__(self, on_execute):
        self.on_execute = on_execute
        self.executed = 0

    def Execute(self):
        self.executed += 1
        self.on_execute()


class FakeVBE:
    def __init__(self, control, pane_module="modBroken"):
        project = SimpleNamespace(FileName="C:\\db.accdb")
        self.VBProjects = SimpleNamespace(Count=1)
        self.VBProjects = _Projects([project])
        self.ActiveVBProject = None
        self.MainWindow = SimpleNamespace(Visible=False)
        self._control = control
        module = FakeModule(pane_module, ["Option Explicit", "Sub A()", "    y = 1", "End Sub"])
        self.ActiveCodePane = SimpleNamespace(CodeModule=module, GetSelection=lambda: (3, 5, 3, 6))
        self.CommandBars = SimpleNamespace(FindControl=self._find)

    def _find(self, kind, control_id):
        assert (kind, control_id) == (1, cc.COMPILE_CONTROL_ID)
        return self._control


class _Projects:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __call__(self, index):
        return self._items[index - 1]


def fake_app(control, compiled=False, pane_module="modBroken"):
    return SimpleNamespace(
        VBE=FakeVBE(control, pane_module),
        IsCompiled=compiled,
        CurrentProject=SimpleNamespace(FullName="C:\\db.accdb"),
    )


class FakeProbe:
    def __init__(self, compiled=False):
        self.compiled = threading.Event()
        if compiled:
            self.compiled.set()
        self.closed = False

    def start(self):
        return self

    def close(self):
        self.closed = True


def vbe_dialog(hwnd=9) -> WindowInfo:
    frame = Rect(0, 0, 300, 150)
    return WindowInfo(hwnd=hwnd, pid=PID, cls="#32770", title=cc.VBE_DIALOG_TITLE,
                      rect=frame, frame=frame, owner=VBE_HWND)


def vbe_window(title="Microsoft Visual Basic for Applications - [modBroken (Code)]"):
    frame = Rect(0, 0, 800, 600)
    return WindowInfo(hwnd=VBE_HWND, pid=PID, cls="wndclass_desked_gsk", title=title,
                      rect=frame, frame=frame)


COMPILE_ENTRY = {"dialog_id": "x", "title": cc.VBE_DIALOG_TITLE, "class": "#32770",
                 "text": "Compile error:\n\nVariable not defined", "buttons": ["OK", "Help"]}


def run(app, probe, reads, clicked=None, gone=True):
    clicks = [] if clicked is None else clicked
    sequence = iter(reads)
    last = {"value": ([], [])}

    def read_windows(pid):
        last["value"] = next(sequence, last["value"])
        return last["value"]

    return cc.compile_with_location(
        app, PID, timeout=2,
        probe_factory=lambda a: probe,
        read_windows=read_windows,
        click=lambda hwnd, label: clicks.append((hwnd, label)) or True,
        window_gone=lambda hwnd: gone,
    )


def test_already_compiled_skips_the_compile():
    control = FakeControl(lambda: pytest.fail("compiled again"))
    result = run(fake_app(control, compiled=True), FakeProbe(), [])
    assert result == {"success": True, "already_compiled": True}


def test_no_compile_control_falls_back():
    app = fake_app(None)
    assert run(app, FakeProbe(), []) is None


def test_probe_reports_success():
    control = FakeControl(lambda: None)
    probe = FakeProbe(compiled=True)
    assert run(fake_app(control), probe, []) == {"success": True}
    assert control.executed == 1 and probe.closed


def test_compile_error_is_located_from_the_editor():
    clicks = []
    reads = [([], []), ([vbe_window(), vbe_dialog()], [COMPILE_ENTRY])]
    result = run(fake_app(FakeControl(lambda: None)), FakeProbe(), reads, clicks)

    error = result["compile_error"]
    assert result["success"] is False
    assert error["message"] == "Compile error: Variable not defined"
    assert (error["module"], error["line"], error["start_column"]) == ("modBroken", 3, 5)
    assert error["code"] == "    y = 1"
    assert error["location_source"] == "editor"
    assert result["vbe_opened_by_compile"] is True
    assert clicks == [(9, "OK")]


def test_module_disagreement_reports_the_window_title():
    reads = [([vbe_window(), vbe_dialog()], [COMPILE_ENTRY])]
    result = run(fake_app(FakeControl(lambda: None), pane_module="modOther"), FakeProbe(), reads)

    error = result["compile_error"]
    assert error["module"] == "modBroken"
    assert error["location_source"] == "window"
    assert "line" not in error
    assert any("modOther" in w for w in result["warnings"])


def test_undismissable_dialog_returns_window_evidence():
    reads = [([vbe_window(), vbe_dialog()], [COMPILE_ENTRY])]
    result = run(fake_app(FakeControl(lambda: None)), FakeProbe(), reads, gone=False)
    assert result["compile_error"]["location_source"] == "window"
    assert result["warnings"]


def test_another_dialog_is_left_for_the_agent():
    other = {"dialog_id": "y", "title": "Microsoft Access", "class": "#32770",
             "text": "Something else", "buttons": ["OK"]}
    frame = Rect(0, 0, 10, 10)
    window = WindowInfo(hwnd=5, pid=PID, cls="#32770", title="Microsoft Access",
                        rect=frame, frame=frame)
    clicks = []
    result = run(fake_app(FakeControl(lambda: None)), FakeProbe(), [([window], [other])], clicks)

    assert result["error_pattern"] == "dialog_shown"
    assert result["dialogs"] == [other]
    assert clicks == []


def test_timeout_without_dialog_or_success(monkeypatch):
    monkeypatch.setattr(cc, "POLL_SEC", 0.01)
    result = cc.compile_with_location(
        fake_app(FakeControl(lambda: None)), PID, timeout=0.05,
        probe_factory=lambda a: FakeProbe(), read_windows=lambda pid: ([], []),
    )
    assert result["error_pattern"] == "compile_timeout"
