"""Dialog dismissal, break-mode reset, and compile location against real Access.

Each test starts its own Access through COM, records its PID and creation
time, registers it as server-created, and at the end stops only that PID
after re-checking the creation time. No other Access window is touched.
"""

from __future__ import annotations

import gc
import subprocess
import sys
import threading
import time

import pytest

pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(sys.platform != "win32", reason="Access automation"),
]

AC_MODULE = 5
AC_QUIT_SAVE_NONE = 2


@pytest.fixture
def scratch(tmp_path):
    """A visible Access with a new database, owned by this test."""
    win32com = pytest.importorskip("win32com.client")
    from msaccess_vcs_mcp.access_com.instance_registry import (
        process_create_time, register_owned, unregister_owned,
    )
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none

    before = set(list_access_pids_or_none() or [])
    try:
        app = win32com.DispatchEx("Access.Application")
    except Exception as exc:
        pytest.skip(f"Access is not available: {exc}")
    new = set(list_access_pids_or_none() or []) - before
    if len(new) != 1:
        app.Quit()
        pytest.skip("Could not tell which Access process this test started.")
    pid = new.pop()
    created = process_create_time(pid)
    db_path = str(tmp_path / "dialogs.accdb")
    app.NewCurrentDatabase(db_path)
    app.Visible = True
    app.UserControl = True
    register_owned(pid, db_path, create_time=created)
    try:
        yield app, pid, db_path
    finally:
        unregister_owned(pid, created)
        # Quit and release the proxy first: a proxy left pointing at a killed
        # process fails with RPC_S_CALL_FAILED when it is collected.
        try:
            app.Quit(AC_QUIT_SAVE_NONE)
        except Exception:
            pass
        del app
        gc.collect()
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline and process_create_time(pid) == created:
            time.sleep(0.2)
        if created is not None and process_create_time(pid) == created:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


def _add_module(app, name, code):
    component = app.VBE.ActiveVBProject.VBComponents.Add(1)
    component.Name = name
    component.CodeModule.AddFromString(code)
    app.DoCmd.Save(AC_MODULE, name)


def _call_in_worker(db_path, action):
    """Run ``action(app)`` on a separate COM thread; it blocks behind dialogs."""
    import pythoncom
    import win32com.client

    box: dict = {}

    def worker():
        pythoncom.CoInitialize()
        try:
            box["result"] = action(win32com.client.GetObject(db_path))
        except Exception as exc:
            box["error"] = exc
        finally:
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return thread, box


def _wait_for_dialogs(pid, timeout=8.0):
    from msaccess_vcs_mcp import vba_control

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        _, dialogs = vba_control.current_dialogs(pid)
        if dialogs:
            return dialogs
        time.sleep(0.1)
    return []


def test_msgbox_is_dismissed_by_its_button(scratch):
    from msaccess_vcs_mcp import vba_control

    app, pid, db_path = scratch
    _add_module(app, "modDialogs", (
        "Public Function AskUser() As Long\r\n"
        "    AskUser = MsgBox(\"Continue?\", vbYesNo, \"Integration\")\r\n"
        "End Function\r\n"
    ))
    thread, box = _call_in_worker(db_path, lambda a: a.Run("AskUser"))
    dialogs = _wait_for_dialogs(pid)
    assert dialogs and dialogs[0]["buttons"] == ["Yes", "No"]

    result = vba_control.dismiss_dialog(db_path, None, dialogs[0]["dialog_id"], "No")
    thread.join(5)

    assert result["success"] and result["closed"]
    assert not thread.is_alive()
    value = box["result"][0] if isinstance(box.get("result"), tuple) else box.get("result")
    assert value == 7


def test_office_dialog_is_read_and_dismissed_through_accessibility(scratch):
    from msaccess_vcs_mcp import vba_control

    app, pid, db_path = scratch
    _add_module(app, "modBroken", "Public Sub Broken()\r\n    undefinedVar = 1\r\nEnd Sub\r\n")
    if "Option Explicit" not in app.VBE.ActiveVBProject.VBComponents("modBroken").CodeModule.Lines(1, 2):
        app.VBE.ActiveVBProject.VBComponents("modBroken").CodeModule.InsertLines(2, "Option Explicit")
    thread, _box = _call_in_worker(db_path, lambda a: a.DoCmd.RunCommand(125))
    dialogs = _wait_for_dialogs(pid)
    assert dialogs and dialogs[0]["class"] == "NUIDialog"
    assert dialogs[0]["text_source"] == "accessibility"
    assert "OK" in dialogs[0]["buttons"]

    result = vba_control.dismiss_dialog(db_path, None, dialogs[0]["dialog_id"], "OK")
    thread.join(5)
    assert result["success"] and not thread.is_alive()


def test_break_mode_reset_reports_the_stop_line(scratch):
    from msaccess_vcs_mcp import screenshot, vba_control

    app, pid, db_path = scratch
    _add_module(app, "modBreak", (
        "Public Function StopHere() As Long\r\n    Stop\r\n    StopHere = 1\r\nEnd Function\r\n"
    ))
    app.VBE.MainWindow.Visible = True
    thread, _box = _call_in_worker(db_path, lambda a: a.Run("StopHere"))
    deadline = time.monotonic() + 8
    while time.monotonic() < deadline:
        if screenshot.capture_process(pid, include_image=False).meta["vbe_state"] == "break":
            break
        time.sleep(0.1)
    assert screenshot.win32_precheck(pid) is not None

    result = vba_control.reset_vba(db_path, None)
    thread.join(5)

    assert result["success"] is True
    assert result["stopped_at"]["module"] == "modBreak"
    assert result["stopped_at"]["code"].strip() == "Stop"
    assert not thread.is_alive()


def test_compile_reports_module_and_line(scratch):
    from msaccess_vcs_mcp import compile_check

    app, pid, _db_path = scratch
    _add_module(app, "modBroken", "Public Sub Broken()\r\n    undefinedVar = 1\r\nEnd Sub\r\n")
    module = app.VBE.ActiveVBProject.VBComponents("modBroken").CodeModule
    if "Option Explicit" not in module.Lines(1, module.CountOfDeclarationLines or 1):
        module.InsertLines(module.CountOfDeclarationLines + 1, "Option Explicit")

    result = compile_check.compile_with_location(app, pid, timeout=60)

    error = result["compile_error"]
    assert result["success"] is False
    assert error["location_source"] == "editor"
    assert error["module"] == "modBroken"
    assert error["code"].strip() == "undefinedVar = 1"
    assert "Variable not defined" in error["message"]
