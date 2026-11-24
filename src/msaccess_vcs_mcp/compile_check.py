"""Compile from the VBA editor and report where the first error is.

``acCmdCompileAllModules`` (the add-in's ``CompileVBA``) shows Office's
"Cannot Compile Project." dialog, blocks the caller until someone clicks
OK, and never moves to the error. The editor's own Compile command
(control 578) returns at once, then navigates to the error and shows a
standard dialog: "Compile error:" and the reason, with OK and Help.

So the compile is started from the editor, and the gate thread watches for
that dialog with window calls only. While the dialog is up a COM call into
Access waits behind it, so nothing on this thread calls Access until the
dialog is dismissed. Whether the compile succeeded is asked on a separate
COM thread, where a call stuck behind the dialog blocks only that thread.

The location comes from two places: while the dialog is up, its message
and the code window's title (the module); after OK, the code pane's
selection over COM, which gives the exact line and columns.
"""

from __future__ import annotations

import os
import re
import threading
import time
from typing import Any, Callable

from . import screen_capture as sc
from .access_com.vba_project import current_vb_project

COMPILE_CONTROL_ID = 578
VBE_DIALOG_TITLE = "Microsoft Visual Basic for Applications"
DEFAULT_TIMEOUT_SEC = 120.0
POLL_SEC = 0.1
PROBE_INTERVAL_SEC = 0.25
DISMISS_WAIT_SEC = 3.0
CODE_WINDOW_TITLE = re.compile(r"^(?P<module>.+?) \((?:Code)\)$")


def compile_timeout() -> float:
    try:
        value = float(os.environ.get("ACCESS_VCS_COMPILE_TIMEOUT_SEC", DEFAULT_TIMEOUT_SEC))
    except ValueError:
        return DEFAULT_TIMEOUT_SEC
    return value if value > 0 else DEFAULT_TIMEOUT_SEC


def normalize_message(text: str) -> str:
    """"Compile error:\\n\\nVariable not defined" -> "Compile error: Variable not defined"."""
    return " ".join(part.strip() for part in text.splitlines() if part.strip())


def module_from_title(title: str) -> str | None:
    """Module name from a code window title ("modX (Code)") or the VBE title
    when that window is maximized ("... - [modX (Code)]")."""
    bracket = re.search(r"\[([^\[\]]+)\]\s*$", title)
    candidate = bracket.group(1) if bracket else title
    match = CODE_WINDOW_TITLE.match(candidate.strip())
    return match.group("module") if match else None


def active_code_window_module(windows: list[sc.WindowInfo]) -> str | None:
    vbe = sc.find_window(windows, sc.VBE_WINDOW_CLASS)
    if vbe is None:
        return None
    module = module_from_title(vbe.title)
    if module:
        return module
    child = sc.mdi_active_child(vbe.hwnd)
    return module_from_title(sc._stored_text(child)) if child else None


def compile_dialog(
    windows: list[sc.WindowInfo], dialogs: list[dict[str, Any]]
) -> tuple[sc.WindowInfo, dict[str, Any]] | None:
    for window, entry in zip(sc.dialog_windows(windows), dialogs):
        if entry.get("title") == VBE_DIALOG_TITLE and "OK" in entry.get("buttons", []):
            return window, entry
    return None


class _CompiledProbe:
    """Polls ``IsCompiled`` from its own COM apartment."""

    def __init__(self, app: Any) -> None:
        import pythoncom

        self.compiled = threading.Event()
        self.stop = threading.Event()
        self._stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, app._oleobj_
        )
        self.thread = threading.Thread(target=self._run, daemon=True, name="vcs-compile-probe")

    def _run(self) -> None:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch(
                pythoncom.CoGetInterfaceAndReleaseStream(self._stream, pythoncom.IID_IDispatch)
            )
            while not self.stop.is_set():
                try:
                    if app.IsCompiled:
                        self.compiled.set()
                        return
                except Exception:
                    pass
                self.stop.wait(PROBE_INTERVAL_SEC)
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()

    def start(self) -> "_CompiledProbe":
        self.thread.start()
        return self

    def close(self) -> None:
        self.stop.set()


def _windows(pid: int) -> list[sc.WindowInfo]:
    with sc.dpi_aware():
        return sc.enum_windows({pid})


def _read(pid: int) -> tuple[list[sc.WindowInfo], list[dict[str, Any]]]:
    windows = _windows(pid)
    main = sc.find_window(windows, sc.MAIN_WINDOW_CLASS)
    vbe = sc.find_window(windows, sc.VBE_WINDOW_CLASS)
    dialogs, _ = sc.read_dialogs(windows, main.hwnd if main else 0, vbe.hwnd if vbe else 0)
    return windows, dialogs


def compile_with_location(
    app: Any,
    pid: int,
    *,
    timeout: float | None = None,
    probe_factory: Callable[[Any], Any] = _CompiledProbe,
    read_windows: Callable[[int], tuple[list[sc.WindowInfo], list[dict[str, Any]]]] = _read,
    click: Callable[[int, str], bool] = sc.click_button,
    window_gone: Callable[[int], bool] = sc.window_gone,
) -> dict[str, Any] | None:
    """Compile the current project; None when the editor has no Compile command.

    On a compile error the result carries ``compile_error`` with the message,
    module, line, columns, and code. The editor is left open on the error.
    """
    from .vba_control import code_location

    vbe = app.VBE
    vbe.ActiveVBProject = current_vb_project(app)
    if app.IsCompiled:
        return {"success": True, "already_compiled": True}
    control = vbe.CommandBars.FindControl(1, COMPILE_CONTROL_ID)
    if control is None:
        return None
    vbe_was_visible = bool(vbe.MainWindow.Visible)

    control.Execute()
    probe = probe_factory(app).start()
    deadline = time.monotonic() + (timeout or compile_timeout())
    found = None
    other: list[dict[str, Any]] = []
    try:
        while time.monotonic() < deadline:
            if probe.compiled.is_set():
                return {"success": True}
            windows, dialogs = read_windows(pid)
            found = compile_dialog(windows, dialogs)
            if found is not None:
                break
            if dialogs:
                other = dialogs
                break
            time.sleep(POLL_SEC)
    finally:
        probe.close()

    if found is None:
        if other:
            return {
                "success": False,
                "error": "Access showed a dialog during the compile.",
                "error_pattern": "dialog_shown",
                "dialogs": other,
            }
        return {
            "success": False,
            "error": "The compile did not finish or report an error in time.",
            "error_pattern": "compile_timeout",
        }

    window, entry = found
    evidence = {
        "message": normalize_message(str(entry.get("text", ""))),
        "module": active_code_window_module(windows),
    }
    clicked = click(window.hwnd, "OK")
    deadline = time.monotonic() + DISMISS_WAIT_SEC
    while clicked and not window_gone(window.hwnd) and time.monotonic() < deadline:
        time.sleep(POLL_SEC)
    if not clicked or not window_gone(window.hwnd):
        return {
            "success": False,
            "compile_error": {**evidence, "location_source": "window"},
            "warnings": ["The compile error dialog could not be dismissed."],
        }

    location = None
    warnings: list[str] = []
    try:
        location = code_location(vbe, context=3)
    except Exception as exc:
        warnings.append(f"Could not read the error location from the editor: {exc}")
    if location and evidence["module"] and location["module"] != evidence["module"]:
        warnings.append(
            f"The editor's code pane is '{location['module']}' but the code window title "
            f"names '{evidence['module']}'; reporting the title."
        )
        location = None
    compile_error: dict[str, Any] = {"message": evidence["message"]}
    if location:
        compile_error.update(location, location_source="editor")
    else:
        compile_error.update(module=evidence["module"], location_source="window")
    result: dict[str, Any] = {
        "success": False,
        "compile_error": compile_error,
        "vbe_opened_by_compile": not vbe_was_visible,
    }
    if warnings:
        result["warnings"] = warnings
    return result
