"""Answer Access dialogs and leave VBA break mode on server-created instances.

Both actions change what a running Access does next, so they are limited to
instances the server created (a confirmed PID plus creation-time match in
the ownership registry) and refused when writes are disabled. A user's own
Access window is never clicked or reset; the result says to ask the user.

Neither waits for the Access gate. The call that is stuck behind the dialog
or the break usually holds it, and answering the dialog is what releases it.
Dismissal uses window messages and MSAA only. A reset needs COM, which Access
still services in break mode, and runs on its own thread with a deadline.
"""

from __future__ import annotations

import threading
import time
from typing import Any, Callable

from . import screen_capture as sc
from .screenshot import ScreenshotError, resolve_target_pid, server_owns

DISMISS_WAIT_SEC = 2.0
RESET_TIMEOUT_SEC = 15.0
RESET_SETTLE_SEC = 5.0
VBE_RESET_CONTROL_ID = 228
CONTEXT_LINES = 2

USER_OWNED_GUIDANCE = (
    "This Access window was not opened by this server, so it will not click its "
    "dialogs or reset its code. Ask the user to respond, then retry."
)


def _error(pattern: str, message: str, **extra: Any) -> dict[str, Any]:
    return ScreenshotError(pattern, message, **extra).meta()


def _owned_or_error(pid: int, owns: Callable[[int], bool]) -> dict[str, Any] | None:
    if owns(pid):
        return None
    return _error(
        "user_owned_instance",
        f"Access PID {pid} was not created by this server.",
        pid=pid,
        owned_by_server=False,
        agent_guidance=USER_OWNED_GUIDANCE,
    )


def _process_windows(pid: int) -> list[sc.WindowInfo]:
    with sc.dpi_aware():
        return sc.enum_windows({pid})


def current_dialogs(pid: int) -> tuple[list[sc.WindowInfo], list[dict[str, Any]]]:
    """Visible dialog windows and their entries, in the same order."""
    windows = _process_windows(pid)
    main = sc.find_window(windows, sc.MAIN_WINDOW_CLASS)
    vbe = sc.find_window(windows, sc.VBE_WINDOW_CLASS)
    dialogs, _ = sc.read_dialogs(windows, main.hwnd if main else 0, vbe.hwnd if vbe else 0)
    return sc.dialog_windows(windows), dialogs


def _press(hwnd: int, entry: dict[str, Any], button: str) -> bool:
    if entry.get("text_source") != "accessibility" and sc.click_button(hwnd, button):
        return True
    from . import accessibility

    return accessibility.press_button(hwnd, button)


def _wait_closed(hwnd: int, timeout: float) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if sc.window_gone(hwnd):
            return True
        time.sleep(0.05)
    return sc.window_gone(hwnd)


def dismiss_dialog(
    database_path: str | None,
    pid: int | None,
    dialog_id: str,
    button: str,
    *,
    owns: Callable[[int], bool] = server_owns,
) -> dict[str, Any]:
    """Click ``button`` on the dialog the agent read as ``dialog_id``."""
    try:
        resolved, _source = resolve_target_pid(database_path, pid)
    except ScreenshotError as exc:
        return exc.meta()
    refused = _owned_or_error(resolved, owns)
    if refused:
        return refused

    windows, dialogs = current_dialogs(resolved)
    match = next(
        ((w, d) for w, d in zip(windows, dialogs) if d.get("dialog_id") == dialog_id), None
    )
    if match is None:
        return _error(
            "dialog_changed",
            "No dialog with that dialog_id is showing. It was answered or replaced; read "
            "`dialogs` and pass the current dialog_id.",
            pid=resolved,
            dialogs=dialogs,
        )
    window, entry = match
    wanted = sc.strip_accelerators(button).strip().casefold()
    if wanted not in {b.casefold() for b in entry.get("buttons", [])}:
        return _error(
            "invalid_button",
            f"The dialog has no '{button}' button.",
            pid=resolved,
            buttons=entry.get("buttons", []),
        )
    if not _press(window.hwnd, entry, button):
        return _error(
            "dismiss_failed",
            f"The '{button}' button could not be pressed.",
            pid=resolved,
            dialogs=dialogs,
        )
    closed = _wait_closed(window.hwnd, DISMISS_WAIT_SEC)
    _, remaining = current_dialogs(resolved)
    result: dict[str, Any] = {
        "success": True,
        "pid": resolved,
        "dismissed": {"dialog_id": dialog_id, "title": entry.get("title"), "button": button},
        "closed": closed,
        "dialogs": remaining,
        "vbe_state": sc.vbe_state(_process_windows(resolved)),
    }
    if not closed:
        result["warnings"] = [
            "The dialog is still showing after the click. Access may be busy; take a "
            "screenshot before trying again."
        ]
    return result


# ---------------------------------------------------------------------------
# Code location and break-mode reset
# ---------------------------------------------------------------------------


def code_location(vbe: Any, context: int = CONTEXT_LINES) -> dict[str, Any] | None:
    """Module, selection, and surrounding code of the VBE's active code pane."""
    pane = vbe.ActiveCodePane
    if pane is None:
        return None
    module = pane.CodeModule
    start_line, start_col, end_line, end_col = (int(v) for v in pane.GetSelection())
    total = int(module.CountOfLines)
    first = max(1, start_line - context)
    last = min(total, end_line + context)
    lines = [
        {"line": n, "code": str(module.Lines(n, 1))} for n in range(first, last + 1)
    ] if total else []
    return {
        "module": str(module.Parent.Name),
        "line": start_line,
        "start_column": start_col,
        "end_line": end_line,
        "end_column": end_col,
        "code": str(module.Lines(start_line, 1)) if 0 < start_line <= total else "",
        "context": lines,
    }


def _database_for(pid: int, database_path: str | None) -> str | None:
    if database_path:
        return database_path
    from .access_com.instance_registry import list_owned

    try:
        paths = [r.database_path for r in list_owned() if r.pid == pid and r.database_path]
    except Exception:
        return None
    return paths[0] if len(paths) == 1 else None


def _reset_session(database_path: str, pid: int, box: dict[str, Any]) -> None:
    import pythoncom

    from .access_com.process_qos import pid_from_access_app
    from .screenshot import _find_app

    pythoncom.CoInitialize()
    try:
        app = _find_app(database_path)
        if app is None:
            box["error"] = _error("database_not_open", "The database is not open in Access.")
            return
        app_pid = pid_from_access_app(app)
        if app_pid != pid:
            box["error"] = _error(
                "database_not_open",
                f"The database is open in Access PID {app_pid}, not PID {pid}.",
            )
            return
        vbe = app.VBE
        try:
            box["stopped_at"] = code_location(vbe)
        except Exception as exc:
            box["location_error"] = str(exc)
        control = vbe.CommandBars.FindControl(1, VBE_RESET_CONTROL_ID)
        if control is None:
            box["error"] = _error("reset_failed", "The VBA editor has no Reset command.")
            return
        control.Execute()
        box["executed"] = True
    except Exception as exc:
        box["error"] = _error("reset_failed", f"The reset failed: {exc}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def reset_vba(
    database_path: str | None,
    pid: int | None,
    *,
    owns: Callable[[int], bool] = server_owns,
    session: Callable[[str, int, dict[str, Any]], None] = _reset_session,
    timeout: float = RESET_TIMEOUT_SEC,
) -> dict[str, Any]:
    """Report where VBA stopped, then run the VBE Reset command."""
    try:
        resolved, _source = resolve_target_pid(database_path, pid)
    except ScreenshotError as exc:
        return exc.meta()
    refused = _owned_or_error(resolved, owns)
    if refused:
        return refused
    state = sc.vbe_state(_process_windows(resolved))
    if state != "break":
        return _error(
            "not_in_break_mode",
            "VBA is not in break mode, so there is nothing to reset.",
            pid=resolved,
            vbe_state=state,
        )
    target_db = _database_for(resolved, database_path)
    if not target_db:
        return _error(
            "invalid_argument",
            "Pass database_path: the server could not tell which database this instance has open.",
            pid=resolved,
        )

    box: dict[str, Any] = {}
    worker = threading.Thread(
        target=session, args=(target_db, resolved, box), daemon=True, name="vcs-reset-vba"
    )
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        return _error(
            "access_unresponsive",
            f"Access did not answer the reset within {timeout:.0f} s.",
            pid=resolved,
        )
    if "error" in box:
        return {**box["error"], "pid": resolved, "stopped_at": box.get("stopped_at")}

    deadline = time.monotonic() + RESET_SETTLE_SEC
    state = sc.vbe_state(_process_windows(resolved))
    while state == "break" and time.monotonic() < deadline:
        time.sleep(0.1)
        state = sc.vbe_state(_process_windows(resolved))
    result: dict[str, Any] = {
        "success": state != "break",
        "pid": resolved,
        "stopped_at": box.get("stopped_at"),
        "vbe_state": state,
    }
    if "location_error" in box:
        result["warnings"] = [f"Could not read the break location: {box['location_error']}"]
    if state == "break":
        result.update(
            error="The Reset command ran but VBA is still in break mode.",
            error_pattern="reset_failed",
        )
    return result
