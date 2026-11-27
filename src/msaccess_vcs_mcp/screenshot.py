"""Screenshot orchestration: choose the Access process, capture, build the result.

``application`` and ``vbe`` captures never call into Access. The process is
found through the Restart Manager, the owned-instance registry, and window
titles, and the pixels come from ``screen_capture``. That is the path that
must keep working while Access is blocked.

``active`` and ``object`` captures need COM to find the object's window.
They run inside the Access gate, refuse to call into an Access that is
already showing a dialog, and run every COM step on a worker with a
deadline, because a call queued into a blocked Access runs later, after the
tool has given up.
"""

from __future__ import annotations

import asyncio
import json
import os
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable

from mcp.server.fastmcp import Image
from mcp.types import CallToolResult, TextContent

from . import screen_capture as sc
from . import screenshot_store
from . import user_settings
from .access_com.instance_registry import list_owned, paths_match
from .access_com.process_qos import list_access_pids_or_none, pid_from_access_app

TARGETS = ("application", "vbe", "active", "object")
OBJECT_TYPES = ("form", "report", "table", "query")
MIN_DIMENSION = 320
MAX_DIMENSION = 3840
DEFAULT_MAX_DIMENSION = 1568
CAPTURE_TIMEOUT_SEC = 30.0
DEFAULT_OPEN_TIMEOUT_SEC = 10.0
COM_STEP_TIMEOUT_SEC = 10.0
CAPTURE_STEP_TIMEOUT_SEC = 20.0
DEFAULT_SETTLE_MS = 150
MAX_SETTLE_MS = 5000
VIEWS = ("normal", "design")

RETENTION_UNDECIDED_GUIDANCE = (
    f"Saved screenshots are kept for the {user_settings.SCREENSHOT_RETENTION.default}-day "
    "default because the user has not chosen "
    "a retention period. Ask the user how long to keep them (\"never\", 7 or 30 days, or "
    "\"forever\"), then call vcs_configure(screenshot_retention=...) with their answer."
)

SAVE_NEVER_WARNING = (
    "Nothing was saved: the user chose never to keep screenshots (screenshot_retention is "
    "\"never\"). Use include_image=True to see the capture instead."
)

# Access object types (AcObjectType), views, and save options.
AC_TABLE = 0
AC_QUERY = 1
AC_FORM = 2
AC_REPORT = 3
AC_NORMAL = 0
AC_DESIGN = 1
AC_VIEW_NORMAL = 0
AC_VIEW_DESIGN = 1
AC_VIEW_PREVIEW = 2
AC_READ_ONLY = 2
AC_SAVE_NO = 2
AC_OBJECT_TYPES = {"table": AC_TABLE, "query": AC_QUERY, "form": AC_FORM, "report": AC_REPORT}

# AcCurrentView, read from Form.CurrentView and Report.CurrentView.
CURRENT_VIEW_NAMES = {
    0: "design", 1: "form", 2: "datasheet", 3: "pivot_table", 4: "pivot_chart",
    5: "print_preview", 6: "report", 7: "layout",
}

# DAO QueryDefTypeEnum values that only return rows. Everything else --
# action, DDL, make-table, pass-through -- can change data when opened.
# access_com.dao_helpers.get_query_type_name uses a different, wrong map.
DB_Q_SELECT = 0
DB_Q_CROSSTAB = 16
DB_Q_SET_OPERATION = 128
SAFE_QUERY_TYPES = frozenset({DB_Q_SELECT, DB_Q_CROSSTAB, DB_Q_SET_OPERATION})

# Access run-time errors raised through COM.
ERR_NO_ACTIVE_FORM = 2475
ERR_NO_ACTIVE_REPORT = 2476
ERR_NO_ACTIVE_DATASHEET = 2484
ERR_OBJECT_NOT_FOUND = 2467
ERR_OPEN_CANCELED = 2501

DIALOG_GUIDANCE = (
    "Access is showing a dialog. Read `dialogs` before the image: it holds the "
    "dialog's text and buttons. When `owned_by_server` is true, answer it with "
    "vcs_dismiss_dialog(dialog_id=..., button=...). Otherwise ask the user to "
    "respond to the dialog, then retry."
)

BREAK_GUIDANCE = (
    "VBA is in break mode (`vbe_state` is 'break'): code stopped on an error or a "
    "Stop statement and every call into Access waits behind it. When "
    "`owned_by_server` is true, vcs_reset_vba() reports where it stopped and ends "
    "the run. Otherwise ask the user to reset the code (Run > Reset in the VBA "
    "editor)."
)


def server_owns(pid: int) -> bool:
    """A confirmed server-created instance; anything uncertain is the user's."""
    from .access_com.instance_registry import is_owned, process_create_time

    try:
        return is_owned(pid, process_create_time(pid))
    except Exception:
        return False


def guidance_for(meta: dict[str, Any]) -> str | None:
    parts = []
    if meta.get("vbe_state") == "break":
        parts.append(BREAK_GUIDANCE)
    if meta.get("dialogs"):
        parts.append(DIALOG_GUIDANCE)
    return " ".join(parts) or None


class ScreenshotError(Exception):
    """A failure the tool reports as ``success: false`` with an error pattern."""

    def __init__(self, error_pattern: str, message: str, **extra: Any) -> None:
        super().__init__(message)
        self.error_pattern = error_pattern
        self.extra = extra

    def meta(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": False,
            "error": str(self),
            "error_pattern": self.error_pattern,
            "recoverable": True,
        }
        result.update(self.extra)
        return result


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------


def build_result(meta: dict[str, Any], png: bytes | None = None) -> CallToolResult:
    """Metadata as the first text block, then the image when there is one.

    No ``structuredContent``: some clients render it in place of ``content``
    and would drop the image.
    """
    content: list[Any] = [TextContent(type="text", text=json.dumps(meta, indent=2, default=str))]
    if png:
        content.append(Image(data=png, format="png").to_image_content())
    return CallToolResult(content=content)


def error_result(pattern: str, message: str, **extra: Any) -> CallToolResult:
    return build_result(ScreenshotError(pattern, message, **extra).meta())


# ---------------------------------------------------------------------------
# Finding the process
# ---------------------------------------------------------------------------


def survey_access() -> dict[int, dict[str, Any]]:
    """Every process owning an Access window: title, hung, and VBE state."""
    found: dict[int, dict[str, Any]] = {}
    for w in sc.enum_windows():
        if w.cls == sc.MAIN_WINDOW_CLASS:
            entry = found.setdefault(w.pid, {"title": "", "hung": False, "vbe_open": False})
            entry["title"] = w.title
            entry["hung"] = w.hung
        elif w.cls == sc.VBE_WINDOW_CLASS:
            entry = found.setdefault(w.pid, {"title": "", "hung": False, "vbe_open": False})
            entry["vbe_open"] = entry["vbe_open"] or w.visible
    return found


def _owned_records() -> list[Any]:
    try:
        return list_owned()
    except Exception:
        return []


def _instances(live: set[int], survey: dict[int, dict[str, Any]], owned: set[int]) -> list[dict]:
    return [
        {
            "pid": pid,
            "title": survey.get(pid, {}).get("title", ""),
            "owned_by_server": pid in owned,
            "vbe_open": survey.get(pid, {}).get("vbe_open", False),
            "hung": survey.get(pid, {}).get("hung", False),
        }
        for pid in sorted(live)
    ]


def _title_has_path(title: str, database_path: str) -> bool:
    return os.path.normcase(os.path.abspath(database_path)).lower() in title.lower()


def _title_has_name(title: str, database_path: str) -> bool:
    return Path(database_path).name.lower() in title.lower()


def _title_match(live: set[int], survey: dict, database_path: str, instances: list) -> int | None:
    """The one window whose title names ``database_path``, or None.

    Access's default title carries the full path, which tells worktrees of
    one project apart. A database that sets its own application title shows
    no path, and every copy shows the same name, so a name shared by several
    windows is ambiguous rather than a match.
    """
    titles = {p: survey.get(p, {}).get("title", "") for p in live}
    by_path = {p for p, t in titles.items() if _title_has_path(t, database_path)}
    if len(by_path) == 1:
        return next(iter(by_path))
    by_name = by_path or {p for p, t in titles.items() if _title_has_name(t, database_path)}
    if len(by_name) == 1:
        return next(iter(by_name))
    if len(by_name) > 1:
        raise ScreenshotError(
            "ambiguous_instance",
            f"Several Access windows show a database named {Path(database_path).name} "
            "(copies of one project in different folders usually do). Pass pid.",
            instances=[i for i in instances if i["pid"] in by_name],
        )
    return None


def resolve_target_pid(database_path: str | None, pid: int | None) -> tuple[int, str]:
    """The Access process to capture, found without calling into Access.

    With a ``database_path``, the single-instance shortcut is never taken:
    an unrelated database's window must not be returned as if it were the
    one asked for.
    """
    survey = survey_access()
    tasklist = list_access_pids_or_none()
    live = set(tasklist) if tasklist is not None else set(survey)
    records = _owned_records()
    owned = {r.pid for r in records}
    instances = _instances(live, survey, owned)

    if pid is not None:
        if pid in live:
            return pid, "argument"
        raise ScreenshotError(
            "invalid_pid",
            f"PID {pid} is not a running Access process.",
            instances=instances,
        )

    if database_path:
        holders = sc.pids_holding_file(database_path)
        if holders:
            matches = holders & live
            if len(matches) == 1:
                return next(iter(matches)), "restart_manager"
            if len(matches) > 1:
                raise ScreenshotError(
                    "ambiguous_instance",
                    "Several Access processes hold this database open. Pass pid.",
                    instances=[i for i in instances if i["pid"] in matches],
                )
        owned_matches = {
            r.pid for r in records if r.pid in live and paths_match(r.database_path, database_path)
        }
        if len(owned_matches) == 1:
            return next(iter(owned_matches)), "owned_registry"
        if len(owned_matches) > 1:
            raise ScreenshotError(
                "ambiguous_instance",
                "Several server-created Access processes have this database open. Pass pid.",
                instances=[i for i in instances if i["pid"] in owned_matches],
            )
        titled = _title_match(live, survey, database_path, instances)
        if titled is not None:
            return titled, "window_title"
        raise ScreenshotError(
            "database_not_open",
            "No running Access process has this database open. Pass pid to capture one "
            "of the instances listed.",
            instances=instances,
        )

    if len(live) == 1:
        return next(iter(live)), "single_instance"
    if not live:
        raise ScreenshotError("access_not_running", "No Access process is running.")
    raise ScreenshotError(
        "ambiguous_instance",
        "Several Access processes are running. Pass database_path or pid.",
        instances=instances,
    )


# ---------------------------------------------------------------------------
# Win32 capture of a process
# ---------------------------------------------------------------------------


@dataclass
class Outcome:
    meta: dict[str, Any] = field(default_factory=dict)
    png: bytes | None = None


def _window_meta(info: sc.WindowInfo, responsive: bool | None = None) -> dict[str, Any]:
    meta: dict[str, Any] = {"title": info.title, "class": info.cls, "minimized": info.minimized}
    if responsive is not None:
        meta["responsive"] = responsive
    return meta


def _popups(windows: list[sc.WindowInfo], skip: set[int]) -> list[dict[str, str]]:
    return [
        {"title": w.title, "class": w.cls}
        for w in sorted(windows, key=lambda w: w.z)
        if w.hwnd not in skip
        and w.visible
        and not w.cloaked
        and not w.frame.empty
        and not sc.is_chrome(w)
        and not sc.is_dialog_class(w.cls)
        and w.cls not in (sc.MAIN_WINDOW_CLASS, sc.VBE_WINDOW_CLASS)
    ]


def object_crop(main: sc.WindowInfo, object_hwnd: int) -> sc.Rect:
    """The object's window, clipped to the MDI client area of the main window."""
    rect = sc.window_rect(object_hwnd)
    client = sc.mdi_client_rect(main.hwnd)
    if client is not None:
        clipped = rect.intersect(client)
        if clipped is not None:
            rect = clipped
    return rect


def capture_process(
    pid: int,
    target: str = "application",
    *,
    max_dimension: int = DEFAULT_MAX_DIMENSION,
    include_image: bool = True,
    object_hwnd: int = 0,
) -> Outcome:
    """Capture one Access process with Win32 only. Never calls into Access.

    ``object_hwnd`` crops to an object: a top-level popup is captured as
    its own base window; an MDI child crops the main window.
    """
    with sc.dpi_aware():
        windows = sc.enum_windows({pid})
        main = sc.find_window(windows, sc.MAIN_WINDOW_CLASS)
        vbe = sc.find_window(windows, sc.VBE_WINDOW_CLASS)
        if vbe is not None and not vbe.visible:
            vbe = None
        main_hwnd = main.hwnd if main else 0
        vbe_hwnd = vbe.hwnd if vbe else 0

        base = main
        crop: sc.Rect | None = None
        if target == "vbe":
            base = vbe
        elif object_hwnd:
            popup = next((w for w in windows if w.hwnd == object_hwnd), None)
            if popup is not None and sc.is_top_level(object_hwnd):
                base = popup
            elif main is not None:
                crop = object_crop(main, object_hwnd)

        dialogs, warnings = sc.read_dialogs(
            windows, main_hwnd, vbe_hwnd, base.frame if base else None
        )
        meta: dict[str, Any] = {
            "modal_blocked": bool(
                (main is not None and not main.enabled) or (vbe is not None and not vbe.enabled)
            ),
            "vbe_state": sc.vbe_state(windows),
            "dialogs": dialogs,
            "popups": _popups(windows, {w.hwnd for w in windows if sc.is_dialog_class(w.cls)}),
            "warnings": warnings,
        }
        guidance = guidance_for(meta)
        if guidance:
            meta["agent_guidance"] = guidance

        if base is None:
            if target == "vbe":
                raise ScreenshotError("vbe_not_open", "The VBA editor is not open.", **meta)
            raise ScreenshotError(
                "capture_failed", f"Access PID {pid} has no main window.", **meta
            )
        meta["window"] = _window_meta(base)
        if not base.visible:
            raise ScreenshotError(
                "window_hidden",
                "The Access window is hidden. Ask the user to show it; this tool never "
                "changes window state.",
                **meta,
            )
        if base.minimized:
            raise ScreenshotError(
                "window_minimized",
                "The window is minimized. Ask the user to restore it; this tool never "
                "restores or focuses windows.",
                **meta,
            )
        if not include_image:
            meta["window"]["responsive"] = sc.is_responsive(base.hwnd)
            return Outcome(meta=meta)

        overlays = sc.select_overlays(windows, base)
        capture = sc.capture_composite(
            base,
            overlays,
            sc.enum_windows(visible_only=True),
            crop=crop,
            max_dimension=max_dimension,
        )
        drawn = {o.hwnd for o in overlays}
        missing = [
            f"Dialog '{w.title or w.cls}' is outside the captured window, so it is not in "
            "the image; its text is in `dialogs`."
            for w in windows
            if sc.is_dialog_class(w.cls) and w.visible and not w.cloaked
            and w.hwnd not in drawn and w.frame.intersect(base.frame) is None
        ]
        meta["window"]["responsive"] = capture.responsive
        meta["image"] = capture.image_meta()
        meta["warnings"] = warnings + missing + capture.warnings
        return Outcome(meta=meta, png=capture.png)


def capture_application_meta(
    pid: int, max_dimension: int, include_image: bool
) -> tuple[dict[str, Any], bytes | None]:
    """A gate-free application capture to attach to an error, never raising."""
    try:
        outcome = capture_process(
            pid, "application", max_dimension=max_dimension, include_image=include_image
        )
        return outcome.meta, outcome.png
    except ScreenshotError as exc:
        meta = dict(exc.extra)
        meta["warnings"] = [*meta.get("warnings", []), f"Application capture: {exc}"]
        return meta, None
    except Exception as exc:
        return {"warnings": [f"Application capture failed: {exc}"]}, None


# ---------------------------------------------------------------------------
# COM object capture
# ---------------------------------------------------------------------------


def _open_timeout() -> float:
    try:
        value = float(
            os.environ.get("ACCESS_VCS_SCREENSHOT_OPEN_TIMEOUT_SEC", DEFAULT_OPEN_TIMEOUT_SEC)
        )
    except ValueError:
        return DEFAULT_OPEN_TIMEOUT_SEC
    return value if value > 0 else DEFAULT_OPEN_TIMEOUT_SEC


def access_error_number(exc: BaseException) -> int | None:
    """Access run-time error number carried by a ``com_error``, if any."""
    excepinfo = getattr(exc, "excepinfo", None)
    candidates: list[int] = []
    if excepinfo and len(excepinfo) > 5 and isinstance(excepinfo[5], int):
        candidates.append(excepinfo[5])
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        candidates.append(hresult)
    for code in candidates:
        if (code & 0xFFFF0000) == 0x800A0000:
            return code & 0xFFFF
    if excepinfo and len(excepinfo) > 2 and isinstance(excepinfo[2], str):
        for number in (ERR_NO_ACTIVE_FORM, ERR_NO_ACTIVE_REPORT, ERR_NO_ACTIVE_DATASHEET,
                       ERR_OBJECT_NOT_FOUND, ERR_OPEN_CANCELED):
            if str(number) in excepinfo[2]:
                return number
    return None


def normalize_hwnd(value: Any) -> int:
    return int(value) & 0xFFFFFFFF


def win32_precheck(pid: int) -> ScreenshotError | None:
    """Refuse COM when Access is visibly blocked. Win32 only.

    A COM call queued into an Access that is showing a dialog runs once the
    user answers it, long after this tool has returned; an OpenForm that
    runs then leaves an object open that nobody closes.
    """
    with sc.dpi_aware():
        windows = sc.enum_windows({pid})
    main = sc.find_window(windows, sc.MAIN_WINDOW_CLASS)
    if main is None:
        return ScreenshotError("database_not_open", f"Access PID {pid} has no main window.")
    if sc.vbe_state(windows) == "break":
        return ScreenshotError(
            "access_blocked",
            "VBA is in break mode, so no COM call was made. The application screenshot "
            "shows the state; see agent_guidance.",
            vbe_state="break",
        )
    dialog = next((w for w in windows if w.visible and sc.is_dialog_class(w.cls)), None)
    if not main.enabled or main.hung or dialog is not None:
        reason = (
            f"a dialog ('{dialog.title or dialog.cls}')" if dialog is not None
            else "a hang" if main.hung else "a modal window"
        )
        return ScreenshotError(
            "access_blocked",
            f"Access is blocked by {reason}, so no COM call was made. The application "
            "screenshot shows what it is waiting on.",
        )
    return None


_pending_lock = threading.Lock()
_pending_cleanup: dict[str, threading.Thread] = {}


def _db_key(database_path: str) -> str:
    return os.path.normcase(os.path.abspath(database_path))


def cleanup_pending(database_path: str) -> bool:
    with _pending_lock:
        worker = _pending_cleanup.get(_db_key(database_path))
        if worker is not None and not worker.is_alive():
            del _pending_cleanup[_db_key(database_path)]
            worker = None
        return worker is not None


def _register_pending(database_path: str, worker: threading.Thread) -> None:
    with _pending_lock:
        _pending_cleanup[_db_key(database_path)] = worker


@dataclass
class ObjectRequest:
    database_path: str
    pid: int
    target: str
    object_type: str | None
    object_name: str | None
    open_if_closed: bool
    max_dimension: int
    include_image: bool
    view: str = "normal"
    settle_ms: int = DEFAULT_SETTLE_MS


class _Progress:
    """Phase reporting from the COM worker, so each step gets its own deadline."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.phase = "attach"
        self.started = time.monotonic()
        self.abandoned = threading.Event()
        self.opened = False

    def enter(self, phase: str) -> None:
        with self._lock:
            self.phase = phase
            self.started = time.monotonic()

    def snapshot(self) -> tuple[str, float]:
        with self._lock:
            return self.phase, time.monotonic() - self.started


def _phase_limit(phase: str) -> float:
    if phase == "open":
        return _open_timeout()
    if phase == "capture":
        return CAPTURE_STEP_TIMEOUT_SEC
    return COM_STEP_TIMEOUT_SEC


def _find_app(database_path: str) -> Any:
    from .addin_integration import VCSAddinIntegration

    return VCSAddinIntegration._find_access_in_rot(database_path)


def _active_object(app: Any) -> tuple[Any, str, str]:
    """The active form, report, or datasheet. Form first, so a focused
    subform datasheet resolves to the whole form."""
    probes = (
        ("ActiveForm", "form", ERR_NO_ACTIVE_FORM),
        ("ActiveReport", "report", ERR_NO_ACTIVE_REPORT),
        ("ActiveDatasheet", "datasheet", ERR_NO_ACTIVE_DATASHEET),
    )
    for attribute, kind, _number in probes:
        try:
            obj = getattr(app.Screen, attribute)
        except Exception as exc:
            if access_error_number(exc) in (
                ERR_NO_ACTIVE_FORM, ERR_NO_ACTIVE_REPORT, ERR_NO_ACTIVE_DATASHEET
            ):
                continue
            raise
        if obj is not None:
            return obj, kind, f"Screen.{attribute}"
    raise ScreenshotError(
        "no_active_object",
        "No form, report, or datasheet is active. Use target='application' or "
        "target='object'.",
    )


def _access_object(app: Any, object_type: str, name: str) -> Any:
    collections = {
        "form": lambda: app.CurrentProject.AllForms,
        "report": lambda: app.CurrentProject.AllReports,
        "table": lambda: app.CurrentData.AllTables,
        "query": lambda: app.CurrentData.AllQueries,
    }
    try:
        return collections[object_type]().Item(name)
    except Exception as exc:
        if access_error_number(exc) == ERR_OBJECT_NOT_FOUND or "2467" in str(exc):
            raise ScreenshotError(
                "object_not_found", f"No {object_type} named '{name}' exists."
            ) from exc
        raise


def _current_view(obj: Any) -> str | None:
    try:
        value = int(obj.CurrentView)
    except Exception:
        return None
    return CURRENT_VIEW_NAMES.get(value, str(value))


def _datasheet_hwnd(app: Any, main_hwnd: int) -> int:
    try:
        return normalize_hwnd(app.Screen.ActiveDatasheet.Hwnd)
    except Exception as exc:
        if access_error_number(exc) != ERR_NO_ACTIVE_DATASHEET:
            raise
    with sc.dpi_aware():
        hwnd = sc.mdi_active_child(main_hwnd)
    if not hwnd:
        raise ScreenshotError(
            "capture_failed", "The object is open but no window for it could be found."
        )
    return hwnd


def _object_session(request: ObjectRequest, progress: _Progress, box: dict[str, Any]) -> None:
    """The COM half of an object capture, run on its own worker thread."""
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    ac_type = AC_OBJECT_TYPES.get(request.object_type or "", AC_FORM)
    name = request.object_name or ""
    object_meta: dict[str, Any] = {
        "type": request.object_type,
        "name": name,
        "source": None,
        "is_popup": False,
        "activated_by_tool": False,
        "opened_by_tool": False,
        "closed": False,
        "cleanup_pending": False,
    }
    previous: tuple[int, str] | None = None
    try:
        app = _find_app(request.database_path)
        if app is None:
            raise ScreenshotError("database_not_open", "The database is not open in Access.")
        app_pid = pid_from_access_app(app)
        if app_pid != request.pid:
            raise ScreenshotError(
                "database_not_open",
                f"The database is open in Access PID {app_pid}, not PID {request.pid}.",
            )
        with sc.dpi_aware():
            main = sc.find_window(sc.enum_windows({request.pid}), sc.MAIN_WINDOW_CLASS)
        main_hwnd = main.hwnd if main else 0

        progress.enter("resolve")
        if request.target == "active":
            obj, kind, source = _active_object(app)
            object_meta.update(type=kind, name=str(obj.Name), source=source)
            hwnd = normalize_hwnd(obj.Hwnd)
        else:
            access_obj = _access_object(app, request.object_type, name)
            if access_obj.IsLoaded:
                if request.object_type == "form":
                    form = app.Forms(name)
                    if not form.Visible:
                        raise ScreenshotError(
                            "object_hidden", f"The form '{name}' is open but hidden."
                        )
                    hwnd = normalize_hwnd(form.Hwnd)
                    object_meta["source"] = "Forms"
                    object_meta["current_view"] = _current_view(form)
                elif request.object_type == "report":
                    report = app.Reports(name)
                    hwnd = normalize_hwnd(report.Hwnd)
                    object_meta["source"] = "Reports"
                    object_meta["current_view"] = _current_view(report)
                else:
                    previous = (int(app.CurrentObjectType), str(app.CurrentObjectName))
                    app.DoCmd.SelectObject(ac_type, name, False)
                    object_meta["activated_by_tool"] = True
                    hwnd = _datasheet_hwnd(app, main_hwnd)
                    object_meta["source"] = "Screen.ActiveDatasheet"
                current = object_meta.get("current_view")
                if request.view == "design" and current not in (None, "design"):
                    object_meta["note"] = (
                        f"The {request.object_type} was already open in {current} view, so it "
                        "was captured as it is. view='design' applies only to objects this "
                        "tool opens."
                    )
                if (
                    request.object_type in ("form", "report")
                    and main_hwnd
                    and not sc.is_top_level(hwnd)
                ):
                    with sc.dpi_aware():
                        active = sc.mdi_active_child(main_hwnd)
                    if active != hwnd:
                        previous = (int(app.CurrentObjectType), str(app.CurrentObjectName))
                        app.DoCmd.SelectObject(ac_type, name, False)
                        object_meta["activated_by_tool"] = True
            elif not request.open_if_closed:
                raise ScreenshotError(
                    "object_not_open",
                    f"The {request.object_type} '{name}' is not open. Pass open_if_closed=True "
                    "to open it for the capture; that runs its event code.",
                )
            else:
                if request.object_type == "query":
                    query_type = int(app.CurrentDb().QueryDefs(name).Type)
                    if query_type not in SAFE_QUERY_TYPES:
                        raise ScreenshotError(
                            "unsupported_query_type",
                            f"Query '{name}' is not a select, crosstab, or union query "
                            f"(DAO type {query_type}); opening it could change data.",
                        )
                progress.enter("open")
                design = request.view == "design"
                try:
                    if request.object_type == "form":
                        app.DoCmd.OpenForm(name, AC_DESIGN if design else AC_NORMAL)
                    elif request.object_type == "report":
                        app.DoCmd.OpenReport(name, AC_VIEW_DESIGN if design else AC_VIEW_PREVIEW)
                    elif request.object_type == "table":
                        app.DoCmd.OpenTable(name, AC_VIEW_NORMAL, AC_READ_ONLY)
                    else:
                        app.DoCmd.OpenQuery(name, AC_VIEW_NORMAL, AC_READ_ONLY)
                except Exception as exc:
                    if access_error_number(exc) == ERR_OPEN_CANCELED:
                        raise ScreenshotError(
                            "open_canceled",
                            f"Opening '{name}' was canceled by the object itself "
                            "(for example a report with no data).",
                        ) from exc
                    raise
                progress.opened = True
                object_meta["opened_by_tool"] = True
                object_meta["view"] = request.view
                if progress.abandoned.is_set():
                    return
                progress.enter("settle")
                try:
                    app.DoCmd.RepaintObject(ac_type, name)
                except Exception:
                    pass
                time.sleep(request.settle_ms / 1000)
                if request.object_type == "form":
                    hwnd = normalize_hwnd(app.Forms(name).Hwnd)
                elif request.object_type == "report":
                    hwnd = normalize_hwnd(app.Reports(name).Hwnd)
                else:
                    hwnd = _datasheet_hwnd(app, main_hwnd)
                object_meta["source"] = "opened"

        progress.enter("capture")
        with sc.dpi_aware():
            object_meta["is_popup"] = sc.is_top_level(hwnd)
            if not object_meta["is_popup"] and not sc.is_visible(hwnd):
                raise ScreenshotError(
                    "object_hidden", "The object's window is not visible."
                )
        outcome = capture_process(
            request.pid,
            "application",
            max_dimension=request.max_dimension,
            include_image=request.include_image,
            object_hwnd=hwnd,
        )
        box["outcome"] = outcome
    except ScreenshotError as exc:
        box["error"] = exc
    except Exception as exc:
        box["error"] = ScreenshotError("capture_failed", f"Object capture failed: {exc}")
    finally:
        progress.enter("cleanup")
        if app is not None:
            if object_meta["opened_by_tool"]:
                try:
                    app.DoCmd.Close(ac_type, name, AC_SAVE_NO)
                    object_meta["closed"] = True
                except Exception:
                    pass
            elif previous is not None and previous[1]:
                try:
                    app.DoCmd.SelectObject(previous[0], previous[1], False)
                except Exception:
                    pass
        box["object"] = object_meta
        app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def capture_object_sync(request: ObjectRequest) -> dict[str, Any]:
    """Object capture, run on the gate's thread. Returns a plain result dict.

    ``fallback: True`` asks the caller to attach an application capture,
    taken after the gate is released.
    """
    if cleanup_pending(request.database_path):
        return {
            "fallback": True,
            "error": ScreenshotError(
                "cleanup_pending",
                "An object opened by an earlier capture is still waiting to be closed, "
                "usually behind a dialog. Answer the dialog first.",
            ),
        }
    blocked = win32_precheck(request.pid)
    if blocked is not None:
        return {"fallback": True, "error": blocked}

    progress = _Progress()
    box: dict[str, Any] = {}
    worker = threading.Thread(
        target=_object_session, args=(request, progress, box), daemon=True,
        name="vcs-screenshot-com",
    )
    worker.start()
    while worker.is_alive():
        worker.join(0.05)
        phase, elapsed = progress.snapshot()
        if worker.is_alive() and elapsed > _phase_limit(phase):
            progress.abandoned.set()
            if progress.opened or phase in ("open", "settle", "capture"):
                _register_pending(request.database_path, worker)
            object_meta = {
                "type": request.object_type,
                "name": request.object_name,
                "opened_by_tool": progress.opened,
                "closed": False,
                "cleanup_pending": phase in ("open", "settle", "capture") or progress.opened,
            }
            if phase == "open":
                error = ScreenshotError(
                    "open_blocked",
                    f"Opening the object did not finish within {_phase_limit(phase):.0f} s, "
                    "usually because its Open or Load event is showing a dialog, or, in "
                    "design view, because Access is warning that the database is not open "
                    "exclusively. If the object opens once the dialog is answered, this "
                    "server closes it without saving.",
                    object=object_meta,
                )
            else:
                error = ScreenshotError(
                    "access_unresponsive",
                    f"Access did not answer during '{phase}' within "
                    f"{_phase_limit(phase):.0f} s.",
                    object=object_meta,
                )
            return {"fallback": True, "error": error}
    if "error" in box:
        error = box["error"]
        error.extra.setdefault("object", box.get("object"))
        return {"fallback": False, "error": error}
    outcome: Outcome = box["outcome"]
    outcome.meta["object"] = box.get("object")
    return {"fallback": False, "outcome": outcome}


# ---------------------------------------------------------------------------
# Tool entry point
# ---------------------------------------------------------------------------


def validate_arguments(
    database_path: str | None,
    target: str,
    object_type: str | None,
    object_name: str | None,
    open_if_closed: bool,
    pid: int | None,
    max_dimension: int,
    view: str = "normal",
    settle_ms: int = DEFAULT_SETTLE_MS,
) -> None:
    if target not in TARGETS:
        raise ScreenshotError("invalid_argument", f"target must be one of {', '.join(TARGETS)}.")
    if not MIN_DIMENSION <= max_dimension <= MAX_DIMENSION:
        raise ScreenshotError(
            "invalid_argument",
            f"max_dimension must be between {MIN_DIMENSION} and {MAX_DIMENSION}.",
        )
    if open_if_closed and target != "object":
        raise ScreenshotError("invalid_argument", "open_if_closed applies only to target='object'.")
    if target in ("active", "object"):
        if not database_path:
            raise ScreenshotError("invalid_argument", f"target='{target}' needs database_path.")
        if pid is not None:
            raise ScreenshotError(
                "invalid_argument", f"target='{target}' finds the process from database_path; omit pid."
            )
    if target == "object":
        if object_type not in OBJECT_TYPES:
            raise ScreenshotError(
                "invalid_argument", f"object_type must be one of {', '.join(OBJECT_TYPES)}."
            )
        if not object_name:
            raise ScreenshotError("invalid_argument", "target='object' needs object_name.")
    if view not in VIEWS:
        raise ScreenshotError("invalid_argument", f"view must be one of {', '.join(VIEWS)}.")
    if view == "design" and (target != "object" or object_type not in ("form", "report")):
        raise ScreenshotError(
            "invalid_argument", "view='design' applies only to target='object' forms and reports."
        )
    if not 0 <= settle_ms <= MAX_SETTLE_MS:
        raise ScreenshotError(
            "invalid_argument", f"settle_ms must be between 0 and {MAX_SETTLE_MS}."
        )


async def _in_thread(fn: Callable[[], Any], timeout: float = CAPTURE_TIMEOUT_SEC) -> Any:
    try:
        return await asyncio.wait_for(asyncio.to_thread(fn), timeout)
    except asyncio.TimeoutError as exc:
        raise ScreenshotError(
            "capture_timeout", f"The capture did not finish within {timeout:.0f} s."
        ) from exc


def _add_warning(meta: dict[str, Any], warning: str) -> None:
    meta["warnings"] = [*(meta.get("warnings") or []), warning]


def _add_retention(meta: dict[str, Any], image: dict[str, Any]) -> None:
    """When the saved file expires, and who should decide that if nobody has."""
    days = screenshot_store.retention_days()
    decided = user_settings.is_decided(user_settings.SCREENSHOT_RETENTION.key)
    expires = None if days is None else datetime.now().astimezone() + timedelta(days=days)
    image["retention"] = {
        "days": None if days is None else int(days),
        "forever": days is None,
        "expires_at": expires.isoformat(timespec="seconds") if expires else None,
        "decided": decided,
    }
    if not decided and not meta.get("agent_guidance"):
        meta["agent_guidance"] = RETENTION_UNDECIDED_GUIDANCE


async def _finish(
    meta: dict[str, Any],
    png: bytes | None,
    *,
    include_image: bool,
    save_image: bool,
    captured: str,
) -> CallToolResult:
    """Save the PNG when asked, then build the result; inline only when asked.

    ``captured`` is the target the pixels show, which for a fallback is
    ``application`` whatever was requested.
    """
    if save_image:
        if png and not screenshot_store.saving_allowed():
            _add_warning(meta, SAVE_NEVER_WARNING)
            folder = screenshot_store.screenshot_dir()
            await _in_thread(lambda: screenshot_store.prune_expired(folder, 0))
        elif png:
            try:
                path = await _in_thread(lambda: screenshot_store.save_and_prune(png, captured))
            except Exception as exc:
                _add_warning(meta, f"The screenshot could not be saved: {exc}")
            else:
                image = meta.setdefault("image", {})
                image["saved_path"] = str(path)
                image["markdown"] = screenshot_store.markdown_link(path, captured)
                image["target"] = captured
                _add_retention(meta, image)
        elif meta.get("error_pattern") != "invalid_argument":
            _add_warning(meta, "save_image: no image was captured, so nothing was saved.")
    return build_result(meta, png if include_image else None)


async def _with_application_capture(
    error: ScreenshotError, base: dict[str, Any], pid: int, max_dimension: int,
    include_image: bool, save_image: bool = False,
) -> CallToolResult:
    """An error result carrying a gate-free application capture and dialogs."""
    try:
        meta, png = await _in_thread(
            lambda: capture_application_meta(pid, max_dimension, include_image or save_image)
        )
    except ScreenshotError as exc:
        meta, png = {"warnings": [str(exc)]}, None
    result = dict(base)
    result.update(meta)
    result.update(error.meta())
    return await _finish(
        result, png, include_image=include_image, save_image=save_image, captured="application"
    )


async def capture_screenshot(
    *,
    database_path: str | None,
    target: str,
    object_type: str | None,
    object_name: str | None,
    open_if_closed: bool,
    pid: int | None,
    max_dimension: int,
    include_image: bool,
    get_gate: Callable[[], Any],
    view: str = "normal",
    settle_ms: int = DEFAULT_SETTLE_MS,
    save_image: bool = False,
) -> CallToolResult:
    base: dict[str, Any] = {"target": target}
    pixels = include_image or save_image

    async def finish(meta: dict[str, Any], png: bytes | None = None) -> CallToolResult:
        return await _finish(
            meta, png, include_image=include_image, save_image=save_image, captured=target
        )

    try:
        validate_arguments(
            database_path, target, object_type, object_name, open_if_closed, pid, max_dimension,
            view, settle_ms,
        )
        resolved_pid, source = await _in_thread(lambda: resolve_target_pid(database_path, pid))
    except ScreenshotError as exc:
        return await finish({**base, **exc.meta()})
    owned = await _in_thread(lambda: server_owns(resolved_pid))
    base.update(pid=resolved_pid, pid_source=source, owned_by_server=owned)

    if target in ("application", "vbe"):
        try:
            outcome = await _in_thread(
                lambda: capture_process(
                    resolved_pid, target, max_dimension=max_dimension, include_image=pixels
                )
            )
        except ScreenshotError as exc:
            return await finish({**base, **exc.meta()})
        except Exception as exc:
            return await finish({**base, **ScreenshotError("capture_failed", str(exc)).meta()})
        return await finish({"success": True, **base, **outcome.meta}, outcome.png)

    gate = get_gate()
    in_flight = gate.current_in_flight(str(database_path))
    if in_flight is not None:
        busy = ScreenshotError(
            "server_busy",
            f"Another Access operation is in progress on this database ({in_flight.tool}). "
            "Object captures wait for it; the application capture below shows what Access "
            "is doing.",
            busy_with={"tool": in_flight.tool, "database": in_flight.database},
        )
        return await _with_application_capture(
            busy, base, resolved_pid, max_dimension, include_image, save_image
        )

    request = ObjectRequest(
        database_path=str(database_path),
        pid=resolved_pid,
        target=target,
        object_type=object_type,
        object_name=object_name,
        open_if_closed=open_if_closed,
        max_dimension=max_dimension,
        include_image=pixels,
        view=view,
        settle_ms=settle_ms,
    )
    result = await gate.run_exclusive(
        "vcs_capture_screenshot", str(database_path), capture_object_sync, False, request
    )
    if isinstance(result, dict) and result.get("error_pattern") == "server_busy":
        busy = ScreenshotError(
            "server_busy", result.get("error", "Server busy."), busy_with=result.get("busy_with")
        )
        return await _with_application_capture(
            busy, base, resolved_pid, max_dimension, include_image, save_image
        )
    if result.get("fallback"):
        return await _with_application_capture(
            result["error"], base, resolved_pid, max_dimension, include_image, save_image
        )
    if "error" in result:
        return await finish({**base, **result["error"].meta()})
    outcome = result["outcome"]
    return await finish({"success": True, **base, **outcome.meta}, outcome.png)
