"""Tests for vcs_capture_screenshot orchestration.

Process resolution, gate routing, and the object-open safety rules run
against fakes: a fake Access application object, patched Win32 window
lists, and a stub gate. Nothing here needs Access or a desktop.
"""

from __future__ import annotations

import asyncio
import json
import threading
import time
from dataclasses import dataclass
from types import SimpleNamespace

import pytest
from mcp.types import CallToolResult

import msaccess_vcs_mcp.screen_capture as sc
import msaccess_vcs_mcp.screenshot as shot
from msaccess_vcs_mcp.screen_capture import Rect, WindowInfo

DB = r"C:\path\to\example.accdb"
OTHER_DB = r"C:\path\to\other.accdb"
PNG = b"\x89PNG\r\n\x1a\nfake"


def meta_of(result: CallToolResult) -> dict:
    assert result.content[0].type == "text"
    return json.loads(result.content[0].text)


# ---------------------------------------------------------------------------
# Process resolution
# ---------------------------------------------------------------------------


@dataclass
class _Record:
    pid: int
    database_path: str


@pytest.fixture
def processes(monkeypatch):
    """Configure the processes resolve_target_pid sees."""
    state = SimpleNamespace(
        survey={}, tasklist=[], records=[], holders=None, rm_calls=[],
    )

    def holders(path):
        state.rm_calls.append(path)
        return state.holders

    monkeypatch.setattr(shot, "survey_access", lambda: state.survey)
    monkeypatch.setattr(shot, "list_access_pids_or_none", lambda: state.tasklist)
    monkeypatch.setattr(shot, "_owned_records", lambda: state.records)
    monkeypatch.setattr(shot, "paths_match", lambda a, b: a.lower() == b.lower())
    monkeypatch.setattr(sc, "pids_holding_file", holders)
    return state


def _survey(**titles):
    return {int(pid[1:]): {"title": t, "hung": False, "vbe_open": False} for pid, t in titles.items()}


def test_pid_argument_wins(processes):
    processes.tasklist = [10, 20]
    assert shot.resolve_target_pid(DB, 20) == (20, "argument")
    assert processes.rm_calls == []


def test_pid_argument_must_be_running_access(processes):
    processes.tasklist = [10]
    with pytest.raises(shot.ScreenshotError) as info:
        shot.resolve_target_pid(None, 99)
    assert info.value.error_pattern == "invalid_pid"
    assert [i["pid"] for i in info.value.extra["instances"]] == [10]


def test_restart_manager_first(processes):
    processes.tasklist = [10, 20]
    processes.holders = {20, 4444}
    processes.records = [_Record(10, DB)]
    assert shot.resolve_target_pid(DB, None) == (20, "restart_manager")


def test_restart_manager_ambiguous(processes):
    processes.tasklist = [10, 20]
    processes.holders = {10, 20}
    with pytest.raises(shot.ScreenshotError) as info:
        shot.resolve_target_pid(DB, None)
    assert info.value.error_pattern == "ambiguous_instance"


def test_owned_registry_when_restart_manager_is_blind(processes):
    processes.tasklist = [10, 20]
    processes.holders = set()
    processes.records = [_Record(10, OTHER_DB), _Record(20, DB)]
    assert shot.resolve_target_pid(DB, None) == (20, "owned_registry")


def test_owned_registry_ignores_dead_processes(processes):
    processes.tasklist = [10]
    processes.records = [_Record(20, DB)]
    processes.survey = _survey(p10="example : Database- C:\\path\\to\\example.accdb")
    assert shot.resolve_target_pid(DB, None) == (10, "window_title")


def test_window_title(processes):
    processes.tasklist = [10, 20]
    processes.survey = _survey(p10="Access - other.accdb", p20="Access - example.accdb")
    assert shot.resolve_target_pid(DB, None) == (20, "window_title")


def test_database_path_never_falls_back_to_the_only_instance(processes):
    """An unrelated database's window must not pass for the one asked for."""
    processes.tasklist = [10]
    processes.survey = _survey(p10="Access - other.accdb")
    with pytest.raises(shot.ScreenshotError) as info:
        shot.resolve_target_pid(DB, None)
    assert info.value.error_pattern == "database_not_open"
    assert info.value.extra["instances"][0]["pid"] == 10


def test_single_instance_without_database_path(processes):
    processes.tasklist = [10]
    assert shot.resolve_target_pid(None, None) == (10, "single_instance")


def test_no_access(processes):
    with pytest.raises(shot.ScreenshotError) as info:
        shot.resolve_target_pid(None, None)
    assert info.value.error_pattern == "access_not_running"


def test_several_instances_need_a_choice(processes):
    processes.tasklist = [10, 20]
    processes.survey = _survey(p10="A", p20="B")
    processes.records = [_Record(20, OTHER_DB)]
    with pytest.raises(shot.ScreenshotError) as info:
        shot.resolve_target_pid(None, None)
    assert info.value.error_pattern == "ambiguous_instance"
    instances = info.value.extra["instances"]
    assert [(i["pid"], i["title"], i["owned_by_server"]) for i in instances] == [
        (10, "A", False), (20, "B", True)
    ]


def test_failed_tasklist_uses_window_survey(processes):
    processes.tasklist = None
    processes.survey = _survey(p30="Access")
    assert shot.resolve_target_pid(None, None) == (30, "single_instance")


# ---------------------------------------------------------------------------
# Arguments and error numbers
# ---------------------------------------------------------------------------


def _validate(**overrides):
    args = dict(
        database_path=DB, target="application", object_type=None, object_name=None,
        open_if_closed=False, pid=None, max_dimension=1568,
    )
    args.update(overrides)
    shot.validate_arguments(**args)


@pytest.mark.parametrize(
    "overrides",
    [
        {"target": "window"},
        {"max_dimension": 100},
        {"max_dimension": 5000},
        {"open_if_closed": True},
        {"target": "active", "database_path": None},
        {"target": "active", "pid": 10},
        {"target": "object", "object_type": "macro", "object_name": "x"},
        {"target": "object", "object_type": "form", "object_name": ""},
    ],
)
def test_invalid_arguments(overrides):
    with pytest.raises(shot.ScreenshotError) as info:
        _validate(**overrides)
    assert info.value.error_pattern == "invalid_argument"


def test_valid_arguments():
    _validate()
    _validate(database_path=None, pid=10, target="vbe")
    _validate(target="object", object_type="query", object_name="qryExample", open_if_closed=True)


class _ComError(Exception):
    def __init__(self, hresult=None, excepinfo=None):
        super().__init__("com error")
        self.hresult = hresult
        self.excepinfo = excepinfo


def vba_hresult(number: int) -> int:
    """The signed HRESULT pywin32 reports for an Access run-time error."""
    return (0x800A0000 | number) - 0x100000000


def test_access_error_number():
    assert shot.access_error_number(_ComError(excepinfo=(0, "Access", "", None, 0, vba_hresult(2501)))) == 2501
    assert shot.access_error_number(_ComError(hresult=vba_hresult(2475))) == 2475
    assert shot.access_error_number(
        _ComError(excepinfo=(0, "Access", "Error 2467: object not found", None, 0, 0))
    ) == 2467
    assert shot.access_error_number(ValueError("x")) is None


def test_normalize_hwnd_handles_negative_handles():
    assert shot.normalize_hwnd(-2) == 0xFFFFFFFE


def test_open_timeout_env(monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_SCREENSHOT_OPEN_TIMEOUT_SEC", "3.5")
    assert shot._open_timeout() == 3.5
    monkeypatch.setenv("ACCESS_VCS_SCREENSHOT_OPEN_TIMEOUT_SEC", "nope")
    assert shot._open_timeout() == shot.DEFAULT_OPEN_TIMEOUT_SEC


# ---------------------------------------------------------------------------
# Win32 capture of a process
# ---------------------------------------------------------------------------


class _NullContext:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _win(hwnd, cls, *, z=5, owner=0, **kwargs):
    frame = kwargs.pop("frame", Rect(0, 0, 800, 600))
    return WindowInfo(hwnd=hwnd, pid=10, cls=cls, title=kwargs.pop("title", cls), rect=frame,
                      frame=frame, owner=owner, z=z, **kwargs)


@pytest.fixture
def desktop(monkeypatch):
    state = SimpleNamespace(windows=[], composites=[])

    def composite(base, overlays, all_windows, *, crop=None, max_dimension=1568):
        state.composites.append(SimpleNamespace(base=base, overlays=overlays, crop=crop))
        return sc.Capture(png=PNG, width=10, height=10, original_width=10, original_height=10,
                          scale=1.0, method="print_window")

    monkeypatch.setattr(sc, "dpi_aware", _NullContext)
    monkeypatch.setattr(sc, "enum_windows", lambda pids=None, visible_only=False: state.windows)
    monkeypatch.setattr(sc, "is_responsive", lambda hwnd: True)
    monkeypatch.setattr(sc, "capture_composite", composite)
    monkeypatch.setattr(sc, "read_child_controls", lambda hwnd: [sc.ChildControl("Static", "Hello")])
    return state


def test_capture_application_lists_dialogs(desktop):
    main = _win(0x100, sc.MAIN_WINDOW_CLASS, enabled=False)
    box = _win(1, "#32770", owner=0x100, z=0, title="Example")
    desktop.windows = [box, main]

    outcome = shot.capture_process(10, "application")

    assert outcome.png == PNG
    assert outcome.meta["modal_blocked"] is True
    assert outcome.meta["dialogs"][0]["text"] == "Hello"
    assert outcome.meta["agent_guidance"] == shot.DIALOG_GUIDANCE
    assert [w.hwnd for w in desktop.composites[0].overlays] == [1]


def test_capture_without_image_skips_pixels(desktop):
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS)]
    outcome = shot.capture_process(10, "application", include_image=False)
    assert outcome.png is None
    assert "image" not in outcome.meta
    assert desktop.composites == []


def test_vbe_not_open(desktop):
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS), _win(0x200, sc.VBE_WINDOW_CLASS, visible=False)]
    with pytest.raises(shot.ScreenshotError) as info:
        shot.capture_process(10, "vbe")
    assert info.value.error_pattern == "vbe_not_open"


@pytest.mark.parametrize(
    "state,pattern", [({"minimized": True}, "window_minimized"), ({"visible": False}, "window_hidden")]
)
def test_minimized_or_hidden_window_is_not_touched(desktop, state, pattern):
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS, **state)]
    with pytest.raises(shot.ScreenshotError) as info:
        shot.capture_process(10, "application")
    assert info.value.error_pattern == pattern
    assert desktop.composites == []


def test_popup_object_is_its_own_base(desktop, monkeypatch):
    main = _win(0x100, sc.MAIN_WINDOW_CLASS)
    popup = _win(0x300, "OForm", owner=0x100, z=1, frame=Rect(50, 50, 250, 200))
    desktop.windows = [popup, main]
    monkeypatch.setattr(sc, "is_top_level", lambda hwnd: True)

    shot.capture_process(10, "application", object_hwnd=0x300)

    assert desktop.composites[0].base is popup
    assert desktop.composites[0].crop is None


def test_mdi_child_object_crops_to_client_area(desktop, monkeypatch):
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS)]
    monkeypatch.setattr(sc, "is_top_level", lambda hwnd: False)
    monkeypatch.setattr(sc, "window_rect", lambda hwnd: Rect(-20, 100, 400, 500))
    monkeypatch.setattr(sc, "mdi_client_rect", lambda hwnd: Rect(0, 90, 800, 580))

    shot.capture_process(10, "application", object_hwnd=0x400)

    assert desktop.composites[0].crop == Rect(0, 100, 400, 500)


def test_win32_precheck(desktop):
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS)]
    assert shot.win32_precheck(10) is None
    desktop.windows = [_win(1, "#32770", owner=0x100, z=0, title="Example"),
                       _win(0x100, sc.MAIN_WINDOW_CLASS)]
    assert shot.win32_precheck(10).error_pattern == "access_blocked"
    desktop.windows = [_win(0x100, sc.MAIN_WINDOW_CLASS, hung=True)]
    assert "hang" in str(shot.win32_precheck(10))


# ---------------------------------------------------------------------------
# Object capture against a fake Access
# ---------------------------------------------------------------------------


class _Item:
    def __init__(self, loaded):
        self.IsLoaded = loaded


class _Collection:
    def __init__(self, items):
        self.items = items

    def Item(self, name):
        if name not in self.items:
            raise _ComError(excepinfo=(0, "Access", "Error 2467", None, 0, vba_hresult(2467)))
        return _Item(self.items[name])


class FakeDoCmd:
    def __init__(self, app):
        self.app = app
        self.calls: list[tuple] = []
        self.open_error: Exception | None = None
        self.open_delay = 0.0

    def _open(self, *call):
        self.calls.append(call)
        if self.open_delay:
            time.sleep(self.open_delay)
        if self.open_error:
            raise self.open_error
        self.app.loaded.add(call[1])

    def OpenForm(self, name, view):
        self._open("OpenForm", name, view)

    def OpenReport(self, name, view):
        self._open("OpenReport", name, view)

    def OpenTable(self, name, view, mode):
        self._open("OpenTable", name, view, mode)

    def OpenQuery(self, name, view, mode):
        self._open("OpenQuery", name, view, mode)

    def Close(self, ac_type, name, save):
        self.calls.append(("Close", ac_type, name, save))

    def SelectObject(self, ac_type, name, in_db_window):
        self.calls.append(("SelectObject", ac_type, name))

    def RepaintObject(self, ac_type, name):
        pass


class FakeAccess:
    def __init__(self, forms=(), reports=(), tables=(), queries=None, loaded=()):
        self.loaded = set(loaded)
        self.forms, self.reports, self.tables = set(forms), set(reports), set(tables)
        self.queries = queries or {}
        self.DoCmd = FakeDoCmd(self)
        self.CurrentObjectType = 2
        self.CurrentObjectName = "frmPrevious"
        self.Screen = SimpleNamespace(ActiveDatasheet=SimpleNamespace(Hwnd=0x500))

    def _items(self, names):
        return _Collection({n: n in self.loaded for n in names})

    @property
    def CurrentProject(self):
        return SimpleNamespace(AllForms=self._items(self.forms), AllReports=self._items(self.reports))

    @property
    def CurrentData(self):
        return SimpleNamespace(AllTables=self._items(self.tables), AllQueries=self._items(self.queries))

    def Forms(self, name):
        return SimpleNamespace(Hwnd=0x600, Visible=True, Name=name)

    def Reports(self, name):
        return SimpleNamespace(Hwnd=0x700, Name=name)

    def CurrentDb(self):
        defs = {n: SimpleNamespace(Type=t) for n, t in self.queries.items()}
        return SimpleNamespace(QueryDefs=defs.__getitem__)

    def opens(self):
        return [c for c in self.DoCmd.calls if c[0].startswith("Open")]

    def closes(self):
        return [c for c in self.DoCmd.calls if c[0] == "Close"]


@pytest.fixture
def fake_access(monkeypatch):
    holder = SimpleNamespace(app=FakeAccess(), captured=[])

    def capture(pid, target, *, max_dimension, include_image, object_hwnd=0):
        holder.captured.append(object_hwnd)
        return shot.Outcome(meta={"window": {"title": "Access"}}, png=PNG)

    monkeypatch.setattr(shot, "_find_app", lambda path: holder.app)
    monkeypatch.setattr(shot, "pid_from_access_app", lambda app: 10)
    monkeypatch.setattr(shot, "win32_precheck", lambda pid: None)
    monkeypatch.setattr(shot, "capture_process", capture)
    monkeypatch.setattr(sc, "dpi_aware", _NullContext)
    monkeypatch.setattr(sc, "enum_windows", lambda pids=None, visible_only=False: [
        _win(0x100, sc.MAIN_WINDOW_CLASS)
    ])
    monkeypatch.setattr(sc, "is_top_level", lambda hwnd: False)
    monkeypatch.setattr(sc, "is_visible", lambda hwnd: True)
    monkeypatch.setattr(sc, "mdi_active_child", lambda hwnd: 0x600)
    shot._pending_cleanup.clear()
    yield holder
    shot._pending_cleanup.clear()


def _request(object_type="form", name="frmExample", open_if_closed=False, target="object",
             view="normal", settle_ms=0):
    return shot.ObjectRequest(
        database_path=DB, pid=10, target=target, object_type=object_type, object_name=name,
        open_if_closed=open_if_closed, max_dimension=1568, include_image=True,
        view=view, settle_ms=settle_ms,
    )


@pytest.mark.parametrize(
    "object_type,name,expected_open",
    [
        ("form", "frmExample", ("OpenForm", "frmExample", shot.AC_DESIGN)),
        ("report", "rptExample", ("OpenReport", "rptExample", shot.AC_VIEW_DESIGN)),
    ],
)
def test_design_view_opens_in_design_and_closes_without_saving(
    fake_access, object_type, name, expected_open
):
    fake_access.app = FakeAccess(forms=["frmExample"], reports=["rptExample"])
    result = shot.capture_object_sync(
        _request(object_type, name, open_if_closed=True, view="design")
    )
    assert fake_access.app.opens() == [expected_open]
    ac_type = shot.AC_OBJECT_TYPES[object_type]
    assert fake_access.app.closes() == [("Close", ac_type, name, shot.AC_SAVE_NO)]
    assert result["outcome"].meta["object"]["view"] == "design"


def test_open_form_reports_current_view(fake_access, monkeypatch):
    fake_access.app = FakeAccess(forms=["frmExample"], loaded=["frmExample"])
    monkeypatch.setattr(
        fake_access.app, "Forms",
        lambda name: SimpleNamespace(Hwnd=0x600, Visible=True, Name=name, CurrentView=1),
    )
    result = shot.capture_object_sync(_request(view="design"))
    meta = result["outcome"].meta["object"]
    assert meta["current_view"] == "form"
    assert "already open in form view" in meta["note"]
    assert fake_access.app.opens() == []


def test_open_form_in_design_has_no_note(fake_access, monkeypatch):
    fake_access.app = FakeAccess(forms=["frmExample"], loaded=["frmExample"])
    monkeypatch.setattr(
        fake_access.app, "Forms",
        lambda name: SimpleNamespace(Hwnd=0x600, Visible=True, Name=name, CurrentView=0),
    )
    meta = shot.capture_object_sync(_request(view="design"))["outcome"].meta["object"]
    assert meta["current_view"] == "design"
    assert "note" not in meta


def test_settle_ms_waits_after_opening(fake_access, monkeypatch):
    waits = []
    monkeypatch.setattr(shot.time, "sleep", lambda seconds: waits.append(seconds))
    fake_access.app = FakeAccess(forms=["frmExample"])
    shot.capture_object_sync(_request(open_if_closed=True, settle_ms=1200))
    assert 1.2 in waits


@pytest.mark.parametrize("kwargs, message", [
    ({"view": "sideways"}, "view must be"),
    ({"view": "design", "object_type": "table"}, "forms and reports"),
    ({"settle_ms": -1}, "settle_ms"),
    ({"settle_ms": shot.MAX_SETTLE_MS + 1}, "settle_ms"),
])
def test_view_and_settle_validation(kwargs, message):
    args = {"view": "normal", "settle_ms": 0, "object_type": "form", **kwargs}
    with pytest.raises(shot.ScreenshotError, match=message):
        shot.validate_arguments(
            DB, "object", args["object_type"], "frmExample", True, None, 1568,
            args["view"], args["settle_ms"],
        )


def test_open_form_is_captured(fake_access):
    fake_access.app = FakeAccess(forms=["frmExample"], loaded=["frmExample"])
    result = shot.capture_object_sync(_request())
    assert "outcome" in result
    assert fake_access.captured == [0x600]
    assert fake_access.app.opens() == []
    assert fake_access.app.closes() == []
    assert result["outcome"].meta["object"]["opened_by_tool"] is False


def test_closed_object_is_not_opened_by_default(fake_access):
    fake_access.app = FakeAccess(forms=["frmExample"])
    result = shot.capture_object_sync(_request())
    assert result["error"].error_pattern == "object_not_open"
    assert fake_access.app.opens() == []


def test_missing_object(fake_access):
    fake_access.app = FakeAccess(forms=["frmOther"])
    result = shot.capture_object_sync(_request(open_if_closed=True))
    assert result["error"].error_pattern == "object_not_found"


@pytest.mark.parametrize(
    "object_type,name,expected_open",
    [
        ("form", "frmExample", ("OpenForm", "frmExample", shot.AC_NORMAL)),
        ("report", "rptExample", ("OpenReport", "rptExample", shot.AC_VIEW_PREVIEW)),
        ("table", "tblExample", ("OpenTable", "tblExample", shot.AC_VIEW_NORMAL, shot.AC_READ_ONLY)),
        ("query", "qryExample", ("OpenQuery", "qryExample", shot.AC_VIEW_NORMAL, shot.AC_READ_ONLY)),
    ],
)
def test_open_if_closed_opens_safely_and_closes_without_saving(
    fake_access, object_type, name, expected_open
):
    fake_access.app = FakeAccess(
        forms=["frmExample"], reports=["rptExample"], tables=["tblExample"],
        queries={"qryExample": shot.DB_Q_SELECT},
    )
    result = shot.capture_object_sync(_request(object_type, name, open_if_closed=True))

    assert "outcome" in result, result
    assert fake_access.app.opens() == [expected_open]
    ac_type = shot.AC_OBJECT_TYPES[object_type]
    assert fake_access.app.closes() == [("Close", ac_type, name, shot.AC_SAVE_NO)]
    meta = result["outcome"].meta["object"]
    assert meta["opened_by_tool"] is True and meta["closed"] is True


@pytest.mark.parametrize("query_type", [32, 48, 64, 80, 96, 112, 240])
def test_action_queries_are_never_opened(fake_access, query_type):
    fake_access.app = FakeAccess(queries={"qryExample": query_type})
    result = shot.capture_object_sync(_request("query", "qryExample", open_if_closed=True))
    assert result["error"].error_pattern == "unsupported_query_type"
    assert fake_access.app.opens() == []


def test_canceled_open(fake_access):
    fake_access.app = FakeAccess(reports=["rptExample"])
    fake_access.app.DoCmd.open_error = _ComError(
        excepinfo=(0, "Access", "", None, 0, vba_hresult(2501))
    )
    result = shot.capture_object_sync(_request("report", "rptExample", open_if_closed=True))
    assert result["error"].error_pattern == "open_canceled"
    assert fake_access.app.closes() == []


def test_loaded_datasheet_is_selected_then_restored(fake_access):
    fake_access.app = FakeAccess(tables=["tblExample"], loaded=["tblExample"])
    result = shot.capture_object_sync(_request("table", "tblExample"))
    assert "outcome" in result
    selects = [c for c in fake_access.app.DoCmd.calls if c[0] == "SelectObject"]
    assert selects == [
        ("SelectObject", shot.AC_TABLE, "tblExample"),
        ("SelectObject", 2, "frmPrevious"),
    ]
    assert fake_access.captured == [0x500]


def test_wrong_process(fake_access, monkeypatch):
    fake_access.app = FakeAccess(forms=["frmExample"], loaded=["frmExample"])
    monkeypatch.setattr(shot, "pid_from_access_app", lambda app: 99)
    result = shot.capture_object_sync(_request())
    assert result["error"].error_pattern == "database_not_open"
    assert fake_access.captured == []


def test_blocked_access_gets_no_com_call(fake_access, monkeypatch):
    monkeypatch.setattr(shot, "win32_precheck",
                        lambda pid: shot.ScreenshotError("access_blocked", "dialog"))
    monkeypatch.setattr(shot, "_find_app", lambda path: pytest.fail("COM call into blocked Access"))
    result = shot.capture_object_sync(_request())
    assert result["fallback"] is True
    assert result["error"].error_pattern == "access_blocked"


def test_open_timeout_leaves_cleanup_to_the_worker(fake_access, monkeypatch):
    """An open stuck behind an event dialog is abandoned; when it finishes,
    the worker closes the object, and object captures wait until then."""
    fake_access.app = FakeAccess(forms=["frmExample"])
    fake_access.app.DoCmd.open_delay = 0.6
    monkeypatch.setattr(shot, "_open_timeout", lambda: 0.2)

    result = shot.capture_object_sync(_request(open_if_closed=True))

    assert result["fallback"] is True
    assert result["error"].error_pattern == "open_blocked"
    assert result["error"].extra["object"]["cleanup_pending"] is True
    assert shot.cleanup_pending(DB)

    second = shot.capture_object_sync(_request(open_if_closed=True))
    assert second["error"].error_pattern == "cleanup_pending"

    deadline = time.monotonic() + 5
    while shot.cleanup_pending(DB) and time.monotonic() < deadline:
        time.sleep(0.05)
    assert not shot.cleanup_pending(DB)
    assert fake_access.app.closes() == [("Close", shot.AC_FORM, "frmExample", shot.AC_SAVE_NO)]
    assert fake_access.captured == []


def test_active_object(fake_access):
    app = FakeAccess()
    app.Screen = SimpleNamespace(ActiveForm=SimpleNamespace(Name="frmExample", Hwnd=0x600))
    fake_access.app = app
    result = shot.capture_object_sync(_request(target="active", object_type=None, name=None))
    assert result["outcome"].meta["object"]["source"] == "Screen.ActiveForm"
    assert fake_access.captured == [0x600]


def test_no_active_object(fake_access):
    class _Screen:
        def __getattr__(self, name):
            number = {"ActiveForm": 2475, "ActiveReport": 2476, "ActiveDatasheet": 2484}[name]
            raise _ComError(excepinfo=(0, "Access", "", None, 0, vba_hresult(number)))

    app = FakeAccess()
    app.Screen = _Screen()
    fake_access.app = app
    result = shot.capture_object_sync(_request(target="active", object_type=None, name=None))
    assert result["error"].error_pattern == "no_active_object"


# ---------------------------------------------------------------------------
# Gate routing
# ---------------------------------------------------------------------------


class _Gate:
    def __init__(self, in_flight=None, result=None):
        self.in_flight = in_flight
        self.result = result
        self.calls = []

    def current_in_flight(self):
        return self.in_flight

    async def run_exclusive(self, tool, database, fn, is_async, *args):
        self.calls.append((tool, database, fn, is_async, args))
        return self.result if self.result is not None else fn(*args)


@pytest.fixture
def routing(monkeypatch):
    state = SimpleNamespace(captures=[])

    def capture(pid, target, *, max_dimension, include_image, object_hwnd=0):
        state.captures.append(target)
        return shot.Outcome(meta={"dialogs": [], "window": {"title": "Access"}},
                            png=PNG if include_image else None)

    monkeypatch.setattr(shot, "resolve_target_pid", lambda db, pid: (10, "argument"))
    monkeypatch.setattr(shot, "capture_process", capture)
    return state


def _run(**kwargs):
    args = dict(database_path=DB, target="application", object_type=None, object_name=None,
                open_if_closed=False, pid=None, max_dimension=1568, include_image=True)
    args.update(kwargs)
    return asyncio.run(shot.capture_screenshot(**args))


def test_application_capture_never_touches_the_gate(routing):
    result = _run(get_gate=lambda: pytest.fail("application capture used the gate"))
    meta = meta_of(result)
    assert meta["success"] is True
    assert meta["pid"] == 10 and meta["pid_source"] == "argument"
    assert [c.type for c in result.content] == ["text", "image"]


def test_metadata_only(routing):
    result = _run(include_image=False, get_gate=lambda: None)
    assert [c.type for c in result.content] == ["text"]


def test_object_capture_busy_returns_application_capture(routing):
    gate = _Gate(in_flight=SimpleNamespace(tool="vcs_export_database", database=DB))
    result = _run(target="active", get_gate=lambda: gate)
    meta = meta_of(result)
    assert meta["success"] is False
    assert meta["error_pattern"] == "server_busy"
    assert meta["busy_with"]["tool"] == "vcs_export_database"
    assert gate.calls == []
    assert routing.captures == ["application"]
    assert result.content[1].type == "image"


def test_object_capture_enters_gate_on_com_thread(routing, monkeypatch):
    gate = _Gate()
    monkeypatch.setattr(
        shot, "capture_object_sync",
        lambda request: {"fallback": False,
                         "outcome": shot.Outcome(meta={"window": {}}, png=PNG)},
    )
    result = _run(target="object", object_type="form", object_name="frmExample", get_gate=lambda: gate)
    assert meta_of(result)["success"] is True
    tool, database, _fn, is_async, args = gate.calls[0]
    assert (tool, database, is_async) == ("vcs_capture_screenshot", DB, False)
    assert args[0].object_name == "frmExample"


def test_object_fallback_attaches_application_capture(routing):
    error = shot.ScreenshotError("access_blocked", "Access is blocked by a dialog.")
    gate = _Gate(result={"fallback": True, "error": error})
    result = _run(target="active", get_gate=lambda: gate)
    meta = meta_of(result)
    assert meta["error_pattern"] == "access_blocked"
    assert meta["window"] == {"title": "Access"}
    assert routing.captures == ["application"]


def test_gate_busy_while_waiting(routing):
    gate = _Gate(result={"success": False, "error": "busy", "error_pattern": "server_busy",
                         "busy_with": {"tool": "vcs_run_tests"}})
    meta = meta_of(_run(target="active", get_gate=lambda: gate))
    assert meta["error_pattern"] == "server_busy"
    assert meta["busy_with"] == {"tool": "vcs_run_tests"}


def test_invalid_argument_result(routing):
    meta = meta_of(_run(target="object", get_gate=lambda: None))
    assert meta["success"] is False and meta["error_pattern"] == "invalid_argument"


def test_capture_failure_is_reported(routing, monkeypatch):
    def boom(*args, **kwargs):
        raise RuntimeError("GDI failure")

    monkeypatch.setattr(shot, "capture_process", boom)
    meta = meta_of(_run(get_gate=lambda: None))
    assert meta["error_pattern"] == "capture_failed"
    assert "GDI failure" in meta["error"]


# ---------------------------------------------------------------------------
# FastMCP end to end
# ---------------------------------------------------------------------------


def test_fastmcp_returns_text_then_image(routing, monkeypatch):
    import msaccess_vcs_mcp.tools as tools_module

    monkeypatch.setattr(tools_module, "validate_database_path", lambda p: p)
    tool = tools_module.mcp._tool_manager.get_tool("vcs_capture_screenshot")
    assert tool.fn_metadata.output_schema is None
    assert tool.parameters["properties"]["target"]["enum"] == list(shot.TARGETS)

    result = asyncio.run(
        tools_module.mcp.call_tool("vcs_capture_screenshot", {"database_path": DB})
    )

    assert isinstance(result, CallToolResult)
    assert [c.type for c in result.content] == ["text", "image"]
    assert result.content[1].mimeType == "image/png"
    assert meta_of(result)["success"] is True
    assert result.structuredContent is None


def test_fastmcp_invalid_database_path(monkeypatch):
    import msaccess_vcs_mcp.tools as tools_module

    def reject(path):
        raise ValueError("Database not found: example")

    monkeypatch.setattr(tools_module, "validate_database_path", reject)
    result = asyncio.run(
        tools_module.mcp.call_tool("vcs_capture_screenshot", {"database_path": DB})
    )
    meta = meta_of(result)
    assert meta["error_pattern"] == "invalid_argument"
    assert "Database not found" in meta["error"]
