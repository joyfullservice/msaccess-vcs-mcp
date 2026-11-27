"""Several Access instances driven through one server at once (opt-in).

Skipped unless pytest runs with ``--run-concurrency``:

    pytest tests/test_concurrency_integration.py --run-concurrency -v

Run it after changing the Access gate, instance creation or ownership, the VBA
worker, dialog control, screenshots, the callback server, the ``vcs_tool``
wrapper, or the add-in's headless export, merge, build, or test paths.

It builds three throwaway databases in a temp folder, all named ``App.accdb``
in different folders as worktrees of one project are, and drives them through
the tools as parallel agents would. Up to three visible Access windows appear
for a few minutes. Every instance it starts is recorded in a temporary
ownership registry and closed at the end by PID and creation time; Access
windows it did not start are never touched.
"""

from __future__ import annotations

import asyncio
import gc
import json
import subprocess
import sys
import time
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.integration,
    pytest.mark.concurrency,
    pytest.mark.skipif(sys.platform != "win32", reason="Access automation"),
]

AC_MODULE = 5
AC_FORM = 2
AC_LABEL = 100
AC_DETAIL = 0
AC_SAVE_YES = 1
AC_QUIT_SAVE_NONE = 2
FORM_NAME = "frmConcurrency"
# Worktrees of one project: the same file name, and the same VBA project name,
# in different folders.
FOLDERS = ("wt-alpha", "wt-bravo", "wt-charlie")
FILE_NAME = "App.accdb"

MODULE_CODE = (
    "Option Compare Database\r\n"
    "Option Explicit\r\n"
    "\r\n"
    "Private Declare PtrSafe Function GetCurrentProcessId Lib \"kernel32\" () As Long\r\n"
    "\r\n"
    "Public Function ProcessId() As Long\r\n"
    "    ProcessId = GetCurrentProcessId()\r\n"
    "End Function\r\n"
    "\r\n"
    "Public Function SlowWork(ByVal Seconds As Variant) As String\r\n"
    "    Dim sngEnd As Single\r\n"
    "    sngEnd = Timer + CSng(Seconds)\r\n"
    "    Do While Timer < sngEnd\r\n"
    "        DoEvents\r\n"
    "    Loop\r\n"
    "    SlowWork = CurrentProject.Name\r\n"
    "End Function\r\n"
    "\r\n"
    "Public Function TimedWork(ByVal Seconds As Variant) As String\r\n"
    "    Dim dblStart As Double\r\n"
    "    dblStart = Timer\r\n"
    "    SlowWork Seconds\r\n"
    "    TimedWork = Format(dblStart, \"0.000\") & \"|\" & Format(Timer, \"0.000\")\r\n"
    "End Function\r\n"
    "\r\n"
    "Public Function ShowBox() As String\r\n"
    "    MsgBox \"Concurrency suite\", vbOKOnly, \"Concurrency\"\r\n"
    "    ShowBox = \"closed\"\r\n"
    "End Function\r\n"
)


def _run(coro):
    return asyncio.run(coro)


def _meta(result) -> dict:
    return json.loads(result.content[0].text)


def _build_databases(folder: Path) -> list[str]:
    """Create the three databases in one Access instance this suite starts and quits."""
    win32com = pytest.importorskip("win32com.client")
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none

    before = set(list_access_pids_or_none() or [])
    try:
        app = win32com.DispatchEx("Access.Application")
    except Exception as exc:
        pytest.skip(f"Access is not available: {exc}")
    new = set(list_access_pids_or_none() or []) - before
    pid = new.pop() if len(new) == 1 else None
    created = process_create_time(pid) if pid else None

    paths = []
    try:
        for name in FOLDERS:
            path = str(folder / name / FILE_NAME)
            Path(path).parent.mkdir(parents=True)
            app.NewCurrentDatabase(path)
            component = app.VBE.ActiveVBProject.VBComponents.Add(1)
            component.Name = "modConcurrency"
            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(MODULE_CODE)
            app.DoCmd.Save(AC_MODULE, "modConcurrency")
            form = app.CreateForm()
            default_name = form.Name
            label = app.CreateControl(default_name, AC_LABEL, AC_DETAIL, "", "", 300, 300, 4000, 600)
            label.Properties("Caption").Value = "Concurrency suite"
            app.DoCmd.Close(AC_FORM, default_name, AC_SAVE_YES)
            app.DoCmd.Rename(FORM_NAME, AC_FORM, default_name)
            app.CloseCurrentDatabase()
            paths.append(path)
    finally:
        try:
            app.Quit(AC_QUIT_SAVE_NONE)
        except Exception:
            pass
        del app
        gc.collect()
        if pid and created is not None:
            deadline = time.monotonic() + 15
            while time.monotonic() < deadline and process_create_time(pid) == created:
                time.sleep(0.2)
            if process_create_time(pid) == created:
                subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    return paths


@pytest.fixture(scope="module")
def databases(tmp_path_factory):
    folder = tmp_path_factory.mktemp("concurrency")
    with pytest.MonkeyPatch.context() as mp:
        mp.setenv("ACCESS_VCS_OWNED_INSTANCES_PATH", str(folder / "owned-instances.json"))
        mp.setenv("ACCESS_VCS_LOG_DIR", str(folder / "logs"))
        mp.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(folder / "logs"))
        mp.setenv("ACCESS_VCS_LEAVE_ACCESS_OPEN", "true")
        mp.delenv("ACCESS_VCS_DISABLE_WRITES", raising=False)
        mp.delenv("ACCESS_VCS_MAX_PARALLEL", raising=False)

        from msaccess_vcs_mcp.access_gate import reset_access_gate

        reset_access_gate()
        paths = _build_databases(folder)
        try:
            yield paths
        finally:
            _close_suite_instances(paths)
            reset_access_gate()


def _close_suite_instances(paths: list[str]) -> None:
    """Close only the instances recorded in this suite's temporary registry."""
    from msaccess_vcs_mcp.access_com.connection import close_owned_instances_holding
    from msaccess_vcs_mcp.access_com.instance_registry import list_owned, process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import terminate_pid

    close_owned_instances_holding(paths)
    for record in list_owned():
        if record.create_time is not None and process_create_time(record.pid) == record.create_time:
            terminate_pid(record.pid)


@pytest.fixture
def cold(databases):
    """The three databases with no Access instance holding them."""
    _close_suite_instances(databases)
    return databases


def _call(db, function, *args, timeout=None):
    from msaccess_vcs_mcp import tools

    return tools.vcs_call_vba(db, function, [str(a) for a in args], timeout)


async def _pid(db) -> int:
    result = await _call(db, "ProcessId")
    assert result["success"], result
    return int(result["result"])


def _span(result) -> tuple[float, float]:
    assert result["success"], result
    start, end = (float(x) for x in result["result"].split("|"))
    return start, end


def test_cold_start_gives_each_database_its_own_instance(cold):
    from msaccess_vcs_mcp.access_com.instance_registry import list_owned

    async def run():
        return await asyncio.gather(*(_pid(db) for db in cold))

    pids = _run(run())

    assert len(set(pids)) == 3
    owned = {record.pid: record.database_path for record in list_owned()}
    for pid, db in zip(pids, cold):
        assert Path(owned[pid]).resolve() == Path(db).resolve()


def test_different_databases_run_in_parallel(databases):
    async def run():
        await asyncio.gather(*(_pid(db) for db in databases))
        started = time.perf_counter()
        results = await asyncio.gather(*(_call(db, "SlowWork", 3) for db in databases))
        return time.perf_counter() - started, results

    elapsed, results = _run(run())

    assert all(r["success"] for r in results), results
    assert elapsed < 7, f"three 3-second calls took {elapsed:.1f}s"


def test_same_database_stays_serial(databases):
    db = databases[0]

    async def run():
        await _pid(db)
        return await asyncio.gather(_call(db, "TimedWork", 1.5), _call(db, "TimedWork", 1.5))

    first, second = sorted(_span(r) for r in _run(run()))
    assert first[1] <= second[0] + 0.05


def test_busy_is_per_database(databases, monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_BUSY_WAIT_SEC", "1")
    a, b = databases[0], databases[1]

    async def run():
        await asyncio.gather(_pid(a), _pid(b))
        slow = asyncio.create_task(_call(a, "SlowWork", 4))
        await asyncio.sleep(1)
        other = await _call(b, "SlowWork", 0)
        same = await _call(a, "SlowWork", 0)
        await slow
        return other, same

    other, same = _run(run())

    assert other["success"], other
    assert same["error_pattern"] == "server_busy"
    assert Path(same["busy_with"]["database"]).resolve() == Path(a).resolve()


def test_exclusive_hold_waits_for_lanes(databases):
    from msaccess_vcs_mcp.access_gate import get_access_gate

    async def run():
        await asyncio.gather(*(_pid(db) for db in databases[:2]))
        lanes = [asyncio.create_task(_call(db, "SlowWork", 2)) for db in databases[:2]]
        await asyncio.sleep(0.5)
        exclusive_at = await get_access_gate().run_exclusive_all(
            "concurrency-suite", None, time.perf_counter, False
        )
        lane_results = await asyncio.gather(*lanes)
        return exclusive_at, time.perf_counter(), lane_results

    started = time.perf_counter()
    exclusive_at, _, lane_results = _run(run())

    assert all(r["success"] for r in lane_results)
    assert exclusive_at - started >= 1.9


def test_parallel_cap(databases, monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_MAX_PARALLEL", "2")
    monkeypatch.setenv("ACCESS_VCS_BUSY_WAIT_SEC", "1")

    async def run():
        for db in databases:
            await _pid(db)
        return await asyncio.gather(*(_call(db, "SlowWork", 3) for db in databases))

    results = _run(run())

    busy = [r for r in results if r.get("error_pattern") == "server_busy"]
    assert len(busy) == 1
    assert busy[0]["busy_with"]["reason"] == "max_parallel"


def test_stuck_instance_does_not_stall_the_others(databases):
    from msaccess_vcs_mcp import tools

    stuck, others = databases[0], databases[1:]

    async def run():
        await asyncio.gather(*(_pid(db) for db in databases))
        blocked = asyncio.create_task(_call(stuck, "ShowBox", timeout=6))
        await asyncio.sleep(1)
        started = time.perf_counter()
        free = await asyncio.gather(*(_call(db, "SlowWork", 1) for db in others))
        free_elapsed = time.perf_counter() - started
        timed_out = await blocked

        meta = _meta(await tools.vcs_capture_screenshot(stuck, include_image=False))
        dialog = next(d for d in meta["dialogs"] if "Concurrency suite" in (d.get("text") or ""))
        dismissed = await tools.vcs_dismiss_dialog(dialog["dialog_id"], "OK", database_path=stuck)
        await asyncio.sleep(1)
        after = await _call(stuck, "SlowWork", 0)
        return free, free_elapsed, timed_out, dismissed, after

    free, free_elapsed, timed_out, dismissed, after = _run(run())

    assert all(r["success"] for r in free), free
    assert free_elapsed < 5
    assert timed_out["success"] is False
    assert dismissed["success"], dismissed
    assert after["success"], after


def test_parallel_screenshots_find_the_right_window(databases):
    from msaccess_vcs_mcp import tools

    async def run():
        pids = await asyncio.gather(*(_pid(db) for db in databases))
        windows = await asyncio.gather(*(
            tools.vcs_capture_screenshot(db, include_image=False) for db in databases
        ))
        objects = await asyncio.gather(*(
            tools.vcs_capture_screenshot(
                db, target="object", object_type="form", object_name=FORM_NAME,
                open_if_closed=True, view="design", save_image=True,
            )
            for db in databases
        ))
        return pids, windows, objects

    pids, windows, objects = _run(run())

    for pid, window, capture in zip(pids, windows, objects):
        assert _meta(window)["pid"] == pid
        meta = _meta(capture)
        assert meta["success"], meta
        assert meta["pid"] == pid
        assert Path(meta["image"]["saved_path"]).is_file()


def test_parallel_addin_exports(databases):
    from msaccess_vcs_mcp import tools
    from msaccess_vcs_mcp.config import get_default_addin_path

    addin = get_default_addin_path()
    if not addin or not Path(addin).is_file():
        pytest.skip("The VCS add-in is not installed.")

    outputs = [Path(db).parent / f"{Path(db).name}.src" for db in databases]
    elsewhere = [Path(db).parent / "export-elsewhere" for db in databases]

    async def run():
        results = await asyncio.gather(*(tools.vcs_export_database(db) for db in databases))
        refused = await tools.vcs_export_database(databases[0], str(outputs[1]))
        redirected = await asyncio.gather(*(
            tools.vcs_export_database(db, str(out)) for db, out in zip(databases, elsewhere)
        ))
        return results, refused, redirected

    results, refused, redirected = _run(run())

    assert all(r["success"] for r in results), results
    assert [Path(r["export_path"]) for r in results] == outputs
    # Another database's export folder holds its change index.
    assert refused["error_pattern"] == "export_folder_has_index", refused
    assert all(r["success"] for r in redirected), redirected
    for result, out in zip(redirected, elsewhere):
        assert Path(result["export_path"]) == out
        assert (out / "modules" / "modConcurrency.bas").is_file()
        assert not (out / "vcs-index.idx").exists()

    def tree(folder: Path) -> dict[str, bytes | None]:
        # A form carries an identifier Access assigns per database, so only its
        # presence is compared; module text must match exactly.
        return {
            str(p.relative_to(folder)): p.read_bytes() if p.suffix == ".bas" else None
            for p in folder.rglob("*")
            if p.is_file() and p.suffix in (".bas", ".cls", ".form")
        }

    first, *rest = (tree(out) for out in outputs)
    assert str(Path("modules") / "modConcurrency.bas") in first
    assert str(Path("forms") / f"{FORM_NAME}.form") in first
    for other in rest:
        assert other == first


def test_server_created_window_closes_gracefully(databases):
    from msaccess_vcs_mcp.access_com.connection import close_owned_instances_holding

    db = databases[-1]
    _run(_pid(db))

    closed = close_owned_instances_holding([db])

    assert len(closed) == 1
    assert closed[0]["quit"] is True
    assert closed[0]["terminated"] is False
