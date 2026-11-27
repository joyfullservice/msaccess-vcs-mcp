"""Tests for the Access operation gate."""

from __future__ import annotations

import asyncio
import threading
import time
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp.access_gate import (
    AccessGate,
    EXEMPT_TOOLS,
    InFlight,
    _busy_error,
    reset_access_gate,
)


@pytest.fixture(autouse=True)
def _reset_gate():
    reset_access_gate()
    yield
    reset_access_gate()


def test_run_exclusive_executes_sync_fn_in_apartment():
    gate = AccessGate()
    seen: list[str] = []

    def work():
        seen.append("done")
        return 42

    result = asyncio.run(
        gate.run_exclusive("vcs_test", r"C:\db.accdb", work, False)
    )
    assert result == 42
    assert seen == ["done"]


def test_run_exclusive_executes_async_fn_on_loop():
    gate = AccessGate()

    async def work():
        await asyncio.sleep(0)
        return "async"

    result = asyncio.run(gate.run_exclusive("vcs_test", None, work, True))
    assert result == "async"


def test_run_exclusive_serializes_concurrent_calls():
    gate = AccessGate()
    order: list[int] = []

    async def runner():
        async def first():
            order.append(1)
            await asyncio.sleep(0.05)
            order.append(2)
            return 1

        async def second():
            order.append(3)
            return 2

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=2.0):
            t1 = asyncio.create_task(
                gate.run_exclusive("first", None, first, True)
            )
            await asyncio.sleep(0.01)
            t2 = asyncio.create_task(
                gate.run_exclusive("second", None, second, True)
            )
            return await asyncio.gather(t1, t2)

    r1, r2 = asyncio.run(runner())
    assert r1 == 1
    assert r2 == 2
    assert order == [1, 2, 3]


def test_run_exclusive_returns_busy_when_slot_unavailable():
    gate = AccessGate()
    entered = asyncio.Event()

    async def runner():
        async def slow():
            entered.set()
            await asyncio.sleep(0.3)
            return "slow"

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.05):
            slow_task = asyncio.create_task(
                gate.run_exclusive("vcs_run_tests", r"C:\big.accdb", slow, True)
            )
            await entered.wait()
            busy = await gate.run_exclusive(
                "vcs_call_vba", r"C:\BIG.accdb", lambda: None, False
            )
            await slow_task
            return busy

    busy = asyncio.run(runner())
    assert busy["success"] is False
    assert busy["error_pattern"] == "server_busy"
    assert busy["recoverable"] is True
    assert busy["busy_with"]["tool"] == "vcs_run_tests"
    assert busy["busy_with"]["database"] == r"C:\big.accdb"
    assert busy["busy_with"]["reason"] == "same_database"


def _timed(gate, calls, wait=2.0):
    """Start each (tool, database, seconds) call at once; return elapsed and results."""

    async def runner():
        def sleeper(seconds):
            async def work():
                await asyncio.sleep(seconds)
                return seconds
            return work

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=wait):
            started = time.perf_counter()
            results = await asyncio.gather(*(
                gate.run_exclusive(tool, db, sleeper(seconds), True)
                for tool, db, seconds in calls
            ))
            return time.perf_counter() - started, results

    return asyncio.run(runner())


def test_different_databases_run_in_parallel():
    elapsed, results = _timed(AccessGate(), [
        ("a", r"C:\a.accdb", 0.3), ("b", r"C:\b.accdb", 0.3), ("c", r"C:\c.accdb", 0.3),
    ])
    assert results == [0.3, 0.3, 0.3]
    assert elapsed < 0.6


def test_same_database_runs_in_order():
    elapsed, _ = _timed(AccessGate(), [("a", r"C:\a.accdb", 0.2), ("b", r"c:\A.ACCDB", 0.2)])
    assert elapsed >= 0.38


def test_sync_work_runs_on_one_thread_per_database():
    gate = AccessGate()

    async def runner():
        async def thread_of(db):
            return await gate.run_exclusive("t", db, threading.get_ident, False)
        return [await thread_of(db) for db in (r"C:\a.accdb", r"C:\b.accdb", r"C:\a.accdb")]

    first_a, b, second_a = asyncio.run(runner())
    assert first_a == second_a
    assert first_a != b


def test_parallel_cap_returns_busy(monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_MAX_PARALLEL", "2")
    _, results = _timed(AccessGate(), [
        ("a", r"C:\a.accdb", 0.3), ("b", r"C:\b.accdb", 0.3), ("c", r"C:\c.accdb", 0.3),
    ], wait=0.05)
    busy = [r for r in results if isinstance(r, dict)]
    assert len(busy) == 1
    assert busy[0]["busy_with"]["reason"] == "max_parallel"


def test_parallel_cap_waits_when_allowed(monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_MAX_PARALLEL", "2")
    elapsed, results = _timed(AccessGate(), [
        ("a", r"C:\a.accdb", 0.2), ("b", r"C:\b.accdb", 0.2), ("c", r"C:\c.accdb", 0.2),
    ])
    assert results == [0.2, 0.2, 0.2]
    assert elapsed >= 0.38


def test_exclusive_waits_for_every_lane_and_holds_them_off():
    gate = AccessGate()
    order: list[str] = []

    async def runner():
        def step(name, seconds):
            async def work():
                order.append(f"{name} start")
                await asyncio.sleep(seconds)
                order.append(f"{name} end")
            return work

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=2.0):
            lanes = [
                asyncio.create_task(gate.run_exclusive("a", r"C:\a.accdb", step("a", 0.2), True)),
                asyncio.create_task(gate.run_exclusive("b", r"C:\b.accdb", step("b", 0.1), True)),
            ]
            await asyncio.sleep(0.05)
            exclusive = asyncio.create_task(
                gate.run_exclusive_all("all", None, step("all", 0.1), True)
            )
            await asyncio.sleep(0.05)
            late = asyncio.create_task(
                gate.run_exclusive("c", r"C:\c.accdb", step("c", 0.0), True)
            )
            await asyncio.gather(*lanes, exclusive, late)

    asyncio.run(runner())
    assert order.index("all start") > order.index("a end")
    assert order.index("all start") > order.index("b end")
    assert order.index("c start") > order.index("all end")


def test_call_without_database_takes_the_whole_gate():
    gate = AccessGate()

    async def runner():
        async def slow():
            await asyncio.sleep(0.3)

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.05):
            task = asyncio.create_task(gate.run_exclusive("keyless", None, slow, True))
            await asyncio.sleep(0.05)
            busy = await gate.run_exclusive("lane", r"C:\a.accdb", lambda: None, False)
            await task
            return busy

    busy = asyncio.run(runner())
    assert busy["busy_with"]["tool"] == "keyless"
    assert busy["busy_with"]["reason"] == "exclusive"


def test_cancelled_waiter_does_not_keep_the_lane():
    gate = AccessGate()

    async def runner():
        async def slow():
            await asyncio.sleep(0.2)

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=2.0):
            holder = asyncio.create_task(gate.run_exclusive("a", r"C:\a.accdb", slow, True))
            await asyncio.sleep(0.05)
            waiter = asyncio.create_task(
                gate.run_exclusive("b", r"C:\a.accdb", lambda: "late", False)
            )
            await asyncio.sleep(0.05)
            waiter.cancel()
            await holder
            await asyncio.sleep(0.3)
            return await gate.run_exclusive("c", r"C:\a.accdb", lambda: "free", False)

    with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.5):
        assert asyncio.run(runner()) == "free"


def test_rebuild_database_lane_is_the_database_it_builds():
    from msaccess_vcs_mcp import tools

    def rebuild(source_dir, output_path=None, template_path=None, ctx=None):
        pass

    target = tools._lane_target
    assert target("vcs_rebuild_database", rebuild, (r"C:\wt\App.accdb.src",), {}) == r"C:\wt\App.accdb"
    assert target("vcs_rebuild_database", rebuild, (r"C:\wt\App.accdb.src\\",), {}) == r"C:\wt\App.accdb"
    assert target(
        "vcs_rebuild_database", rebuild, (r"C:\wt\App.accdb.src",), {"output_path": r"C:\out\x.accdb"}
    ) == r"C:\out\x.accdb"
    assert target("vcs_rebuild_database", rebuild, (r"C:\wt\source",), {}) == r"C:\wt\source"
    assert target("vcs_export_database", rebuild, (r"C:\wt\App.accdb.src",), {}) is None


def test_current_in_flight_is_per_database():
    gate = AccessGate()

    async def runner():
        seen = {}

        async def work():
            seen["a"] = gate.current_in_flight(r"C:\a.accdb")
            seen["b"] = gate.current_in_flight(r"C:\b.accdb")
            seen["any"] = gate.current_in_flight()

        await gate.run_exclusive("vcs_run_tests", r"C:\a.accdb", work, True)
        return seen

    seen = asyncio.run(runner())
    assert seen["a"].tool == "vcs_run_tests"
    assert seen["b"] is None
    assert seen["any"].tool == "vcs_run_tests"


def test_run_exclusive_releases_slot_after_exception():
    gate = AccessGate()

    def boom():
        raise RuntimeError("fail")

    with pytest.raises(RuntimeError, match="fail"):
        asyncio.run(gate.run_exclusive("vcs_test", None, boom, False))

    result = asyncio.run(
        gate.run_exclusive("vcs_test", None, lambda: "ok", False)
    )
    assert result == "ok"


def test_busy_error_shape():
    err = _busy_error(InFlight("vcs_export_database", r"C:\db.accdb", time.perf_counter()))
    assert err["error_pattern"] == "server_busy"
    assert err["busy_with"]["tool"] == "vcs_export_database"
    assert "retry_after_seconds" in err


def test_exempt_tools_include_status_queries():
    assert "vcs_get_version_info" in EXEMPT_TOOLS
    assert "vcs_get_recent_calls" in EXEMPT_TOOLS
    assert "vcs_cancel_operation" in EXEMPT_TOOLS
    assert "vcs_rebuild_addin" in EXEMPT_TOOLS


def test_screenshot_tool_is_exempt():
    """Window captures must work while another call holds the gate; object
    captures enter the gate themselves."""
    assert "vcs_capture_screenshot" in EXEMPT_TOOLS


def test_dialog_and_reset_tools_are_exempt():
    """The call stuck behind a dialog or a break holds the gate; these release it."""
    assert "vcs_dismiss_dialog" in EXEMPT_TOOLS
    assert "vcs_reset_vba" in EXEMPT_TOOLS


def test_com_initializer_runs_for_sync_work():
    gate = AccessGate()
    asyncio.run(gate.run_exclusive("vcs_test", None, lambda: None, False))
    assert gate.com_initialized is True
