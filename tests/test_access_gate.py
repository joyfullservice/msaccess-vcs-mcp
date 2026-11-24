"""Tests for the Access operation gate."""

from __future__ import annotations

import asyncio
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
                "vcs_call_vba", r"C:\other.accdb", lambda: None, False
            )
            await slow_task
            return busy

    busy = asyncio.run(runner())
    assert busy["success"] is False
    assert busy["error_pattern"] == "server_busy"
    assert busy["recoverable"] is True
    assert busy["busy_with"]["tool"] == "vcs_run_tests"
    assert busy["busy_with"]["database"] == r"C:\big.accdb"


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
