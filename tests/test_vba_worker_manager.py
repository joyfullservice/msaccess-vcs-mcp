"""Tests for thread-based VBA worker management and COM recovery."""

from __future__ import annotations

import asyncio
import json
import threading
import time
from unittest.mock import Mock, patch

import pytest

import msaccess_vcs_mcp.tools as tools_module
import msaccess_vcs_mcp.vba_worker_manager as worker_module
from msaccess_vcs_mcp.com_recovery import classify_com_error, get_recovery_manager
from msaccess_vcs_mcp.vba_worker_manager import VBAWorkerManager, _parse_reset_result


@pytest.fixture(autouse=True)
def _reset_recovery_and_worker():
    get_recovery_manager().reset()
    VBAWorkerManager._active_worker = None
    yield
    get_recovery_manager().reset()
    VBAWorkerManager._active_worker = None


def _patch_worker_thread(monkeypatch, results: list[dict]):
    """Replace the worker's thread body with a synchronous result feeder.

    Each call to ``_run_worker`` pops the next result dict from *results*
    and stuffs it into the ``result_box`` that the real thread body would
    write.  The thread itself is a no-op so the test never touches COM.
    """
    call_log: list[dict] = []
    real_run_worker = VBAWorkerManager._run_worker

    def fake_run_worker(self, *, operation, database_path, addin_path,
                        code, timeout_seconds, retry=False):
        call_log.append({
            "operation": operation,
            "database_path": database_path,
            "code": code,
            "retry": retry,
        })
        if results and results[0].get("timeout"):
            result = results.pop(0)
            return {
                "success": False,
                "error": f"VBA worker timed out after {timeout_seconds} seconds",
                "error_pattern": "timeout",
                "recoverable": True,
                "timed_out": True,
                "duration_ms": 100.0,
            }
        if results:
            response = dict(results.pop(0))
            response.setdefault("duration_ms", 10.0)
            return response
        return {"success": False, "error": "No more fake results", "error_pattern": "unknown"}

    monkeypatch.setattr(VBAWorkerManager, "_run_worker", fake_run_worker)
    return call_log


def _install_fake_com_worker(
    monkeypatch,
    *,
    reset_response: str,
    initial_barrier_error: Exception | None = None,
    fresh_barrier_error: Exception | None = None,
):
    """Install deterministic COM/add-in doubles for the real worker body."""
    events: list[str] = []

    class FakeProject:
        def __init__(self, label: str, error: Exception | None = None):
            self.label = label
            self.error = error

        @property
        def FullName(self):
            events.append(f"barrier:{self.label}")
            if self.error is not None:
                raise self.error
            return "C:\\db.accdb"

    class FakeApp:
        def __init__(self, label: str, error: Exception | None = None):
            self.label = label
            self.CurrentProject = FakeProject(label, error)

    initial_app = FakeApp("initial", initial_barrier_error)
    fresh_app = FakeApp("fresh", fresh_barrier_error)
    find_count = 0

    class FakeIntegration:
        def __init__(self, addin_path=None):
            self.label = ""

        @staticmethod
        def _find_access_in_rot(database_path):
            nonlocal find_count
            find_count += 1
            label = "initial" if find_count == 1 else "fresh"
            events.append(f"find:{label}")
            return initial_app if find_count == 1 else fresh_app

        def load_addin(self, app, db_path=None):
            self.label = app.label
            events.append(f"load:{self.label}")
            return True

        def call_sync(self, command, *args):
            events.append(f"call:{self.label}:{command}")
            if command == "ResetVbaProjectState":
                return reset_response
            if command == "RunVBA":
                return '{"success": true, "result": "ok"}'
            raise AssertionError(f"Unexpected command: {command}")

    fake_pythoncom = Mock()
    monkeypatch.setattr(worker_module, "COM_AVAILABLE", True)
    monkeypatch.setattr(worker_module, "pythoncom", fake_pythoncom)
    monkeypatch.setattr(worker_module, "VCSAddinIntegration", FakeIntegration)
    monkeypatch.setattr(
        worker_module,
        "ensure_access_visible",
        lambda app: events.append(f"visible:{app.label}"),
    )
    return events


# ------------------------------------------------------------------
# COM error classification (unchanged)
# ------------------------------------------------------------------

def test_classifies_recoverable_com_errors():
    assert classify_com_error("The RPC server is unavailable (-2147023174)") == "rpc_unavailable"
    assert classify_com_error("Call was rejected by callee (-2147418111)") == "call_rejected"
    assert classify_com_error("Object invoked has disconnected") == "object_disconnected"
    assert classify_com_error("MCP error -32000: Connection closed") == "connection_closed"
    assert classify_com_error("Operation timed out after 45 seconds") == "timeout"


# ------------------------------------------------------------------
# Happy path
# ------------------------------------------------------------------

def test_run_vba_success(monkeypatch):
    results = [
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": '{"success": true, "result": 42}',
        }
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="MCP_TempFunction = 42",
        addin_path="C:\\addin.accda",
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == '{"success": true, "result": 42}'
    assert calls[0]["operation"] == "run_vba"
    assert calls[0]["code"] == "MCP_TempFunction = 42"


def test_real_worker_orders_reset_barrier_reacquire_then_run(monkeypatch):
    events = _install_fake_com_worker(
        monkeypatch,
        reset_response='{"success": true, "resetQueued": true}',
    )

    result = VBAWorkerManager()._run_worker(
        operation="run_vba",
        database_path="C:\\db.accdb",
        addin_path="C:\\addin.accda",
        code='MCP_TempFunction = "ok"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["phase"] == "run_vba"
    assert result["result"] == '{"success": true, "result": "ok"}'
    assert events == [
        "find:initial",
        "visible:initial",
        "load:initial",
        "call:initial:ResetVbaProjectState",
        "barrier:initial",
        "find:fresh",
        "barrier:fresh",
        "visible:fresh",
        "load:fresh",
        "call:fresh:RunVBA",
    ]


@pytest.mark.parametrize(
    "reset_response,error_pattern",
    [
        (
            '{"success": false, "resetQueued": false, '
            '"error": "unsafe", "error_pattern": "reset_refused"}',
            "reset_refused",
        ),
        (
            '{"success": false, "resetQueued": false, '
            '"error": "failed", "error_pattern": "reset_failed"}',
            "reset_failed",
        ),
        ('{"success": true}', "reset_failed"),
        ("not json", "reset_failed"),
    ],
)
def test_real_worker_fails_closed_when_reset_is_not_confirmed(
    monkeypatch, reset_response, error_pattern
):
    events = _install_fake_com_worker(
        monkeypatch,
        reset_response=reset_response,
    )

    result = VBAWorkerManager()._run_worker(
        operation="run_vba",
        database_path="C:\\db.accdb",
        addin_path="C:\\addin.accda",
        code='MCP_TempFunction = "must not run"',
        timeout_seconds=1,
    )

    logical_result = json.loads(result["result"])
    assert result["success"] is True  # COM succeeded; logical failure is JSON.
    assert logical_result["success"] is False
    assert logical_result["error_pattern"] == error_pattern
    assert not any(event.endswith(":RunVBA") for event in events)
    assert not any(event.startswith("barrier:") for event in events)


def test_real_worker_reacquires_after_initial_barrier_error(monkeypatch):
    events = _install_fake_com_worker(
        monkeypatch,
        reset_response='{"success": true, "resetQueued": true}',
        initial_barrier_error=RuntimeError("reset teardown in progress"),
    )

    result = VBAWorkerManager()._run_worker(
        operation="run_vba",
        database_path="C:\\db.accdb",
        addin_path="C:\\addin.accda",
        code='MCP_TempFunction = "ok"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert "barrier:initial" in events
    assert "barrier:fresh" in events
    assert events[-1] == "call:fresh:RunVBA"


def test_real_worker_stops_when_fresh_barrier_fails(monkeypatch):
    events = _install_fake_com_worker(
        monkeypatch,
        reset_response='{"success": true, "resetQueued": true}',
        fresh_barrier_error=RuntimeError("Access still resetting"),
    )

    result = VBAWorkerManager()._run_worker(
        operation="run_vba",
        database_path="C:\\db.accdb",
        addin_path="C:\\addin.accda",
        code='MCP_TempFunction = "must not run"',
        timeout_seconds=1,
    )

    assert result["success"] is False
    assert result["phase"] == "reset_barrier"
    assert not any(event.endswith(":RunVBA") for event in events)


def test_parse_reset_result_rejects_non_object_json():
    result = _parse_reset_result('["not", "an", "object"]')

    assert result["success"] is False
    assert result["resetQueued"] is False
    assert result["error_pattern"] == "reset_failed"


# ------------------------------------------------------------------
# Timeout
# ------------------------------------------------------------------

def test_run_vba_timeout_returns_recoverable(monkeypatch):
    results = [{"timeout": True}]
    _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="Do While True: Loop",
        timeout_seconds=0.1,
    )

    assert result["success"] is False
    assert result["timed_out"] is True
    assert result["recoverable"] is True
    assert result["error_pattern"] == "timeout"
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "timed_out"


# ------------------------------------------------------------------
# Recovery probe after prior timeout
# ------------------------------------------------------------------

def test_failed_probe_recycles_owned_instance_then_retries(monkeypatch):
    get_recovery_manager().mark_failure(
        "C:\\db.accdb",
        "VBA worker timed out after 45 seconds",
        "timeout",
        timed_out=True,
    )
    results = [
        {
            "success": False,
            "operation": "probe",
            "phase": "connect",
            "error": "Access is not responding",
            "error_pattern": "access_unresponsive",
        },
        {
            "success": True,
            "operation": "probe",
            "phase": "load_addin",
            "result": "ok",
        },
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": "done",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)
    recycle = Mock(return_value=True)
    monkeypatch.setattr(worker_module, "recycle_owned_instance", recycle)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "done"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    recycle.assert_called_once_with("C:\\db.accdb")
    assert [c["operation"] for c in calls] == ["probe", "probe", "run_vba"]


def test_failed_probe_does_not_recycle_user_owned_instance(monkeypatch):
    get_recovery_manager().mark_failure(
        "C:\\db.accdb",
        "VBA worker timed out after 45 seconds",
        "timeout",
        timed_out=True,
    )
    results = [
        {
            "success": False,
            "operation": "probe",
            "phase": "connect",
            "error": "Access is not responding",
            "error_pattern": "access_unresponsive",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)
    recycle = Mock(return_value=False)
    monkeypatch.setattr(worker_module, "recycle_owned_instance", recycle)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "done"',
        timeout_seconds=1,
    )

    assert result["success"] is False
    assert result["error_pattern"] == "access_unresponsive"
    assert "vcs_capture_screenshot" in result["agent_guidance"]
    recycle.assert_called_once_with("C:\\db.accdb")
    assert [c["operation"] for c in calls] == ["probe"]


def test_real_worker_opens_access_when_rot_misses(monkeypatch):
    events: list[str] = []

    class FakeProject:
        @property
        def FullName(self):
            events.append("barrier")
            return "C:\\db.accdb"

    class FakeApp:
        def __init__(self):
            self.CurrentProject = FakeProject()

    app = FakeApp()
    find_count = 0
    opened: list[str] = []

    class FakeIntegration:
        def __init__(self, addin_path=None):
            pass

        @staticmethod
        def _find_access_in_rot(database_path):
            nonlocal find_count
            find_count += 1
            events.append(f"find:{find_count}")
            if find_count == 1:
                return None
            return app

        def load_addin(self, worker_app, db_path=None):
            events.append("load")
            return True

        def call_sync(self, command, *args):
            events.append(f"call:{command}")
            if command == "ResetVbaProjectState":
                return '{"success": true, "resetQueued": true}'
            if command == "RunVBA":
                return '{"success": true, "result": "ok"}'
            raise AssertionError(f"Unexpected command: {command}")

    class FakeConn:
        def __init__(self, path):
            opened.append(path)
            events.append("open")

        def connect(self):
            return app, None

        def close(self):
            events.append("close")

    fake_pythoncom = Mock()
    monkeypatch.setattr(worker_module, "COM_AVAILABLE", True)
    monkeypatch.setattr(worker_module, "pythoncom", fake_pythoncom)
    monkeypatch.setattr(worker_module, "VCSAddinIntegration", FakeIntegration)
    monkeypatch.setattr(worker_module, "AccessConnection", FakeConn)
    monkeypatch.setattr(
        worker_module,
        "ensure_access_visible",
        lambda _app: events.append("visible"),
    )

    result = VBAWorkerManager()._run_worker(
        operation="run_vba",
        database_path="C:\\db.accdb",
        addin_path="C:\\addin.accda",
        code='MCP_TempFunction = "ok"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert opened == ["C:\\db.accdb"]
    assert events[0:3] == ["find:1", "open", "find:2"]
    # The connection stays open for the whole operation. Closing it at the
    # end of the lookup would quit the instance under
    # ACCESS_VCS_LEAVE_ACCESS_OPEN=false, and the very next ROT lookup
    # would find nothing.
    assert events.index("close") == len(events) - 1
    assert "call:RunVBA" in events[: events.index("close")]


def test_next_call_probes_after_prior_timeout(monkeypatch):
    get_recovery_manager().mark_failure(
        "C:\\db.accdb",
        "VBA worker timed out after 45 seconds",
        "timeout",
        timed_out=True,
    )
    results = [
        {
            "success": True,
            "operation": "probe",
            "phase": "load_addin",
            "result": "ok",
        },
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": "done",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "done"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert [c["operation"] for c in calls] == ["probe", "run_vba"]
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "healthy"


# ------------------------------------------------------------------
# Pre-dispatch disconnect -> probe + retry
# ------------------------------------------------------------------

def test_predispatch_disconnect_probes_and_retries_once(monkeypatch):
    results = [
        {
            "success": False,
            "operation": "run_vba",
            "phase": "connect",
            "error": "The RPC server is unavailable",
            "error_pattern": "rpc_unavailable",
        },
        {
            "success": True,
            "operation": "probe",
            "phase": "load_addin",
            "result": "ok",
        },
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": "retried",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "retried"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == "retried"
    assert [c["operation"] for c in calls] == ["run_vba", "probe", "run_vba"]


# ------------------------------------------------------------------
# run_vba phase failure is NOT retried
# ------------------------------------------------------------------

def test_run_phase_failure_is_not_retried(monkeypatch):
    results = [
        {
            "success": False,
            "operation": "run_vba",
            "phase": "run_vba",
            "error": "Call was rejected by callee",
            "error_pattern": "call_rejected",
        }
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="CurrentDb.Execute \"UPDATE T SET X = 1\"",
        timeout_seconds=1,
    )

    assert result["success"] is False
    assert result["recoverable"] is True
    assert result["phase"] == "run_vba"
    assert len(calls) == 1


# ------------------------------------------------------------------
# Active-worker single-flight guard
# ------------------------------------------------------------------

def test_active_worker_guard_returns_error():
    stop = threading.Event()
    holding = threading.Thread(target=stop.wait, daemon=True)
    holding.start()
    try:
        VBAWorkerManager._active_worker = holding
        manager = VBAWorkerManager()
        result = manager._run_worker(
            operation="run_vba",
            database_path="C:\\db.accdb",
            addin_path=None,
            code="x = 1",
            timeout_seconds=1,
        )
        assert result["success"] is False
        assert "still running" in result["error"]
        assert result["error_pattern"] == "access_unresponsive"
    finally:
        stop.set()
        holding.join(timeout=1)
        VBAWorkerManager._active_worker = None


# ------------------------------------------------------------------
# tools.py integration (result semantics preserved)
# ------------------------------------------------------------------

def test_vcs_run_vba_preserves_json_result_semantics(tmp_path, monkeypatch):
    db_path = tmp_path / "test.accdb"
    db_path.write_text("", encoding="utf-8")

    async def fake_ensure(ctx):
        return None

    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)
    monkeypatch.setattr(tools_module.mcp, "get_context", lambda: None)
    monkeypatch.setattr(tools_module, "load_config", lambda: {})
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": "C:\\addin.accda"},
    )
    monkeypatch.setattr(
        tools_module,
        "run_vba_resilient",
        lambda **kwargs: {
            "success": True,
            "result": '{"success": true, "result": 7}',
        },
    )

    result = asyncio.run(
        tools_module.vcs_run_vba(
            database_path=str(db_path),
            code="MCP_TempFunction = 7",
            timeout_seconds=12,
        )
    )

    assert result == {"success": True, "result": 7}


def test_vcs_run_vba_preserves_cleanup_diagnostics(tmp_path, monkeypatch):
    db_path = tmp_path / "test.accdb"
    db_path.write_text("", encoding="utf-8")

    async def fake_ensure(ctx):
        return None

    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)
    monkeypatch.setattr(tools_module.mcp, "get_context", lambda: None)
    monkeypatch.setattr(tools_module, "load_config", lambda: {})
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": "C:\\addin.accda"},
    )
    monkeypatch.setattr(
        tools_module,
        "run_vba_resilient",
        lambda **kwargs: {
            "success": True,
            "result": json.dumps(
                {
                    "success": False,
                    "error_pattern": "temp_module_cleanup_failed",
                    "cleanupFailed": True,
                    "orphanModule": "MCP_Temp_123",
                    "payloadResult": "done",
                }
            ),
        },
    )

    result = asyncio.run(
        tools_module.vcs_run_vba(
            database_path=str(db_path),
            code='MCP_TempFunction = "done"',
        )
    )

    assert result == {
        "success": False,
        "error_pattern": "temp_module_cleanup_failed",
        "cleanupFailed": True,
        "orphanModule": "MCP_Temp_123",
        "payloadResult": "done",
    }
