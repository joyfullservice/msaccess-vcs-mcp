"""Tests for cancelling a running operation from MCP."""

from __future__ import annotations

import asyncio
import json
import os
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.operation_manager import OperationManager
from msaccess_vcs_mcp.rebuild_watcher import wait_for_rebuild_status


def _unwrap(tool_fn):
    while hasattr(tool_fn, "__wrapped__"):
        tool_fn = tool_fn.__wrapped__
    return tool_fn


cancel_operation = _unwrap(tools.vcs_cancel_operation)


@pytest.fixture(autouse=True)
def _diag_dir(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


@pytest.fixture
def manager():
    OperationManager._instance = None
    mgr = OperationManager.get_instance()
    with patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=mgr):
        yield mgr
    OperationManager._instance = None


class TestCancelTool:
    def test_nothing_running(self, manager):
        result = cancel_operation()
        assert result["success"] is False
        assert "No operation is running" in result["error"]

    def test_only_running_operation_is_cancelled(self, manager):
        op_id, _ = manager.register_operation(database_path=r"C:\db\one.accdb", command="Export")
        result = cancel_operation()
        assert result["success"] is True
        assert result["operation_id"] == op_id
        assert result["command"] == "Export"
        assert manager.is_cancelled(op_id)

    def test_several_running_is_ambiguous(self, manager):
        first, _ = manager.register_operation(database_path=r"C:\db\one.accdb", command="Export")
        second, _ = manager.register_operation(database_path=r"C:\db\two.accdb", command="BuildHeadless")
        result = cancel_operation()
        assert result["success"] is False
        assert {op["operation_id"] for op in result["operations"]} == {first, second}
        assert not manager.is_cancelled(first)
        assert not manager.is_cancelled(second)

    def test_operation_id_wins(self, manager):
        manager.register_operation(database_path=r"C:\db\one.accdb", command="Export")
        second, _ = manager.register_operation(database_path=r"C:\db\two.accdb", command="Export")
        result = cancel_operation(operation_id=second)
        assert result["success"] is True
        assert result["operation_id"] == second

    def test_unknown_operation_id(self, manager):
        result = cancel_operation(operation_id="no-such-id")
        assert result["success"] is False
        assert result["operations"] == []

    def test_database_path_is_normalized(self, manager):
        op_id, _ = manager.register_operation(database_path=r"C:\DB\Folder\One.accdb", command="Export")
        result = cancel_operation(database_path=r"c:/db/folder/sub/../one.accdb")
        assert result["success"] is True
        assert result["operation_id"] == op_id

    def test_database_path_with_nothing_running(self, manager):
        other, _ = manager.register_operation(database_path=r"C:\db\one.accdb", command="Export")
        result = cancel_operation(database_path=r"C:\db\two.accdb")
        assert result["success"] is False
        assert [op["operation_id"] for op in result["operations"]] == [other]

    def test_abandoned_operation_still_cancellable_by_id(self, manager):
        op_id, _ = manager.register_operation(database_path=r"C:\db\one.accdb")
        manager.abandon_operation(op_id)
        assert cancel_operation()["success"] is False
        assert cancel_operation(operation_id=op_id)["success"] is True


class TestTombstones:
    def test_abandon_keeps_answering_cancelled(self, manager):
        op_id, _ = manager.register_operation(database_path=r"C:\db\one.accdb")
        manager.abandon_operation(op_id)
        assert manager.get_operation(op_id) is None
        assert not manager.is_database_busy(r"C:\db\one.accdb")
        assert manager.is_cancelled(op_id)

    def test_terminal_callback_clears_tombstone(self, manager):
        op_id, _ = manager.register_operation()
        manager.abandon_operation(op_id)
        manager.route_callback(op_id, {"type": "log", "message": "still going"})
        assert manager.is_cancelled(op_id)
        manager.route_callback(op_id, {"type": "cancelled"})
        assert not manager.is_cancelled(op_id)

    def test_tombstone_expires(self, manager):
        op_id, _ = manager.register_operation()
        manager.abandon_operation(op_id)
        manager._tombstones[op_id] = 0.0
        assert not manager.is_cancelled(op_id)

    def test_wait_timeout_abandons(self, manager):
        op_id, _ = manager.register_operation()
        result = asyncio.run(manager.wait_for_completion(op_id, timeout_seconds=0.01))
        assert result["success"] is False
        assert result["error_pattern"] == "timeout"
        assert manager.get_operation(op_id) is None
        assert manager.is_cancelled(op_id)

    def test_client_abandonment_abandons(self, manager):
        async def run():
            op_id, _ = manager.register_operation()
            task = asyncio.create_task(manager.wait_for_completion(op_id, timeout_seconds=30))
            await asyncio.sleep(0.01)
            task.cancel()
            with pytest.raises(asyncio.CancelledError):
                await task
            return op_id

        op_id = asyncio.run(run())
        assert manager.is_cancelled(op_id)

    def test_cancelled_callback_result(self, manager):
        async def run():
            manager.set_event_loop(asyncio.get_running_loop())
            op_id, _ = manager.register_operation()
            manager.route_callback(op_id, {"type": "cancelled", "message": "Operation was cancelled"})
            return await manager.wait_for_completion(op_id, timeout_seconds=5)

        result = asyncio.run(run())
        assert result["cancelled"] is True
        assert result["error_pattern"] == "operation_cancelled"
        assert result["error"]

    def test_callback_info_opts_into_polling(self, manager):
        info = json.loads(manager.create_callback_info("op", "http://localhost:1/cb"))
        assert info["cancel_poll"] is True


class _Thread:
    def __init__(self, alive: bool):
        self.alive = alive

    def is_alive(self):
        return self.alive


class TestRunApiWithCallbacks:
    def _run(self, manager, thread, box):
        calls = []

        def fake_call(_acquire, run_name, call_args):
            calls.append((run_name, call_args))
            return thread, box

        with (
            patch("msaccess_vcs_mcp.tools._start_application_call", fake_call),
            patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        ):
            run = asyncio.run(tools._run_api_with_callbacks(
                None,
                r"C:\addin\Version Control.accda",
                "BuildHeadless",
                [r"C:\src"],
                database_path=r"C:\db\one.accdb",
                timeout=0.05,
                ctx=None,
                heartbeat="Building",
            ))
        return run, calls

    def test_goes_through_api_async_and_unwraps(self, manager):
        inner = json.dumps({"success": False, "cancelled": True})
        box = {"success": True, "result": json.dumps({"sync": True, "result": inner})}
        run, calls = self._run(manager, _Thread(False), box)

        run_name, call_args = calls[0]
        assert run_name == os.path.splitext(os.path.abspath(r"C:\addin\Version Control.accda"))[0] + ".APIAsync"
        assert json.loads(call_args[0])["cancel_poll"] is True
        assert call_args[1:] == ["BuildHeadless", r"C:\src"]
        assert run["run_result"]["result"] == inner
        assert not run["timed_out"]
        assert manager.list_active_operations() == []
        assert not manager.is_cancelled(run["operation_id"])

        result = tools._mark_cancelled(tools._addin_json_result(run["run_result"]["result"]), None)
        assert result["error_pattern"] == "operation_cancelled"

    def test_timeout_abandons(self, manager):
        run, _ = self._run(manager, _Thread(True), {})
        assert run["timed_out"]
        assert manager.list_active_operations() == []
        assert manager.is_cancelled(run["operation_id"])


class TestCancelledResults:
    def test_test_run_cancel_without_results(self):
        result = tools._test_results_from_completion(
            {"success": False, "cancelled": True, "message": "Operation was cancelled"}
        )
        assert result["cancelled"] is True
        assert result["error_pattern"] == "operation_cancelled"

    def test_timeout_keeps_error_pattern(self):
        result = tools._test_results_from_completion(
            {"success": False, "error": "timed out", "error_pattern": "timeout", "operation_id": "op"}
        )
        assert result["error_pattern"] == "timeout"
        assert result["operation_id"] == "op"

    def test_mark_cancelled_from_completion(self):
        result = tools._mark_cancelled({"success": False, "error": "x"}, {"cancelled": True})
        assert result["cancelled"] is True
        assert result["error_pattern"] == "operation_cancelled"

    def test_mark_cancelled_leaves_other_results(self):
        result = tools._mark_cancelled({"success": True}, None)
        assert result == {"success": True}

    def test_rebuild_watcher_treats_cancelled_as_terminal(self, tmp_path):
        status_file = str(tmp_path / "rebuild-status.json")
        result = asyncio.run(wait_for_rebuild_status(
            status_file,
            "phase-1",
            timeout_sec=5,
            read_status=lambda _p: {"phaseStarted": "phase-1", "status": "cancelled"},
            wait_for_change=lambda _d, _t: False,
            processes_alive=lambda: True,
            promote_new_access=lambda pids: pids,
        ))
        assert result["success"] is False
        assert result["cancelled"] is True
        assert result["error_pattern"] == "operation_cancelled"
