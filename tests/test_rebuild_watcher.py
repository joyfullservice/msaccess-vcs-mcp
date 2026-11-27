"""Tests for rebuild-status.json watching and vcs_rebuild_addin launch/wait."""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from msaccess_vcs_mcp.rebuild_watcher import (
    get_rebuild_timeout,
    read_rebuild_status,
    wait_for_rebuild_status,
)
from msaccess_vcs_mcp.operation_manager import MonotonicProgressReporter
import msaccess_vcs_mcp.tools as tools_module


def _write_status(path: Path, **fields):
    payload = {
        "status": "building",
        "error": "",
        "buildLog": "",
        "phaseStarted": "2026-08-27 10:00:00",
        "updated": "2026-08-27 10:00:01",
    }
    payload.update(fields)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload), encoding="utf-8-sig")
    return payload


def test_read_rebuild_status_accepts_bom(tmp_path):
    path = tmp_path / "logs" / "rebuild-status.json"
    _write_status(path, status="complete")
    record = read_rebuild_status(str(path))
    assert record["status"] == "complete"
    assert record["phaseStarted"] == "2026-08-27 10:00:00"


def test_read_rebuild_status_partial_write(tmp_path):
    path = tmp_path / "rebuild-status.json"
    path.write_text("{ truncated", encoding="utf-8")
    assert read_rebuild_status(str(path)) is None


def test_get_rebuild_timeout_env(monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_REBUILD_TIMEOUT_SEC", "30")
    assert get_rebuild_timeout() == 30.0
    assert get_rebuild_timeout(12) == 12.0


def test_wait_wakes_on_injected_change(tmp_path):
    status = tmp_path / "logs" / "rebuild-status.json"
    _write_status(status, status="building")
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    events = ["building", "compiling", "complete"]
    seen: list[str] = []

    def wait_for_change(_directory: str, _timeout: float) -> bool:
        if len(seen) < len(events):
            next_status = events[len(seen)]
            seen.append(next_status)
            _write_status(status, status=next_status)
        return True

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=5,
        ctx=ctx,
        wait_for_change=wait_for_change,
        processes_alive=lambda: True,
    ))
    assert result["success"] is True
    assert result["status"] == "complete"
    messages = [c.kwargs["message"] for c in ctx.report_progress.await_args_list]
    assert "Rebuild building" in messages
    assert "Rebuild compiling" in messages
    assert "Rebuild complete" in messages
    progresses = [c.kwargs["progress"] for c in ctx.report_progress.await_args_list]
    assert progresses == sorted(progresses)


def test_wait_promotes_access_started_after_snapshot(tmp_path):
    status = tmp_path / "logs" / "rebuild-status.json"
    _write_status(status, status="complete")
    seen: list[set[int]] = []

    def promote(known: set[int]) -> set[int]:
        seen.append(set(known))
        return known | {99}

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=5,
        wait_for_change=lambda *_a: True,
        processes_alive=lambda: True,
        preexisting_access_pids={7, 8},
        promote_new_access=promote,
    ))
    assert result["success"] is True
    assert seen
    assert seen[0] == {7, 8}


def test_wait_reuses_existing_progress_sequence(tmp_path):
    status = tmp_path / "logs" / "rebuild-status.json"
    _write_status(status, status="complete")
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    reporter = MonotonicProgressReporter()

    async def _run():
        await reporter.emit(ctx, message="Rebuild launched")
        return await wait_for_rebuild_status(
            str(status),
            "2026-08-27 10:00:00",
            timeout_sec=5,
            ctx=ctx,
            reporter=reporter,
            wait_for_change=lambda *_a: True,
            processes_alive=lambda: True,
        )

    result = asyncio.run(_run())
    assert result["success"] is True
    progresses = [c.kwargs["progress"] for c in ctx.report_progress.await_args_list]
    assert progresses == [1.0, 2.0]


def test_forward_rebuild_callbacks_uses_existing_vba_stream():
    queue = asyncio.Queue()
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    reporter = MonotonicProgressReporter()
    state = {"log_messages": []}

    async def _run():
        await queue.put({
            "type": "log",
            "message": "Importing modules...",
        })
        await queue.put({
            "type": "progress",
            "progress": 28,
            "total": 30,
            "message": "Importing queries",
        })
        await queue.put({
            "type": "complete",
            "message": "Operation completed successfully",
            "log_path": r"C:\src\logs\Build_1.log",
        })
        await tools_module._forward_rebuild_callbacks(
            queue,
            ctx,
            reporter,
            state,
        )

    asyncio.run(_run())
    messages = [c.kwargs["message"] for c in ctx.report_progress.await_args_list]
    assert messages == [
        "Importing modules...",
        "Importing queries (28/30)",
        "Build phase complete: Operation completed successfully",
    ]
    assert state["log_messages"] == ["Importing modules..."]
    assert state["log_path"] == r"C:\src\logs\Build_1.log"


def test_wait_ignores_stale_phase_started(tmp_path):
    status = tmp_path / "rebuild-status.json"
    _write_status(
        status,
        status="complete",
        phaseStarted="2026-08-17 16:45:36",
    )
    calls = {"n": 0}

    def wait_for_change(_directory: str, _timeout: float) -> bool:
        calls["n"] += 1
        if calls["n"] == 2:
            _write_status(
                status,
                status="complete",
                phaseStarted="2026-08-27 10:00:00",
            )
        return True

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=5,
        wait_for_change=wait_for_change,
        processes_alive=lambda: True,
    ))
    assert result["success"] is True
    assert result["phaseStarted"] == "2026-08-27 10:00:00"


def test_wait_timeout(tmp_path):
    status = tmp_path / "rebuild-status.json"
    _write_status(status, status="building")

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=0.01,
        wait_for_change=lambda *_a: True,
        processes_alive=lambda: True,
    ))
    assert result["success"] is False
    assert result["error_pattern"] == "timeout"


def test_wait_cancelled(tmp_path):
    status = tmp_path / "rebuild-status.json"
    _write_status(status, status="building")

    async def _run():
        cancel = asyncio.Event()
        cancel.set()
        return await wait_for_rebuild_status(
            str(status),
            "2026-08-27 10:00:00",
            timeout_sec=5,
            cancel_event=cancel,
            wait_for_change=lambda *_a: True,
            processes_alive=lambda: True,
        )

    result = asyncio.run(_run())
    assert result["cancelled"] is True
    assert result["success"] is False


def test_wait_stalled_when_processes_gone(tmp_path):
    status = tmp_path / "rebuild-status.json"
    _write_status(
        status,
        status="building",
        updated="2000-01-01 00:00:00",
    )

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=5,
        wait_for_change=lambda *_a: True,
        processes_alive=lambda: False,
        stall_seconds=1,
    ))
    assert result["success"] is False
    assert result["error_pattern"] == "rebuild_stalled"


def test_wait_terminal_failure(tmp_path):
    status = tmp_path / "rebuild-status.json"
    _write_status(status, status="compile-failed", error="does not compile")

    result = asyncio.run(wait_for_rebuild_status(
        str(status),
        "2026-08-27 10:00:00",
        timeout_sec=5,
        wait_for_change=lambda *_a: True,
        processes_alive=lambda: True,
    ))
    assert result["success"] is False
    assert result["status"] == "compile-failed"
    assert "does not compile" in result["error"]


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


def test_development_addin_from_source(tmp_path):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    host = tmp_path / "Version Control.accda"
    host.write_bytes(b"x")
    assert tools_module._development_addin_from_source(source) == host


def test_rebuild_addin_returns_refusal_without_watch(tmp_path, monkeypatch):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"x")
    logs = source / "logs"
    logs.mkdir()

    refused = json.dumps({
        "success": False,
        "status": "refused",
        "phaseStarted": "2026-08-27 10:00:00",
        "error": "other Access instance",
    })

    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(tools_module, "get_config", lambda: {})
    monkeypatch.setattr(tools_module, "get_callback_url", lambda: None)
    monkeypatch.setattr(
        tools_module,
        "_is_installed_addin_path",
        lambda _p: False,
    )
    monkeypatch.setattr(
        tools_module,
        "_execute_call_vba",
        lambda *_a, **_k: {"success": True, "result": refused},
    )

    gate = MagicMock()

    async def _exclusive(_tool, _db, fn, _is_async, /, *args, **kwargs):
        return await fn(*args, **kwargs)

    gate.run_exclusive_all = _exclusive
    monkeypatch.setattr(tools_module, "get_access_gate", lambda: gate)

    watched = AsyncMock(side_effect=AssertionError("should not watch after refusal"))
    monkeypatch.setattr(tools_module, "wait_for_rebuild_status", watched)

    result = asyncio.run(_unwrap(tools_module.vcs_rebuild_addin)(str(source)))
    assert result["success"] is False
    assert result["status"] == "refused"
    watched.assert_not_awaited()


def test_rebuild_addin_emits_starting_access_before_launch(tmp_path, monkeypatch):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"x")
    launched = json.dumps({
        "success": True,
        "status": "launched",
        "phaseStarted": "2026-08-27 10:00:00",
        "statusFile": str(source / "logs" / "rebuild-status.json"),
    })
    order: list[str] = []
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()

    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(tools_module, "get_config", lambda: {})
    monkeypatch.setattr(tools_module, "get_callback_url", lambda: None)
    monkeypatch.setattr(tools_module, "_is_installed_addin_path", lambda _p: False)

    def _execute(*_a, **_k):
        started = [c.kwargs["message"] for c in ctx.report_progress.await_args_list]
        assert started == ["Starting Access..."]
        order.append("execute")
        return {"success": True, "result": launched}

    monkeypatch.setattr(tools_module, "_execute_call_vba", _execute)

    gate = MagicMock()

    async def _exclusive(_tool, _db, fn, _is_async, /, *args, **kwargs):
        return await fn(*args, **kwargs)

    gate.run_exclusive_all = _exclusive
    monkeypatch.setattr(tools_module, "get_access_gate", lambda: gate)

    async def _watch(*_a, **_k):
        order.append("watch")
        return {
            "success": True,
            "status": "complete",
            "phaseStarted": "2026-08-27 10:00:00",
        }

    monkeypatch.setattr(tools_module, "wait_for_rebuild_status", _watch)

    result = asyncio.run(
        _unwrap(tools_module.vcs_rebuild_addin)(str(source), ctx=ctx)
    )
    assert result["success"] is True
    assert order == ["execute", "watch"]
    messages = [c.kwargs["message"] for c in ctx.report_progress.await_args_list]
    assert messages[0] == "Starting Access..."
    assert "Rebuild launched; waiting for build callbacks" in messages


def test_rebuild_addin_watches_after_launch(tmp_path, monkeypatch):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"x")

    launched = json.dumps({
        "success": True,
        "status": "launched",
        "phaseStarted": "2026-08-27 10:00:00",
        "statusFile": str(source / "logs" / "rebuild-status.json"),
    })
    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(tools_module, "get_config", lambda: {})
    monkeypatch.setattr(tools_module, "get_callback_url", lambda: None)
    monkeypatch.setattr(tools_module, "_is_installed_addin_path", lambda _p: False)
    monkeypatch.setattr(
        tools_module,
        "_execute_call_vba",
        lambda *_a, **_k: {"success": True, "result": launched},
    )

    gate = MagicMock()
    released = {"held": True}

    async def _exclusive(_tool, _db, fn, _is_async, /, *args, **kwargs):
        result = await fn(*args, **kwargs)
        released["held"] = False
        return result

    gate.run_exclusive_all = _exclusive
    monkeypatch.setattr(tools_module, "get_access_gate", lambda: gate)

    async def _watch(*_a, **_k):
        assert released["held"] is False
        return {
            "success": True,
            "status": "complete",
            "phaseStarted": "2026-08-27 10:00:00",
        }

    monkeypatch.setattr(tools_module, "wait_for_rebuild_status", _watch)

    result = asyncio.run(_unwrap(tools_module.vcs_rebuild_addin)(str(source)))
    assert result["success"] is True
    assert result["status"] == "complete"
    assert result["rebuild_phase_started"] == "2026-08-27 10:00:00"


def test_rebuild_addin_passes_callback_identity_to_worker(tmp_path, monkeypatch):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"x")
    callback_queue = asyncio.Queue()
    captured: dict[str, object] = {}

    launched = json.dumps({
        "success": True,
        "status": "launched",
        "phaseStarted": "2026-08-27 10:00:00",
        "statusFile": str(source / "logs" / "rebuild-status.json"),
    })

    class FakeManager:
        def set_event_loop(self, loop):
            captured["loop"] = loop

        def register_operation(self, **kwargs):
            captured["register"] = kwargs
            return "operation-123", callback_queue

        def create_callback_info(self, operation_id, callback_url, client):
            return json.dumps({
                "callback_url": callback_url,
                "operation_id": operation_id,
                "client": client,
            })

        def unregister_operation(self, operation_id):
            captured["unregistered"] = operation_id

    monkeypatch.setattr(tools_module, "check_write_permission", lambda _c: None)
    monkeypatch.setattr(tools_module, "get_config", lambda: {})
    monkeypatch.setattr(
        tools_module,
        "get_callback_url",
        lambda: "http://127.0.0.1:54321/callback",
    )
    monkeypatch.setattr(tools_module, "_get_operation_manager", FakeManager)
    monkeypatch.setattr(tools_module, "_is_installed_addin_path", lambda _p: False)

    def _execute(_database, _function, args, _timeout=None):
        captured["call_args"] = args
        return {"success": True, "result": launched}

    monkeypatch.setattr(tools_module, "_execute_call_vba", _execute)

    gate = MagicMock()

    async def _exclusive(_tool, _db, fn, _is_async, /, *args, **kwargs):
        return await fn(*args, **kwargs)

    gate.run_exclusive_all = _exclusive
    monkeypatch.setattr(tools_module, "get_access_gate", lambda: gate)

    async def _watch(*_args, **_kwargs):
        await callback_queue.put({
            "type": "log",
            "message": "Importing project...",
        })
        await callback_queue.put({
            "type": "complete",
            "message": "Operation completed successfully",
            "log_path": r"C:\src\logs\Build_1.log",
        })
        await asyncio.sleep(0)
        return {
            "success": True,
            "status": "complete",
            "phaseStarted": "2026-08-27 10:00:00",
        }

    monkeypatch.setattr(tools_module, "wait_for_rebuild_status", _watch)

    result = asyncio.run(_unwrap(tools_module.vcs_rebuild_addin)(str(source)))
    call_args = captured["call_args"]
    callback_info = json.loads(call_args[2])
    assert call_args[:2] == ["RebuildAddIn", str(source.resolve())]
    assert callback_info["operation_id"] == "operation-123"
    assert result["log_messages"] == ["Importing project..."]
    assert result["log_path"] == r"C:\src\logs\Build_1.log"
    assert captured["unregistered"] == "operation-123"
