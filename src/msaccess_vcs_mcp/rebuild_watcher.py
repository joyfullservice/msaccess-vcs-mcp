"""Watch ``rebuild-status.json`` after an add-in self-rebuild launches.

``RebuildAddIn`` returns once the worker is confirmed running, then the host
Access instance exits so the files it held can be replaced. Live COM
callbacks die with that process. The worker writes coarse phases to
``<source>/logs/rebuild-status.json``; this module observes that file with
Windows directory-change notifications (poll fallback) and reports each
phase via MCP progress until a terminal status matching this attempt's
``phaseStarted``.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import time
from datetime import datetime
from typing import Any, Callable

from .access_com.process_qos import prefer_full_power_new_access
from .operation_manager import MonotonicProgressReporter

logger = logging.getLogger(__name__)

DEFAULT_REBUILD_TIMEOUT_SEC = 1200.0
STALL_SECONDS = 90.0
POLL_INTERVAL_SEC = 0.5
WATCH_CHUNK_SEC = 2.0
DEBOUNCE_SEC = 0.05

TERMINAL_STATUSES = frozenset({
    "complete",
    "refused",
    "launch-failed",
    "build-failed",
    "compile-failed",
    "install-failed",
    "cancelled",
})

WaitForChange = Callable[[str, float], bool]
ReadStatus = Callable[[str], dict[str, Any] | None]
ProcessesAlive = Callable[[], bool]
PromoteNewAccess = Callable[[set[int]], set[int]]


def get_rebuild_timeout(timeout_seconds: float | None = None) -> float:
    """Resolve the watch-phase timeout for one add-in rebuild."""
    if timeout_seconds is not None and timeout_seconds > 0:
        return float(timeout_seconds)
    try:
        value = float(os.environ.get(
            "ACCESS_VCS_REBUILD_TIMEOUT_SEC",
            str(DEFAULT_REBUILD_TIMEOUT_SEC),
        ))
    except ValueError:
        return DEFAULT_REBUILD_TIMEOUT_SEC
    return value if value > 0 else DEFAULT_REBUILD_TIMEOUT_SEC


def read_rebuild_status(status_path: str) -> dict[str, Any] | None:
    """Parse ``rebuild-status.json``, tolerating a UTF-8 BOM and partial writes."""
    if not status_path or not os.path.isfile(status_path):
        return None
    try:
        with open(status_path, encoding="utf-8-sig") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError, UnicodeError):
        return None
    if not isinstance(data, dict):
        return None
    return data


def rebuild_processes_alive() -> bool:
    """True when MSACCESS.EXE or wscript.exe is running, or the query failed.

    A live rebuild always has at least one of those processes. Failure to ask
    is treated as alive so a query problem cannot be read as a stalled run.
    """
    try:
        import subprocess

        for image in ("MSACCESS.EXE", "wscript.exe"):
            result = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {image}", "/NH"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            if image.lower() in (result.stdout or "").lower():
                return True
        return False
    except Exception:
        return True


def poll_directory(directory: str, timeout_sec: float) -> bool:
    """Sleep up to ``timeout_sec``. Always returns True so the caller re-reads."""
    if timeout_sec > 0:
        time.sleep(timeout_sec)
    return True


def native_watch_directory(directory: str, timeout_sec: float) -> bool:
    """Block until the directory changes or ``timeout_sec`` elapses.

    Uses ``ReadDirectoryChangesW`` when pywin32 is available. Closing the
    directory handle unblocks a timed-out wait. Falls back to polling if
    the directory is missing or native watching cannot start.
    """
    if timeout_sec <= 0:
        return os.path.isdir(directory)
    if not os.path.isdir(directory):
        return poll_directory(directory, min(timeout_sec, POLL_INTERVAL_SEC))

    try:
        import win32con
        import win32event
        import win32file
        import pywintypes
    except ImportError:
        return poll_directory(directory, min(timeout_sec, POLL_INTERVAL_SEC))

    handle = None
    try:
        handle = win32file.CreateFile(
            directory,
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ
            | win32con.FILE_SHARE_WRITE
            | win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_FLAG_BACKUP_SEMANTICS
            | win32con.FILE_FLAG_OVERLAPPED,
            None,
        )
    except Exception:
        return poll_directory(directory, min(timeout_sec, POLL_INTERVAL_SEC))

    try:
        buffer = win32file.AllocateReadBuffer(8192)
        overlapped = pywintypes.OVERLAPPED()
        overlapped.hEvent = win32event.CreateEvent(None, True, False, None)
        win32file.ReadDirectoryChangesW(
            handle,
            buffer,
            False,
            win32con.FILE_NOTIFY_CHANGE_LAST_WRITE
            | win32con.FILE_NOTIFY_CHANGE_FILE_NAME
            | win32con.FILE_NOTIFY_CHANGE_SIZE,
            overlapped,
            None,
        )
        wait_result = win32event.WaitForSingleObject(
            overlapped.hEvent,
            max(1, int(timeout_sec * 1000)),
        )
        if wait_result == win32event.WAIT_OBJECT_0:
            win32file.GetOverlappedResult(handle, overlapped, False)
            return True
        if wait_result == win32event.WAIT_TIMEOUT:
            win32file.CancelIo(handle)
            return False
        win32file.CancelIo(handle)
        return False
    except Exception:
        return poll_directory(directory, min(timeout_sec, POLL_INTERVAL_SEC))
    finally:
        try:
            handle.Close()
        except Exception:
            pass


def _parse_updated(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def _is_stale(updated: Any, now: Callable[[], float], stall_seconds: float) -> bool:
    parsed = _parse_updated(updated)
    if parsed is None:
        return False
    age = now() - parsed.timestamp()
    return age > stall_seconds


async def wait_for_rebuild_status(
    status_path: str,
    phase_started: str,
    *,
    timeout_sec: float | None = None,
    ctx: Any = None,
    reporter: MonotonicProgressReporter | None = None,
    cancel_event: asyncio.Event | None = None,
    wait_for_change: WaitForChange | None = None,
    read_status: ReadStatus | None = None,
    processes_alive: ProcessesAlive | None = None,
    now: Callable[[], float] | None = None,
    stall_seconds: float = STALL_SECONDS,
    preexisting_access_pids: set[int] | None = None,
    promote_new_access: PromoteNewAccess | None = None,
) -> dict[str, Any]:
    """Wait until this rebuild attempt reaches a terminal status.

    Records whose ``phaseStarted`` does not match ``phase_started`` belong to
    another run and are ignored. Cancellation abandons the wait only; the
    worker is left running.

    When ``preexisting_access_pids`` is set, Access processes that appear
    after that snapshot (builder, silent installer) get full-power QoS.
    Access that was already running is left alone.
    """
    timeout = get_rebuild_timeout(timeout_sec)
    wait_fn = wait_for_change or native_watch_directory
    read_fn = read_status or read_rebuild_status
    alive_fn = processes_alive if processes_alive is not None else rebuild_processes_alive
    now_fn = now or time.time
    progress_reporter = reporter or MonotonicProgressReporter()
    deadline = time.monotonic() + timeout
    last_status: str | None = None
    logs_dir = os.path.dirname(status_path) or "."
    known_pids = (
        set(preexisting_access_pids) if preexisting_access_pids is not None else None
    )
    promote_fn = promote_new_access or prefer_full_power_new_access

    while True:
        if known_pids is not None:
            known_pids = promote_fn(known_pids)

        if cancel_event is not None and cancel_event.is_set():
            return {
                "success": False,
                "cancelled": True,
                "message": "Wait cancelled; worker may still be running",
                "status_file": status_path,
            }

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return {
                "success": False,
                "error": f"Rebuild timed out after {timeout} seconds",
                "error_pattern": "timeout",
                "status_file": status_path,
            }

        record = read_fn(status_path)
        if record and record.get("phaseStarted") == phase_started:
            status = str(record.get("status") or "")
            if status and status != last_status:
                last_status = status
                await progress_reporter.emit(ctx, message=f"Rebuild {status}")
                if status in TERMINAL_STATUSES:
                    result = dict(record)
                    result["status_file"] = status_path
                    if status == "complete":
                        result["success"] = True
                    elif status == "cancelled":
                        result["success"] = False
                        result["cancelled"] = True
                        result["error_pattern"] = "operation_cancelled"
                        result.setdefault(
                            "error", record.get("error") or "The rebuild was canceled."
                        )
                    else:
                        result["success"] = False
                        result.setdefault(
                            "error",
                            record.get("error") or f"Rebuild ended with status {status}",
                        )
                    return result

            if _is_stale(record.get("updated"), now_fn, stall_seconds) and not alive_fn():
                return {
                    "success": False,
                    "error": (
                        "Rebuild status stopped advancing and no MSACCESS.EXE or "
                        "wscript.exe process is running. The worker likely died "
                        "without writing a terminal status."
                    ),
                    "error_pattern": "rebuild_stalled",
                    "status": status or None,
                    "status_file": status_path,
                    "updated": record.get("updated"),
                }

        chunk = min(WATCH_CHUNK_SEC, remaining)
        if DEBOUNCE_SEC > 0:
            await asyncio.sleep(min(DEBOUNCE_SEC, chunk))
        await asyncio.to_thread(wait_fn, logs_dir, max(chunk - DEBOUNCE_SEC, 0.0))
