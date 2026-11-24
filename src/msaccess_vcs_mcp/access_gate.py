"""Serialize Access COM work across concurrent MCP tool calls.

One MCP server process is shared by every Cursor window. Sync tool bodies
previously ran on the asyncio event loop and blocked all other requests.
This module fronts Access-touching tools with a single COM apartment thread
and a one-at-a-time slot so callers get a fast ``server_busy`` answer
instead of queueing into a client-side timeout.
"""

from __future__ import annotations

import asyncio
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, Callable

try:
    import pythoncom

    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

DEFAULT_BUSY_WAIT_SEC = 15.0

# Tools that never touch an Access instance — they stay responsive while
# a long export or test run holds the gate.
EXEMPT_TOOLS = frozenset({
    "vcs_get_version_info",
    "vcs_cancel_operation",
    "vcs_get_recent_calls",
    # Reads packaged text, so it answers while Access is blocked.
    "vcs_tips",
    # Launch holds the gate itself; the subsequent status-file wait must not.
    "vcs_rebuild_addin",
    # Window captures use Win32 only so they work while Access is blocked;
    # object captures take the gate themselves.
    "vcs_capture_screenshot",
    # The call blocked behind a dialog or a break usually holds the gate, and
    # answering the dialog or resetting the run is what releases it.
    "vcs_dismiss_dialog",
    "vcs_reset_vba",
})


def _read_busy_wait_sec() -> float:
    try:
        value = float(os.environ.get("ACCESS_VCS_BUSY_WAIT_SEC", str(DEFAULT_BUSY_WAIT_SEC)))
    except ValueError:
        return DEFAULT_BUSY_WAIT_SEC
    return value if value > 0 else DEFAULT_BUSY_WAIT_SEC


def _init_com_apartment() -> None:
    if COM_AVAILABLE:
        pythoncom.CoInitialize()


@dataclass(frozen=True)
class InFlight:
    tool: str
    database: str | None
    started_at: float


def _busy_error(in_flight: InFlight) -> dict[str, Any]:
    elapsed_ms = round((time.perf_counter() - in_flight.started_at) * 1000, 2)
    db_hint = f" on {in_flight.database}" if in_flight.database else ""
    return {
        "success": False,
        "error": (
            f"Another Access operation is in progress ({in_flight.tool}{db_hint}). "
            "The MCP server runs one Access operation at a time across all Cursor "
            "windows sharing this server process. Retry after the in-flight call "
            "completes, or call vcs_get_recent_calls() to see what is running."
        ),
        "error_pattern": "server_busy",
        "recoverable": True,
        "busy_with": {
            "tool": in_flight.tool,
            "database": in_flight.database,
            "elapsed_ms": elapsed_ms,
        },
        "retry_after_seconds": 5,
    }


class AccessGate:
    """One COM apartment thread and one in-flight Access operation at a time."""

    def __init__(self) -> None:
        self._state_lock = threading.Lock()
        self._in_flight: InFlight | None = None
        self._slot = threading.Lock()
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="vcs-access-apartment",
            initializer=_init_com_apartment,
        )
        self._com_initialized = False

    @property
    def com_initialized(self) -> bool:
        return self._com_initialized

    def _mark_com_initialized(self) -> None:
        self._com_initialized = True

    def current_in_flight(self) -> InFlight | None:
        with self._state_lock:
            return self._in_flight

    async def run_exclusive(
        self,
        tool: str,
        database: str | None,
        fn: Callable[..., Any],
        is_async: bool,
        /,
        *args: Any,
        **kwargs: Any,
    ) -> Any:
        wait_sec = _read_busy_wait_sec()
        acquired = await asyncio.to_thread(self._slot.acquire, True, wait_sec)
        if not acquired:
            current = self.current_in_flight()
            if current is not None:
                return _busy_error(current)
            # Slot may have freed between timeout and the read — one short retry.
            acquired = await asyncio.to_thread(self._slot.acquire, True, 0.1)
            if not acquired:
                return _busy_error(
                    InFlight(tool="unknown", database=None, started_at=time.perf_counter())
                )

        with self._state_lock:
            self._in_flight = InFlight(
                tool=tool,
                database=database,
                started_at=time.perf_counter(),
            )

        try:
            if is_async:
                return await fn(*args, **kwargs)

            def _run_sync() -> Any:
                if COM_AVAILABLE and not self._com_initialized:
                    # Executor initializer runs once per worker thread; record it
                    # for tests that assert COM was initialized.
                    self._mark_com_initialized()
                return fn(*args, **kwargs)

            loop = asyncio.get_running_loop()
            return await loop.run_in_executor(self._executor, _run_sync)
        finally:
            with self._state_lock:
                self._in_flight = None
            self._slot.release()


_gate: AccessGate | None = None
_gate_lock = threading.Lock()


def get_access_gate() -> AccessGate:
    global _gate
    with _gate_lock:
        if _gate is None:
            _gate = AccessGate()
        return _gate


def reset_access_gate() -> None:
    """Reset the module singleton (tests only)."""
    global _gate
    with _gate_lock:
        if _gate is not None:
            _gate._executor.shutdown(wait=False, cancel_futures=True)
        _gate = None
