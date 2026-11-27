"""Schedule Access COM work across concurrent MCP tool calls.

One MCP server process is shared by every Cursor window and every agent in
them. Each database gets a lane: a one-at-a-time slot and its own COM
apartment thread, so calls on the same database run in order while calls on
different databases (each in its own Access instance) run in parallel, up to
``ACCESS_VCS_MAX_PARALLEL`` at once. A call that names no database, and
``vcs_rebuild_addin``'s launch, take the whole gate: they wait for every lane
to finish and hold new lane calls off until they are done. A caller that
waits longer than ``ACCESS_VCS_BUSY_WAIT_SEC`` gets a fast ``server_busy``
answer instead of queueing into a client-side timeout.
"""

from __future__ import annotations

import asyncio
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, Callable

from .com_recovery import normalize_database_key

try:
    import pythoncom

    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

DEFAULT_BUSY_WAIT_SEC = 15.0
DEFAULT_MAX_PARALLEL = 3
# A lane's apartment thread is shut down after this long without a call.
LANE_IDLE_SEC = 600.0

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
    # Reads and writes only the server's own settings file, and may wait on
    # the user's answer to a form.
    "vcs_configure",
})


def _read_busy_wait_sec() -> float:
    try:
        value = float(os.environ.get("ACCESS_VCS_BUSY_WAIT_SEC", str(DEFAULT_BUSY_WAIT_SEC)))
    except ValueError:
        return DEFAULT_BUSY_WAIT_SEC
    return value if value > 0 else DEFAULT_BUSY_WAIT_SEC


def read_max_parallel() -> int:
    try:
        value = int(os.environ.get("ACCESS_VCS_MAX_PARALLEL", str(DEFAULT_MAX_PARALLEL)))
    except ValueError:
        return DEFAULT_MAX_PARALLEL
    return value if value > 0 else DEFAULT_MAX_PARALLEL


def _init_com_apartment() -> None:
    if COM_AVAILABLE:
        pythoncom.CoInitialize()


def _apartment() -> ThreadPoolExecutor:
    return ThreadPoolExecutor(
        max_workers=1,
        thread_name_prefix="vcs-access-apartment",
        initializer=_init_com_apartment,
    )


@dataclass(frozen=True)
class InFlight:
    tool: str
    database: str | None
    started_at: float


_BUSY_REASONS = {
    "same_database": (
        "Calls on the same database run one at a time. Retry after it completes."
    ),
    "exclusive": (
        "That operation needs every database to itself, so no other Access call "
        "runs until it completes. Retry after it completes."
    ),
    "exclusive_waiting": (
        "An operation that needs every database to itself is waiting for running "
        "calls to finish, and new calls wait behind it. Retry shortly."
    ),
    "max_parallel": (
        "The server is already running Access calls on {max_parallel} databases, "
        "the most it runs at once (ACCESS_VCS_MAX_PARALLEL). Retry after one completes."
    ),
}


def _busy_error(in_flight: InFlight, reason: str = "same_database") -> dict[str, Any]:
    elapsed_ms = round((time.perf_counter() - in_flight.started_at) * 1000, 2)
    db_hint = f" on {in_flight.database}" if in_flight.database else ""
    detail = _BUSY_REASONS.get(reason, _BUSY_REASONS["same_database"]).format(
        max_parallel=read_max_parallel()
    )
    return {
        "success": False,
        "error": (
            f"Another Access operation is in progress ({in_flight.tool}{db_hint}). "
            f"{detail} Call vcs_get_recent_calls() to see what is running."
        ),
        "error_pattern": "server_busy",
        "recoverable": True,
        "busy_with": {
            "tool": in_flight.tool,
            "database": in_flight.database,
            "elapsed_ms": elapsed_ms,
            "reason": reason,
        },
        "retry_after_seconds": 5,
    }


class _Lane:
    """One database's slot and COM apartment thread."""

    def __init__(self, key: str) -> None:
        self.key = key
        self.in_flight: InFlight | None = None
        self.executor = _apartment()
        self.idle_since = time.monotonic()


class AccessGate:
    """One lane per database, plus an exclusive hold over every lane."""

    def __init__(self) -> None:
        self._cond = threading.Condition()
        self._lanes: dict[str, _Lane] = {}
        self._running = 0
        self._exclusive: InFlight | None = None
        self._exclusive_waiting = 0
        self._exclusive_executor: ThreadPoolExecutor | None = None
        self._com_initialized = False

    @property
    def com_initialized(self) -> bool:
        return self._com_initialized

    def _mark_com_initialized(self) -> None:
        self._com_initialized = True

    def current_in_flight(self, database: str | None = None) -> InFlight | None:
        """What would make a call on ``database`` wait, or any call when None."""
        with self._cond:
            if self._exclusive is not None:
                return self._exclusive
            if database is None:
                return self._oldest_running()
            lane = self._lanes.get(normalize_database_key(database))
            return lane.in_flight if lane else None

    def all_in_flight(self) -> list[InFlight]:
        with self._cond:
            flights = [lane.in_flight for lane in self._lanes.values() if lane.in_flight]
            if self._exclusive is not None:
                flights.insert(0, self._exclusive)
            return flights

    def _oldest_running(self) -> InFlight | None:
        flights = [lane.in_flight for lane in self._lanes.values() if lane.in_flight]
        return min(flights, key=lambda f: f.started_at) if flights else None

    # -- acquiring and releasing (run on worker threads, never the event loop) --

    def _lane_blocker(self, lane: _Lane | None) -> tuple[InFlight, str] | None:
        if self._exclusive is not None:
            return self._exclusive, "exclusive"
        if self._exclusive_waiting:
            waiting_on = self._oldest_running() or InFlight(
                tool="an exclusive operation", database=None, started_at=time.perf_counter()
            )
            return waiting_on, "exclusive_waiting"
        if lane is not None and lane.in_flight is not None:
            return lane.in_flight, "same_database"
        if self._running >= read_max_parallel():
            return self._oldest_running(), "max_parallel"
        return None

    def _acquire_lane(self, key: str, flight: InFlight, wait_sec: float) -> _Lane | dict[str, Any]:
        deadline = time.monotonic() + wait_sec
        with self._cond:
            while True:
                lane = self._lanes.get(key)
                blocker = self._lane_blocker(lane)
                if blocker is None:
                    if lane is None:
                        lane = _Lane(key)
                        self._lanes[key] = lane
                    lane.in_flight = flight
                    self._running += 1
                    return lane
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return _busy_error(*blocker)
                self._cond.wait(remaining)

    def _release_lane(self, lane: _Lane) -> None:
        with self._cond:
            lane.in_flight = None
            lane.idle_since = time.monotonic()
            self._running -= 1
            self._shut_idle_lanes()
            self._cond.notify_all()

    def _shut_idle_lanes(self) -> None:
        cutoff = time.monotonic() - LANE_IDLE_SEC
        for key, lane in list(self._lanes.items()):
            if lane.in_flight is None and lane.idle_since < cutoff:
                lane.executor.shutdown(wait=False)
                del self._lanes[key]

    def _acquire_all(self, flight: InFlight, wait_sec: float) -> ThreadPoolExecutor | dict[str, Any]:
        deadline = time.monotonic() + wait_sec
        with self._cond:
            self._exclusive_waiting += 1
            try:
                while self._exclusive is not None or self._running > 0:
                    remaining = deadline - time.monotonic()
                    if remaining <= 0:
                        if self._exclusive is not None:
                            return _busy_error(self._exclusive, "exclusive")
                        return _busy_error(self._oldest_running(), "same_database")
                    self._cond.wait(remaining)
                self._exclusive = flight
                if self._exclusive_executor is None:
                    self._exclusive_executor = _apartment()
                return self._exclusive_executor
            finally:
                self._exclusive_waiting -= 1
                self._cond.notify_all()

    def _release_all(self) -> None:
        with self._cond:
            self._exclusive = None
            self._cond.notify_all()

    async def _acquire(self, acquire: Callable[[], Any], release: Callable[[Any], None]) -> Any:
        """Wait on a worker thread; if the caller is cancelled, release what it got."""
        pending = asyncio.ensure_future(asyncio.to_thread(acquire))
        try:
            return await asyncio.shield(pending)
        except asyncio.CancelledError:
            def _undo(task: asyncio.Future) -> None:
                if not task.cancelled() and task.exception() is None:
                    got = task.result()
                    if not isinstance(got, dict):
                        release(got)
            pending.add_done_callback(_undo)
            raise

    async def _run(
        self,
        executor: ThreadPoolExecutor,
        fn: Callable[..., Any],
        is_async: bool,
        args: tuple,
        kwargs: dict,
    ) -> Any:
        if is_async:
            return await fn(*args, **kwargs)

        def _run_sync() -> Any:
            if COM_AVAILABLE and not self._com_initialized:
                # Executor initializer runs once per worker thread; record it
                # for tests that assert COM was initialized.
                self._mark_com_initialized()
            return fn(*args, **kwargs)

        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(executor, _run_sync)

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
        """Run ``fn`` in ``database``'s lane, or holding every lane when None."""
        if database is None:
            return await self.run_exclusive_all(tool, None, fn, is_async, *args, **kwargs)

        flight = InFlight(tool=tool, database=database, started_at=time.perf_counter())
        key = normalize_database_key(database)
        wait_sec = _read_busy_wait_sec()
        lane = await self._acquire(
            lambda: self._acquire_lane(key, flight, wait_sec), self._release_lane
        )
        if isinstance(lane, dict):
            return lane
        try:
            return await self._run(lane.executor, fn, is_async, args, kwargs)
        finally:
            self._release_lane(lane)

    async def run_exclusive_all(
        self,
        tool: str,
        database: str | None,
        fn: Callable[..., Any],
        is_async: bool,
        /,
        *args: Any,
        **kwargs: Any,
    ) -> Any:
        """Run ``fn`` once no lane is busy, holding new lane calls off meanwhile.

        ``database`` only labels the call in ``busy_with``.
        """
        flight = InFlight(tool=tool, database=database, started_at=time.perf_counter())
        wait_sec = _read_busy_wait_sec()
        executor = await self._acquire(
            lambda: self._acquire_all(flight, wait_sec), lambda _: self._release_all()
        )
        if isinstance(executor, dict):
            return executor
        try:
            return await self._run(executor, fn, is_async, args, kwargs)
        finally:
            self._release_all()

    def shutdown(self) -> None:
        with self._cond:
            for lane in self._lanes.values():
                lane.executor.shutdown(wait=False, cancel_futures=True)
            self._lanes.clear()
            if self._exclusive_executor is not None:
                self._exclusive_executor.shutdown(wait=False, cancel_futures=True)
                self._exclusive_executor = None


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
            _gate.shutdown()
        _gate = None
