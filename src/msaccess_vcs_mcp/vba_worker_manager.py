"""Thread-based management for isolated VBA execution with timeout."""

from __future__ import annotations

import json
import os
import threading
import time
from typing import Any, Optional

try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

from .access_com.connection import (
    AccessConnection,
    ensure_access_visible,
    recycle_owned_instance,
)
from .addin_integration import VCSAddinIntegration
from .com_recovery import (
    classify_com_error,
    get_recovery_manager,
    is_recoverable_pattern,
)
from .usage_logging import (
    log_com_recovery_event,
    log_diagnostic_event,
    log_vba_worker_event,
)


DEFAULT_RUN_VBA_TIMEOUT_SEC = 45.0
DEFAULT_CALL_VBA_TIMEOUT_SEC = 45.0
DEFAULT_BUILD_TIMEOUT_SEC = 600.0
DEFAULT_RECOVERY_PROBE_TIMEOUT_SEC = 10.0


def _reset_failure(error: str) -> dict[str, Any]:
    """Build a fail-closed reset result for the public tool response."""
    return {
        "success": False,
        "resetQueued": False,
        "error": error,
        "error_pattern": "reset_failed",
        "phase": "reset_state",
    }


def _parse_reset_result(raw_result: Any) -> dict[str, Any]:
    """Parse and validate the add-in's pre-RunVBA reset response."""
    if not isinstance(raw_result, str):
        return _reset_failure(
            "ResetVbaProjectState returned a non-JSON response."
        )

    try:
        result = json.loads(raw_result)
    except (TypeError, json.JSONDecodeError) as exc:
        return _reset_failure(
            f"ResetVbaProjectState returned invalid JSON: {exc}"
        )

    if not isinstance(result, dict):
        return _reset_failure(
            "ResetVbaProjectState returned JSON that was not an object."
        )

    result.setdefault("phase", "reset_state")
    if result.get("success") is True and result.get("resetQueued") is True:
        return result

    result["success"] = False
    result["resetQueued"] = False
    result.setdefault(
        "error",
        "ResetVbaProjectState did not confirm that the reset was queued.",
    )
    result.setdefault("error_pattern", "reset_failed")
    return result


def _read_timeout_env(name: str, default: float) -> float:
    try:
        value = float(os.environ.get(name, str(default)))
    except ValueError:
        return default
    return value if value > 0 else default


def get_run_vba_timeout(timeout_seconds: float | None = None) -> float:
    """Resolve the effective timeout for one `vcs_run_vba` call."""
    if timeout_seconds is not None and timeout_seconds > 0:
        return float(timeout_seconds)
    return _read_timeout_env("ACCESS_VCS_RUN_VBA_TIMEOUT_SEC", DEFAULT_RUN_VBA_TIMEOUT_SEC)


def get_call_vba_timeout(timeout_seconds: float | None = None) -> float:
    """Resolve the effective timeout for one `vcs_call_vba` call."""
    if timeout_seconds is not None and timeout_seconds > 0:
        return float(timeout_seconds)
    return _read_timeout_env("ACCESS_VCS_CALL_VBA_TIMEOUT_SEC", DEFAULT_CALL_VBA_TIMEOUT_SEC)


def get_build_timeout(timeout_seconds: float | None = None) -> float:
    """Resolve the effective timeout for one `vcs_rebuild_database` build."""
    if timeout_seconds is not None and timeout_seconds > 0:
        return float(timeout_seconds)
    return _read_timeout_env("ACCESS_VCS_BUILD_TIMEOUT_SEC", DEFAULT_BUILD_TIMEOUT_SEC)


def get_recovery_probe_timeout() -> float:
    """Resolve the timeout for short recovery health probes."""
    return _read_timeout_env(
        "ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC",
        DEFAULT_RECOVERY_PROBE_TIMEOUT_SEC,
    )


class VBAWorkerManager:
    """Run VBA in a daemon thread with a hard timeout.

    Modelled on ``_probe_with_timeout`` in ``addin_integration.py`` and
    ``_run_dao_with_timeout`` in db-inspector-mcp.  The worker thread
    creates its own COM apartment via ``pythoncom.CoInitialize()`` and
    re-acquires the Access instance through the Running Object Table so
    that ``thread.join(timeout)`` can fire even if the COM call blocks.
    """

    _active_worker: Optional[threading.Thread] = None

    def __init__(self) -> None:
        self.recovery = get_recovery_manager()

    def run_vba(
        self,
        database_path: str,
        code: str,
        addin_path: str | None = None,
        timeout_seconds: float | None = None,
    ) -> dict[str, Any]:
        """Run VBA in a timeout-controlled worker thread with recovery."""
        timeout = get_run_vba_timeout(timeout_seconds)
        preflight = self._probe_if_needed(database_path, addin_path)
        if preflight is not None:
            return preflight

        result = self._run_worker(
            operation="run_vba",
            database_path=database_path,
            addin_path=addin_path,
            code=code,
            timeout_seconds=timeout,
        )
        if result.get("success"):
            self.recovery.mark_healthy(database_path)
            return result

        return self._handle_run_failure(
            database_path=database_path,
            addin_path=addin_path,
            code=code,
            timeout_seconds=timeout,
            result=result,
        )

    def probe(self, database_path: str, addin_path: str | None = None) -> dict[str, Any]:
        """Run a short Access/add-in health probe in a worker thread."""
        return self._run_worker(
            operation="probe",
            database_path=database_path,
            addin_path=addin_path,
            code=None,
            timeout_seconds=get_recovery_probe_timeout(),
        )

    # ------------------------------------------------------------------
    # Recovery helpers (unchanged from subprocess version)
    # ------------------------------------------------------------------

    def _probe_if_needed(
        self,
        database_path: str,
        addin_path: str | None,
    ) -> dict[str, Any] | None:
        if not self.recovery.should_probe(database_path):
            return None

        state = self.recovery.mark_probing(database_path)
        log_com_recovery_event(
            "com_recovery_probe_start",
            database_path=database_path,
            status=state.status,
        )
        log_diagnostic_event(
            "com_recovery_probe_start",
            database=str(database_path),
            status=state.status,
        )

        probe_result = self.probe(database_path, addin_path)
        if probe_result.get("success"):
            recovered = self.recovery.mark_healthy(database_path, recovered=True)
            log_com_recovery_event(
                "com_recovery_probe_result",
                database_path=database_path,
                status=recovered.status,
                success=True,
            )
            log_diagnostic_event(
                "com_recovery_probe_result",
                database=str(database_path),
                status=recovered.status,
                success=True,
            )
            return None

        state = self.recovery.mark_failure(
            database_path,
            probe_result.get("error", "Access recovery probe failed"),
            probe_result.get("error_pattern"),
            timed_out=probe_result.get("timed_out", False),
        )
        log_com_recovery_event(
            "com_recovery_probe_result",
            database_path=database_path,
            status=state.status,
            success=False,
            error=probe_result.get("error"),
            error_pattern=state.error_pattern,
        )
        log_diagnostic_event(
            "com_recovery_probe_result",
            database=str(database_path),
            status=state.status,
            success=False,
            error=probe_result.get("error"),
            error_pattern=state.error_pattern,
        )

        recycled = False
        try:
            recycled = recycle_owned_instance(database_path)
        except Exception as exc:
            log_diagnostic_event(
                "owned_instance_recycle_failed",
                database=str(database_path),
                error=str(exc),
            )
            recycled = False

        if recycled:
            log_diagnostic_event(
                "owned_instance_recycled",
                database=str(database_path),
            )
            retry_probe = self.probe(database_path, addin_path)
            if retry_probe.get("success"):
                recovered = self.recovery.mark_healthy(database_path, recovered=True)
                log_com_recovery_event(
                    "com_recovery_probe_result",
                    database_path=database_path,
                    status=recovered.status,
                    success=True,
                )
                log_diagnostic_event(
                    "com_recovery_probe_result",
                    database=str(database_path),
                    status=recovered.status,
                    success=True,
                    recycled=True,
                )
                return None
            state = self.recovery.mark_failure(
                database_path,
                retry_probe.get("error", "Access recovery probe failed after recycle"),
                retry_probe.get("error_pattern"),
                timed_out=retry_probe.get("timed_out", False),
            )

        return self._recoverable_error(
            "Access is still not responding after a recovery probe.",
            state.error_pattern or "access_unresponsive",
            agent_guidance=(
                "Call vcs_capture_screenshot(database_path, include_image=False) to see "
                "whether a dialog or VBA break mode is holding Access. Answer it with "
                "vcs_dismiss_dialog or vcs_reset_vba when owned_by_server is true; "
                "otherwise ask the user to. Then retry. For more, call vcs_tips(\"stuck\")."
            ),
            recovery_status=state.status,
        )

    def _handle_run_failure(
        self,
        database_path: str,
        addin_path: str | None,
        code: str,
        timeout_seconds: float,
        result: dict[str, Any],
    ) -> dict[str, Any]:
        error = result.get("error", "VBA worker failed")
        pattern = result.get("error_pattern") or classify_com_error(error)
        timed_out = result.get("timed_out", False)
        phase = result.get("phase")
        state = self.recovery.mark_failure(
            database_path,
            error,
            pattern,
            timed_out=timed_out,
        )

        if timed_out:
            return self._recoverable_error(
                error,
                "timeout",
                agent_guidance=(
                    "Access may still be running the submitted VBA, or waiting on a "
                    "dialog or break mode. Call vcs_capture_screenshot(database_path, "
                    "include_image=False) to see which before retrying. For more, call "
                    "vcs_tips(\"stuck\")."
                ),
                recovery_status=state.status,
                timed_out=True,
            )

        if (
            is_recoverable_pattern(pattern)
            and phase in {"connect", "load_addin"}
        ):
            probe_result = self.probe(database_path, addin_path)
            if probe_result.get("success"):
                self.recovery.mark_healthy(database_path, recovered=True)
                retry_result = self._run_worker(
                    operation="run_vba",
                    database_path=database_path,
                    addin_path=addin_path,
                    code=code,
                    timeout_seconds=timeout_seconds,
                    retry=True,
                )
                if retry_result.get("success"):
                    self.recovery.mark_healthy(database_path)
                    return retry_result

                retry_error = retry_result.get("error", "VBA worker retry failed")
                retry_pattern = retry_result.get("error_pattern") or classify_com_error(retry_error)
                retry_state = self.recovery.mark_failure(
                    database_path,
                    retry_error,
                    retry_pattern,
                    timed_out=retry_result.get("timed_out", False),
                )
                return self._recoverable_error(
                    retry_error,
                    retry_pattern,
                    agent_guidance=(
                        "Access responded to the recovery probe, but the retry failed. "
                        "Call vcs_capture_screenshot(database_path, include_image=False) "
                        "to confirm Access is responsive before retrying."
                    ),
                    recovery_status=retry_state.status,
                    timed_out=retry_result.get("timed_out", False),
                    phase=retry_result.get("phase"),
                    recoverable=is_recoverable_pattern(retry_pattern),
                )

        return self._recoverable_error(
            error,
            pattern,
            agent_guidance=(
                "The failed call was not retried automatically because the VBA "
                "may have started. Call vcs_capture_screenshot(database_path, "
                "include_image=False) to confirm Access is responsive, then retry."
            ),
            recovery_status=state.status,
            phase=phase,
            recoverable=is_recoverable_pattern(pattern),
        )

    # ------------------------------------------------------------------
    # Thread-based worker (replaces subprocess)
    # ------------------------------------------------------------------

    def _run_worker(
        self,
        operation: str,
        database_path: str,
        addin_path: str | None,
        code: str | None,
        timeout_seconds: float,
        retry: bool = False,
    ) -> dict[str, Any]:
        cls = type(self)
        if cls._active_worker is not None and cls._active_worker.is_alive():
            return {
                "success": False,
                "error": (
                    "A previous VBA worker thread is still running. "
                    "Access may be in VBA break mode or blocked on a modal dialog."
                ),
                "error_pattern": "access_unresponsive",
                "recoverable": True,
                "phase": "guard",
            }

        start = time.perf_counter()
        log_vba_worker_event(
            "vba_worker_start",
            database_path=database_path,
            operation=operation,
            retry=retry,
        )
        log_diagnostic_event(
            "vba_worker_start",
            database=str(database_path),
            operation=operation,
            timeout_seconds=timeout_seconds,
            retry=retry,
        )

        result_box: dict[str, Any] = {}

        def worker() -> None:
            phase = "start"

            def set_phase(next_phase: str) -> None:
                nonlocal phase
                phase = next_phase
                log_diagnostic_event(
                    "vba_worker_phase",
                    database=str(database_path),
                    operation=operation,
                    phase=phase,
                    retry=retry,
                )

            held_conn = None

            try:
                if not COM_AVAILABLE:
                    raise ImportError("pywin32 is required for COM automation")
                pythoncom.CoInitialize()
                try:
                    set_phase("connect")
                    worker_app, held_conn = _find_or_open_access(database_path)
                    if worker_app is None:
                        raise RuntimeError(
                            f"Cannot find Access instance for {database_path} "
                            f"from worker thread. The Access application "
                            f"may have been closed."
                        )
                    # Submitted code can break into the VBE or raise a dialog;
                    # both are only recoverable in a window someone can see.
                    ensure_access_visible(worker_app)

                    set_phase("load_addin")
                    addin = VCSAddinIntegration(addin_path)
                    addin.load_addin(worker_app, db_path=database_path)

                    if operation == "probe":
                        result_box["result"] = {
                            "success": True,
                            "operation": operation,
                            "phase": phase,
                            "result": "ok",
                        }
                        return

                    set_phase("reset_state")
                    reset_result = _parse_reset_result(
                        addin.call_sync("ResetVbaProjectState")
                    )
                    if not reset_result.get("success"):
                        log_diagnostic_event(
                            "vba_worker_reset_failed",
                            database=str(database_path),
                            operation=operation,
                            phase=phase,
                            error=reset_result.get("error"),
                            error_pattern=reset_result.get("error_pattern"),
                            retry=retry,
                        )
                        result_box["result"] = {
                            # The COM round trip completed. Return the add-in's
                            # logical failure through the normal JSON channel.
                            "success": True,
                            "operation": operation,
                            "phase": phase,
                            "result": json.dumps(reset_result),
                        }
                        return

                    set_phase("reset_barrier")
                    try:
                        # A built-in COM property read gives the queued VBE
                        # teardown a safe message pump with no host VBA payload
                        # running beneath it.
                        _ = worker_app.CurrentProject.FullName
                    except Exception as barrier_error:
                        log_diagnostic_event(
                            "vba_worker_reset_barrier_retry",
                            database=str(database_path),
                            operation=operation,
                            phase=phase,
                            error=str(barrier_error),
                            error_pattern=classify_com_error(barrier_error),
                            retry=retry,
                        )

                    # Never carry pre-reset COM proxies into payload execution.
                    addin = None
                    worker_app, reacquired_conn = _find_or_open_access(database_path)
                    if reacquired_conn is not None:
                        held_conn = reacquired_conn
                    if worker_app is None:
                        raise RuntimeError(
                            f"Cannot reacquire Access instance for {database_path} "
                            "after resetting the VBA project."
                        )
                    _ = worker_app.CurrentProject.FullName
                    ensure_access_visible(worker_app)

                    set_phase("reacquire_addin")
                    addin = VCSAddinIntegration(addin_path)
                    addin.load_addin(worker_app, db_path=database_path)

                    set_phase("run_vba")
                    vba_result = addin.call_sync("RunVBA", code)
                    result_box["result"] = {
                        "success": True,
                        "operation": operation,
                        "phase": phase,
                        "result": vba_result,
                    }
                finally:
                    if held_conn is not None:
                        try:
                            held_conn.close()
                        except Exception:
                            pass
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                result_box["result"] = {
                    "success": False,
                    "operation": operation,
                    "phase": phase,
                    "error": str(exc),
                    "error_type": type(exc).__name__,
                    "error_pattern": classify_com_error(exc),
                }

        thread = threading.Thread(target=worker, daemon=True, name="vcs-vba-worker")
        cls._active_worker = thread
        thread.start()
        thread.join(timeout=timeout_seconds)

        duration_ms = round((time.perf_counter() - start) * 1000, 2)

        if thread.is_alive():
            error = f"VBA worker timed out after {timeout_seconds} seconds"
            log_vba_worker_event(
                "vba_worker_timeout",
                database_path=database_path,
                operation=operation,
                duration_ms=duration_ms,
                success=False,
                timed_out=True,
                error=error,
                error_pattern="timeout",
                retry=retry,
            )
            log_diagnostic_event(
                "vba_worker_timeout",
                database=str(database_path),
                operation=operation,
                duration_ms=duration_ms,
                error=error,
                retry=retry,
            )
            return {
                "success": False,
                "error": error,
                "error_pattern": "timeout",
                "recoverable": True,
                "timed_out": True,
                "duration_ms": duration_ms,
            }

        cls._active_worker = None

        response = result_box.get("result", {
            "success": False,
            "error": "VBA worker thread completed without producing a result",
            "error_pattern": "unknown",
        })
        response["duration_ms"] = duration_ms
        if not response.get("success"):
            response.setdefault("recoverable", is_recoverable_pattern(response.get("error_pattern")))

        log_vba_worker_event(
            "vba_worker_result",
            database_path=database_path,
            operation=operation,
            duration_ms=duration_ms,
            success=response.get("success"),
            timed_out=False,
            error=response.get("error"),
            error_pattern=response.get("error_pattern"),
            phase=response.get("phase"),
            retry=retry,
        )
        log_diagnostic_event(
            "vba_worker_result",
            database=str(database_path),
            operation=operation,
            duration_ms=duration_ms,
            success=response.get("success"),
            phase=response.get("phase"),
            error=response.get("error"),
            error_pattern=response.get("error_pattern"),
            retry=retry,
        )
        return response

    @staticmethod
    def _recoverable_error(
        error: str,
        error_pattern: str,
        agent_guidance: str,
        recovery_status: str,
        timed_out: bool = False,
        phase: str | None = None,
        recoverable: bool = True,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": False,
            "error": error,
            "error_pattern": error_pattern,
            "agent_guidance": agent_guidance,
            "recoverable": recoverable,
            "recovery_status": recovery_status,
        }
        if timed_out:
            result["timed_out"] = True
        if phase:
            result["phase"] = phase
        return result


def _find_or_open_access(database_path: str) -> tuple[Any, Any]:
    """Return ``(app, connection)`` for ``database_path``, opening if needed.

    The worker originally only looked in the Running Object Table. When
    nothing held the file open -- including after the server quit its own
    instance -- every ``vcs_run_vba`` failed with "Cannot find Access
    instance".

    Any connection opened here is returned rather than closed, and the
    caller keeps it alive for the whole operation. Closing it immediately
    would quit the instance under ``ACCESS_VCS_LEAVE_ACCESS_OPEN=false``,
    leaving the ROT lookup on the next line to find nothing.
    """
    worker_app = VCSAddinIntegration._find_access_in_rot(database_path)
    if worker_app is not None:
        return worker_app, None
    conn = AccessConnection(database_path)
    try:
        conn.connect()
    except Exception:
        conn.close()
        raise
    return VCSAddinIntegration._find_access_in_rot(database_path), conn


_worker_manager = VBAWorkerManager()


def run_vba_resilient(
    database_path: str,
    code: str,
    addin_path: str | None = None,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """Run VBA through the thread-isolated worker manager."""
    return _worker_manager.run_vba(
        database_path=database_path,
        code=code,
        addin_path=addin_path,
        timeout_seconds=timeout_seconds,
    )
