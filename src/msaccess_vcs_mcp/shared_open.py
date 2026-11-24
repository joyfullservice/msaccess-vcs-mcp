"""Warn when another session holds a database the server is about to merge into.

Design and VBA changes made while another session has the database open can
fail to save, or be dropped when that session saves its own copy. The loss
cannot be confirmed in advance, so imports still run; the result says who
else holds the file and what to do.

Local holders come from the Restart Manager, which asks the kernel and so
answers while Access is hung. Other computers appear only in the lock file
(``.laccdb`` / ``.ldb``), whose entries outlive the sessions that wrote
them, so those warnings say the entry may be stale.
"""

from __future__ import annotations

import os
import socket
from pathlib import Path
from typing import Any, Callable

from .guidance import insert_after

LOCK_RECORD_BYTES = 64
LOCK_NAME_BYTES = 32
LOCK_EXTENSIONS = {".accdb": ".laccdb", ".accda": ".laccdb", ".accde": ".laccdb", ".mdb": ".ldb"}
CONSEQUENCE = "design and VBA changes may not save"
ACTION = "Ask the user to close it, then verify with vcs_diff_database."


def lock_file_for(database_path: str) -> Path | None:
    path = Path(database_path)
    extension = LOCK_EXTENSIONS.get(path.suffix.lower())
    return path.with_suffix(extension) if extension else None


def lock_file_machines(lock_path: Path) -> list[str]:
    """Machine names in a lock file, in order, without duplicates."""
    try:
        data = lock_path.read_bytes()
    except OSError:
        return []
    machines: list[str] = []
    for offset in range(0, len(data) - LOCK_NAME_BYTES + 1, LOCK_RECORD_BYTES):
        raw = data[offset:offset + LOCK_NAME_BYTES].split(b"\x00", 1)[0]
        name = raw.decode("latin-1", errors="replace").strip()
        if name and name not in machines:
            machines.append(name)
    return machines


def _local_names() -> set[str]:
    names = {socket.gethostname(), os.environ.get("COMPUTERNAME", "")}
    return {n.upper() for n in names if n}


def _default_holders(path: str) -> set[int] | None:
    from .screen_capture import pids_holding_file

    return pids_holding_file(path)


def _default_access_pids() -> set[int]:
    from .access_com.process_qos import list_access_pids

    return set(list_access_pids())


def _default_driven_pid(app: Any) -> int | None:
    from .access_com.process_qos import pid_from_access_app

    return pid_from_access_app(app)


def shared_open_warnings(
    database_path: str,
    app: Any,
    *,
    holders: Callable[[str], set[int] | None] = _default_holders,
    access_pids: Callable[[], set[int]] = _default_access_pids,
    driven_pid: Callable[[Any], int | None] = _default_driven_pid,
    local_names: Callable[[], set[str]] = _local_names,
) -> list[str]:
    """One sentence per other holder: who, the consequence, and the action. Never raises."""
    warnings: list[str] = []
    try:
        found = holders(database_path) or set()
        others = found - {driven_pid(app), os.getpid()}
        access = access_pids() if others else set()
        for pid in sorted(others):
            who = (
                f"Another Access session (PID {pid})" if pid in access
                else f"Another process (PID {pid})"
            )
            warnings.append(f"{who} has this database open; {CONSEQUENCE}. {ACTION}")
    except Exception:
        pass
    try:
        lock_path = lock_file_for(database_path)
        if lock_path is not None:
            local = local_names()
            for machine in lock_file_machines(lock_path):
                if machine.upper() not in local:
                    warnings.append(
                        f"The lock file lists computer {machine} as using this database "
                        f"(the entry may be stale); if that session is open, {CONSEQUENCE}. "
                        f"{ACTION}"
                    )
    except Exception:
        pass
    return warnings


def with_warnings_first(result: Any, warnings: list[str]) -> Any:
    """``result`` with these warnings ahead of its own and of the per-object detail."""
    if not warnings or not isinstance(result, dict):
        return result
    existing = result.get("warnings")
    combined = [*warnings, *(existing if isinstance(existing, list) else [])]
    anchor = "error_pattern" if "error_pattern" in result else (
        "error" if "error" in result else "success"
    )
    return insert_after(result, anchor, "warnings", combined)
