"""What Access is waiting on, attached to a timeout before the agent asks.

When a call times out or Access stops answering, the server can usually see
why: a dialog, or VBA stopped in break mode. Reading that takes window calls
only (the screenshot tool's metadata path), so it works while Access is
blocked, and it saves the agent a round trip.

The probe runs only when the target database is known. Without one, the
screenshot resolver's single-instance shortcut could describe an unrelated
Access window as if it were the one that timed out.

Dialog text goes in ``access_state``, never in ``error``: the usage log
records ``error``.
"""

from __future__ import annotations

import asyncio
import inspect
from pathlib import Path
from typing import Any, Callable

from .guidance import insert_after

TRIGGER_PATTERNS = frozenset({"timeout", "access_unresponsive"})
PROBE_TIMEOUT_SEC = 5.0
TEXT_CAP = 500
SKIPPED_TOOLS = frozenset({"vcs_capture_screenshot"})


def target_database(tool: str, func: Callable[..., Any], args: tuple, kwargs: dict) -> str | None:
    """The database a call acted on, or None when the call names none."""
    try:
        bound = inspect.signature(func).bind_partial(*args, **kwargs).arguments
    except TypeError:
        return None
    database = bound.get("database_path")
    if isinstance(database, str) and database:
        return database
    if tool == "vcs_rebuild_addin":
        source = bound.get("source_dir")
        if isinstance(source, str) and source:
            folder = Path(source.rstrip("\\/"))
            if folder.suffix.lower() == ".src":
                return str(folder.with_suffix(""))
    return None


def _cap(value: Any) -> Any:
    if isinstance(value, str) and len(value) > TEXT_CAP:
        return value[: TEXT_CAP - 1] + "…"
    return value


def read_access_state(database_path: str) -> dict[str, Any] | None:
    """Window-level state of the Access process holding ``database_path``."""
    from . import screenshot

    try:
        pid, _ = screenshot.resolve_target_pid(database_path, None)
    except Exception:
        return None
    try:
        meta = screenshot.capture_process(pid, include_image=False).meta
    except screenshot.ScreenshotError as exc:
        meta = dict(exc.extra)
    except Exception:
        return None
    dialogs = [
        {
            "dialog_id": d.get("dialog_id"),
            "title": _cap(d.get("title", "")),
            "text": _cap(d.get("text", "")),
            "buttons": [_cap(b) for b in d.get("buttons", [])],
        }
        for d in meta.get("dialogs", [])
    ]
    return {
        "pid": pid,
        "owned_by_server": screenshot.server_owns(pid),
        "vbe_state": meta.get("vbe_state"),
        "modal_blocked": bool(meta.get("modal_blocked")),
        "dialogs": dialogs,
    }


def state_guidance(state: dict[str, Any]) -> str | None:
    owned = state.get("owned_by_server")
    dialogs = state.get("dialogs") or []
    if dialogs:
        first = dialogs[0]
        if owned:
            buttons = ", ".join(f'"{b}"' for b in first.get("buttons") or []) or "none read"
            return (
                "Access is waiting on the dialog in access_state. Read its text, then "
                f"answer it with vcs_dismiss_dialog(dialog_id=\"{first.get('dialog_id')}\", "
                f"button=...) using one of its buttons ({buttons}), and retry."
            )
        return (
            "Access is waiting on the dialog in access_state, in a window this server did "
            "not create. Ask the user to answer it, then retry."
        )
    if state.get("vbe_state") == "break":
        if owned:
            return (
                "VBA is stopped in break mode. Call vcs_reset_vba() to see where it "
                "stopped and end the run, then retry."
            )
        return (
            "VBA is stopped in break mode in a window this server did not create. Ask the "
            "user to reset the code (Run > Reset in the VBA editor), then retry."
        )
    return None


async def attach_access_state(
    tool: str,
    func: Callable[..., Any],
    args: tuple,
    kwargs: dict,
    result: Any,
    reader: Callable[[str], dict[str, Any] | None] = read_access_state,
) -> Any:
    if tool in SKIPPED_TOOLS or not isinstance(result, dict):
        return result
    if result.get("success") is not False or result.get("error_pattern") not in TRIGGER_PATTERNS:
        return result
    database = target_database(tool, func, args, kwargs)
    if database is None:
        return result
    try:
        state = await asyncio.wait_for(asyncio.to_thread(reader, database), PROBE_TIMEOUT_SEC)
    except Exception:
        return result
    if not state:
        return result
    guidance = state_guidance(state)
    if guidance:
        result = insert_after(result, "error_pattern", "agent_guidance", guidance)
    return {**result, "access_state": state}
