"""The next step for an agent, attached to failed tool results.

Some clients never show the model the server instructions, and Claude Code
truncates them, so the action that recovers from a failure has to arrive in
the result itself. Every result gets one field, ``agent_guidance``: a full
sentence naming the exact next call. A result that already carries guidance
from the code that built it keeps its own.

Guidance never goes into ``error``: the usage log records that field, and
guidance can quote dialog text.
"""

from __future__ import annotations

from typing import Any

STUCK = (
    "Call vcs_capture_screenshot(database_path, include_image=False) to see whether "
    "Access is showing a dialog or is in VBA break mode before retrying. For more, "
    "call vcs_tips(\"stuck\")."
)

BY_ERROR_PATTERN = {
    "timeout": (
        "The server stopped waiting, but Access was left running and may still be "
        "working or waiting on a dialog. " + STUCK
    ),
    "access_unresponsive": (
        "Access is not answering automation, usually because a dialog is open or VBA "
        "is in break mode. " + STUCK
    ),
    "empty_result": (
        "Nothing was built. The add-in returned no result, which usually means it "
        "raised an error that Access showed as a dialog. " + STUCK
    ),
    "server_busy": (
        "Another call holds Access (see busy_with). Retry this call in a few seconds, "
        "and issue Access calls one at a time rather than in parallel."
    ),
    "installed_addin_refused": (
        "Pass the development copy instead: the Version Control.accda beside "
        "Version Control.accda.src in the add-in's repository. For rebuilding and "
        "testing the add-in, call vcs_tips(\"addin-development\")."
    ),
}

REBUILD_REFUSED = (
    "Nothing was rebuilt. Ask the user to close the Access windows listed in "
    "otherInstances (never close them yourself), then call vcs_rebuild_addin again."
)


def guidance_for(tool: str, result: dict[str, Any]) -> str | None:
    if tool == "vcs_rebuild_addin" and result.get("status") == "refused":
        return REBUILD_REFUSED
    pattern = result.get("error_pattern")
    return BY_ERROR_PATTERN.get(pattern) if isinstance(pattern, str) else None


def insert_after(result: dict[str, Any], anchor: str, key: str, value: Any) -> dict[str, Any]:
    """A copy of ``result`` with ``key`` placed right after ``anchor`` (or last)."""
    if anchor not in result:
        return {**result, key: value}
    ordered: dict[str, Any] = {}
    for existing, existing_value in result.items():
        if existing == key:
            continue
        ordered[existing] = existing_value
        if existing == anchor:
            ordered[key] = value
    return ordered


def with_guidance(tool: str, result: Any) -> Any:
    if not isinstance(result, dict) or result.get("success") is not False:
        return result
    if result.get("agent_guidance"):
        return result
    text = guidance_for(tool, result)
    if text is None:
        return result
    return insert_after(result, "error_pattern", "agent_guidance", text)
