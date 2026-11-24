"""Situational guidance served by ``vcs_tips``, one markdown file per topic.

The server instructions hold only the rules every call needs; depth that
matters only in one situation lives here and is read on demand.
"""

from __future__ import annotations

from importlib import resources
from typing import Any


def _topic_files() -> dict[str, str]:
    folder = resources.files("msaccess_vcs_mcp") / "tips"
    return {
        entry.name[:-3]: entry.read_text(encoding="utf-8")
        for entry in folder.iterdir()
        if entry.name.endswith(".md")
    }


def topic_names() -> list[str]:
    return sorted(_topic_files())


def _summary(text: str) -> str:
    for line in text.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return ""


def get_tips(topic: str | None) -> dict[str, Any]:
    topics = _topic_files()
    listing = [{"topic": name, "about": _summary(topics[name])} for name in sorted(topics)]

    key = (topic or "").strip().lower()
    if not key:
        return {"success": True, "topics": listing}
    if key in topics:
        return {"success": True, "topic": key, "content": topics[key]}

    matches = [name for name, text in sorted(topics.items()) if key in name or key in text.lower()]
    if len(matches) == 1:
        return {"success": True, "topic": matches[0], "content": topics[matches[0]]}
    return {
        "success": False,
        "error": f"No tips topic named '{topic}'.",
        "matching_topics": matches,
        "topics": listing,
    }
