"""Saved screenshots: where they go, what they are called, and when they expire.

A saved file exists so a chat can show it, and the chat refers to it only by
path. Files therefore live for a fixed period rather than a fixed count: a
busy day must not break the images in today's chat.

Only names this module wrote (``vcs-shot-*.png``) are ever deleted, and only
once they are past the retention period. Several server processes can prune
one folder at the same time without removing a file a chat was just given,
and a folder shared with other files loses nothing else.
"""

from __future__ import annotations

import os
import re
import secrets
import threading
import time
from datetime import datetime
from pathlib import Path, PurePath
from urllib.parse import quote

from .config import _strip_quotes

DEFAULT_RETENTION_DAYS = 30.0
SECONDS_PER_DAY = 86400
NAME_PREFIX = "vcs-shot-"
NAME_PATTERN = re.compile(r"^vcs-shot-\d{8}-\d{6}-[a-z]+-[0-9a-f]{8}\.png$")
ALT_TEXT = {
    "application": "Access window",
    "vbe": "VBA editor",
    "active": "Access object",
    "object": "Access object",
}

_prune_lock = threading.Lock()
_prune_started = False


def screenshot_dir() -> Path:
    raw = _strip_quotes(os.environ.get("ACCESS_VCS_SCREENSHOT_DIR", "").strip()).strip()
    if raw:
        return Path(raw).expanduser().resolve()
    return Path.home() / ".msaccess-vcs-mcp" / "screenshots"


def retention_days() -> float:
    """Days a saved file is kept; 0 keeps files forever."""
    try:
        value = float(
            os.environ.get("ACCESS_VCS_SCREENSHOT_RETENTION_DAYS", DEFAULT_RETENTION_DAYS)
        )
    except ValueError:
        return DEFAULT_RETENTION_DAYS
    return value if value >= 0 else DEFAULT_RETENTION_DAYS


def save_png(png: bytes, target: str) -> Path:
    """Write ``png`` under a new name. Never overwrites, never leaves a partial file."""
    folder = screenshot_dir()
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    while True:
        path = folder / f"{NAME_PREFIX}{stamp}-{target}-{secrets.token_hex(4)}.png"
        try:
            handle = open(path, "xb")
        except FileExistsError:
            continue
        break
    try:
        with handle:
            handle.write(png)
    except BaseException:
        path.unlink(missing_ok=True)
        raise
    return path


def prune_expired(folder: Path, days: float) -> int:
    """Delete this module's files older than ``days``. Returns how many went."""
    if days <= 0:
        return 0
    cutoff = time.time() - days * SECONDS_PER_DAY
    try:
        entries = list(os.scandir(folder))
    except OSError:
        return 0
    removed = 0
    for entry in entries:
        if not NAME_PATTERN.match(entry.name):
            continue
        try:
            if entry.is_file() and entry.stat().st_mtime < cutoff:
                os.unlink(entry.path)
                removed += 1
        except OSError:
            continue
    return removed


def save_and_prune(png: bytes, target: str) -> Path:
    path = save_png(png, target)
    try:
        prune_expired(path.parent, retention_days())
    except Exception:
        pass
    return path


def start_prune_once() -> threading.Thread | None:
    """Prune once per process, in the background, with the settings in force now.

    Called after the first ``load_config()``, so a retention period set in a
    workspace ``.env`` is honoured. The folder and period are read here, not
    in the thread, so the prune cannot act on settings that change later.
    """
    global _prune_started
    with _prune_lock:
        if _prune_started:
            return None
        _prune_started = True
    folder, days = screenshot_dir(), retention_days()
    if days <= 0 or not folder.is_dir():
        return None
    thread = threading.Thread(
        target=_prune_quietly, args=(folder, days), name="vcs-screenshot-prune", daemon=True
    )
    thread.start()
    return thread


def _prune_quietly(folder: Path, days: float) -> None:
    try:
        prune_expired(folder, days)
    except Exception:
        pass


def markdown_link(path: PurePath, target: str) -> str:
    """An image link Cursor's chat renders: ``/C:/...`` with every special character encoded."""
    url = quote(path.as_posix(), safe="/:")
    if not url.startswith("/"):
        url = "/" + url
    return f"![{ALT_TEXT.get(target, 'Access screenshot')}]({url})"
