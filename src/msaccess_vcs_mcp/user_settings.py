"""Settings the user decides once, kept per user rather than per project.

Some settings deserve an explicit choice instead of a silent default. Their
answers live in ``~/.msaccess-vcs-mcp/settings.json`` because one server
process serves every Cursor window: a per-project file would ask the same
question in every repository.

A value in the environment (the ``mcp.json`` env section or any ``.env``)
always wins over the file, and nothing here writes into ``os.environ``, so
the ``.env`` precedence and hot reload in :mod:`config` are untouched. The
file is re-read whenever its modification time changes, so a change made by
the CLI or another server process applies on the next read.

Each entry records who decided it. ``{"decision": "default"}`` means the user
chose the default rather than a number, so a later change to the default
still reaches them. A key with no entry is pending; that is also how a
setting added in a later release is detected, with no version bookkeeping.
"""

from __future__ import annotations

import json
import os
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from .config import _strip_quotes

SETTINGS_PATH_ENV = "ACCESS_VCS_SETTINGS_PATH"
DEFAULT_CHOICE = "default"
NEVER = "never"
FOREVER = "forever"
DECIDED_BY = ("form", "agent", "cli")

SOURCE_ENV = "env"
SOURCE_FILE = "settings_file"
SOURCE_DEFAULT = "default"

LOCK_TIMEOUT_SEC = 5.0
STALE_LOCK_SEC = 30.0


RETENTION_WORDS = 'use "never", "forever", or a whole number of days (1 or more)'


def _parse_retention(text: str) -> str:
    """``never``, ``forever``, or whole days (1 or more).

    Numbers never stand for a policy: 0 or -1 would have to be learned, and an
    agent guessing their meaning would get it backwards.
    """
    word = text.strip().lower()
    if word in (NEVER, FOREVER):
        return word
    try:
        value = float(word)
    except ValueError:
        raise ValueError(RETENTION_WORDS) from None
    if not (value >= 1 and value.is_integer() and value != float("inf")):
        raise ValueError(RETENTION_WORDS)
    return str(int(value))


def _display_retention(value: str) -> str:
    return value if value in (NEVER, FOREVER) else f"{value} days"


@dataclass(frozen=True)
class DecisionSetting:
    key: str
    name: str
    default: str
    title: str
    prompt: str
    description: str
    relevant_to: str
    parse: Callable[[str], str]
    display: Callable[[str], str]


SCREENSHOT_RETENTION = DecisionSetting(
    key="ACCESS_VCS_SCREENSHOT_RETENTION",
    name="screenshot_retention",
    default="30",
    title="Screenshot retention",
    prompt='"never", "forever", or a number of days',
    description=(
        "How long a screenshot saved for the chat (save_image=True) is kept. "
        '"never" saves none, so the image appears only in the reply, and deletes any '
        'already saved. "forever" never deletes them. With a number of days, a chat '
        "older than that shows the image as broken."
    ),
    relevant_to="vcs_capture_screenshot",
    parse=_parse_retention,
    display=_display_retention,
)

# Audit logging (ACCESS_VCS_ENABLE_LOGGING) is deliberately absent: an agent
# answering for the user could switch off its own audit trail.
DECISION_SETTINGS: tuple[DecisionSetting, ...] = (SCREENSHOT_RETENTION,)

_cache_lock = threading.Lock()
_cache: tuple[str, int, int, dict[str, Any]] | None = None


def settings_path() -> Path:
    raw = _strip_quotes(os.environ.get(SETTINGS_PATH_ENV, "").strip()).strip()
    if raw:
        return Path(raw).expanduser()
    return Path.home() / ".msaccess-vcs-mcp" / "settings.json"


def find_setting(name_or_key: str) -> DecisionSetting:
    wanted = name_or_key.strip().lower()
    for setting in DECISION_SETTINGS:
        if wanted in (setting.name, setting.key.lower()):
            return setting
    names = ", ".join(s.name for s in DECISION_SETTINGS)
    raise KeyError(f"Unknown setting '{name_or_key}'. Known settings: {names}.")


def _env_value(setting: DecisionSetting) -> str | None:
    raw = _strip_quotes(os.environ.get(setting.key, "").strip()).strip()
    return raw or None


def _read_disk(path: Path) -> dict[str, Any]:
    """The file's contents; an unreadable file is moved aside and treated as empty."""
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return {}
    except OSError:
        return {}
    try:
        data = json.loads(text)
        if not isinstance(data, dict):
            raise ValueError("not an object")
    except ValueError as exc:
        _set_aside(path, exc)
        return {}
    return data


def _set_aside(path: Path, exc: Exception) -> None:
    from .usage_logging import log_diagnostic_event

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    aside = path.with_name(f"{path.name}.corrupt-{stamp}")
    try:
        os.replace(path, aside)
    except OSError:
        aside = None
    log_diagnostic_event(
        "user_settings_unreadable",
        path=str(path),
        moved_to=str(aside) if aside else None,
        error_type=type(exc).__name__,
    )


def _read() -> dict[str, Any]:
    global _cache
    path = settings_path()
    try:
        stat = path.stat()
    except OSError:
        return {}
    stamp = (str(path), stat.st_mtime_ns, stat.st_size)
    with _cache_lock:
        if _cache is not None and _cache[:3] == stamp:
            return _cache[3]
    data = _read_disk(path)
    with _cache_lock:
        _cache = (*stamp, data)
    return data


def _entry(setting: DecisionSetting, data: dict[str, Any]) -> dict[str, Any] | None:
    entry = data.get(setting.key)
    if not isinstance(entry, dict):
        return None
    if entry.get("decision") == DEFAULT_CHOICE:
        return entry
    try:
        setting.parse(str(entry.get("value")))
    except (TypeError, ValueError):
        return None
    return entry


def get_setting(key: str) -> str:
    """The effective value: environment, then the settings file, then the default."""
    setting = find_setting(key)
    env = _env_value(setting)
    if env is not None:
        try:
            return setting.parse(env)
        except ValueError:
            return env
    entry = _entry(setting, _read())
    if entry is not None and "value" in entry:
        return setting.parse(str(entry["value"]))
    return setting.default


def get_setting_source(key: str) -> str:
    """Which of ``env``, ``settings_file``, or ``default`` supplied the value."""
    setting = find_setting(key)
    if _env_value(setting) is not None:
        return SOURCE_ENV
    entry = _entry(setting, _read())
    if entry is not None and "value" in entry:
        return SOURCE_FILE
    return SOURCE_DEFAULT


def is_decided(key: str) -> bool:
    setting = find_setting(key)
    return _env_value(setting) is not None or _entry(setting, _read()) is not None


def describe(setting: DecisionSetting) -> dict[str, Any]:
    entry = None if _env_value(setting) is not None else _entry(setting, _read())
    source = get_setting_source(setting.key)
    value = get_setting(setting.key)
    return {
        "name": setting.name,
        "env_var": setting.key,
        "value": value,
        "display": setting.display(value),
        "default": setting.default,
        "source": source,
        "decided": source == SOURCE_ENV or entry is not None,
        "decided_by": SOURCE_ENV if source == SOURCE_ENV else (entry or {}).get("decided_by"),
        "description": setting.description,
    }


def describe_all() -> list[dict[str, Any]]:
    return [describe(s) for s in DECISION_SETTINGS]


def pending_settings() -> list[str]:
    return [s.name for s in DECISION_SETTINGS if not is_decided(s.key)]


def normalize(setting: DecisionSetting, value: Any) -> str:
    """The value to store, or ``DEFAULT_CHOICE``. Raises ``ValueError`` when invalid."""
    text = str(value).strip()
    if text.lower() == DEFAULT_CHOICE:
        return DEFAULT_CHOICE
    try:
        return setting.parse(text)
    except ValueError as exc:
        raise ValueError(f"{setting.name}: '{text}' is not valid ({exc}).") from None


def save(decisions: dict[str, Any], decided_by: str) -> list[str]:
    """Validate and store ``{name_or_key: value}``; returns the names saved.

    Every value is validated before anything is written, so a bad value
    saves nothing.
    """
    if decided_by not in DECIDED_BY:
        raise ValueError(f"decided_by must be one of {', '.join(DECIDED_BY)}")
    entries: dict[str, dict[str, Any]] = {}
    names: list[str] = []
    stamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    for name, value in decisions.items():
        setting = find_setting(name)
        stored = normalize(setting, value)
        entry: dict[str, Any] = (
            {"decision": DEFAULT_CHOICE} if stored == DEFAULT_CHOICE else {"value": stored}
        )
        entry.update(decided_by=decided_by, decided_at=stamp)
        entries[setting.key] = entry
        names.append(setting.name)
    if entries:
        _update(lambda data: data.update(entries))
    return names


def reset(name_or_key: str) -> bool:
    """Forget a decision, making the setting pending again."""
    setting = find_setting(name_or_key)
    removed: list[bool] = []
    _update(lambda data: removed.append(data.pop(setting.key, None) is not None))
    return bool(removed and removed[0])


def _update(change: Callable[[dict[str, Any]], Any]) -> None:
    """Read, change, and replace the file under a lock shared by every process."""
    global _cache
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with _file_lock(path.with_name(path.name + ".lock")):
        data = _read_disk(path)
        change(data)
        temp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
        temp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        os.replace(temp, path)
    with _cache_lock:
        _cache = None


class _file_lock:
    """An exclusive-create lock file; one left by a crashed process expires."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def __enter__(self) -> "_file_lock":
        deadline = time.monotonic() + LOCK_TIMEOUT_SEC
        while True:
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                self._clear_if_stale()
                if time.monotonic() >= deadline:
                    raise TimeoutError(f"Settings file is locked: {self.path}") from None
                time.sleep(0.05)
                continue
            os.write(fd, str(os.getpid()).encode("ascii"))
            os.close(fd)
            return self

    def __exit__(self, *exc: Any) -> None:
        try:
            os.unlink(self.path)
        except OSError:
            pass

    def _clear_if_stale(self) -> None:
        try:
            if time.time() - self.path.stat().st_mtime > STALE_LOCK_SEC:
                os.unlink(self.path)
        except OSError:
            pass
