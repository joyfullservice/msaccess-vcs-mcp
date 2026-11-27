"""Fixtures shared by every test module."""

import pytest

import msaccess_vcs_mcp.screenshot_store as store
import msaccess_vcs_mcp.user_settings as user_settings


@pytest.fixture(autouse=True)
def screenshot_dir(tmp_path, monkeypatch):
    """Keep saves and prunes out of the real screenshot folder.

    Every tool call starts the once-per-process prune, so any test that
    calls a tool would otherwise reach ``~/.msaccess-vcs-mcp/screenshots``.
    """
    folder = tmp_path / "screenshots"
    monkeypatch.setenv("ACCESS_VCS_SCREENSHOT_DIR", str(folder))
    monkeypatch.delenv("ACCESS_VCS_SCREENSHOT_RETENTION", raising=False)
    monkeypatch.setattr(store, "_prune_started", False)
    return folder


@pytest.fixture(autouse=True)
def settings_file(tmp_path, monkeypatch):
    """Keep reads and writes out of the real ``~/.msaccess-vcs-mcp/settings.json``."""
    path = tmp_path / "user-settings" / "settings.json"
    monkeypatch.setenv(user_settings.SETTINGS_PATH_ENV, str(path))
    monkeypatch.setattr(user_settings, "_cache", None)
    return path
