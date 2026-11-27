"""Fixtures shared by every test module."""

import pytest

import msaccess_vcs_mcp.screenshot_store as store


@pytest.fixture(autouse=True)
def screenshot_dir(tmp_path, monkeypatch):
    """Keep saves and prunes out of the real screenshot folder.

    Every tool call starts the once-per-process prune, so any test that
    calls a tool would otherwise reach ``~/.msaccess-vcs-mcp/screenshots``.
    """
    folder = tmp_path / "screenshots"
    monkeypatch.setenv("ACCESS_VCS_SCREENSHOT_DIR", str(folder))
    monkeypatch.delenv("ACCESS_VCS_SCREENSHOT_RETENTION_DAYS", raising=False)
    monkeypatch.setattr(store, "_prune_started", False)
    return folder
