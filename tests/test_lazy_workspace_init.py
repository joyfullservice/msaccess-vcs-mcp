"""Tests for lazy MCP workspace-roots .env discovery in tools.py."""

import asyncio
import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

import msaccess_vcs_mcp.config as config_module
import msaccess_vcs_mcp.tools as tools_module
from msaccess_vcs_mcp.tools import _ensure_env_loaded, _file_uri_to_path


@pytest.fixture(autouse=True)
def _reset_state(monkeypatch):
    """Clear the lazy-init flag and cached config state between tests."""
    tools_module._lazy_init_attempted = False
    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    for key in list(os.environ):
        if key.startswith("ACCESS_VCS_"):
            monkeypatch.delenv(key, raising=False)

    yield

    tools_module._lazy_init_attempted = False
    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}


def _make_ctx(root_uris: list[str]) -> MagicMock:
    """Build a fake Context exposing ``session.list_roots()`` -> roots."""
    roots = [SimpleNamespace(uri=uri) for uri in root_uris]
    list_roots_result = SimpleNamespace(roots=roots)

    session = MagicMock()
    session.list_roots = AsyncMock(return_value=list_roots_result)

    ctx = MagicMock()
    ctx.session = session
    return ctx


class TestFileUriToPath:
    def test_converts_unix_uri(self):
        assert _file_uri_to_path("file:///home/me/proj") == Path("/home/me/proj")

    def test_converts_windows_uri(self):
        # file:///C:/foo -> /C:/foo, leading slash stripped on Windows-style paths.
        result = _file_uri_to_path("file:///C:/Users/me/proj")
        assert str(result) in (r"C:\Users\me\proj", "C:/Users/me/proj")

    def test_decodes_percent_encoding(self):
        result = _file_uri_to_path("file:///tmp/with%20space")
        assert "with space" in str(result)

    def test_returns_none_for_non_file_uri(self):
        assert _file_uri_to_path("http://example.com") is None


class TestEnsureEnvLoaded:
    def test_loads_env_from_first_root_with_env(self, tmp_path, monkeypatch):
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        # Two workspace roots; only the second has an .env file.
        empty_root = tmp_path / "empty_workspace"
        empty_root.mkdir()
        real_root = tmp_path / "real_workspace"
        real_root.mkdir()
        (real_root / ".env").write_text(
            "ACCESS_VCS_ENABLE_LOGGING=true\n"
            "ACCESS_VCS_DATABASE=ws.accdb\n",
            encoding="utf-8",
        )

        ctx = _make_ctx([
            empty_root.as_uri(),
            real_root.as_uri(),
        ])

        asyncio.run(_ensure_env_loaded(ctx))

        assert os.environ.get("ACCESS_VCS_ENABLE_LOGGING") == "true"
        assert os.environ.get("ACCESS_VCS_DATABASE") == "ws.accdb"
        assert tools_module._lazy_init_attempted is True

    def test_skips_when_ctx_is_none(self):
        asyncio.run(_ensure_env_loaded(None))
        # No exception, flag NOT flipped (so a real ctx call later still works).
        assert tools_module._lazy_init_attempted is False

    def test_only_runs_once(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)

        ctx = _make_ctx([])
        asyncio.run(_ensure_env_loaded(ctx))
        assert tools_module._lazy_init_attempted is True

        # Second call must not invoke list_roots again.
        ctx.session.list_roots.reset_mock()
        asyncio.run(_ensure_env_loaded(ctx))
        ctx.session.list_roots.assert_not_called()

    def test_skips_when_startup_env_already_found(
        self, tmp_path, monkeypatch,
    ):
        """If .env was already discovered at startup (CWD walk), don't override."""
        startup_project = tmp_path / "startup"
        startup_project.mkdir()
        (startup_project / ".env").write_text(
            "ACCESS_VCS_DATABASE=startup.accdb\n", encoding="utf-8",
        )
        monkeypatch.chdir(startup_project)

        # Prime config so _project_root points at the startup dir.
        from msaccess_vcs_mcp.config import load_config
        load_config()
        assert os.environ.get("ACCESS_VCS_DATABASE") == "startup.accdb"

        # Now simulate a workspace root with a different .env.
        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text(
            "ACCESS_VCS_DATABASE=workspace.accdb\n", encoding="utf-8",
        )

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        # The startup-discovered value should still be in place.
        assert os.environ["ACCESS_VCS_DATABASE"] == "startup.accdb"
        ctx.session.list_roots.assert_not_called()

    def test_handles_list_roots_failure_gracefully(
        self, tmp_path, monkeypatch, capsys,
    ):
        monkeypatch.chdir(tmp_path)

        ctx = MagicMock()
        ctx.session = MagicMock()
        ctx.session.list_roots = AsyncMock(side_effect=RuntimeError("not supported"))

        asyncio.run(_ensure_env_loaded(ctx))

        captured = capsys.readouterr()
        assert "Could not request workspace roots" in captured.err
        assert tools_module._lazy_init_attempted is True
