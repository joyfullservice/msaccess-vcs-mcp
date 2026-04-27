"""Tests for .env file discovery, loading, and hot-reload behaviour."""

import json
import os
import time
from pathlib import Path
from unittest.mock import patch

import pytest

import msaccess_vcs_mcp.config as config_module
from msaccess_vcs_mcp.config import (
    RESOLUTION_CWD_ENV,
    RESOLUTION_PROJECT_DIR_ENV,
    RESOLUTION_WORKSPACE_ROOTS,
    _check_env_reload,
    _find_project_root,
    _load_env_files,
    _load_env_from_directory,
    get_project_root_info,
    initialize_from_workspace,
    load_config,
)


@pytest.fixture(autouse=True)
def _reset_config_state(monkeypatch):
    """Clear cached env-load state and ACCESS_VCS_* env vars between tests.

    Also disables the always-on diagnostic stream so config-loading
    tests don't write to the user's real ``~/.msaccess-vcs-mcp/logs/``
    directory while exercising ``_load_env_files``.
    """
    import msaccess_vcs_mcp.usage_logging as logging_module

    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None

    for key in list(os.environ):
        if key.startswith("ACCESS_VCS_"):
            monkeypatch.delenv(key, raising=False)

    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

    yield

    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None


def _write_env(path: Path, contents: str) -> Path:
    """Write a .env file and return its path."""
    env_path = path / ".env"
    env_path.write_text(contents, encoding="utf-8")
    return env_path


class TestFindProjectRoot:
    """Tests for ``_find_project_root`` resolution order."""

    def test_explicit_project_dir_wins(self, tmp_path, monkeypatch):
        """ACCESS_VCS_PROJECT_DIR overrides CWD-based discovery."""
        project = tmp_path / "explicit_project"
        project.mkdir()
        _write_env(project, "ACCESS_VCS_DATABASE=ignored\n")

        # Make CWD point to a totally unrelated directory.
        other = tmp_path / "elsewhere"
        other.mkdir()
        monkeypatch.chdir(other)

        monkeypatch.setenv("ACCESS_VCS_PROJECT_DIR", str(project))

        assert _find_project_root() == project.resolve()

    def test_explicit_project_dir_missing_falls_through(
        self, tmp_path, monkeypatch, capsys,
    ):
        """A missing ACCESS_VCS_PROJECT_DIR is warned about and skipped."""
        project = tmp_path / "real_project"
        project.mkdir()
        _write_env(project, "X=1\n")
        monkeypatch.chdir(project)

        monkeypatch.setenv(
            "ACCESS_VCS_PROJECT_DIR", str(tmp_path / "does_not_exist"),
        )

        result = _find_project_root()

        assert result == project.resolve()
        captured = capsys.readouterr()
        assert "ACCESS_VCS_PROJECT_DIR" in captured.err
        assert "non-existent" in captured.err

    def test_walks_up_from_cwd(self, tmp_path, monkeypatch):
        """Discovery walks up from CWD until it finds a .env file."""
        project = tmp_path / "proj"
        project.mkdir()
        _write_env(project, "X=1\n")

        nested = project / "a" / "b" / "c"
        nested.mkdir(parents=True)
        monkeypatch.chdir(nested)

        assert _find_project_root() == project.resolve()


class TestLoadEnvFiles:
    """Tests for ``_load_env_files`` first-load and reload behaviour."""

    def test_first_load_sets_env_var(self, tmp_path, monkeypatch):
        project = tmp_path / "proj"
        project.mkdir()
        _write_env(project, "ACCESS_VCS_ENABLE_LOGGING=true\n")
        monkeypatch.chdir(project)

        _load_env_files()

        assert os.environ.get("ACCESS_VCS_ENABLE_LOGGING") == "true"
        assert config_module._env_loaded is True

    def test_first_load_respects_existing_env(self, tmp_path, monkeypatch):
        """On first load, MCP env-section values (already in os.environ) win."""
        project = tmp_path / "proj"
        project.mkdir()
        _write_env(project, "ACCESS_VCS_ENABLE_LOGGING=true\n")
        monkeypatch.chdir(project)

        monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "from_mcp_env")

        _load_env_files()

        assert os.environ["ACCESS_VCS_ENABLE_LOGGING"] == "from_mcp_env"

    def test_reload_overrides_existing(self, tmp_path, monkeypatch):
        """After ``_check_env_reload`` triggers, edits override env values."""
        project = tmp_path / "proj"
        project.mkdir()
        env_path = _write_env(project, "ACCESS_VCS_ENABLE_LOGGING=false\n")
        monkeypatch.chdir(project)

        _load_env_files()
        assert os.environ.get("ACCESS_VCS_ENABLE_LOGGING") == "false"

        # Edit the .env file and bump the mtime so reload detection fires.
        env_path.write_text("ACCESS_VCS_ENABLE_LOGGING=true\n", encoding="utf-8")
        new_mtime = time.time() + 5
        os.utime(env_path, (new_mtime, new_mtime))

        assert _check_env_reload() is True
        _load_env_files()

        assert os.environ["ACCESS_VCS_ENABLE_LOGGING"] == "true"

    def test_caches_after_first_load(self, tmp_path, monkeypatch):
        """Repeated calls without an mtime change don't re-walk discovery."""
        project = tmp_path / "proj"
        project.mkdir()
        _write_env(project, "X=1\n")
        monkeypatch.chdir(project)

        _load_env_files()
        first_root = config_module._project_root

        # Move CWD elsewhere -- second call should still use cached root.
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        _load_env_files()
        assert config_module._project_root == first_root


class TestLoadEnvFromDirectory:
    """Tests for the explicit-directory loader used by lazy workspace init."""

    def test_loads_env_from_arbitrary_directory(self, tmp_path):
        _write_env(tmp_path, "ACCESS_VCS_DATABASE=workspace.accdb\n")

        loaded = _load_env_from_directory(tmp_path)

        assert loaded is True
        assert os.environ.get("ACCESS_VCS_DATABASE") == "workspace.accdb"
        assert config_module._project_root == tmp_path.resolve()

    def test_returns_false_when_no_env(self, tmp_path):
        loaded = _load_env_from_directory(tmp_path)

        assert loaded is False


class TestInitializeFromWorkspace:
    """Tests for the public lazy-init entry point."""

    def test_loads_env_and_returns_config(self, tmp_path, monkeypatch):
        # CWD intentionally elsewhere -- the workspace path is what matters.
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        _write_env(
            workspace,
            "ACCESS_VCS_ENABLE_LOGGING=true\n"
            "ACCESS_VCS_DATABASE=ws.accdb\n",
        )

        config = initialize_from_workspace(workspace)

        assert os.environ.get("ACCESS_VCS_ENABLE_LOGGING") == "true"
        assert os.environ.get("ACCESS_VCS_DATABASE") == "ws.accdb"
        assert config["ACCESS_VCS_ENABLE_LOGGING"] is True
        assert config["ACCESS_VCS_DATABASE"] == "ws.accdb"
        assert config_module._project_root == workspace.resolve()

    def test_subsequent_load_config_does_not_overwrite_workspace_root(
        self, tmp_path, monkeypatch,
    ):
        """After lazy init, calling load_config() must not re-discover from CWD."""
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        _write_env(workspace, "ACCESS_VCS_DATABASE=ws.accdb\n")

        initialize_from_workspace(workspace)
        assert config_module._project_root == workspace.resolve()

        load_config()

        assert config_module._project_root == workspace.resolve()


class TestProjectRootResolutionMethod:
    """Tests for resolution-method tracking surfaced via get_project_root_info()."""

    def test_unset_before_discovery(self):
        info = get_project_root_info()
        assert info["project_root"] is None
        assert info["resolution_method"] is None

    def test_records_explicit_project_dir(self, tmp_path, monkeypatch):
        project = tmp_path / "explicit"
        project.mkdir()
        _write_env(project, "X=1\n")

        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)
        monkeypatch.setenv("ACCESS_VCS_PROJECT_DIR", str(project))

        _load_env_files()

        info = get_project_root_info()
        assert info["resolution_method"] == RESOLUTION_PROJECT_DIR_ENV
        assert info["project_root"] == str(project.resolve())

    def test_records_cwd_walk_with_env(self, tmp_path, monkeypatch):
        project = tmp_path / "proj"
        project.mkdir()
        _write_env(project, "X=1\n")
        monkeypatch.chdir(project)

        _load_env_files()

        info = get_project_root_info()
        assert info["resolution_method"] == RESOLUTION_CWD_ENV
        assert info["project_root"] == str(project.resolve())

    def test_records_workspace_roots(self, tmp_path, monkeypatch):
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        workspace = tmp_path / "ws"
        workspace.mkdir()
        _write_env(workspace, "X=1\n")

        initialize_from_workspace(workspace)

        info = get_project_root_info()
        assert info["resolution_method"] == RESOLUTION_WORKSPACE_ROOTS
        assert info["project_root"] == str(workspace.resolve())

    def test_explicit_dir_takes_precedence_over_cwd_with_env(
        self, tmp_path, monkeypatch,
    ):
        explicit = tmp_path / "explicit"
        explicit.mkdir()
        _write_env(explicit, "X=1\n")

        cwd_proj = tmp_path / "cwd_proj"
        cwd_proj.mkdir()
        _write_env(cwd_proj, "Y=2\n")
        monkeypatch.chdir(cwd_proj)

        monkeypatch.setenv("ACCESS_VCS_PROJECT_DIR", str(explicit))

        _load_env_files()

        info = get_project_root_info()
        assert info["resolution_method"] == RESOLUTION_PROJECT_DIR_ENV
        assert info["project_root"] == str(explicit.resolve())


class TestResolutionRecordedInUsageLog:
    """The logging_initialized event must capture project-root provenance."""

    def test_logging_initialized_event_includes_resolution(
        self, tmp_path, monkeypatch,
    ):
        import msaccess_vcs_mcp.usage_logging as logging_module
        from msaccess_vcs_mcp.usage_logging import (
            _initialize_logging,
            reset_logging,
        )

        # Reset logging globals so this test is hermetic.
        logging_module._logging_enabled = None
        logging_module._log_handler = None
        logging_module._log_file = None

        project = tmp_path / "proj"
        project.mkdir()
        log_dir = tmp_path / "logs"
        _write_env(
            project,
            f"ACCESS_VCS_ENABLE_LOGGING=true\n"
            f"ACCESS_VCS_LOG_DIR={log_dir}\n",
        )
        monkeypatch.chdir(project)

        load_config()

        try:
            assert _initialize_logging() is True

            entries = [
                json.loads(line)
                for line in (log_dir / "vcs-mcp-usage.jsonl").read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            init_events = [e for e in entries if e.get("event") == "logging_initialized"]
            assert init_events, "expected a logging_initialized event"

            event = init_events[0]
            assert event["project_root"] == str(project.resolve())
            assert event["project_root_resolution"] == RESOLUTION_CWD_ENV
        finally:
            reset_logging()
