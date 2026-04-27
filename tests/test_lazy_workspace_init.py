"""Tests for lazy MCP workspace-roots .env discovery in tools.py."""

import asyncio
import json
import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

import msaccess_vcs_mcp.config as config_module
import msaccess_vcs_mcp.tools as tools_module
import msaccess_vcs_mcp.usage_logging as logging_module
from msaccess_vcs_mcp.tools import _ensure_env_loaded, _file_uri_to_path


@pytest.fixture(autouse=True)
def _reset_state(monkeypatch, tmp_path_factory):
    """Clear the lazy-init flag and cached config state between tests.

    Also redirects the diagnostic stream to a per-test temp directory
    so the always-on diagnostic log doesn't write to the user's real
    ``~/.msaccess-vcs-mcp/logs/`` during the suite. Tests can read
    ``logging_module._diagnostic_file`` to inspect the events that
    fired.
    """
    tools_module._lazy_init_attempted = False
    tools_module._lazy_init_skip_logged = False
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

    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))

    yield

    tools_module._lazy_init_attempted = False
    tools_module._lazy_init_skip_logged = False
    config_module._env_loaded = False
    config_module._is_reload = False
    config_module._project_root = None
    config_module._project_root_method = None
    config_module._env_mtimes = {}

    if logging_module._diagnostic_handler is not None:
        try:
            logging_module._diagnostic_handler.close()
        except Exception:
            pass
    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None


def _read_diagnostic_events() -> list[dict]:
    """Return all events recorded in the active diagnostic log file.

    Reads from the configured ``ACCESS_VCS_DIAGNOSTIC_LOG_DIR`` rather
    than ``logging_module._diagnostic_file`` so the helper still works
    after ``reset_logging()`` has cleared the in-memory handle (which
    happens whenever ``initialize_from_workspace`` is called and
    triggers a logging reload mid-test).
    """
    diag_dir_str = os.environ.get("ACCESS_VCS_DIAGNOSTIC_LOG_DIR")
    if not diag_dir_str:
        return []
    path = Path(diag_dir_str) / "vcs-mcp-diagnostic.jsonl"
    if not path.exists():
        return []
    return [
        json.loads(line)
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


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
        self, tmp_path, monkeypatch,
    ):
        """list_roots() failure is captured as a diagnostic event with the
        underlying error type, instead of a stderr-only print.
        """
        monkeypatch.chdir(tmp_path)

        ctx = MagicMock()
        ctx.session = MagicMock()
        ctx.session.list_roots = AsyncMock(side_effect=RuntimeError("not supported"))

        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        event_names = [e["event"] for e in events]
        assert "lazy_init_started" in event_names
        assert "list_roots_failed" in event_names
        failure = next(e for e in events if e["event"] == "list_roots_failed")
        assert failure["error_type"] == "RuntimeError"
        assert "not supported" in failure["error"]
        assert tools_module._lazy_init_attempted is True


class TestDiagnosticEvents:
    """Verify every branch of _ensure_env_loaded emits the right event."""

    def test_emits_no_ctx_skip(self):
        """ctx=None records lazy_init_skipped(reason='no_ctx')."""
        asyncio.run(_ensure_env_loaded(None))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "no_ctx" for e in skipped)

    def test_emits_already_attempted_skip(self, tmp_path, monkeypatch):
        """A repeated call records lazy_init_skipped(reason='already_attempted')."""
        monkeypatch.chdir(tmp_path)
        ctx = _make_ctx([])

        asyncio.run(_ensure_env_loaded(ctx))
        # Re-run; the flag is already set so we should hit the early
        # 'already_attempted' branch.
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "already_attempted" for e in skipped)

    def test_emits_lazy_init_started_and_list_roots_response(
        self, tmp_path, monkeypatch,
    ):
        """A successful path emits started + list_roots_response + loaded."""
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text("ACCESS_VCS_DATABASE=ws.accdb\n", encoding="utf-8")

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        names = [e["event"] for e in events]
        assert "lazy_init_started" in names
        assert "list_roots_response" in names
        assert "lazy_init_loaded" in names

        started = next(e for e in events if e["event"] == "lazy_init_started")
        assert started["startup_root_has_env"] is False

        response = next(e for e in events if e["event"] == "list_roots_response")
        assert workspace.as_uri() in response["roots"]

        loaded = next(e for e in events if e["event"] == "lazy_init_loaded")
        assert loaded["workspace"] == str(workspace)

    def test_emits_no_env_in_roots(self, tmp_path, monkeypatch):
        """When list_roots returns workspaces without .env, we record
        lazy_init_no_env_in_roots with the URIs we tried.
        """
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir()
        monkeypatch.chdir(elsewhere)

        empty_a = tmp_path / "empty_a"
        empty_b = tmp_path / "empty_b"
        empty_a.mkdir()
        empty_b.mkdir()

        ctx = _make_ctx([empty_a.as_uri(), empty_b.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        names = [e["event"] for e in events]
        assert "lazy_init_no_env_in_roots" in names
        assert "lazy_init_loaded" not in names

        no_env = next(e for e in events if e["event"] == "lazy_init_no_env_in_roots")
        assert empty_a.as_uri() in no_env["roots"]
        assert empty_b.as_uri() in no_env["roots"]

    def test_emits_startup_env_present_skip(self, tmp_path, monkeypatch):
        """If startup discovery already found an .env, we skip with
        reason='startup_env_present' (and never call list_roots).
        """
        startup_project = tmp_path / "startup"
        startup_project.mkdir()
        (startup_project / ".env").write_text(
            "ACCESS_VCS_DATABASE=startup.accdb\n", encoding="utf-8",
        )
        monkeypatch.chdir(startup_project)

        from msaccess_vcs_mcp.config import load_config
        load_config()

        workspace = tmp_path / "workspace"
        workspace.mkdir()
        (workspace / ".env").write_text("X=1\n", encoding="utf-8")

        ctx = _make_ctx([workspace.as_uri()])
        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "startup_env_present" for e in skipped)
        ctx.session.list_roots.assert_not_called()

    def test_emits_no_session_skip_for_request_unavailable_ctx(self, tmp_path, monkeypatch):
        """Context.request_context raises ValueError outside a request.

        Simulates ``mcp.get_context()`` being called outside a real MCP
        request (e.g. unit tests calling the wrapper directly). The
        helper records ``reason='no_session'`` and the lazy-init flag is
        NOT flipped, so a later real request can still drive discovery.
        """
        monkeypatch.chdir(tmp_path)

        ctx = MagicMock()
        type(ctx).session = property(
            lambda self: (_ for _ in ()).throw(
                ValueError("Context is not available outside of a request")
            )
        )

        asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [e for e in events if e["event"] == "lazy_init_skipped"]
        assert any(e.get("reason") == "no_session" for e in skipped)
        assert tools_module._lazy_init_attempted is False


class TestVcsToolWrapper:
    """Verify ``vcs_tool`` wires lazy init for *every* registered tool.

    The contract: regardless of whether the wrapped function is sync or
    async, and regardless of whether it declares a ``ctx: Context``
    parameter, the wrapper must:
      1. Pull the active request context via ``mcp.get_context()``.
      2. Invoke ``_ensure_env_loaded`` with it.
      3. Refresh ``load_config()``.
      4. Dispatch the tool body (await for async, call for sync).

    These tests stub ``mcp.get_context`` and ``_ensure_env_loaded`` to
    verify the orchestration without actually calling MCP/Access COM.
    """

    def _get_registered_callable(self, fn_name: str):
        """Pull a FastMCP tool by registered name.

        ``mcp.tool()`` defaults the tool name to the wrapped callable's
        ``__name__`` (which ``functools.wraps`` preserves), so we look up
        by the wrapped function's name.
        """
        manager = tools_module.mcp._tool_manager
        for tool in manager.list_tools():
            if tool.name == fn_name:
                return tool.fn
        raise AssertionError(f"tool {fn_name!r} not registered")

    def test_sync_tool_triggers_lazy_init(self, monkeypatch):
        """A sync tool registered via @vcs_tool still drives _ensure_env_loaded."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        sentinel_ctx = MagicMock(name="ctx")
        monkeypatch.setattr(
            tools_module.mcp, "get_context", lambda: sentinel_ctx
        )
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        body_calls = []

        @tools_module.vcs_tool("test_sync_tool")
        def test_sync_tool_fn(arg: str = "default") -> dict:
            body_calls.append(arg)
            return {"ok": True, "arg": arg}

        wrapper = self._get_registered_callable("test_sync_tool_fn")
        result = asyncio.run(wrapper(arg="hello"))

        assert result == {"ok": True, "arg": "hello"}
        assert ensure_calls == [sentinel_ctx]
        assert body_calls == ["hello"]

    def test_async_tool_triggers_lazy_init(self, monkeypatch):
        """Async tools (with or without ctx param) drive _ensure_env_loaded."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        sentinel_ctx = MagicMock(name="ctx")
        monkeypatch.setattr(
            tools_module.mcp, "get_context", lambda: sentinel_ctx
        )
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        @tools_module.vcs_tool("test_async_tool")
        async def test_async_tool_fn(arg: str = "default") -> dict:
            return {"ok": True, "arg": arg}

        wrapper = self._get_registered_callable("test_async_tool_fn")
        result = asyncio.run(wrapper(arg="world"))

        assert result == {"ok": True, "arg": "world"}
        assert ensure_calls == [sentinel_ctx]

    def test_wrapper_handles_get_context_failure(self, monkeypatch):
        """If mcp.get_context() raises, wrapper passes None and proceeds."""
        ensure_calls = []

        async def fake_ensure(ctx):
            ensure_calls.append(ctx)

        monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)

        def boom():
            raise RuntimeError("no request context")

        monkeypatch.setattr(tools_module.mcp, "get_context", boom)
        monkeypatch.setattr(tools_module, "load_config", lambda: {})

        @tools_module.vcs_tool("test_no_ctx_tool")
        def test_no_ctx_tool_fn() -> dict:
            return {"ok": True}

        wrapper = self._get_registered_callable("test_no_ctx_tool_fn")
        result = asyncio.run(wrapper())

        assert result == {"ok": True}
        assert ensure_calls == [None]

    def test_repeated_skip_logged_only_once(self, tmp_path, monkeypatch):
        """After lazy init runs, repeated 'already_attempted' skips don't
        spam the diagnostic log on every tool call. The first repeat
        emits an event (so operators can verify the cache is working);
        subsequent repeats are silent.
        """
        monkeypatch.chdir(tmp_path)

        ctx = _make_ctx([])
        asyncio.run(_ensure_env_loaded(ctx))

        # Repeat enough times to dwarf any noisy implementation.
        for _ in range(5):
            asyncio.run(_ensure_env_loaded(ctx))

        events = _read_diagnostic_events()
        skipped = [
            e for e in events
            if e["event"] == "lazy_init_skipped"
            and e.get("reason") == "already_attempted"
        ]
        assert len(skipped) == 1, (
            f"expected exactly one 'already_attempted' diagnostic; got {len(skipped)}"
        )
