"""Tests for usage logging lifecycle and hot-reload interactions."""

import json
import os
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import msaccess_vcs_mcp.usage_logging as logging_module
from msaccess_vcs_mcp.usage_logging import (
    _initialize_diagnostic_logging,
    _initialize_logging,
    get_diagnostic_log_path,
    is_diagnostic_logging_enabled,
    log_diagnostic_event,
    reset_logging,
    log_tool_call,
    log_code_execution,
    log_addin_probe,
    with_logging,
    is_logging_enabled,
    get_log_file_path,
    _extract_error_pattern,
    _sanitize_parameters,
    _truncate_string,
)


@pytest.fixture(autouse=True)
def _clean_logging_state(monkeypatch):
    """Reset module-level logging state before and after every test.

    Also disables the always-on diagnostic stream by default so unit
    tests don't write to the user's real ``~/.msaccess-vcs-mcp/logs/``
    directory during the suite. Tests that *want* to exercise the
    diagnostic stream override the env var explicitly via ``monkeypatch``.
    """
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

    logging_module._logging_enabled = None
    logging_module._log_handler = None
    logging_module._log_file = None
    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None
    yield
    logging_module._logging_enabled = None
    logging_module._log_handler = None
    logging_module._log_file = None
    logging_module._diagnostic_handler = None
    logging_module._diagnostic_file = None
    logging_module._diagnostic_initialized = False
    logging_module._diagnostic_enabled = False
    logging_module._diagnostic_disabled_reason = None


class TestInitializeLogging:
    """Tests for _initialize_logging() caching behaviour."""

    def test_disabled_not_cached(self):
        """When logging is disabled, _logging_enabled stays None so the
        check re-runs on the next call (allows lazy .env loading to
        enable it later).
        """
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            result = _initialize_logging()
            assert result is False
            assert logging_module._logging_enabled is None

    def test_enabled_cached(self, tmp_path):
        """When logging is enabled and initialisation succeeds, the True
        state is cached so subsequent calls skip re-init.
        """
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            result = _initialize_logging()
            assert result is True
            assert logging_module._logging_enabled is True

            second = _initialize_logging()
            assert second is True

    def test_works_after_lazy_env_load(self, tmp_path):
        """Simulates the lazy-init scenario: an explicit opt-out is in
        effect on the first call, then the .env enables logging and the
        next call initialises successfully without a process restart.
        """
        log_dir = tmp_path / "logs"

        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "false"},
            clear=False,
        ):
            assert _initialize_logging() is False
            # Disabled state is intentionally NOT cached so a later .env
            # load can flip the switch on.
            assert logging_module._logging_enabled is None

        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            assert _initialize_logging() is True
            assert logging_module._logging_enabled is True

    def test_failure_cached_to_avoid_retry_spam(self, tmp_path):
        """When init fails (e.g. can't create log dir), False is cached
        so we don't retry and spam stderr on every tool call.
        """
        bad_dir = tmp_path / "nonexistent" / "deeply" / "nested"

        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(bad_dir)},
            clear=False,
        ), patch.object(
            logging_module, "_ensure_log_dir", return_value=False
        ):
            result = _initialize_logging()
            assert result is False
            assert logging_module._logging_enabled is False

    def test_creates_log_file(self, tmp_path):
        """Successful init creates the vcs-mcp-usage.jsonl file (renamed
        from ``usage.jsonl`` so it self-identifies in shared logs dirs).
        """
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            assert (log_dir / "vcs-mcp-usage.jsonl").exists()


class TestResetLogging:
    """Tests for reset_logging() state cleanup."""

    def test_clears_all_state(self, tmp_path):
        """reset_logging() sets all module globals back to None."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            assert logging_module._logging_enabled is True
            assert logging_module._log_handler is not None
            assert logging_module._log_file is not None

        reset_logging()
        assert logging_module._logging_enabled is None
        assert logging_module._log_handler is None
        assert logging_module._log_file is None

    def test_closes_handler(self):
        """reset_logging() calls close() on the existing file handler."""
        mock_handler = MagicMock()
        logging_module._logging_enabled = True
        logging_module._log_handler = mock_handler
        logging_module._log_file = Path("/fake/log.jsonl")

        reset_logging()
        mock_handler.close.assert_called_once()

    def test_safe_when_never_initialized(self):
        """reset_logging() does not raise when logging was never init'd."""
        assert logging_module._log_handler is None
        reset_logging()
        assert logging_module._logging_enabled is None

    def test_reinitializes_after_reset(self, tmp_path):
        """After reset, the next _initialize_logging() call picks up
        fresh env vars and re-creates the handler.
        """
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            old_handler = logging_module._log_handler

        reset_logging()

        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            assert logging_module._log_handler is not None
            assert logging_module._log_handler is not old_handler


class TestLogToolCall:
    """Tests for log_tool_call() output."""

    def test_logs_successful_call(self, tmp_path):
        """A successful tool call produces a JSON line with success=True."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_tool_call(
                tool_name="vcs_list_objects",
                parameters={"database_path": "C:\\test.accdb"},
                execution_time_ms=42.5,
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        # First line is init, second is the tool call
        entry = json.loads(lines[-1])
        assert entry["event"] == "tool_call"
        assert entry["tool"] == "vcs_list_objects"
        assert entry["success"] is True
        assert entry["execution_time_ms"] == 42.5
        assert entry["parameters"]["database_path"] == "C:\\test.accdb"

    def test_logs_error_call(self, tmp_path):
        """A failed tool call includes error and error_pattern."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_tool_call(
                tool_name="vcs_export_database",
                parameters={"database_path": "C:\\test.accdb"},
                error="File not found: C:\\test.accdb",
                execution_time_ms=1.2,
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["success"] is False
        assert "File not found" in entry["error"]
        assert entry["error_pattern"] == "file_not_found"

    def test_detects_error_in_result(self, tmp_path):
        """When result dict contains 'error', success is set to False."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_tool_call(
                tool_name="vcs_compile_vba",
                parameters={},
                result={"error": "Compilation failed"},
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["success"] is False


class TestSanitizeParameters:
    """Tests for parameter sanitization."""

    def test_truncates_long_strings(self):
        long_sql = "SELECT " + "x" * 1000
        result = _sanitize_parameters({"query": long_sql}, max_string_length=100)
        assert len(result["query"]) < len(long_sql)
        assert "truncated" in result["query"]

    def test_limits_list_length(self):
        long_list = [f"item_{i}" for i in range(20)]
        result = _sanitize_parameters({"items": long_list})
        assert len(result["items"]) == 11  # 10 items + "... (10 more items)"

    def test_preserves_short_values(self):
        params = {"path": "C:\\db.accdb", "count": 5, "flag": True}
        result = _sanitize_parameters(params)
        assert result == params


class TestTruncateString:

    def test_short_string_unchanged(self):
        assert _truncate_string("hello", 100) == "hello"

    def test_long_string_truncated(self):
        result = _truncate_string("a" * 200, 50)
        assert len(result) < 200
        assert result.startswith("a" * 50)
        assert "truncated" in result


class TestErrorPatterns:
    """Tests for _extract_error_pattern categorization."""

    @pytest.mark.parametrize("error,expected", [
        ("pywintypes.com_error: (-2147352567, ...)", "com_error"),
        ("Database file not found: C:\\test.accdb", "file_not_found"),
        ("Object not found: MyForm", "object_not_found"),
        ("Permission denied accessing file", "permission_denied"),
        ("Write operations disabled", "write_disabled"),
        ("Operation timed out after 30s", "timeout"),
        ("VBA compile error in module", "vba_compile_error"),
        ("Add-in not installed", "addin_error"),
        ("Operation cancelled by user", "operation_cancelled"),
        ("Database is busy", "database_busy"),
        ("Something went wrong", "unknown"),
    ])
    def test_pattern_extraction(self, error, expected):
        assert _extract_error_pattern(error) == expected


class TestWithLoggingDecorator:
    """Tests for the with_logging decorator."""

    def test_sync_function_logged(self, tmp_path):
        """Sync functions are wrapped and logged."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            @with_logging("test_tool")
            def my_tool(name: str, count: int = 1) -> dict:
                return {"result": name, "count": count}

            result = my_tool("hello", count=3)
            assert result == {"result": "hello", "count": 3}

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["tool"] == "test_tool"
        assert entry["success"] is True
        assert entry["parameters"]["name"] == "hello"
        assert entry["parameters"]["count"] == 3

    def test_async_function_logged(self, tmp_path):
        """Async functions are wrapped and logged."""
        import asyncio

        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            @with_logging("test_async_tool")
            async def my_async_tool(path: str) -> dict:
                return {"path": path}

            result = asyncio.run(my_async_tool("C:\\test.accdb"))
            assert result == {"path": "C:\\test.accdb"}

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["tool"] == "test_async_tool"
        assert entry["success"] is True

    def test_exception_logged(self, tmp_path):
        """Exceptions are logged and re-raised."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            @with_logging("failing_tool")
            def my_failing_tool() -> dict:
                raise ValueError("something broke")

            with pytest.raises(ValueError, match="something broke"):
                my_failing_tool()

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["success"] is False
        assert "something broke" in entry["error"]

    def test_noop_when_disabled(self):
        """When logging is disabled, the decorator is transparent."""
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            @with_logging("test_tool")
            def my_tool() -> dict:
                return {"ok": True}

            result = my_tool()
            assert result == {"ok": True}


class TestHelpers:
    """Tests for is_logging_enabled and get_log_file_path."""

    def test_is_logging_enabled_false(self):
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            assert is_logging_enabled() is False

    def test_is_logging_enabled_true(self, tmp_path):
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            assert is_logging_enabled() is True

    def test_get_log_file_path_when_enabled(self, tmp_path):
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            path = get_log_file_path()
            assert path is not None
            assert path.name == "vcs-mcp-usage.jsonl"

    def test_get_log_file_path_when_disabled(self):
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            assert get_log_file_path() is None


class TestLogCodeExecution:
    """Tests for log_code_execution() audit logging.

    Default behaviour: ``code_length`` is recorded but the ``code``
    body is omitted. Set ``ACCESS_VCS_LOG_CODE_CONTENT=true`` to opt in
    to full-body capture for forensic replay.
    """

    def test_default_logs_only_code_length(self, tmp_path):
        """Without LOG_CODE_CONTENT, the body is redacted to length only."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            os.environ.pop("ACCESS_VCS_LOG_CODE_CONTENT", None)
            _initialize_logging()
            sql = "SELECT * FROM Customers WHERE Active = True"
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code=sql,
                code_type="sql",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["tool"] == "vcs_execute_sql"
        assert entry["database"] == "C:\\test.accdb"
        assert entry["code_type"] == "sql"
        assert entry["code_length"] == len(sql)
        assert "code" not in entry

    def test_explicit_opt_out_redacts_code(self, tmp_path):
        """LOG_CODE_CONTENT=false also redacts -- explicit opt-out."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {
                "ACCESS_VCS_ENABLE_LOGGING": "true",
                "ACCESS_VCS_LOG_DIR": str(log_dir),
                "ACCESS_VCS_LOG_CODE_CONTENT": "false",
            },
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_run_vba",
                database_path="C:\\test.accdb",
                code="Dim qd As DAO.QueryDef",
                code_type="vba",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["code_length"] == len("Dim qd As DAO.QueryDef")
        assert "code" not in entry

    def test_opt_in_logs_full_sql_body(self, tmp_path):
        """LOG_CODE_CONTENT=true records the full SQL body."""
        log_dir = tmp_path / "logs"
        sql = "SELECT * FROM Customers WHERE Active = True"
        with patch.dict(
            os.environ,
            {
                "ACCESS_VCS_ENABLE_LOGGING": "true",
                "ACCESS_VCS_LOG_DIR": str(log_dir),
                "ACCESS_VCS_LOG_CODE_CONTENT": "true",
            },
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code=sql,
                code_type="sql",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["code"] == sql
        assert entry["code_length"] == len(sql)

    def test_opt_in_logs_full_vba_body(self, tmp_path):
        """LOG_CODE_CONTENT=true records the full VBA body unredacted."""
        log_dir = tmp_path / "logs"
        vba_code = (
            "Dim qd As DAO.QueryDef\n"
            "Set qd = CurrentDb.QueryDefs(\"qryTest\")\n"
            "MCP_TempFunction = qd.SQL"
        )
        with patch.dict(
            os.environ,
            {
                "ACCESS_VCS_ENABLE_LOGGING": "true",
                "ACCESS_VCS_LOG_DIR": str(log_dir),
                "ACCESS_VCS_LOG_CODE_CONTENT": "true",
            },
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_run_vba",
                database_path="C:\\test.accdb",
                code=vba_code,
                code_type="vba",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["code_type"] == "vba"
        assert entry["code"] == vba_code
        assert "DAO.QueryDef" in entry["code"]

    def test_opt_in_logs_vba_call_unredacted(self, tmp_path):
        """LOG_CODE_CONTENT=true keeps vba_call invocations intact."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {
                "ACCESS_VCS_ENABLE_LOGGING": "true",
                "ACCESS_VCS_LOG_DIR": str(log_dir),
                "ACCESS_VCS_LOG_CODE_CONTENT": "true",
            },
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_call_vba",
                database_path="C:\\test.accdb",
                code="MyModule.DeleteAllRecords('tblCustomers')",
                code_type="vba_call",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["code_type"] == "vba_call"
        assert "DeleteAllRecords" in entry["code"]

    def test_opt_in_does_not_truncate_long_code(self, tmp_path):
        """When opted in, code content is preserved in full -- not
        truncated like tool parameters.
        """
        log_dir = tmp_path / "logs"
        long_sql = "SELECT " + ", ".join(f"field_{i}" for i in range(200))
        assert len(long_sql) > 500  # Exceeds normal truncation limit
        with patch.dict(
            os.environ,
            {
                "ACCESS_VCS_ENABLE_LOGGING": "true",
                "ACCESS_VCS_LOG_DIR": str(log_dir),
                "ACCESS_VCS_LOG_CODE_CONTENT": "true",
            },
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code=long_sql,
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["code"] == long_sql
        assert "truncated" not in entry["code"]
        assert entry["code_length"] == len(long_sql)

    def test_noop_when_disabled(self):
        """No error when logging is disabled — call is silently skipped."""
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code="SELECT 1",
            )


class TestLogAddinProbe:
    """Tests for log_addin_probe() probe instrumentation logging."""

    def test_logs_successful_probe(self, tmp_path):
        """A successful probe is logged with success=True and timed_out=False."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_addin_probe(
                addin_path="C:\\AppData\\MSAccessVCS\\Version Control.accda",
                duration_ms=12.5,
                success=True,
                timed_out=False,
                error=None,
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "addin_probe"
        assert entry["addin_path"].endswith("Version Control.accda")
        assert entry["duration_ms"] == 12.5
        assert entry["success"] is True
        assert entry["timed_out"] is False
        assert "error" not in entry

    def test_logs_timeout_probe(self, tmp_path):
        """A timed-out probe records timed_out=True and the error message."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_addin_probe(
                addin_path="C:\\AppData\\MSAccessVCS\\Version Control.accda",
                duration_ms=10000.0,
                success=False,
                timed_out=True,
                error="VCS add-in probe timed out after 10.0s ...",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "addin_probe"
        assert entry["success"] is False
        assert entry["timed_out"] is True
        assert "timed out" in entry["error"]
        # Timeout should categorize cleanly so analytics queries can group it.
        assert entry["error_pattern"] == "timeout"

    def test_logs_failure_probe(self, tmp_path):
        """A non-timeout failure records success=False, timed_out=False."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_addin_probe(
                addin_path="C:\\bad\\addin.accda",
                duration_ms=5.3,
                success=False,
                timed_out=False,
                error="Failed to load VCS add-in: COM error",
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["success"] is False
        assert entry["timed_out"] is False
        assert "COM error" in entry["error"]

    def test_noop_when_disabled(self):
        """No error when logging is disabled — call is silently skipped."""
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            log_addin_probe(
                addin_path="C:\\addin.accda",
                duration_ms=1.0,
                success=True,
            )


class TestDiagnosticLogging:
    """Tests for the always-on diagnostic stream (vcs-mcp-diagnostic.jsonl)."""

    def test_enabled_by_default(self, tmp_path, monkeypatch):
        """Without the opt-out env var the diagnostic stream initialises."""
        diag_dir = tmp_path / "diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)

        assert _initialize_diagnostic_logging() is True
        assert is_diagnostic_logging_enabled() is True
        path = get_diagnostic_log_path()
        assert path is not None
        assert path.name == "vcs-mcp-diagnostic.jsonl"

    def test_opt_out_via_env_var(self, tmp_path, monkeypatch):
        """ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true disables the stream
        cleanly, with no file created and no exceptions raised.
        """
        diag_dir = tmp_path / "diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

        assert _initialize_diagnostic_logging() is False
        assert is_diagnostic_logging_enabled() is False
        assert get_diagnostic_log_path() is None
        log_diagnostic_event("server_start", cwd="/somewhere")
        assert not diag_dir.exists() or not any(diag_dir.iterdir())

    def test_honors_custom_log_dir(self, tmp_path, monkeypatch):
        """ACCESS_VCS_DIAGNOSTIC_LOG_DIR overrides the default location."""
        diag_dir = tmp_path / "custom_diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)

        log_diagnostic_event("server_start", cwd="C:\\proj")

        path = get_diagnostic_log_path()
        assert path is not None
        assert path.parent == diag_dir.resolve() or path.parent == diag_dir

    def test_writes_valid_jsonl(self, tmp_path, monkeypatch):
        """Each log_diagnostic_event call appends one parseable JSON line."""
        diag_dir = tmp_path / "diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)

        log_diagnostic_event("server_start", cwd="C:\\proj", session_id="abc123")
        log_diagnostic_event("startup_env_load", project_root="C:\\proj", env_loaded=True)

        path = get_diagnostic_log_path()
        assert path is not None
        lines = path.read_text(encoding="utf-8").strip().split("\n")
        assert len(lines) == 2
        first = json.loads(lines[0])
        second = json.loads(lines[1])

        assert first["event"] == "server_start"
        assert first["cwd"] == "C:\\proj"
        assert first["session_id"] == "abc123"
        assert "timestamp" in first
        assert "version" in first

        assert second["event"] == "startup_env_load"
        assert second["env_loaded"] is True

    def test_independent_of_usage_logging(self, tmp_path, monkeypatch):
        """Diagnostic events are written even when usage logging is off
        -- this is the whole point of the diagnostic stream.
        """
        diag_dir = tmp_path / "diag"
        usage_dir = tmp_path / "usage"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)
        monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
        monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(usage_dir))

        log_diagnostic_event("server_start", cwd="C:\\proj")

        diag_path = get_diagnostic_log_path()
        assert diag_path is not None and diag_path.exists()
        # Usage stream stays untouched.
        assert get_log_file_path() is None
        assert not (usage_dir / "vcs-mcp-usage.jsonl").exists()

    def test_initialization_is_idempotent(self, tmp_path, monkeypatch):
        """A second call returns the cached state without reopening the
        handler.
        """
        diag_dir = tmp_path / "diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)

        assert _initialize_diagnostic_logging() is True
        first = logging_module._diagnostic_handler
        assert _initialize_diagnostic_logging() is True
        assert logging_module._diagnostic_handler is first

    def test_reset_logging_clears_diagnostic_state(self, tmp_path, monkeypatch):
        """reset_logging() also closes and clears the diagnostic handler
        so a config reload picks up new diagnostic env vars.
        """
        diag_dir = tmp_path / "diag"
        monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))
        monkeypatch.delenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", raising=False)

        _initialize_diagnostic_logging()
        assert logging_module._diagnostic_handler is not None

        reset_logging()
        assert logging_module._diagnostic_handler is None
        assert logging_module._diagnostic_initialized is False
        assert logging_module._diagnostic_enabled is False


class TestSecretKeyMasking:
    """Tests for the secret-key auto-mask in _sanitize_parameters.

    Defense-in-depth: parameter keys whose names *look* like credentials
    are masked regardless of any other logging switch.
    """

    @pytest.mark.parametrize("key", [
        "password",
        "Password",
        "user_password",
        "secret",
        "API_SECRET",
        "token",
        "auth_token",
        "api_key",
        "apiKey",
        "apikey",
        "API-KEY",
        "connection_string",
        "ConnectionString",
        "connectionstring",
    ])
    def test_redacts_credential_shaped_keys(self, key):
        """Each variant matches the SECRET_KEY_PATTERN and is masked."""
        result = _sanitize_parameters({key: "supersecretvalue"})
        assert result[key] == "<redacted>"

    def test_does_not_redact_innocent_keys(self):
        """Keys that don't match the pattern pass through unchanged."""
        params = {
            "database_path": "C:\\db.accdb",
            "tool_name": "vcs_export",
            "count": 5,
        }
        result = _sanitize_parameters(params)
        assert result == params

    def test_redacts_in_nested_dict(self):
        """Nested dicts also have their credential-shaped keys masked."""
        params = {
            "config": {
                "user": "alice",
                "password": "hunter2",
            },
        }
        result = _sanitize_parameters(params)
        assert result["config"]["user"] == "alice"
        assert result["config"]["password"] == "<redacted>"

    def test_logged_through_tool_call(self, tmp_path):
        """End-to-end: log_tool_call masks credential-shaped params before
        they reach disk.
        """
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_tool_call(
                tool_name="vcs_test",
                parameters={
                    "database_path": "C:\\test.accdb",
                    "password": "hunter2",
                    "API_KEY": "sk-abc123",
                },
            )

        lines = (log_dir / "vcs-mcp-usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["parameters"]["database_path"] == "C:\\test.accdb"
        assert entry["parameters"]["password"] == "<redacted>"
        assert entry["parameters"]["API_KEY"] == "<redacted>"
        # Sanity check: the literal secret values do not appear anywhere.
        raw = lines[-1]
        assert "hunter2" not in raw
        assert "sk-abc123" not in raw


class TestDefaultLogDirResolution:
    """Tests for ``_get_default_log_dir`` workspace-aware resolution.

    When the user's project root has been discovered (via env-var, CWD
    walk, or MCP ``roots/list``), tool-call audit logs should land in
    that workspace -- not in the MCP server's own clone or the user's
    home directory. This keeps logging behaviour predictable regardless
    of where the MCP server happens to be installed.
    """

    def setup_method(self):
        # Snapshot config module state so we can restore it.
        import msaccess_vcs_mcp.config as _config
        self._saved_root = _config._project_root
        self._saved_method = _config._project_root_method

    def teardown_method(self):
        import msaccess_vcs_mcp.config as _config
        _config._project_root = self._saved_root
        _config._project_root_method = self._saved_method

    def test_prefers_workspace_root_over_dev_install(self, tmp_path):
        """A discovered workspace root wins over the MCP server's own repo."""
        import msaccess_vcs_mcp.config as _config
        from msaccess_vcs_mcp.usage_logging import _get_default_log_dir

        workspace = tmp_path / "user_project"
        workspace.mkdir()
        _config._project_root = workspace
        _config._project_root_method = _config.RESOLUTION_PROJECT_DIR_ENV

        result = _get_default_log_dir()
        assert result == workspace / "logs"

    def test_workspace_method_cwd_walk_marker_is_used(self, tmp_path):
        """All real resolution methods (env-var, CWD walk, package walk,
        MCP roots) qualify as 'a real workspace'; only the explicit
        no-env CWD fallback is treated as unreliable.
        """
        import msaccess_vcs_mcp.config as _config
        from msaccess_vcs_mcp.usage_logging import _get_default_log_dir

        workspace = tmp_path / "ws"
        workspace.mkdir()

        for method in (
            _config.RESOLUTION_PROJECT_DIR_ENV,
            _config.RESOLUTION_WORKSPACE_ROOTS,
            _config.RESOLUTION_CWD_ENV,
            _config.RESOLUTION_CWD_MARKER,
            _config.RESOLUTION_PACKAGE_ENV,
            _config.RESOLUTION_PACKAGE_MARKER,
        ):
            _config._project_root = workspace
            _config._project_root_method = method
            assert _get_default_log_dir() == workspace / "logs", (
                f"method={method} should be treated as a real workspace"
            )

    def test_cwd_fallback_does_not_count_as_workspace(self, tmp_path, monkeypatch):
        """``RESOLUTION_CWD_FALLBACK`` means 'no .env or markers found,
        falling back to CWD'. We refuse to write logs there because the
        CWD may be the user's home or a temp dir. Caller falls through
        to dev-install or ``~/.msaccess-vcs-mcp/logs/``.
        """
        import msaccess_vcs_mcp.config as _config
        from msaccess_vcs_mcp.usage_logging import _get_default_log_dir

        bogus_cwd = tmp_path / "bogus"
        bogus_cwd.mkdir()
        _config._project_root = bogus_cwd
        _config._project_root_method = _config.RESOLUTION_CWD_FALLBACK

        result = _get_default_log_dir()
        assert result != bogus_cwd / "logs"
        # Falls back to either the dev-install logs/ or ~/.msaccess-vcs-mcp/logs/.
        assert "logs" in str(result)

    def test_falls_back_when_no_workspace_discovered(self, monkeypatch):
        """With no workspace discovered, behaviour matches the prior
        defaults: dev-install logs/ in development, user-home in package
        mode.
        """
        import msaccess_vcs_mcp.config as _config
        from msaccess_vcs_mcp.usage_logging import _get_default_log_dir

        _config._project_root = None
        _config._project_root_method = None

        result = _get_default_log_dir()
        # The result is one of the two non-workspace fallbacks; we just
        # verify it doesn't crash and lands in a reasonable place.
        assert result.name == "logs" or result.name.endswith(".msaccess-vcs-mcp")

    def test_rejects_user_home_as_workspace_root(self, monkeypatch):
        """Cursor's user-level ``~/.cursor/mcp.json`` causes the startup
        CWD walk to identify the user's home as the project root. We
        refuse to write tool-call logs into ``~/logs/`` and fall back
        instead -- otherwise the MCP server litters the user's home
        directory with audit data.
        """
        import msaccess_vcs_mcp.config as _config
        from msaccess_vcs_mcp.usage_logging import (
            _get_default_log_dir,
            _get_workspace_project_root,
        )

        _config._project_root = Path.home()
        _config._project_root_method = _config.RESOLUTION_CWD_MARKER

        # Workspace resolver refuses the home dir.
        assert _get_workspace_project_root() is None

        # Default log-dir falls back to dev-install or ~/.msaccess-vcs-mcp/.
        result = _get_default_log_dir()
        assert result != Path.home() / "logs"
