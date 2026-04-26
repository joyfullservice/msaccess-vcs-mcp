"""Tests for usage logging lifecycle and hot-reload interactions."""

import json
import os
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import msaccess_vcs_mcp.usage_logging as logging_module
from msaccess_vcs_mcp.usage_logging import (
    _initialize_logging,
    reset_logging,
    log_tool_call,
    log_code_execution,
    with_logging,
    is_logging_enabled,
    get_log_file_path,
    _extract_error_pattern,
    _sanitize_parameters,
    _truncate_string,
)


@pytest.fixture(autouse=True)
def _clean_logging_state():
    """Reset module-level logging state before and after every test."""
    logging_module._logging_enabled = None
    logging_module._log_handler = None
    logging_module._log_file = None
    yield
    logging_module._logging_enabled = None
    logging_module._log_handler = None
    logging_module._log_file = None


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
        """Simulates the lazy-init scenario: first call with no env var
        returns False, then the env var is set, and the next call
        initialises successfully.
        """
        log_dir = tmp_path / "logs"

        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("ACCESS_VCS_ENABLE_LOGGING", None)
            os.environ.pop("ACCESS_VCS_LOG_DIR", None)

            assert _initialize_logging() is False
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
        """Successful init creates the usage.jsonl file."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            assert (log_dir / "usage.jsonl").exists()


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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
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
            assert path.name == "usage.jsonl"

    def test_get_log_file_path_when_disabled(self):
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            assert get_log_file_path() is None


class TestLogCodeExecution:
    """Tests for log_code_execution() audit logging."""

    def test_logs_sql_execution(self, tmp_path):
        """SQL code execution is logged with event=code_execution."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code="SELECT * FROM Customers WHERE Active = True",
                code_type="sql",
            )

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["tool"] == "vcs_execute_sql"
        assert entry["database"] == "C:\\test.accdb"
        assert entry["code_type"] == "sql"
        assert entry["code"] == "SELECT * FROM Customers WHERE Active = True"

    def test_logs_vba_execution(self, tmp_path):
        """VBA code execution is logged with code_type=vba."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            vba_code = (
                "Dim qd As DAO.QueryDef\n"
                "Set qd = CurrentDb.QueryDefs(\"qryTest\")\n"
                "MCP_TempFunction = qd.SQL"
            )
            log_code_execution(
                tool_name="vcs_run_vba",
                database_path="C:\\test.accdb",
                code=vba_code,
                code_type="vba",
            )

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["code_type"] == "vba"
        assert "DAO.QueryDef" in entry["code"]

    def test_logs_vba_call(self, tmp_path):
        """VBA function calls are logged with code_type=vba_call."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            log_code_execution(
                tool_name="vcs_call_vba",
                database_path="C:\\test.accdb",
                code="MyModule.DeleteAllRecords('tblCustomers')",
                code_type="vba_call",
            )

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["event"] == "code_execution"
        assert entry["code_type"] == "vba_call"
        assert "DeleteAllRecords" in entry["code"]

    def test_code_not_truncated(self, tmp_path):
        """Code content is preserved in full, not truncated like tool parameters."""
        log_dir = tmp_path / "logs"
        with patch.dict(
            os.environ,
            {"ACCESS_VCS_ENABLE_LOGGING": "true", "ACCESS_VCS_LOG_DIR": str(log_dir)},
            clear=False,
        ):
            _initialize_logging()
            long_sql = "SELECT " + ", ".join(f"field_{i}" for i in range(200))
            assert len(long_sql) > 500  # Exceeds normal truncation limit
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code=long_sql,
            )

        lines = (log_dir / "usage.jsonl").read_text().strip().split("\n")
        entry = json.loads(lines[-1])
        assert entry["code"] == long_sql
        assert "truncated" not in entry["code"]

    def test_noop_when_disabled(self):
        """No error when logging is disabled — call is silently skipped."""
        with patch.dict(os.environ, {"ACCESS_VCS_ENABLE_LOGGING": "false"}, clear=False):
            log_code_execution(
                tool_name="vcs_execute_sql",
                database_path="C:\\test.accdb",
                code="SELECT 1",
            )
