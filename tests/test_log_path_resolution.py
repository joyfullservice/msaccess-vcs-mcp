"""Tests for VCS operation log path resolution and reporting.

Covers the boundary between the add-in's camelCase JSON API and the
snake_case MCP result contract, the newest-log lookup that replaced the old
root-level ``Export.log``/``Build.log`` fallbacks, and the failure excerpt
that spares agents a follow-up call into a gitignored folder.
"""

import asyncio
import json
from contextlib import contextmanager
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from msaccess_vcs_mcp.addin_integration import MERGE_NOT_AVAILABLE_MESSAGE
from msaccess_vcs_mcp.tools import (
    _addin_json_result,
    _attach_log_context,
    _newest_log,
    _read_log_excerpt,
)


def _unwrap_sync(tool_fn):
    sync_fn = tool_fn
    while hasattr(sync_fn, "__wrapped__"):
        sync_fn = sync_fn.__wrapped__
    return sync_fn


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


def _write_log(source_dir, name, content="line one\nline two\n"):
    logs = source_dir / "logs"
    logs.mkdir(exist_ok=True)
    path = logs / name
    path.write_text(content, encoding="utf-8")
    return path


@contextmanager
def _patch_sync_tool(tmp_path, call_sync_result):
    """Patch the COM plumbing so a sync passthrough tool can run offline."""
    mock_conn = MagicMock()
    mock_conn.__enter__ = MagicMock(return_value=mock_conn)
    mock_conn.__exit__ = MagicMock(return_value=False)
    mock_conn.connect.return_value = (MagicMock(), MagicMock())

    mock_addin = MagicMock()
    mock_addin.call_sync.return_value = call_sync_result

    db_path = tmp_path / "test.accdb"
    db_path.touch()

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db_path),
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db_path), mock_addin


class TestAddinJsonResult:
    """The add-in speaks camelCase; MCP results are snake_case."""

    def test_camel_case_log_path_becomes_snake_case(self):
        result = _addin_json_result(json.dumps({"success": True, "logPath": "C:\\x\\logs\\Merge_1.log"}))

        assert result["log_path"] == "C:\\x\\logs\\Merge_1.log"

    def test_original_camel_case_key_retained_as_alias(self):
        result = _addin_json_result(json.dumps({"success": True, "logPath": "C:\\x.log"}))

        assert result["logPath"] == "C:\\x.log"

    def test_snake_case_from_newer_addin_passes_through(self):
        """A future add-in that emits log_path must keep working."""
        result = _addin_json_result(json.dumps({"success": True, "log_path": "C:\\x.log"}))

        assert result["log_path"] == "C:\\x.log"
        assert result["logPath"] == "C:\\x.log"

    def test_result_without_log_path_is_untouched(self):
        result = _addin_json_result(json.dumps({"success": True, "other": 1}))

        assert result == {"success": True, "other": 1}
        assert "log_path" not in result

    def test_non_json_string_uses_raw_key(self):
        result = _addin_json_result("not json at all")

        assert result == {"success": True, "result": "not json at all"}

    def test_non_json_string_honours_custom_raw_key(self):
        result = _addin_json_result("plain text log body", raw_key="content")

        assert result == {"success": True, "content": "plain text log body"}

    def test_json_scalar_uses_raw_key(self):
        """A bare JSON scalar is not a result object."""
        result = _addin_json_result("42")

        assert result == {"success": True, "result": "42"}

    def test_non_string_input_uses_raw_key(self):
        result = _addin_json_result(True)

        assert result == {"success": True, "result": True}

    def test_failure_gains_excerpt(self, tmp_path):
        log = _write_log(tmp_path, "Merge_20260807_120000_000.log", "boom\nit broke\n")
        payload = json.dumps({"success": False, "error": "nope", "logPath": str(log)})

        result = _addin_json_result(payload)

        assert "it broke" in result["log_excerpt"]

    def test_success_has_no_excerpt(self, tmp_path):
        log = _write_log(tmp_path, "Merge_20260807_120000_000.log")
        payload = json.dumps({"success": True, "logPath": str(log)})

        result = _addin_json_result(payload)

        assert "log_excerpt" not in result


class TestNewestLog:
    """Replaces the dead root-level Export.log / Build.log fallbacks."""

    def test_picks_newest_by_sortable_timestamp(self, tmp_path):
        _write_log(tmp_path, "Merge_20260806_160652_473.log")
        newest = _write_log(tmp_path, "Merge_20260807_095558_258.log")
        _write_log(tmp_path, "Merge_20260721_170955_695.log")

        assert _newest_log(tmp_path, "Merge") == str(newest)

    def test_base_names_do_not_cross_contaminate(self, tmp_path):
        _write_log(tmp_path, "Build_20260807_235959_999.log")
        merge = _write_log(tmp_path, "Merge_20260101_000000_000.log")

        assert _newest_log(tmp_path, "Merge") == str(merge)

    def test_missing_logs_folder_returns_none(self, tmp_path):
        assert _newest_log(tmp_path, "Merge") is None

    def test_empty_logs_folder_returns_none(self, tmp_path):
        (tmp_path / "logs").mkdir()

        assert _newest_log(tmp_path, "Merge") is None

    def test_root_level_legacy_log_is_not_matched(self, tmp_path):
        """The add-in migrates these into logs/; they must not be picked up."""
        (tmp_path / "Build.log").write_text("legacy", encoding="utf-8")

        assert _newest_log(tmp_path, "Build") is None


class TestReadLogExcerpt:
    def test_returns_tail_only(self, tmp_path):
        body = "".join(f"line {i}\n" for i in range(200))
        log = _write_log(tmp_path, "Build_20260807_120000_000.log", body)

        excerpt = _read_log_excerpt(str(log), tail_lines=10)

        assert "line 199" in excerpt
        assert "line 100" not in excerpt

    def test_truncates_to_max_chars(self, tmp_path):
        body = "".join(f"{'x' * 200}\n" for _ in range(50))
        log = _write_log(tmp_path, "Build_20260807_120000_000.log", body)

        excerpt = _read_log_excerpt(str(log), tail_lines=50, max_chars=500)

        assert "truncated" in excerpt
        assert len(excerpt) < 700

    def test_missing_file_returns_none(self, tmp_path):
        assert _read_log_excerpt(str(tmp_path / "nope.log")) is None

    def test_empty_path_returns_none(self):
        assert _read_log_excerpt("") is None

    def test_empty_file_returns_none(self, tmp_path):
        log = _write_log(tmp_path, "Build_20260807_120000_000.log", "")

        assert _read_log_excerpt(str(log)) is None

    def test_undecodable_bytes_do_not_raise(self, tmp_path):
        logs = tmp_path / "logs"
        logs.mkdir()
        log = logs / "Build_20260807_120000_000.log"
        log.write_bytes(b"\x80\x81 bad bytes\n")

        assert _read_log_excerpt(str(log)) is not None


class TestAttachLogContext:
    def test_completion_log_path_wins_over_disk_lookup(self, tmp_path):
        from_callback = _write_log(tmp_path, "Merge_20260101_000000_000.log")
        _write_log(tmp_path, "Merge_20260807_235959_999.log")

        result = _attach_log_context(
            {"success": True}, tmp_path, "Merge", {"log_path": str(from_callback)}
        )

        assert result["log_path"] == str(from_callback)

    def test_result_log_path_wins_when_no_completion(self, tmp_path):
        """Sync ImportByType / ExportByType put log_path on the result first."""
        from_result = _write_log(tmp_path, "Merge_20260101_000000_000.log")
        _write_log(tmp_path, "Merge_20260807_235959_999.log")

        result = _attach_log_context(
            {"success": True, "log_path": str(from_result)}, tmp_path, "Merge"
        )

        assert result["log_path"] == str(from_result)

    def test_falls_back_to_disk_when_callback_has_no_path(self, tmp_path):
        newest = _write_log(tmp_path, "Merge_20260807_235959_999.log")

        result = _attach_log_context({"success": True}, tmp_path, "Merge", {})

        assert result["log_path"] == str(newest)

    def test_stale_callback_path_falls_back_to_disk(self, tmp_path):
        newest = _write_log(tmp_path, "Merge_20260807_235959_999.log")

        result = _attach_log_context(
            {"success": True}, tmp_path, "Merge", {"log_path": str(tmp_path / "gone.log")}
        )

        assert result["log_path"] == str(newest)

    def test_stale_callback_path_with_no_logs_is_none(self, tmp_path):
        result = _attach_log_context(
            {"success": True}, tmp_path, "Merge", {"log_path": str(tmp_path / "gone.log")}
        )

        assert result["log_path"] is None

    def test_failure_includes_excerpt(self, tmp_path):
        _write_log(tmp_path, "Merge_20260807_120000_000.log", "Error: merge blew up\n")

        result = _attach_log_context({"success": False, "error": "x"}, tmp_path, "Merge")

        assert "merge blew up" in result["log_excerpt"]

    def test_success_omits_excerpt(self, tmp_path):
        _write_log(tmp_path, "Merge_20260807_120000_000.log")

        result = _attach_log_context({"success": True}, tmp_path, "Merge")

        assert "log_excerpt" not in result
        assert result["log_path"] is not None

    def test_failure_without_any_log_omits_excerpt(self, tmp_path):
        result = _attach_log_context({"success": False, "error": "x"}, tmp_path, "Merge")

        assert result["log_path"] is None
        assert "log_excerpt" not in result


class TestSyncPassthroughTools:
    """vcs_import_object regression: the add-in fix must reach the caller."""

    def test_import_object_surfaces_log_path(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_object

        payload = json.dumps({"success": True, "logPath": "C:\\src\\logs\\Merge_1.log"})
        with _patch_sync_tool(tmp_path, payload) as (db_path, _):
            result = _unwrap_sync(vcs_import_object)(db_path, "module", "modQueueDb")

        assert result["log_path"] == "C:\\src\\logs\\Merge_1.log"

    def test_import_object_empty_log_path_stays_falsy(self, tmp_path):
        """The pre-fix add-in returned "" here; it must not become truthy."""
        from msaccess_vcs_mcp.tools import vcs_import_object

        payload = json.dumps({"success": True, "logPath": ""})
        with _patch_sync_tool(tmp_path, payload) as (db_path, _):
            result = _unwrap_sync(vcs_import_object)(db_path, "module", "modQueueDb")

        assert not result.get("log_path")

    def test_export_object_surfaces_log_path(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_export_object

        payload = json.dumps({"success": True, "logPath": "C:\\src\\logs\\Export_1.log"})
        with _patch_sync_tool(tmp_path, payload) as (db_path, _):
            result = _unwrap_sync(vcs_export_object)(db_path, "query", "qryFoo")

        assert result["log_path"] == "C:\\src\\logs\\Export_1.log"

    def test_get_log_surfaces_log_path_and_content(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_get_log

        payload = json.dumps(
            {"success": True, "logPath": "C:\\src\\logs\\Merge_1.log", "content": "log body"}
        )
        with _patch_sync_tool(tmp_path, payload) as (db_path, _):
            result = _unwrap_sync(vcs_get_log)(db_path, log_type="Merge")

        assert result["log_path"] == "C:\\src\\logs\\Merge_1.log"
        assert result["content"] == "log body"

    def test_get_log_passes_requested_type_to_addin(self, tmp_path):
        """A merge must not be looked up under the Build base name."""
        from msaccess_vcs_mcp.tools import vcs_get_log

        payload = json.dumps({"success": True, "content": "x"})
        with _patch_sync_tool(tmp_path, payload) as (db_path, mock_addin):
            _unwrap_sync(vcs_get_log)(db_path, log_type="Merge")

        mock_addin.call_sync.assert_called_with("GetLogContent", "Merge")

    def test_get_log_non_json_body_uses_content_key(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_get_log

        with _patch_sync_tool(tmp_path, "raw log text") as (db_path, _):
            result = _unwrap_sync(vcs_get_log)(db_path, log_type="Export")

        assert result["content"] == "raw log text"


@contextmanager
def _patch_merge_tool(tmp_path, async_result, completion=None):
    """Patch the async merge plumbing for vcs_import_objects."""
    mock_conn = MagicMock()
    mock_conn.__enter__ = MagicMock(return_value=mock_conn)
    mock_conn.__exit__ = MagicMock(return_value=False)
    mock_conn.connect.return_value = (MagicMock(), MagicMock())

    mock_addin = MagicMock()
    mock_addin.call_async.return_value = async_result
    mock_addin.merge_build.return_value = {"success": True, "message": "ok"}

    db_path = tmp_path / "test.accdb"
    db_path.touch()
    src_path = tmp_path / "src"
    src_path.mkdir(exist_ok=True)
    mock_addin.call_sync.side_effect = (
        lambda command, *_a: str(src_path) + "\\" if command == "GetExportFolder" else None
    )

    op_manager = MagicMock()
    op_manager.register_operation.return_value = ("op-1", MagicMock())
    op_manager.create_callback_info.return_value = "{}"
    if completion is not None:
        op_manager.wait_for_completion = AsyncMock(return_value=completion)

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db_path),
        patch("msaccess_vcs_mcp.tools.validate_source_directory", return_value=src_path),
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://localhost:1/cb"),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=op_manager),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db_path), src_path, mock_addin


class TestImportObjectsAsyncPaths:
    """The async dispatch must never report success for work that never ran."""

    def test_missing_async_flag_falls_back_to_sync_merge(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_merge_tool(tmp_path, {}) as (db_path, src_path, mock_addin):
            _write_log(src_path, "Merge_20260807_120000_000.log")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        mock_addin.merge_build.assert_called_once()
        assert result["success"] is True

    def test_sync_marker_does_not_merge_twice(self, tmp_path):
        """{"sync": true} means the add-in already ran it inline."""
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_merge_tool(tmp_path, {"sync": True, "result": "done"}) as (
            db_path,
            src_path,
            mock_addin,
        ):
            _write_log(src_path, "Merge_20260807_120000_000.log")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        mock_addin.merge_build.assert_not_called()
        assert result["success"] is True

    def test_sync_fallback_still_resolves_log_path(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_merge_tool(tmp_path, {}) as (db_path, src_path, _):
            log = _write_log(src_path, "Merge_20260807_120000_000.log")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        assert result["log_path"] == str(log)

    def test_failed_sync_fallback_returns_excerpt(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_merge_tool(tmp_path, {}) as (db_path, src_path, mock_addin):
            mock_addin.merge_build.return_value = {"success": False, "message": "merge failed"}
            _write_log(src_path, "Merge_20260807_120000_000.log", "Error: bad source file\n")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        assert result["success"] is False
        assert "bad source file" in result["log_excerpt"]

    def test_merge_not_available_callback_returns_without_waiting(self, tmp_path):
        """The add-in's error callback ends the wait, and an old log is not blamed."""
        from msaccess_vcs_mcp.tools import vcs_import_objects

        completion = {
            "success": False,
            "error": "Merge Build Not Available: ...",
            "code": "merge_not_available",
            "log_path": None,
        }
        with _patch_merge_tool(
            tmp_path, {"async": True, "timeout_ms": 300000}, completion
        ) as (db_path, src_path, mock_addin):
            _write_log(src_path, "Merge_20260807_120000_000.log", "an earlier merge\n")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        mock_addin.merge_build.assert_not_called()
        assert result["success"] is False
        assert result["error"] == MERGE_NOT_AVAILABLE_MESSAGE
        assert result["error_pattern"] == "merge_not_available"
        assert "log_path" not in result
        assert "log_excerpt" not in result

    def test_merge_not_available_in_sync_fallback_skips_stale_log(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_import_objects

        with _patch_merge_tool(tmp_path, {}) as (db_path, src_path, mock_addin):
            mock_addin.merge_build.return_value = {
                "success": False,
                "message": MERGE_NOT_AVAILABLE_MESSAGE,
            }
            _write_log(src_path, "Merge_20260807_120000_000.log", "an earlier merge\n")
            result = asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path)))

        assert result["success"] is False
        assert result["error_pattern"] == "merge_not_available"
        assert "log_path" not in result
