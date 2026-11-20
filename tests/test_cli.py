"""Tests for the streaming msaccess-vcs CLI."""

from __future__ import annotations

import json
from types import SimpleNamespace

from msaccess_vcs_mcp.cli import (
    arguments_for,
    build_parser,
    compact_result_payload,
    completion_message,
    format_duration_ms,
    format_progress_line,
    main,
    parse_result_payload,
    print_progress,
    ProgressPrinter,
    result_succeeded,
    startup_message,
    stdio_server_environment,
)


def test_arguments_for_each_subcommand():
    parser = build_parser()
    export = parser.parse_args(["export", r"C:\db.accdb", r"C:\src", "--full"])
    assert arguments_for(export) == (
        "vcs_export_database",
        {
            "database_path": r"C:\db.accdb",
            "output_dir": r"C:\src",
            "full_export": True,
        },
    )

    merge = parser.parse_args(["merge", r"C:\db.accdb", r"C:\src"])
    assert arguments_for(merge) == (
        "vcs_import_objects",
        {"database_path": r"C:\db.accdb", "source_dir": r"C:\src"},
    )

    rebuild = parser.parse_args([
        "rebuild-database",
        r"C:\src",
        r"C:\out.accdb",
        "--template",
        r"C:\blank.accdb",
    ])
    name, payload = arguments_for(rebuild)
    assert name == "vcs_rebuild_database"
    assert payload["template_path"] == r"C:\blank.accdb"

    rebuild_default = parser.parse_args(["rebuild-database", r"C:\src"])
    assert arguments_for(rebuild_default) == (
        "vcs_rebuild_database",
        {"source_dir": r"C:\src"},
    )

    addin = parser.parse_args(["rebuild-addin", r"C:\src", "--timeout", "90"])
    assert arguments_for(addin) == (
        "vcs_rebuild_addin",
        {"source_dir": r"C:\src", "timeout_seconds": 90.0},
    )

    tests = parser.parse_args([
        "run-tests",
        r"C:\db.accda",
        "--filter",
        "SQL,-slow",
        "--timeout",
        "90",
    ])
    assert arguments_for(tests) == (
        "vcs_run_tests",
        {
            "database_path": r"C:\db.accda",
            "filter": "SQL,-slow",
            "timeout_seconds": 90.0,
        },
    )


def test_startup_message():
    assert startup_message("rebuild-addin") == "Starting rebuild-addin..."
    assert startup_message("export") == "Starting export..."


def test_format_progress_line():
    assert format_progress_line(3, None, "building") == "building"
    assert format_progress_line(2, 10, "queries") == "queries"
    assert format_progress_line(1, None, None) == ""
    assert format_progress_line(1, None, "  ") == ""


def test_print_progress_skips_blank(capsys):
    print_progress(1, None, None)
    print_progress(2, None, "   ")
    print_progress(3, None, "Importing modules...")
    assert capsys.readouterr().out.splitlines() == ["Importing modules..."]


def test_result_succeeded_from_json():
    assert result_succeeded({"success": True})
    assert not result_succeeded({"success": False, "error": "nope"})
    assert result_succeeded({"status": "complete"})
    assert not result_succeeded({"cancelled": True})


def test_parse_result_payload():
    assert parse_result_payload('{"success": true}') == {"success": True}
    assert parse_result_payload("not json") == "not json"


def test_stdio_server_skips_access_session_cleanup(monkeypatch):
    monkeypatch.setenv("EXISTING_VALUE", "kept")
    env = stdio_server_environment()
    assert env["EXISTING_VALUE"] == "kept"
    assert env["ACCESS_VCS_SKIP_SESSION_CLEANUP"] == "true"


def test_main_streams_progress_then_json(capsys):
    async def factory(name, arguments, on_progress):
        assert name == "vcs_rebuild_addin"
        assert arguments["source_dir"] == r"C:\src"
        await on_progress(1, None, "Rebuild launched")
        await on_progress(2, None, "Rebuild complete")
        return SimpleNamespace(
            content=[SimpleNamespace(text=json.dumps({"success": True, "status": "complete"}))],
            isError=False,
        )

    code = main(["rebuild-addin", r"C:\src"], session_factory=factory)
    captured = capsys.readouterr()
    lines = [line for line in captured.out.splitlines() if line]
    assert lines[0] == "Starting rebuild-addin..."
    assert lines[1] == "Rebuild launched"
    assert lines[2] == "Rebuild complete"
    assert lines[-1] == "Rebuild complete."
    assert json.loads("\n".join(lines[3:-1]))["status"] == "complete"
    assert code == 0


def test_main_failure_exit_code(capsys):
    async def factory(_name, _arguments, on_progress):
        await on_progress(1, None, "Rebuild refused")
        return SimpleNamespace(
            content=[SimpleNamespace(text=json.dumps({
                "success": False,
                "status": "refused",
                "error": "busy",
            }))],
            isError=False,
        )

    code = main(["rebuild-addin", r"C:\src"], session_factory=factory)
    assert code == 1
    out = capsys.readouterr().out
    assert "Starting rebuild-addin..." in out
    assert "Rebuild refused" in out
    assert out.strip().splitlines()[-1] == "Rebuild failed."


def test_format_duration_ms():
    assert format_duration_ms(25) == "25ms"
    assert format_duration_ms(1480) == "1.48s"
    assert format_duration_ms(196691) == "3.3m"


def test_completion_message_run_tests():
    payload = {
        "success": True,
        "durationMs": 196691,
        "summary": {
            "subs": 468,
            "assertions": 2281,
            "passed": 2278,
            "failed": 0,
            "errored": 0,
            "empty": 3,
        },
        "tests": {"modFoo.Bar": {"status": "PASSED"}},
    }
    assert completion_message("run-tests", payload, True) == (
        "Tests passed. 468 subs, 2281 assertions, 3 empty in 3.3m"
    )
    payload["success"] = False
    payload["summary"]["failed"] = 2
    assert completion_message("run-tests", payload, False) == (
        "Tests failed. 468 subs, 2281 assertions, 2 failed, 3 empty in 3.3m"
    )


def test_compact_result_payload_strips_tests():
    payload = {
        "success": True,
        "summary": {"subs": 1},
        "tests": {"modFoo.Bar": {"status": "PASSED"}},
        "log_messages": ["."],
        "log_path": r"C:\log.log",
    }
    compact = compact_result_payload("run-tests", payload)
    assert "tests" not in compact
    assert "log_messages" not in compact
    assert compact["summary"] == {"subs": 1}
    assert compact["log_path"] == r"C:\log.log"
    assert compact_result_payload("rebuild-addin", payload)["tests"]


def test_progress_printer_marches_dots_and_names_slow(capsys):
    printer = ProgressPrinter(compact_tests=True)
    printer(1, None, "modFoo.A (1/4)")
    printer(2, None, "....\nPASS  modFoo.Slow  1.20s")
    printer(3, None, ".")
    printer.finish()
    captured = capsys.readouterr()
    assert captured.err == ""
    assert "modFoo.A (1/4)" not in captured.out
    assert captured.out == "....\nPASS  modFoo.Slow  1.20s\n.\n"


def test_progress_printer_finish_emits_last_fast_dot(capsys):
    printer = ProgressPrinter(compact_tests=True)
    printer(1, None, "...")
    printer.finish()
    assert capsys.readouterr().out == "...\n"


def test_main_run_tests_omits_tests_and_prints_summary(capsys):
    async def factory(name, arguments, on_progress):
        assert name == "vcs_run_tests"
        assert arguments["filter"] == "clsTestEncoding"
        await on_progress(1, None, "clsTestEncoding.TestUtf8 (1/3)")
        await on_progress(2, None, "...")
        return SimpleNamespace(
            content=[SimpleNamespace(text=json.dumps({
                "success": True,
                "durationMs": 350,
                "summary": {
                    "subs": 3,
                    "assertions": 12,
                    "passed": 12,
                    "failed": 0,
                    "errored": 0,
                    "empty": 0,
                },
                "tests": {"clsTestEncoding.TestUtf8": {"status": "PASSED"}},
                "log_path": r"C:\TestRun.log",
            }))],
            isError=False,
        )

    code = main(
        ["run-tests", r"C:\db.accda", "--filter", "clsTestEncoding"],
        session_factory=factory,
    )
    captured = capsys.readouterr()
    assert code == 0
    assert "clsTestEncoding.TestUtf8 (1/3)" not in captured.out
    assert captured.err == ""
    lines = [line for line in captured.out.splitlines() if line]
    assert lines[0] == "Starting run-tests..."
    assert lines[1] == "..."
    assert lines[-1] == "Tests passed. 3 subs, 12 assertions in 350ms"
    payload = json.loads("\n".join(lines[2:-1]))
    assert "tests" not in payload
    assert payload["summary"]["subs"] == 3
    assert payload["log_path"] == r"C:\TestRun.log"

