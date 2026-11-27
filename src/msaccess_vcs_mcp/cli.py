"""Terminal client for long Access operations with live MCP progress.

Cursor's chat UI does not reliably show MCP progress notifications. This
CLI launches the same MCP server over stdio, calls one tool, and prints
each ``notifications/progress`` line as it arrives -- the same live
feedback as running a Python or PowerShell script.

``configure`` is the exception: it reads and writes the user-level settings
file directly and starts no server.
"""

from __future__ import annotations

import argparse
import asyncio
import gc
import json
import os
import re
import sys
from collections.abc import Callable
from typing import Any

ProgressCallback = Callable[[float, float | None, str | None], Any]


def stdio_server_environment() -> dict[str, str]:
    """Environment for the CLI's one-shot child MCP server."""
    env = os.environ.copy()
    # CLI commands do not create session option overrides. Starting Access from
    # the child server's atexit hook merely to call EndSession can strand a hidden
    # .accda moniker process after an add-in rebuild.
    env["ACCESS_VCS_SKIP_SESSION_CLEANUP"] = "true"
    return env


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="msaccess-vcs",
        description=(
            "Run a long Access VCS operation with live progress on stdout. "
            "Uses the same MCP server implementation as the Cursor tools."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    export = sub.add_parser("export", help="Export a database to source files")
    export.add_argument("database_path")
    export.add_argument(
        "output_dir",
        nargs="?",
        default=None,
        help="Folder to export to instead of the database's export folder (add-in 5.2.0 or later)",
    )
    export.add_argument(
        "--full",
        action="store_true",
        help="Export every object, not only those marked changed",
    )

    merge = sub.add_parser("merge", help="Merge source files into a database")
    merge.add_argument("database_path")
    merge.add_argument("source_dir")

    rebuild_db = sub.add_parser(
        "rebuild-database",
        help="Build a fresh database from source",
    )
    rebuild_db.add_argument("source_dir")
    rebuild_db.add_argument(
        "output_path",
        nargs="?",
        default=None,
        help="File to build (default: the file the source files name)",
    )
    rebuild_db.add_argument("--template", dest="template_path", default=None)

    rebuild_addin = sub.add_parser(
        "rebuild-addin",
        help="Rebuild the VCS add-in from source and wait for install",
    )
    rebuild_addin.add_argument("source_dir")
    rebuild_addin.add_argument(
        "--timeout",
        dest="timeout_seconds",
        type=float,
        default=None,
        help="Watch timeout in seconds (default ACCESS_VCS_REBUILD_TIMEOUT_SEC)",
    )

    run_tests = sub.add_parser(
        "run-tests",
        help="Run VBA tests with live per-test progress",
    )
    run_tests.add_argument("database_path")
    run_tests.add_argument(
        "--filter",
        default=None,
        help="Comma-separated test filter (module, suite, procedure, or tag)",
    )
    run_tests.add_argument(
        "--timeout",
        dest="timeout_seconds",
        type=float,
        default=None,
        help="Wait timeout in seconds (default 10 minutes from the add-in)",
    )

    configure = sub.add_parser(
        "configure",
        help="Choose user-level settings such as screenshot retention",
        description=(
            "Ask about each setting that needs a decision and save the answers to "
            "~/.msaccess-vcs-mcp/settings.json. A value set in mcp.json or a .env "
            "file wins over the saved one."
        ),
    )
    configure.add_argument(
        "--show", action="store_true", help="Print each setting, its value, and its source"
    )
    configure.add_argument(
        "--set",
        dest="set_values",
        action="append",
        default=[],
        metavar="NAME=VALUE",
        help="Save a value without prompting, e.g. screenshot_retention=7, =never, =forever, or =default",
    )
    configure.add_argument(
        "--reset",
        dest="reset_names",
        action="append",
        default=[],
        metavar="NAME",
        help="Forget a saved decision so the setting is undecided again",
    )
    return parser


def arguments_for(args: argparse.Namespace) -> tuple[str, dict[str, Any]]:
    if args.command == "export":
        payload: dict[str, Any] = {
            "database_path": args.database_path,
            "output_dir": args.output_dir,
            "full_export": bool(args.full),
        }
        return "vcs_export_database", payload
    if args.command == "merge":
        return "vcs_import_objects", {
            "database_path": args.database_path,
            "source_dir": args.source_dir,
        }
    if args.command == "rebuild-database":
        payload = {"source_dir": args.source_dir}
        if args.output_path:
            payload["output_path"] = args.output_path
        if args.template_path:
            payload["template_path"] = args.template_path
        return "vcs_rebuild_database", payload
    if args.command == "rebuild-addin":
        payload = {"source_dir": args.source_dir}
        if args.timeout_seconds is not None:
            payload["timeout_seconds"] = args.timeout_seconds
        return "vcs_rebuild_addin", payload
    if args.command == "run-tests":
        payload = {"database_path": args.database_path}
        if args.filter:
            payload["filter"] = args.filter
        if args.timeout_seconds is not None:
            payload["timeout_seconds"] = args.timeout_seconds
        return "vcs_run_tests", payload
    raise ValueError(f"Unknown command: {args.command}")


def format_progress_line(
    progress: float,
    total: float | None,
    message: str | None,
) -> str:
    """Return the console line for one MCP progress notification.

    The monotonic ``progress`` value is a protocol requirement so MCP
    progress stays strictly increasing. VBA's phase-local ``current/total``
    already lives in ``message`` (for example ``clsDbTheme.cls (24/180)``).
    Print that message only; do not also prefix the protocol sequence.
    """
    del progress, total
    return (message or "").rstrip()


_COUNT_SUFFIX = re.compile(r"\(\d+(?:\.\d+)?/\d+(?:\.\d+)?\)\s*$")
_DOTS_ONLY = re.compile(r"^\.+$")
_DOTS_PER_LINE = 80


class ProgressPrinter:
    """Print MCP progress.

    For ``run-tests``, ``(n/m)`` progress is ignored (not shown). Batched
    ``....`` log callbacks print as pytest-style dots. A named line may share
    that callback (dots, newline, then PASS/FAIL/ERROR/EMPTY).
    """

    def __init__(self, *, compact_tests: bool = False) -> None:
        self.compact_tests = compact_tests
        self._dot_line_len = 0

    def __call__(
        self,
        progress: float,
        total: float | None,
        message: str | None,
        *,
        file=None,
    ) -> None:
        line = format_progress_line(progress, total, message)
        if not line:
            return
        out = file if file is not None else sys.stdout
        if self.compact_tests:
            self._print_test_stream(line, out)
            return
        print(line, file=out, flush=True)

    def _print_test_stream(self, line: str, out) -> None:
        for piece in line.splitlines():
            if not piece.strip():
                continue
            if _COUNT_SUFFIX.search(piece):
                continue
            stripped = piece.strip()
            if _DOTS_ONLY.match(stripped):
                for _ in stripped:
                    self._print_dot(out)
                continue
            self._end_dot_line(out)
            print(piece.rstrip(), file=out, flush=True)

    def _print_dot(self, out=None) -> None:
        dest = out if out is not None else sys.stdout
        print(".", end="", file=dest, flush=True)
        self._dot_line_len += 1
        if self._dot_line_len >= _DOTS_PER_LINE:
            print(file=dest, flush=True)
            self._dot_line_len = 0

    def _end_dot_line(self, out=None) -> None:
        dest = out if out is not None else sys.stdout
        if self._dot_line_len:
            print(file=dest, flush=True)
            self._dot_line_len = 0

    def finish(self, *, file=None) -> None:
        self._end_dot_line(file)

    def close_status(self, *, file=None) -> None:
        self.finish(file=file)


def print_progress(
    progress: float,
    total: float | None,
    message: str | None,
    *,
    file=None,
) -> None:
    line = format_progress_line(progress, total, message)
    if not line:
        return
    print(
        line,
        file=file if file is not None else sys.stdout,
        flush=True,
    )


def format_duration_ms(ms: Any) -> str:
    try:
        value = float(ms)
    except (TypeError, ValueError):
        return ""
    if value < 1000:
        return f"{int(value)}ms"
    if value < 60000:
        return f"{value / 1000:.2f}s"
    return f"{value / 60000:.1f}m"


def compact_result_payload(command: str, payload: Any) -> Any:
    """Drop bulky fields from CLI JSON. The MCP tool still returns the full map."""
    if command != "run-tests" or not isinstance(payload, dict):
        return payload
    compact = dict(payload)
    compact.pop("tests", None)
    compact.pop("log_messages", None)
    return compact


def completion_message(command: str, payload: Any, succeeded: bool) -> str:
    """Human-readable last line so a watcher does not have to parse JSON."""
    if command == "run-tests":
        return _test_completion_message(payload, succeeded)
    if command == "rebuild-addin":
        return "Rebuild complete." if succeeded else "Rebuild failed."
    if command == "rebuild-database":
        return "Database rebuilt." if succeeded else "Rebuild failed."
    if command == "export":
        return "Export complete." if succeeded else "Export failed."
    if command == "merge":
        return "Merge complete." if succeeded else "Merge failed."
    return "Done." if succeeded else "Failed."


def _test_completion_message(payload: Any, succeeded: bool) -> str:
    if not isinstance(payload, dict):
        return "Tests passed." if succeeded else "Tests failed."
    summary = payload.get("summary") or {}
    parts: list[str] = []
    subs = summary.get("subs")
    assertions = summary.get("assertions")
    if subs is not None:
        parts.append(f"{subs} subs")
    if assertions is not None:
        parts.append(f"{assertions} assertions")
    failed = summary.get("failed") or 0
    errored = summary.get("errored") or 0
    empty = summary.get("empty") or 0
    if failed:
        parts.append(f"{failed} failed")
    if errored:
        parts.append(f"{errored} errored")
    if empty:
        parts.append(f"{empty} empty")
    elapsed = format_duration_ms(payload.get("durationMs"))
    verb = "passed" if succeeded else "failed"
    if not parts:
        return f"Tests {verb}."
    body = ", ".join(parts)
    if elapsed:
        return f"Tests {verb}. {body} in {elapsed}"
    return f"Tests {verb}. {body}"


def _tool_result_text(result: Any) -> str:
    parts: list[str] = []
    for item in getattr(result, "content", None) or []:
        text = getattr(item, "text", None)
        if text:
            parts.append(text)
    if parts:
        return "\n".join(parts)
    structured = getattr(result, "structuredContent", None) or getattr(
        result, "data", None
    )
    if structured is not None:
        return json.dumps(structured, indent=2)
    return str(result)


def result_succeeded(payload: Any) -> bool:
    if isinstance(payload, dict):
        if payload.get("success") is False or payload.get("isError") is True:
            return False
        if payload.get("success") is True:
            return True
        if payload.get("status") == "complete":
            return True
        if payload.get("error") or payload.get("cancelled"):
            return False
    if getattr(payload, "isError", False):
        return False
    return True


def parse_result_payload(text: str) -> Any:
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return text


async def run_mcp_tool(
    name: str,
    arguments: dict[str, Any],
    *,
    on_progress: ProgressCallback | None = None,
    session_factory: Callable[..., Any] | None = None,
) -> Any:
    """Call one MCP tool, forwarding progress. ``session_factory`` is for tests."""
    async def _progress(
        progress: float,
        total: float | None,
        message: str | None,
    ) -> None:
        if on_progress is not None:
            maybe = on_progress(progress, total, message)
            if asyncio.iscoroutine(maybe):
                await maybe

    if session_factory is not None:
        return await session_factory(name, arguments, _progress)

    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "msaccess_vcs_mcp"],
        env=stdio_server_environment(),
    )

    with open(os.devnull, "w", encoding="utf-8") as server_stderr:
        async with stdio_client(params, errlog=server_stderr) as streams:
            async with ClientSession(streams[0], streams[1]) as session:
                await session.initialize()
                result = await session.call_tool(
                    name,
                    arguments,
                    progress_callback=_progress,
                )
        # Python 3.14's proactor loop can otherwise defer the subprocess pipe
        # transport finalizer until after asyncio.run() closes the loop. Drop the
        # context-owned references and collect while the loop can still service
        # close callbacks.
        del session, streams
    await asyncio.sleep(0)
    gc.collect()
    await asyncio.sleep(0)
    return result


def startup_message(command: str) -> str:
    """Immediate stdout line before the child MCP server exists.

    Progress notifications cannot arrive until that process starts and
    the tool emits, which is a few seconds of silence without this.
    """
    return f"Starting {command}..."


_DECIDED_BY_TEXT = {
    "form": "in a form",
    "agent": "through an agent",
    "cli": "with msaccess-vcs configure",
}


def print_settings(out=None) -> None:
    from . import user_settings

    out = out if out is not None else sys.stdout
    print(f"Settings file: {user_settings.settings_path()}", file=out)
    for item in user_settings.describe_all():
        if item["source"] == user_settings.SOURCE_ENV:
            origin = f"set by {item['env_var']} in the environment"
        elif item["decided"]:
            origin = "chosen " + _DECIDED_BY_TEXT.get(item["decided_by"], "earlier")
            if item["source"] == user_settings.SOURCE_DEFAULT:
                origin += ", follows the default"
        else:
            origin = "not chosen yet, using the default"
        print(f"  {item['name']} = {item['display']}  ({origin})", file=out)


def run_configure(args: argparse.Namespace, *, input_fn=input, is_tty=None, out=None) -> int:
    """``msaccess-vcs configure``. Talks to the settings file, not to a server."""
    from . import user_settings

    out = out if out is not None else sys.stdout
    if args.set_values or args.reset_names:
        try:
            decisions: dict[str, str] = {}
            for item in args.set_values:
                name, sep, value = item.partition("=")
                if not sep:
                    raise ValueError(f"--set needs NAME=VALUE, not '{item}'.")
                user_settings.normalize(user_settings.find_setting(name), value)
                decisions[name] = value
            for name in args.reset_names:
                user_settings.find_setting(name)
            for name in args.reset_names:
                user_settings.reset(name)
            user_settings.save(decisions, "cli")
        except (KeyError, ValueError) as exc:
            print(str(exc).strip("'\""), file=sys.stderr)
            return 2
        print_settings(out)
        return 0
    if args.show:
        print_settings(out)
        return 0
    if not (sys.stdin.isatty() if is_tty is None else is_tty):
        print(
            "msaccess-vcs configure asks questions and needs a terminal. "
            "Use --show to list settings, or --set NAME=VALUE to save one.",
            file=sys.stderr,
        )
        return 2

    try:
        for setting in user_settings.DECISION_SETTINGS:
            item = user_settings.describe(setting)
            print(f"\n{setting.title}: {setting.description}", file=out)
            if item["source"] == user_settings.SOURCE_ENV:
                print(f"  Set to {item['display']} by {setting.key} in the environment; "
                      "change it there.", file=out)
                continue
            keep = (f"Enter keeps {item['display']}" if item["decided"]
                    else f"Enter uses the default, {setting.display(setting.default)}")
            while True:
                answer = input_fn(f"  {setting.prompt} [{keep}]: ").strip()
                if not answer:
                    if item["decided"]:
                        break
                    answer = user_settings.DEFAULT_CHOICE
                try:
                    user_settings.save({setting.name: answer}, "cli")
                    break
                except ValueError as exc:
                    print(f"  {exc}", file=out)
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.", file=sys.stderr)
        return 130
    print(file=out)
    print_settings(out)
    return 0


def main(argv: list[str] | None = None, *, session_factory=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "configure":
        return run_configure(args)
    name, arguments = arguments_for(args)
    print(startup_message(args.command), flush=True)
    printer = ProgressPrinter(compact_tests=args.command == "run-tests")

    try:
        result = asyncio.run(
            run_mcp_tool(
                name,
                arguments,
                on_progress=printer,
                session_factory=session_factory,
            )
        )
    except KeyboardInterrupt:
        printer.finish()
        print("Cancelled.", file=sys.stderr, flush=True)
        return 130

    printer.finish()
    text = _tool_result_text(result)
    payload = parse_result_payload(text)
    display = compact_result_payload(args.command, payload)
    if isinstance(display, (dict, list)):
        print(json.dumps(display, indent=2), flush=True)
    else:
        print(text, flush=True)

    tool_error = bool(getattr(result, "isError", False))
    succeeded = (not tool_error) and result_succeeded(payload)
    print(completion_message(args.command, payload, succeeded), flush=True)
    return 0 if succeeded else 1


if __name__ == "__main__":
    sys.exit(main())
