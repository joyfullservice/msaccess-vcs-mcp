# Logs

## Operation logs (written by the add-in)

Each export, import, build, or test run writes a log to `<source folder>/logs/<Base>_<timestamp>.log`. `<Base>` follows the operation, not the tool: `Export`, `Merge` (both import tools), `Build`, `TestRun`, or `Other`.

- Every operation tool returns `log_path` for the run it just performed. Read that file directly.
- On failure, the result also carries `log_excerpt`, the tail of the log, so the error is usually already in the result.
- If the path is no longer at hand, call `vcs_get_log(database_path, log_type="Merge")` with the matching base name.

The add-in writes a `.gitignore` that excludes `logs/` and `*.log`, so search tools that honour `.gitignore` (including Cursor's file search) skip these files without saying so. Use the returned path, or list the folder from a shell.

## Server logs

The server writes two JSON Lines files, both prefixed `vcs-mcp-`:

- `vcs-mcp-usage.jsonl`: one entry per tool call (tool, parameters, success, `error`, `error_pattern`, duration). On by default; `ACCESS_VCS_ENABLE_LOGGING=false` turns it off. SQL and VBA bodies are recorded as their length only unless `ACCESS_VCS_LOG_CODE_CONTENT=true`.
- `vcs-mcp-diagnostic.jsonl`: server start-up and configuration events. Always on unless `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`.

`vcs_get_version_info()` returns both paths. `vcs_get_recent_calls()` reads the latest usage entries, which is the way to learn what ran after a client timeout.
