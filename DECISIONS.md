<!-- BEGIN HEADER -->

# Decision Log

A reverse-chronological journal of architectural and strategic decisions.
Maintained by AI coding agents (and human developers) at the end of working
sessions. Each entry captures what was decided, what alternatives were
considered, and why — so future contributors never revisit dead ends or lose
context on trade-offs already evaluated.

Agents: read this file before working on any module referenced here.

### When to log

Log decisions that constrain future design, involved genuine alternatives,
or would be non-obvious to a future contributor. A good litmus test: does
the "What this rules out" section have something meaningful to say?

Do NOT log: bug fixes with obvious solutions, test-only refactors,
documentation updates, or minor config tweaks that don't affect
architecture.

### Entry format

Insert new entries directly below this header, newest first. Do not modify
or reorder existing entries except to add supersession notes (see below).
If a session produced multiple independent decisions, create a separate
entry for each.

```
---

## YYYY-MM-DD — [Short descriptive title]

**Trigger**: What problem, requirement, or situation prompted this work.

**Options explored**:
- For each option, name the approach, its strengths, and why it was or
  wasn't chosen. Include options that were tried and reverted.

**Decision**: What was chosen and the core trade-off.

**What this rules out**: Future directions now constrained or foreclosed.
What would trigger revisiting this decision.

**Relevant files**: Key files created or modified.
```

### Guidelines

- Focus on **why**, not what. The diff shows what changed; this log
  explains the reasoning.
- Capture rejected alternatives with equal care. Future agents need to
  know what was already tried.
- Be specific — name libraries, files, config choices, error messages.
- Aim for 10–50 lines per entry. Reference document, not narrative.
- Plain language. No jargon, no editorializing, no padding.

### Superseded entries

When a new decision invalidates, corrects, or replaces guidance in an older
entry, add a blockquote annotation to the affected older entry — do not
rewrite or delete its original text. Place the note immediately after the
entry's heading or after the paragraph containing the superseded claim.

> **⚠ Superseded** (YYYY-MM-DD): [Brief explanation of what changed and
> why.] See "[title of newer entry]" above.

Use **⚠ Partially superseded** when only specific claims are affected, and
**⚠ Superseded** when the entire entry's premise or decision has been
overturned. Always scan older entries for claims that conflict with the new
decision — agents reading the log linearly will otherwise encounter
contradictory guidance.
<!-- END HEADER -->

---

## 2026-04-25 — Drop silent `Application.Eval` fallback in `_call_addin_function`

**Trigger**: Investigating failing add-in integration tests surfaced a long-standing usability and troubleshooting problem: when `Application.Run("…\Version Control.API", funcName)` failed twice in a row, `_call_addin_function` would silently fall through to `self._app.Eval("CallVcsApi(\"" & funcName & "\")")`, using a hard-coded default wrapper name. In any environment where the user had not implemented a `CallVcsApi` VBA function in the open database, this branch would still succeed against unit-test mocks and against any COM object whose `Eval` happened to return a value, producing false-success results for `vcs_export_*` and friends. The original `Run` error was buried in a three-layer retry/fallback message that obscured root cause during troubleshooting.

**Options explored**:
- **Keep current behavior, update tests to mock `Eval` raising**. Lowest-friction, but bakes the silent-success failure mode into production indefinitely.
- **Make the fallback opt-in via `ACCESS_VCS_API_WRAPPER`**. Removes the silent-success default but adds a configuration knob plus a code path almost no one will exercise; documenting *when* to set it requires explaining a real Access COM bug that most users will never hit.
- **Remove the fallback entirely (chosen)**. Single behavioral contract: if `Application.Run` fails twice, a `RuntimeError` is raised with the underlying COM error verbatim. Simpler call graph, fewer places to look during incident triage.

**Decision**: Removed the `Application.Eval` fallback and the `ACCESS_VCS_API_WRAPPER` env-var hook. The single `Run` retry is preserved -- it covers the documented Access first-call add-in load behavior -- but a second failure now surfaces directly as `RuntimeError("Failed to call add-in function '<name>': <error>")`. The decision was driven by simplification of code and of troubleshooting, not by performance or correctness in any narrow sense.

**What this rules out**: The MCP server will no longer auto-recover from genuine `Application.Run`-from-COM bugs by routing through a per-database VBA wrapper. Users who actually need that workaround must reintroduce it deliberately -- ideally as an explicit opt-in setting with a clear error message when it isn't configured -- not as a hidden default. Re-adding any silent fallback that swallows a documented error path should be rejected on review.

**Relevant files**: `src/msaccess_vcs_mcp/addin_integration.py` (`_call_addin_function`).

---

## 2026-04-25 — Multi-strategy `.env` discovery with workspace-roots lazy init

**Trigger**: A user reported that `ACCESS_VCS_ENABLE_LOGGING=true` in a client project's `.env` had no effect when `msaccess-vcs-mcp` was used from another project. Root cause: `_find_project_root()` only walked up from CWD or the installed package location. When the server was launched from a user-level `mcp.json` (`~/.cursor/mcp.json`) and CWD wasn't the project root, the upward walk failed, and the package-location fallback could match `msaccess-vcs-mcp`'s own `pyproject.toml` instead of the user's project. Compounding this, `_load_env_files()` always called `load_dotenv(..., override=False)`, so even a successful reload would silently fail to apply edited values. The sibling `db-inspector-mcp` project had already solved this with a layered resolution strategy.

**Options explored**:
- **Single env-var override (`ACCESS_VCS_PROJECT_DIR`) only**. Simple, works, but requires per-project user configuration. Doesn't help users who configure the server once at the user level and expect it to "just work" across projects.
- **MCP workspace-roots discovery only** (`ctx.session.list_roots()`). Automatic, no user config required. But only works for tools that accept a `Context` parameter and run async — most existing tools are sync without `ctx`.
- **Layered resolution: env-var → workspace roots → CWD walk → package walk → fallback (chosen)**. Mirrors `db-inspector-mcp`'s proven order. Each strategy compensates for the others' blind spots: explicit env-var for power users, workspace roots for IDE-launched servers, CWD walk for terminal-launched servers, package walk as a last-resort dev-install fallback.
- **Eager workspace-roots probe at startup**. Cleaner than lazy init, but FastMCP's `Context` is not available before the first tool call -- the MCP protocol handshake hasn't completed yet.

**Decision**: Ported `db-inspector-mcp`'s discovery machinery in full, with `ACCESS_VCS_` prefix substitution and three project-specific adaptations:

1. **Resolution-method tracking**. `_find_project_root()` and `initialize_from_workspace()` set a module-level `_project_root_method` (one of `RESOLUTION_PROJECT_DIR_ENV`, `RESOLUTION_WORKSPACE_ROOTS`, `RESOLUTION_CWD_ENV`, `RESOLUTION_CWD_MARKER`, `RESOLUTION_PACKAGE_ENV`, `RESOLUTION_PACKAGE_MARKER`, `RESOLUTION_CWD_FALLBACK`). Surfaced via `get_project_root_info()`, printed to stderr ("Resolved project root: X (via Y)"), and embedded in the `logging_initialized` JSONL event so users can audit *which* mechanism actually populated their config in any given session.
2. **mtime-based hot-reload with `override=True` on reload**. `_check_env_reload()` compares stored `.env`/`.env.local` mtimes against current values; when changed, the next `load_config()` triggers a reload using `override=True` so edited values actually replace old ones. Logging is reset only when a reload was detected (was previously reset on every `load_config()` call -- wasteful).
3. **Lazy init wired through the `vcs_tool` decorator, not individual tools**. The decorator inspects the wrapped handler's signature; for async handlers that accept `ctx: Context`, it calls `await _ensure_env_loaded(ctx)` before `load_config()`. Converted `vcs_get_version_info`, `vcs_list_objects`, `vcs_diff_database`, `vcs_import_objects`, and `vcs_rebuild_database` to async with optional `ctx` so the workspace-roots path triggers on the agent's first realistic call regardless of which tool that is.

**What this rules out**: Sync tools without `ctx` cannot trigger workspace-roots lazy init -- but this is fine because once any async-with-ctx tool runs, the project env is loaded for all subsequent calls (sync included). If a future agent goes straight to a sync tool first (e.g., `vcs_execute_sql` before any read tool), workspace-roots discovery won't fire and the user must rely on `ACCESS_VCS_PROJECT_DIR` or CWD-based discovery; converting more tools to async is the escape hatch. The `RESOLUTION_*` constants are part of an implicit public contract -- renaming them would silently break any log-analysis tooling that filters on `project_root_resolution`. The package-walk fallback can still match `msaccess-vcs-mcp`'s own dev tree when the CWD walk finds nothing; this is intentional for development installs and is the lowest-priority strategy.

**Relevant files**: `src/msaccess_vcs_mcp/config.py` (resolution-method tracking, `_check_env_reload`, `_load_env_from_directory`, `initialize_from_workspace`, `get_project_root_info`), `src/msaccess_vcs_mcp/tools.py` (`_lazy_init_attempted`, `_file_uri_to_path`, `_ensure_env_loaded`, `vcs_tool` decorator, 5 tool signatures), `src/msaccess_vcs_mcp/usage_logging.py` (`logging_initialized` event includes `project_root` + `project_root_resolution`), `.env.example`, `README.md` (new "Using msaccess-vcs-mcp from Another Project" section), `tests/test_config_env_loading.py` (17 tests), `tests/test_lazy_workspace_init.py` (9 tests).

---

## 2026-04-15 — Pre-execution audit logging for code execution tools

**Trigger**: The existing usage logging captures tool parameters but truncates all strings to 500 characters and only writes after execution completes. For the three tools that execute arbitrary code against databases (`vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), this leaves two gaps: (1) a complex SQL query or VBA code block may be truncated beyond usefulness in a forensic review, and (2) if the process crashes during execution, no record of what was attempted exists.

**Options explored**:
- **Raise the truncation limit globally**. Simple, but inflates every log entry (file paths, option names, etc.) unnecessarily. Rotation would trigger sooner.
- **Exempt specific parameter names from truncation** (e.g., `sql`, `code`). Mixes audit concerns into the general sanitization logic. Hard to extend cleanly.
- **Dedicated `log_code_execution()` function with a separate event type (chosen)**. Writes a `"code_execution"` event *before* execution begins, with the full untruncated code/SQL and the target database path. The existing `with_logging` decorator continues to write the post-execution `"tool_call"` event with truncated parameters, success/error, and timing. Two complementary records: the audit trail (what was attempted) and the outcome (what happened).

**Decision**: Added `log_code_execution(tool_name, database_path, code, code_type)` to `usage_logging.py`. Called from `vcs_execute_sql` (code_type=`"sql"`), `vcs_run_vba` (code_type=`"vba"`), and `vcs_call_vba` (code_type=`"vba_call"`) immediately after path validation but before any COM/database interaction. The `code` field is never truncated. The event goes to the same `usage.jsonl` file — no separate audit file — distinguished by `"event": "code_execution"`.

**What this rules out**: Code execution entries have no upper size limit on the `code` field. In practice, VBA code blocks and SQL queries are small (under 10 KB). If an agent somehow generates megabyte-scale code strings, log rotation handles it, but this is not a realistic concern. If a separate audit file is ever wanted (e.g., for compliance), the `log_code_execution` function could be retargeted without changing call sites.

**Relevant files**: `usage_logging.py` (`log_code_execution`), `tools.py` (call sites in `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), `tests/test_usage_logging.py` (5 new tests).

---

## 2026-04-15 — Agents cannot enable McpAllowRunVBA programmatically

**Trigger**: The `vcs_set_option` tool allowed agents to set any VCS option, including `McpAllowRunVBA` which gates arbitrary VBA code execution via `vcs_run_vba`. An agent could autonomously enable this option and then run arbitrary code without user awareness, undermining the security boundary that `McpAllowRunVBA` was designed to provide.

**Options explored**:
- **No guard, rely on default-off**. `McpAllowRunVBA` defaults to False, but nothing prevented an agent from calling `vcs_set_option("db.accdb", "McpAllowRunVBA", True)` as its first action. The docstrings even showed this as an example.
- **Server-side blocklist in `vcs_set_option` (chosen)**. A case-insensitive check against a set of protected option names. Returns a descriptive error directing the user to enable the option manually via the VCS Options form.
- **VBA-side enforcement**. Have the add-in's `SetOption` method refuse `McpAllowRunVBA` when called from MCP. Harder to implement since VBA doesn't know the calling context, and the error would be less clear.

**Decision**: `vcs_set_option` blocks setting `McpAllowRunVBA` with a clear error message. The option requires explicit user consent via the VCS Options form in Access. Docstrings and README updated to stop suggesting agents can self-enable this option.

**What this rules out**: Agents cannot autonomously escalate to arbitrary VBA execution. If future protected options emerge (e.g., a hypothetical `McpAllowDDL`), add them to the `PROTECTED_OPTIONS` set in `vcs_set_option`.

**Relevant files**: `tools.py` (`vcs_set_option`), `README.md` (security section).

---

## 2026-04-15 — Object type normalization lives in VBA, not Python

**Trigger**: `vcs_export_object` and `vcs_import_object` only supported 6 core Access object types (query, form, report, module, table, macro). The add-in's `eDatabaseComponentType` enum defines 24+ types (relations, IMEX specs, VBE project, themes, etc.) that couldn't be exported individually. Additionally, the MCP tools used plural strings (`"queries"`) in `vcs_export_database` but singular (`"query"`) in `vcs_export_object`, creating inconsistency that confused AI agents.

**Options explored**:
- **Python-side normalization map**. A `normalize_object_type()` helper in the MCP server that maps plural/alias forms to canonical singular before passing to VBA. Only benefits MCP callers. Creates a second type map to maintain alongside VBA.
- **VBA-side normalization via `ResolveComponentType` (chosen)**. A `Select Case` function in `modContainers.bas` that accepts singular, plural, and alias forms (50+ strings) and maps to `eDatabaseComponentType`. Benefits all callers — MCP tools, direct `Application.Run` API calls, any future integration. Python becomes a transparent pass-through.
- **Accept both in Python AND VBA**. Redundant and creates maintenance burden keeping two maps in sync.

**Decision**: Type normalization lives entirely in VBA's `ResolveComponentType`. Python passes `object_type` strings through to VBA without validation. VBA returns structured error JSON for unrecognized types. `ExportObject` and `ImportObject` on `clsVersionControl` were extended to handle all 24 component types: core AccessObject types use the existing `ExportSingleObject` path; non-core types use `GetComponentClass` + `GetAllFromDB`. Single-file types (like `vbe_project`) don't require an `object_name` parameter.

**What this rules out**: Adding new component types requires updating VBA's `ResolveComponentType` — the Python MCP layer does not need changes. If a Python-only consumer needs early validation without a COM roundtrip, they would need to maintain their own type list, but this is unlikely since the VBA error response is fast and descriptive.

**Relevant files**: `modContainers.bas` (`ResolveComponentType`), `clsVersionControl.cls` (`ExportObject`, `ImportObject` rewritten), `tools.py` (docstrings updated, `object_name` made optional), `addin_integration.py` (no type-related changes).

---

## 2026-04-15 — Structured JSONL usage logging via composite decorator

> **⚠ Partially superseded** (2026-04-25): The "What this rules out" note about adopting `db-inspector-mcp`'s mtime-based hot-reload pattern "if performance becomes a concern" is now fact -- adopted for correctness rather than performance (the original `override=False` reload was silently failing to apply edits). The `logging_initialized` event also now includes `project_root` and `project_root_resolution` fields. See "Multi-strategy `.env` discovery with workspace-roots lazy init" above.

**Trigger**: Need to troubleshoot and evaluate how AI agents use the MCP tools in practice — which tools are called, with what parameters, how often they fail, and how long they take. The `db-inspector-mcp` sibling project already has a proven logging implementation that was requested as the reference pattern.

**Options explored**:
- **Python stdlib `logging` module**. Standard approach, but produces unstructured text. Not suitable for programmatic analysis of tool call patterns.
- **FastMCP middleware / hooks**. FastMCP doesn't expose a per-tool middleware layer. Would require monkey-patching internals.
- **Composite decorator with JSONL file logging (chosen)**. Follows the exact pattern from `db-inspector-mcp`: a `with_logging(name)` decorator that wraps each tool, writing one JSON object per line to a rotating file. A `vcs_tool("name")` composite decorator chains config reload → usage logging → `mcp.tool()` registration, replacing bare `@mcp.tool()` on all 17 tools. Controlled by `ACCESS_VCS_ENABLE_LOGGING` env var (default: off).

**Decision**: Adopted the `db-inspector-mcp` pattern with project-specific adaptations:
- Env var prefix changed from `DB_MCP_` to `ACCESS_VCS_` for consistency with existing config.
- Removed `database`/`dialect` fields from log entries (VCS tools pass `database_path` as a regular parameter, so it's captured in `parameters` automatically).
- Error pattern categories tailored to VCS-specific errors (COM errors, add-in errors, VBA compile errors, write-disabled, database busy) instead of SQL-specific patterns.
- `Context` objects from FastMCP are filtered out of logged parameters (not serializable).
- Lazy initialization: disabled state is not cached (`_logging_enabled` stays `None`) so the first tool call after `load_config()` populates the env can still enable logging. Failure state _is_ cached to avoid retry spam.

**What this rules out**: Log entries do not include tool return values — only parameters, success/failure, errors, and timing. If result logging is needed later, the `with_logging` decorator already receives the result for serialization checking and could be extended. The `vcs_tool` decorator calls `load_config()` on every tool invocation (one `stat()` call); if this becomes a performance concern, the hot-reload pattern from `db-inspector-mcp` (mtime-based gating) could be adopted.

**Relevant files**: `src/msaccess_vcs_mcp/usage_logging.py` (new), `src/msaccess_vcs_mcp/tools.py` (`vcs_tool` decorator, all 17 tools migrated), `src/msaccess_vcs_mcp/config.py` (logging env vars + `reset_logging` on reload), `src/msaccess_vcs_mcp/main.py` (startup status), `.env.example`, `tests/test_usage_logging.py` (36 tests), `AGENTS.md` (new), `.gitignore` (`logs/`).

---

## 2026-04-14 — Tool naming: `vcs_*` prefix

**Trigger**: Tools were originally named `access_*` (e.g., `access_export_database`). A separate MCP server for Access (`MCP-Access` by bclothier) also uses `access_*` prefixed tools. Both servers might be loaded in the same agent session. Additionally, these tools control the VCS add-in, not the Access application itself, so `access_*` was a misnomer.

**Options explored**:
- **Keep `access_*`**. Familiar, but inaccurate and collides with MCP-Access.
- **Use `vcs_*` (chosen)**. Accurately reflects these tools control the VCS add-in. No namespace collision.
- **Use `msaccess_vcs_*`**. Unambiguous but verbose — wastes tokens on every tool call.

**Decision**: All 9 existing tools renamed from `access_*` to `vcs_*`. All 8 new tools use `vcs_*`. Updated across `tools.py`, `README.md`, `validation.py`, and all docs.

**What this rules out**: Any external references to `access_export_database` etc. break. Acceptable since the server is pre-release with no external consumers.

**Relevant files**: `tools.py`, `README.md`, `validation.py`, `docs/*.md`.

---

## 2026-04-14 — Eight new tools for per-object development workflow

> **⚠ Partially superseded** (2026-04-15): `vcs_export_object` and `vcs_import_object` now support all 24 component types (not just the original 6 core types). Type normalization moved to VBA. See "Object type normalization lives in VBA, not Python" above.

**Trigger**: All existing tools operated at the whole-database level. Agents had no way to export/import a single object, execute SQL, run VBA, or control add-in options — capabilities essential for the tight edit-import-compile-test loop needed during add-in development and general database development.

**Options explored**:
- **Extend existing tools with filters** (e.g., `vcs_export_database` with `object_name` parameter). Conflates bulk and single-object semantics. The existing tools have async/callback infrastructure not needed for quick per-object calls.
- **Build intelligence into the MCP server** (Python-side VBE manipulation, SQL execution via separate ODBC connection). Creates tight coupling to Access internals in Python, duplicates logic better handled in VBA, and opens a second database connection causing file-locking conflicts.
- **Thin MCP tools that delegate to add-in API methods (chosen)**. Each new tool calls a corresponding method on `clsVersionControl` via the existing `API()` dispatcher. The add-in handles all Access interaction. The MCP layer just validates paths, parses JSON results, and formats responses.

**Decision**: 8 new tools added: `vcs_export_object`, `vcs_import_object`, `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`, `vcs_set_option`, `vcs_get_option`, `vcs_get_log`. Total: 17 tools. Each delegates to a public method on `clsVersionControl` in the add-in. The MCP server remains a lightweight wrapper — all business logic lives in VBA.

**What this rules out**: The MCP server does not do database introspection, schema analysis, or complex SQL. Those capabilities stay in `db-inspector-mcp`. If the VCS MCP needs to support operations not expressible through `clsVersionControl` API methods, the add-in must be extended first. This is intentional — it keeps the MCP layer thin and ensures all consumers of the add-in API get the same capabilities.

**Relevant files**: `tools.py` (8 new tool definitions), `addin_integration.py` (`call_sync` used by all new tools).

---

## 2026-04-14 — SQL execution via add-in DAO connection, not separate ODBC

**Trigger**: Usage logs from `db-inspector-mcp` showed 67% of all calls were just running SELECT queries (`db_count_query_results`, `db_preview`). Agents frequently need to inspect `MSysObjects`, `MSysQueries`, table data, and query results. Requiring a second MCP server for this basic need adds configuration overhead and creates a second connection to the same Access file (risking file-locking conflicts).

**Options explored**:
- **Keep SQL execution in db-inspector-mcp only**. Clean separation, but requires agents to have both MCPs configured. Two connections to the same `.accdb` file can cause locking issues. Extra overhead for the dominant use case.
- **Add ODBC connection in the VCS MCP server** (Python-side). Avoids VBA roundtrip but opens a second connection. Would need to handle Access SQL dialect quirks in Python.
- **Route through add-in's existing DAO connection (chosen)**. `ExecuteSQL` method on `clsVersionControl` uses `CurrentDb.OpenRecordset` — the same connection the add-in already holds. No file-locking conflict. Access SQL dialect handled natively. Read-only (SELECT only, enforced in VBA).

**Decision**: `vcs_execute_sql` tool calls the add-in's `ExecuteSQL` API method, which runs the query via `CurrentDb.OpenRecordset`, serializes results as JSON, and returns them. Non-SELECT statements are rejected. Results capped at `max_rows` (default 100). The db-inspector MCP remains available for cross-database comparison and heavy analytical work.

**What this rules out**: No write queries (INSERT/UPDATE/DELETE/DDL) through this tool. If agents need to modify data, they use `vcs_run_vba` or `vcs_call_vba` with appropriate VBA code. The SQL validation is simple (checks for `SELECT` prefix) — a determined agent could bypass it via `vcs_run_vba`, which is why `McpAllowRunVBA` defaults to off.

**Relevant files**: `tools.py` (`vcs_execute_sql`), `clsVersionControl.cls` (`ExecuteSQL`).

---

## 2026-04-14 — Two VBA execution tools with distinct roles

**Trigger**: Agents need to execute VBA code for testing and debugging. Two distinct use cases emerged: calling existing functions by name (safe, predictable) and executing arbitrary agent-generated code (powerful, risky). These have fundamentally different security profiles.

**Options explored**:
- **Single `vcs_run_vba` tool for both**. Agent passes either a function name or a code block, tool detects which. Blurs the security boundary — how do you gate "arbitrary code" while allowing "call existing function"?
- **Two tools with distinct roles (chosen)**. `vcs_call_vba` calls existing named functions via `Application.Run` — no temp module, no compilation, lower risk, separate permission (`McpAllowCallVBA`, default: True). `vcs_run_vba` executes agent-generated code via temp module lifecycle — compilation check, error capture, cleanup, higher risk, separate permission (`McpAllowRunVBA`, default: False).

**Decision**: `vcs_call_vba(database, function_name, args)` for existing functions; `vcs_run_vba(database, code)` for ad-hoc code. The add-in handles `RunVBA`'s full lifecycle (create, compile, execute, capture, cleanup). Error capture in `RunVBA` uses module-level variables with accessor functions rather than embedded JSON string construction in generated code — cleaner and avoids VBA quote-escaping nightmares.

**What this rules out**: `vcs_call_vba` is limited to public functions callable via `Application.Run` (max 3 args in current implementation). Private functions or functions requiring object parameters can't be called directly — use `vcs_run_vba` for those. If the 3-arg limit becomes a problem, the `InvokeTypes` + `pythoncom.Missing` padding pattern (used by MCP-Access) would support up to 30 args.

**Relevant files**: `tools.py` (`vcs_call_vba`, `vcs_run_vba`), `clsVersionControl.cls` (`RunVBA`).

---

## 2026-04-15 — Session-scoped option overrides for MCP/API callers

**Trigger**: `vcs_set_option` changes were silently discarded because the add-in reloads options from `vcs-options.json` at the start of every operation. The agent's overrides never persisted past the first subsequent export/build.

**Decision**: The MCP server generates a session ID at startup (`uuid4().hex[:8]`), registers it with the add-in via `RegisterSession`, and `vcs_set_option` now writes overrides to a session-scoped file (`mcp/options-{session_id}.json`) in the export folder. The add-in's operation entry points load these overrides after `LoadProjectOptions` when `Operation.Source` is API/MCP. On shutdown, `atexit` calls `EndSession` to clean up the file. A `vcs_end_session` tool is also available for explicit mid-session cleanup.

**What this rules out**: Session IDs don't persist across server restarts — the agent must re-set options if the server restarts. Stale override files are auto-cleaned after 30 days on the add-in side.

**Relevant files**: `tools.py` (`vcs_set_option`, `vcs_end_session`), `main.py` (session ID, atexit), `config.py` (`get_session_id`). Add-in side: see `DECISIONS.md` in `msaccess-vcs-addin`.
