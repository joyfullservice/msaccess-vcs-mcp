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

## 2026-04-15 — Structured JSONL usage logging via composite decorator

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
