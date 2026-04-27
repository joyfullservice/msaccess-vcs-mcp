# AGENTS.md - AI Agent Guide to msaccess-vcs-mcp

## Purpose

This repository contains the **msaccess-vcs-mcp** server — a lightweight MCP (Model Context Protocol) bridge that lets AI agents drive the [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin) for version-controlling Microsoft Access databases.

## Development Environment Setup

**Always use the project virtual environment.** The package is installed in editable mode inside `venv/`. Do not install globally or suggest `pip install` outside the venv.

```powershell
# Activate the virtual environment (Windows PowerShell)
cd C:\Repos\msaccess-vcs-mcp
.\venv\Scripts\Activate.ps1

# If the venv doesn't exist yet, create it first:
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

### Running Tests

Tests must run inside the activated virtual environment:

```powershell
.\venv\Scripts\Activate.ps1

# Run all unit tests
pytest

# Run a specific test file
pytest tests/test_usage_logging.py -v

# Skip integration tests (require Access installed)
pytest -m "not integration"

# Run with coverage
pytest --cov=msaccess_vcs_mcp --cov-report=html
```

## Repository Structure

```
msaccess-vcs-mcp/
├── src/msaccess_vcs_mcp/       # Package source
│   ├── __init__.py             # Version (__version__)
│   ├── main.py                 # CLI entry point, startup sequence
│   ├── tools.py                # FastMCP instance + all @vcs_tool handlers
│   ├── config.py               # .env loading, get_config()
│   ├── usage_logging.py        # Structured JSONL usage logging
│   ├── security.py             # Path validation, write guards
│   ├── validation.py           # Startup validation helpers
│   ├── addin_integration.py    # COM calls to VCS add-in
│   ├── operation_manager.py    # Async operation queues
│   ├── callback_server.py      # HTTP callback server for VBA progress
│   └── access_com/             # Low-level COM/DAO helpers
├── tests/                      # pytest test suite
├── docs/                       # Extended documentation
├── .env.example                # Template for environment variables
├── pyproject.toml              # Package metadata & dependencies
└── AGENTS.md                   # This file
```

## Architecture

All tools are registered with the `@vcs_tool("name")` decorator in `tools.py`, which composes three concerns in order:

1. **Config reload** — `load_config()` re-reads `.env` when it changes
2. **Usage logging** — `with_logging(name)` records tool calls to JSONL
3. **MCP registration** — `mcp.tool()` exposes the function to MCP clients

```
AI Agent ──► MCP Server (Python) ──► VCS Add-in (VBA) ──► Access Database
                 │
           Path validation
           Permission checks
           Usage logging
           Async progress tracking
```

## Configuration

All settings come from environment variables (loaded from `.env` / `.env.local` in the project root). See `.env.example` for the full list.

Key variables:
- `ACCESS_VCS_DATABASE` — target database path
- `ACCESS_VCS_DISABLE_WRITES` — set `true` to block write operations
- `ACCESS_VCS_ENABLE_LOGGING` — set `true` to enable usage logging

## Logging

The server writes two parallel JSON Lines streams. Both filenames use the `vcs-mcp-` prefix so they don't collide with other tools that share the same logs directory.

### Diagnostic stream (`vcs-mcp-diagnostic.jsonl`) — always on

Captures server lifecycle events: `server_start`, `startup_env_load`, `lazy_env_load`, `lazy_init_started`, `lazy_init_skipped`, `list_roots_failed`, `list_roots_response`, `lazy_init_loaded`, `lazy_init_no_env_in_roots`, `usage_log_status`. Independent of `ACCESS_VCS_ENABLE_LOGGING` so it answers the "why didn't logging work?" question even when usage logging is silent.

- **Location:** `~/.msaccess-vcs-mcp/logs/vcs-mcp-diagnostic.jsonl`
- **Override:** `ACCESS_VCS_DIAGNOSTIC_LOG_DIR`
- **Opt out:** `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`
- **Rotation:** 1 MB per file, 3 backups
- **Discoverable from agents:** `vcs_get_version_info()` returns the active `diagnostic_log_path`.

### Usage stream (`vcs-mcp-usage.jsonl`) — default on

When `ACCESS_VCS_ENABLE_LOGGING=true` (the default), every tool call writes a structured entry. Set the env var to `false` to opt out.

- **Development installs:** logs to `{project_root}/logs/vcs-mcp-usage.jsonl`
- **Package installs:** logs to `~/.msaccess-vcs-mcp/logs/vcs-mcp-usage.jsonl`
- **Override:** `ACCESS_VCS_LOG_DIR`
- **Rotation:** `ACCESS_VCS_LOG_MAX_SIZE_MB` (default 10 MB), `ACCESS_VCS_LOG_BACKUP_COUNT` (default 5)

Each `tool_call` entry includes: `timestamp`, `version`, `event`, `tool`, `parameters`, `success`, `error`, `error_pattern`, `execution_time_ms`.

### Tiered audit posture

Three sensitivity tiers in the usage stream, each independently controlled:

1. **Audit metadata** — always written when `ENABLE_LOGGING=true`. Tool name, timing, success/error, sanitized parameter dict.
2. **Code-execution bodies** — `vcs_execute_sql`, `vcs_call_vba`, and `vcs_run_vba` write a `"code_execution"` event *before* execution begins. By default only `code_length` is recorded. Set `ACCESS_VCS_LOG_CODE_CONTENT=true` to record the full `code` field for forensic replay (off by default to limit business-data exposure). `code_length` lets analysts spot anomalies (e.g. "an agent ran a 4 KB VBA block") without seeing the body.
3. **Credential-shaped parameter keys** — any parameter whose name matches `password`, `secret`, `token`, `api_key`, `apikey`, `connection_string`, or `connectionstring` (case-insensitive) is replaced with `"<redacted>"` regardless of other switches. Defense in depth.

## Adding a New Tool

1. Write the handler function in `tools.py`
2. Decorate with `@vcs_tool("vcs_your_tool_name")` — this handles config reload, usage logging, and MCP registration automatically
3. Add tests in `tests/`

## Key Conventions

- All tool names use the `vcs_` prefix
- Tool handlers return `dict[str, Any]` with at least a `success` key
- Error results include `"error"` key (detected by usage logging)
- Async tools (`async def`) are supported by `@vcs_tool` transparently
- `Context` parameters from FastMCP are filtered out of usage logs automatically
