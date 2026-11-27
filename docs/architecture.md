# Architecture

How a tool call travels from the MCP client to Access, and where each concern lives. Read this before adding a module or changing the `vcs_tool` wrapper. The steps for adding a tool are in [AGENTS.md](../AGENTS.md#adding-a-tool).

## How a tool call flows

```
AI Agent ──► MCP Server (Python) ──► VCS Add-in (VBA) ──► Access Database
                 │
           Path validation
           Permission checks
           Usage logging
           Async progress tracking
```

All tools are registered with the `@vcs_tool("name")` decorator in `tools.py`, which composes three concerns in order:

1. **Config reload**: `load_config()` re-reads `.env` when it changes ([configuration.md](configuration.md)).
2. **Usage logging**: `with_logging(name)` records tool calls to JSONL ([logging.md](logging.md)).
3. **MCP registration**: `mcp.tool()` exposes the function to MCP clients.

Inside the wrapper, `_refuse_installed_addin_target` runs before anything else touches Access ([addin-rebuild-and-tests.md](addin-rebuild-and-tests.md#the-installed-add-in-is-never-a-target)), the call takes its database's lane ([concurrency.md](concurrency.md)), and `guidance.py` fills in `agent_guidance` on a failed result ([agent-text.md](agent-text.md)).

## Conventions

- All tool names use the `vcs_` prefix.
- Tool handlers return `dict[str, Any]` with at least a `success` key.
- Error results include an `"error"` key, which usage logging detects.
- Async tools (`async def`) are supported by `@vcs_tool` transparently.
- `Context` parameters from FastMCP are filtered out of usage logs automatically.

## Modules

```
msaccess-vcs-mcp/
├── src/msaccess_vcs_mcp/       # Package source
│   ├── __init__.py             # Version (__version__)
│   ├── main.py                 # Server entry point, startup sequence
│   ├── cli.py                  # msaccess-vcs terminal client: configure, rebuild-addin, run-tests
│   ├── tools.py                # FastMCP instance + all @vcs_tool handlers
│   ├── access_gate.py          # One lane per database; the exclusive hold
│   ├── config.py               # .env loading, get_config()
│   ├── usage_logging.py        # Structured JSONL usage logging
│   ├── security.py             # Path validation, write guards
│   ├── validation.py           # Startup validation helpers
│   ├── addin_integration.py    # COM calls to VCS add-in
│   ├── operation_manager.py    # Async operation queues
│   ├── callback_server.py      # HTTP callback server for VBA progress
│   ├── rebuild_watcher.py      # Watches rebuild-status.json after vcs_rebuild_addin launches
│   ├── vba_worker_manager.py   # vcs_run_vba worker thread with a hard timeout
│   ├── com_recovery.py         # COM disconnect classification and recovery state
│   ├── screen_capture.py       # Win32 window capture, dialog text, PNG encoding
│   ├── screenshot.py           # vcs_capture_screenshot targets, gate, COM object path
│   ├── screenshot_store.py     # Saved screenshot names, folder, and retention
│   ├── accessibility.py        # MSAA read and press for Office dialogs (raw ctypes)
│   ├── vba_control.py          # vcs_dismiss_dialog, vcs_reset_vba, code location
│   ├── compile_check.py        # Editor-side compile that reports the failing line
│   ├── guidance.py             # agent_guidance added to failed results
│   ├── access_state.py         # Dialogs and break mode attached to timeouts
│   ├── shared_open.py          # Import warnings when another session holds the database
│   ├── worktree.py             # Git worktree detection; shared-data and worktree-install warnings
│   ├── user_settings.py        # User-level setting decisions and their resolver
│   ├── settings_form.py        # vcs_configure: form or relayed values
│   ├── tips.py                 # vcs_tips topic loader
│   ├── tips/                   # vcs_tips topics, one markdown file each
│   └── access_com/             # Low-level COM/DAO helpers
│       ├── connection.py       # Find, start, open, show, and close Access instances
│       ├── instance_registry.py # Which Access processes the server created
│       ├── process_qos.py      # Full-power cores for server-launched Access
│       └── vba_project.py      # Find the current database's VBA project by file name
├── tests/                      # pytest test suite
├── docs/                       # Extended documentation
├── .env.example                # Template for environment variables
├── pyproject.toml              # Package metadata & dependencies
└── AGENTS.md                   # Entry guide, loaded on every turn
```

The tree is a starting point, not a complete list. Each module's docstring is the authority on what it does.
