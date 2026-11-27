# AGENTS.md — msaccess-vcs-mcp

An MCP server that lets AI agents drive the
[MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin): export Access
databases to source files, merge or rebuild them, run VBA and tests, and capture screenshots.
One server process is shared by every Cursor window and every agent in them, so a mistake
here reaches other people's sessions.

This file loads on every turn. Its content is budgeted at **6,000 characters**; the routing
table has a separate 2,400-character / 20-row budget. `tests/test_agent_text.py` enforces
both, and that every `docs/*.md` is linked here. New depth goes in `docs/` with a routing
row, not in a new section here.

## Where to read next

| Working on | Read |
|---|---|
| Module map; how a call flows through `@vcs_tool` | [docs/architecture.md](docs/architecture.md) |
| Server instructions, tool descriptions, `agent_guidance` | [docs/agent-text.md](docs/agent-text.md) |
| Behaviour a `vcs_tips` topic describes | [tips/](src/msaccess_vcs_mcp/tips/); update the topic in the same change |
| Lanes, `server_busy`, worker timeouts, cancelling | [docs/concurrency.md](docs/concurrency.md) |
| Window visibility; which windows the server may close | [docs/access-instances.md](docs/access-instances.md) |
| Screenshots, dialogs, break mode, compile errors | [docs/screenshots.md](docs/screenshots.md) |
| Environment variables, `.env` loading, user settings | [docs/configuration.md](docs/configuration.md) |
| Rebuilding the add-in; its tests; the installed copy | [docs/addin-rebuild-and-tests.md](docs/addin-rebuild-and-tests.md) |
| Operation logs; diagnostic and usage streams | [docs/logging.md](docs/logging.md) |
| Unit, integration, and end-to-end tests | [docs/TESTING.md](docs/TESTING.md) |
| How agents use the tools, step by step | [docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md) |
| Calling the add-in over COM | [docs/VBA_INTEGRATION.md](docs/VBA_INTEGRATION.md) |
| The add-in's progress and cancel callbacks | [docs/VBA_CALLBACK_API.md](docs/VBA_CALLBACK_API.md) |
| Exported file formats | [docs/EXPORT_FORMATS.md](docs/EXPORT_FORMATS.md) |
| Dependencies, `uv.lock`, releases | [CONTRIBUTING.md](CONTRIBUTING.md) |
| Why something was built this way | [DECISIONS.md](DECISIONS.md) (newest first; search it) |

## Development workflow

Always work in the project virtual environment (`venv/`), where the package is installed in
editable mode. Never `pip install` globally or outside the venv.

```powershell
python -m venv venv; .\venv\Scripts\Activate.ps1; pip install -e ".[dev]"   # first time only
.\venv\Scripts\Activate.ps1
pytest -m "not integration"     # integration tests need Access installed
```

After changing a dependency in `pyproject.toml`, run `uv lock` and commit the lock; CI and the
PyPI workflow install from it.

The multi-instance concurrency suite is skipped by every other run, including `-m integration`.
It drives up to three visible Access windows for a few minutes. Run it after changing the
Access gate (`access_gate.py`), instance creation or ownership (`access_com/connection.py`,
`access_com/instance_registry.py`), the VBA worker, dialog control, screenshots, the callback
server, the `vcs_tool` wrapper, or the add-in's headless export, merge, build, or test paths:

```powershell
pytest tests/test_concurrency_integration.py --run-concurrency -v
```

## Invariants

Nothing enforces these at runtime; missing one destroys someone's work, strands a window, or
resets a VBA project while it runs.

- **Agents never close Access windows.** Never run `Stop-Process -Name MSACCESS`,
  `taskkill /IM MSACCESS.EXE`, or anything else that selects Access by name. A window you did
  not open may belong to another agent's run or to the user, and a process list cannot tell
  you whose it is. When Access hangs or holds a file you need, name the PID and what it holds,
  and let the user close it. The one exception is a process your own script started in this
  session, such as a `New-Object -ComObject Access.Application`: stop it by the PID you
  recorded, after confirming its start time. The server
  itself closes only windows its ownership registry proves it created.
- **The installed add-in is never a target.** No tool accepts it as `database_path`,
  `output_path`, or `template_path`, because opening or writing it resets a VBA project that
  is executing. `_refuse_installed_addin_target` in the `vcs_tool` wrapper holds this for
  every tool. Rebuild the add-in with `vcs_rebuild_addin`, never `vcs_rebuild_database`, and
  run its tests on the development copy, never the installed one.
- **Calls on one database run one at a time.** `access_gate.py` gives each database a lane
  with its own COM apartment thread; different databases run in parallel. Anything a tool
  keeps between calls must be per database or locked. Never call a long add-in method with
  `call_sync` from an async tool: it blocks the event loop, and with it cancellation.
- **Access stays visible.** Open databases through `open_current_database(app, path)`, never
  a bare `OpenCurrentDatabase`, and show the window afterwards. A hidden instance strands the
  user behind a dialog nobody can see.

## Adding a tool

1. Write the handler in `tools.py` and decorate it with `@vcs_tool("vcs_your_tool_name")`,
   which reloads config, logs usage, and registers the tool with MCP. Names use the `vcs_`
   prefix.
2. Return a `dict` with at least `success`. A failure adds `error` and `error_pattern`; next
   steps go in `agent_guidance`, never in `error`, which the usage log records.
3. Add the name to `READ_ONLY_TOOLS` or `DESTRUCTIVE_TOOLS` in `tools.py` when it is either;
   the wrapper sets the MCP `readOnlyHint` / `destructiveHint` annotations from them.
4. Keep the description under 4,000 characters, with what the tool does and its most common
   mistake first.
5. Add tests in `tests/`.
