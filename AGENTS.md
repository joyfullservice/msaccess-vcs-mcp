# AGENTS.md - AI Agent Guide to msaccess-vcs-mcp

## Purpose

This repository contains the **msaccess-vcs-mcp** server — a lightweight MCP (Model Context Protocol) bridge that lets AI agents drive the [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin) for version-controlling Microsoft Access databases.

## Development Environment Setup

**Always use the project virtual environment.** The package is installed in editable mode inside `venv/`. Do not install globally or suggest `pip install` outside the venv.

```powershell
# Activate the virtual environment (Windows PowerShell)
cd C:\path\to\msaccess-vcs-mcp
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
│   ├── screen_capture.py       # Win32 window capture, dialog text, PNG encoding
│   ├── screenshot.py           # vcs_capture_screenshot targets, gate, COM object path
│   ├── accessibility.py        # MSAA read and press for Office dialogs (raw ctypes)
│   ├── vba_control.py          # vcs_dismiss_dialog, vcs_reset_vba, code location
│   ├── compile_check.py        # Editor-side compile that reports the failing line
│   ├── guidance.py             # agent_guidance added to failed results
│   ├── access_state.py         # Dialogs and break mode attached to timeouts
│   ├── shared_open.py          # Import warnings when another session holds the database
│   ├── tips.py                 # vcs_tips topic loader
│   ├── tips/                   # vcs_tips topics, one markdown file each
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

### What agents are told, and where

Claude Code cuts server instructions, and any tool description it loads upfront, at 4,096 characters; some other clients never show the instructions to the model. Guidance is split three ways so nothing an agent needs depends on the instructions surviving:

- **Instructions** (`SERVER_INSTRUCTIONS` in `tools.py`, 3,500 characters or fewer): the rules every call needs, with the installed-add-in rule, no parallel calls, the stuck-Access call, and the `vcs_tips` triggers in the first 2,000 characters. No tool reference; every client lists the tools.
- **Results**: a failed result carries `agent_guidance`, one sentence naming the next call. `guidance.py` fills it in the `vcs_tool` wrapper from `error_pattern` (and a refused `vcs_rebuild_addin`) when the code that built the result did not. Never put guidance or dialog text in `error`; the usage log records that field. On `timeout` and `access_unresponsive`, `access_state.py` also attaches `access_state` (`pid`, `owned_by_server`, `vbe_state`, `modal_blocked`, `dialogs` with text capped at 500 characters) and guidance naming the `dialog_id` or `vcs_reset_vba()`. It reads windows only, through the screenshot resolver, under a 5-second deadline, and only when the call named a database (for `vcs_rebuild_addin`, the development copy beside `source_dir`), because the resolver's single-instance shortcut could describe an unrelated window.
- **`vcs_tips(topic)`**: depth for one situation, from the files in [src/msaccess_vcs_mcp/tips/](src/msaccess_vcs_mcp/tips/): [addin-development](src/msaccess_vcs_mcp/tips/addin-development.md), [stuck](src/msaccess_vcs_mcp/tips/stuck.md), [compile](src/msaccess_vcs_mcp/tips/compile.md), [logs](src/msaccess_vcs_mcp/tips/logs.md), [run-vba](src/msaccess_vcs_mcp/tips/run-vba.md). Gate-exempt, and the only tool registered with `meta={"anthropic/alwaysLoad": True}`.

`tests/test_agent_text.py` enforces the limits and checks that every `vcs_*` name and `vcs_tips` topic mentioned anywhere exists. When a behaviour described in a topic changes, update the topic in the same change.

### Imports into a database held elsewhere

Design and VBA changes merged into a database that another session has open can fail to save, or be dropped when that session saves. `vcs_import_objects` and `vcs_import_object` warn but never refuse, because the loss cannot be confirmed in advance. `shared_open.py` lists other local holders from the Restart Manager (minus the instance being driven and the server itself) and other computers from the `.laccdb` / `.ldb` lock file, whose entries outlive their sessions and so are labelled "may be stale". Each warning is one sentence: who holds it, the consequence, and the action (ask the user to close it, then verify with `vcs_diff_database`). `warnings` sits ahead of the per-object detail.

### Access window visibility

Any Access instance holding a database the server works with is left **visible**, whether the server created it or attached to one already running. A hidden instance strands the user: an error dialog, a VBA break, or a trust prompt blocks every later call with nothing on screen to explain why, and nobody can dismiss what they cannot see. `COM automation` normally starts Access hidden, so this is deliberate, not incidental.

Two rules follow from that, both enforced in `access_com/connection.py`:

- Open databases through `open_current_database(app, path)`, never a bare `app.OpenCurrentDatabase(...)`. It lowers `Application.UserControl` across the open — `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's own `AutoRun` opens its installer form when that flag says a person is watching, stranding the instance the server is about to drive — and it shows the window afterwards.
- Show the window *after* the database opens, via `ensure_access_visible(app)`. Making a window visible can itself set `UserControl`, which is why the order is not interchangeable.
- After that, instances the server created also get `UserControl = True` so the window is a normal interactive Access app. `Visible` alone is not enough on a COM-created process.

The one deliberate exception is `validate_access_installation()` in `config.py`: it opens no database and quits immediately, so a window would only flash on screen with nothing to act on.

### Full-power cores

Access is single-threaded. MCP-launched Access processes disable EcoQoS and run at Above Normal so Windows prefers a performance core. User-owned Access the server attaches to is left alone. Do not pin CPU affinity.

## Configuration

All settings come from environment variables (loaded from `.env` / `.env.local` in the project root). See `.env.example` for the full list.

Key variables:
- `ACCESS_VCS_DATABASE` — target database path
- `ACCESS_VCS_DISABLE_WRITES` — set `true` to block write operations
- `ACCESS_VCS_ENABLE_LOGGING` — set `true` to enable usage logging
- `ACCESS_VCS_RUN_VBA_TIMEOUT_SEC` — parent-side timeout for `vcs_run_vba` worker processes (default 45s)
- `ACCESS_VCS_CALL_VBA_TIMEOUT_SEC` — parent-side timeout for `vcs_call_vba` `Application.Run` calls (default 45s)
- `ACCESS_VCS_BUILD_TIMEOUT_SEC` — how long `vcs_rebuild_database` waits for the add-in's `BuildHeadless` to return (default 600s); on timeout Access is left open
- `ACCESS_VCS_REBUILD_TIMEOUT_SEC` — how long `vcs_rebuild_addin` waits after launch for a terminal status (default 1200s)
- `ACCESS_VCS_SCREENSHOT_OPEN_TIMEOUT_SEC` — how long `vcs_capture_screenshot` waits for an `open_if_closed` open before returning `open_blocked` (default 10s)
- `ACCESS_VCS_COMPILE_TIMEOUT_SEC` — how long `vcs_compile_vba` waits for the editor's compile to succeed or show its error before returning `compile_timeout` (default 120s)
- `ACCESS_VCS_BUSY_WAIT_SEC` — how long a second tool call waits for the Access gate before returning `server_busy` (default 15s)
- `ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC` — timeout for automatic Access/add-in recovery probes after a VBA timeout or COM disconnect (default 10s)
- `ACCESS_VCS_LEAVE_ACCESS_OPEN` — when `true` (the default), a COM-created Access instance is not Quit on disconnect so a later attach can reuse the boosted process. The user closes the window. Set `false` to restore quit-per-call. The server still closes instances it created when a rebuild must replace a file they hold, or when recycling a stuck owned instance after a failed recovery probe. User-owned Access is never closed.
- `ACCESS_VCS_CLOSE_TIMEOUT_SEC` — how long a graceful close of a server-created instance may take before it is force-terminated (default 10s)
- `ACCESS_VCS_OWNED_INSTANCES_PATH` — where the ownership registry lives (default `~/.msaccess-vcs-mcp/owned-instances.json`)

### Which Access windows the server may close

Persistence only pays off if the server can still get a file back when it needs to replace one, so it has to know which windows are its own. That is recorded on disk in the ownership registry, keyed by PID *and* process creation time — a PID alone is reused by Windows, and a reused PID would authorize closing a stranger's process.

The rules, all enforced in `access_com/instance_registry.py` and `access_com/connection.py`:

- **Only a confirmed match is owned.** If the creation stamp cannot be read, or the record does not carry one, the answer is "not ours". Every ambiguity resolves toward leaving the window alone.
- **A failed process query is not an empty one.** `list_access_pids_or_none()` returns `None` when `tasklist` itself fails, and the registry keeps its records rather than forgetting every window it created.
- **Closing re-verifies identity.** Both ways of resolving an instance from a path can land on a different process than the one recorded, and a moniker bind can even start a new one. `pid` and creation time are re-checked against the record before anything is closed.
- **A hung owned instance is terminated.** The graceful `CloseCurrentDatabase` + `Quit` runs in a worker thread with `ACCESS_VCS_CLOSE_TIMEOUT_SEC`; a blocking call here would hang the Access gate. If the process survives, it is force-terminated. This is reachable only for a PID proven to be server-created, where unsaved state is acceptable loss.
- **A live process keeps its claim.** A record is dropped only once the process is confirmed gone. Forgetting a live owned instance would make it permanently uncloseable.
- **A loaded add-in counts as holding a file.** Every tool call loads the add-in as a library, which locks it no matter which database that instance has open. `load_addin` records this, and a rebuild closes those instances too — matching only on the open database would make rebuilds refuse almost every time.

One MCP server process is shared across all Cursor windows. Sync tools run in a single COM apartment thread with one Access operation at a time. A long call in one window causes others to get `error_pattern: server_busy` with `busy_with` naming the in-flight tool — retry rather than waiting for a client `-32001` timeout. After any client timeout, call `vcs_get_recent_calls()` to see what actually executed.

`vcs_run_vba` executes Access COM work in a daemon worker thread with a hard timeout. If a snippet hangs because Access is in break mode, blocked on a modal dialog, or otherwise unresponsive, the MCP server abandons that thread and returns a recoverable timeout. It does **not** kill `MSACCESS.EXE` or close user-owned Access windows; after Access becomes responsive, the next call runs an automatic probe and resumes normal operation.

### Cancelling an operation

`vcs_cancel_operation(operation_id?, database_path?)` stops a running export, import, database rebuild, add-in rebuild, or test run; with no arguments it targets the only one running. It is gate-exempt and runs on the event loop, so every long tool must keep the loop free while Access works: whole-project export, merge, and test runs use `APIAsync` timer continuations, and the synchronous `ExportByType`, `ImportByType`, and `BuildHeadless` go through `_run_api_with_callbacks` on a worker thread. Never call a long add-in method with `call_sync` from an async tool.

- **The add-in polls; the server only answers.** `Operation.Pulse` asks `/cancel-status/{id}` about twice a second, and only when the callback info carries `cancel_poll: true`. The callback server is a `ThreadingHTTPServer` so that poll is answered while another callback is in flight. An add-in built before this support never polls, so its operations cannot be cancelled from MCP.
- **A wait that ends early abandons, never just unregisters.** On a wait timeout or client cancellation, `abandon_operation` marks the operation cancelled and keeps a ten-minute tombstone, so Access stops at its next poll instead of running on unobserved. A terminal callback for the id clears it.
- **Results say so.** A canceled run returns `cancelled: true` and `error_pattern: "operation_cancelled"`; a canceled add-in rebuild ends with status `cancelled` and installs nothing.

### Agents never close Access windows

Agents do not stop, kill, or quit `MSACCESS.EXE`. Never run `Stop-Process -Name MSACCESS`, `Get-Process MSACCESS | Stop-Process`, `taskkill /IM MSACCESS.EXE`, or any other command that selects Access by name. Other Cursor windows share this server, so an Access window you did not open may belong to another agent's run or to the user, and closing it destroys that work. A process list cannot tell you whose window it is.

When Access hangs or holds a file you need, leave the window and say so: name the PID and what it holds, and let the user close it. The one exception is a process your own script started in this session, such as a `New-Object -ComObject Access.Application` you created. Stop that one by the PID you recorded when you created it, and confirm its start time first.

### Rebuilding the VCS add-in

To rebuild `Version Control.accda` from source after editing add-in files, do **not** use `vcs_rebuild_database` (that rebuilds a user project). Call:

```python
vcs_rebuild_addin(r"C:\path\to\msaccess-vcs-addin\Version Control.accda.src")
```

The tool derives the development copy beside that folder, launches
`RebuildAddIn`, and passes its HTTP callback identity through the disconnected
worker to the builder Access process. Detailed `Log.Add` / `Log.Progress`
messages stream during the build; the status file covers compile, install, and
durable terminal recovery. Do **not** poll it yourself in the normal workflow.
`vcs_call_vba(..., ["RebuildAddIn", source])` remains launch-only.

MCP progress notifications are best-effort in Cursor 3.13 (often only "Running..." until the tool returns). For guaranteed live output, run `msaccess-vcs rebuild-addin <source>` from a terminal and keep that command in the foreground so the stream stays in the primary chat. The CLI exits when the operation reaches terminal status; that process exit is the completion signal. Do not background the CLI just to wait on a notification, and do not add a second timer wait, sleep, or `rebuild-status.json` poll after it has already finished.

`refused` and `launch-failed` are returned immediately and mean nothing was rebuilt: before launch, and inside the Access gate so nothing can reopen the file in between, the server closes Access windows **it created** that hold a file the rebuild replaces — including instances whose only claim on the add-in is having loaded it as a library. `refused` when a **user-owned** `MSACCESS.EXE` still holds one of those files or cannot be asked which files it holds (`otherInstances` names what to close; the add-in never closes another process), and `launch-failed` when the helper script never started, which leaves Access open and is safe to retry.

If a client times out (`-32001`) or the tool returns `rebuild_stalled` / `timeout`, recover by reading `<source>/logs/rebuild-status.json` and matching `phaseStarted` against `rebuild_phase_started`. A `complete` whose `phaseStarted` predates the call is an earlier run's record. After any client timeout, call `vcs_get_recent_calls()` before inferring from the status file. A live rebuild always has `MSACCESS.EXE` or `wscript.exe`; neither, with a non-terminal status, means the run died.

### Running the add-in's own tests

Pass the **development copy** in the add-in's repository as `database_path`:

```python
vcs_run_tests(r"C:\path\to\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

A run needs two projects and they are different files: the installed add-in loads as a library and supplies the runner and `TestAssert`, while the code under test is whatever the current database holds. The runner scans the current VBA project, so the host decides which tests are found — aim a run at a user database, or anything in the repo's `Testing` folder, and you get that database's tests reported as a clean pass.

The installed add-in is refused as a host (see below); it also has no source tree beside it for the tests that read one. The add-in refuses such a run itself in `ExecuteTests` via `modInstall.CurrentDbIsInstalledAddIn`, so the server's refusal is the earlier of two.

`AccessConnection` opens the development copy itself (Access refuses to bind a file moniker to an `.accda`, so `GetObject` fails and the explicit `OpenCurrentDatabase` fallback in `_open_as_current_database` takes over), so no manual pre-open is needed.

MCP progress notifications are best-effort in Cursor. For live per-test output, run `msaccess-vcs run-tests <database>` from a terminal and keep that command in the foreground, the same way as `rebuild-addin`. The CLI prints pytest-style dots for fast passes, names tests that take a second or more, then a human completion line; it does not dump the full `tests` map. Headless means no add-in UI, not a hidden Access window.

### The installed add-in is never a target

No tool accepts the installed add-in as `database_path`, `output_path`, or `template_path`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper, ahead of the gate and of any COM work, so the rule holds for every tool rather than the handful that grew their own guards — export, import, rebuild, `run_vba`, `run_tests`, `call_vba`, and the rest. Refusals carry `error_pattern: installed_addin_refused`.

`vcs_get_version_info()` reports the installed add-in's version without opening it, which is the one thing you might legitimately want from that file.

The comparison ignores the file extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`: a compiled install is a `.accde` built from the same `.accda`, and only one of the two is ever named in `ACCESS_VCS_ADDIN_PATH`. Folder parameters (`source_dir`, `output_dir`) are not checked — an export folder beside the install is not the install.

**Resolving the install path.** With `ACCESS_VCS_ADDIN_PATH` unset, `get_default_addin_path()` reads `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`, which is the only place to read it from — the add-in's own `GetInstalledAddInFileName` is built from exactly these two values. `Install Folder` is present only for a folder the user chose (the installer deletes it when the folder is the default, so absence means `%AppData%\MSAccessVCS`, not "not installed"), and `Compile accde` decides the extension. Do not reconstruct the path from `%AppData%` alone or assume `.accda`. The result is cached for the life of the process; `reset_addin_path_cache()` clears it.

Run these tests **through this server**, never from the add-in's own window. `modTestAssert.TestAssert` routes through `Application.Run` to the *installed* add-in path, while the runner singleton that records assertions lives in whichever project received `RunTests`. Invoking them from a development copy puts those in different projects: assertions are discarded and every test reports `EMPTY`. Treat an all-`EMPTY` result as a broken harness, not a pass.

### Screenshots

`vcs_capture_screenshot` returns a PNG of the Access window (`target="application"`), the VBA editor (`"vbe"`), the active object (`"active"`), or a named form, report, table, or query (`"object"`), with a JSON metadata block in front of it. The metadata carries the text, buttons, and `dialog_id` of every dialog Access is showing, `vbe_state` (`break`, `running`, `design`, read from the editor's title), and `owned_by_server`, so `include_image=False` is the cheap first call when Access seems stuck. Win32 capture lives in `screen_capture.py`; target resolution, the gate, and the COM object path live in `screenshot.py`.

- **Office dialogs are read through MSAA.** `NUIDialog` windows ("Cannot Compile Project.") have no child controls, so window calls find nothing. `accessibility.py` walks their accessibility tree with raw ctypes `IAccessible` calls (pywin32 cannot hand `AccessibleChildren` a raw pointer), only for a window that answers messages, on a worker with a deadline. `text_source` says which reader produced the text.

- **Window captures never touch COM or the gate.** The tool is in `EXEMPT_TOOLS`, and the `application` and `vbe` paths use only Win32 calls (on a worker thread, under a timeout), so they work while Access sits on a MsgBox, in break mode, or inside another tool's long call. Do not add a COM call to those paths; a blocked Access would hang them.
- **Object captures take the gate themselves** with `run_exclusive(..., is_async=False)`, so the COM work runs on the apartment thread. A Win32 precheck refuses with `access_blocked` before any COM when a dialog is up or VBA is in break mode, and an immediate `server_busy` check avoids waiting on another call. Both return the application capture instead.
- **The instance is found without COM.** In order: the `pid` argument, the Restart Manager (which process holds the file), the ownership registry, the window title, and the single running instance when no `database_path` is given. Several candidates give `ambiguous_instance` with the list.
- **The tool never changes window state.** It does not restore, focus, show, or close a window, and never clicks a dialog button. `open_if_closed=True` opens a closed object (reports in print preview, tables and queries read-only, only select, crosstab, and union queries) and closes it with `acSaveNo`; if the open times out (`open_blocked`, usually an Open event showing a dialog), the abandoned worker closes the object once the open finishes, and object captures for that database refuse with `cleanup_pending` until it has.
- **A responsive window is rendered with `PrintWindow`;** a hung one is copied from the screen (`capture_method: "screen"`, with `occluded` when another window covers it). In that copy every other application's window above Access is grayed out, applied bottom to top so an Access dialog above a foreign window keeps its pixels. Dialogs owned by Access are composited onto the image only when they touch the captured window; one on another monitor stays in `dialogs` with a warning rather than stretching the canvas.
- **Images are returned inline only, never written to disk.** The result is a `CallToolResult` (text then image, no output schema); `usage_logging` reads its outcome from the first text block and never logs the content.

- **`view="design"` opens a form or report in design view**, which fires no Open or Load events, so a layout changed in source can be checked after import without running its code. An object that is already open is captured as it is, with `object.current_view` read from `CurrentView` (and a `note` when the requested view differs). A design open that stops on a notice lands in the same `open_blocked` and cleanup path as an event-code dialog.
- **Modal and PopUp forms need no pre-check.** Checked on Access 365 (2026-11-24) in a throwaway database: `OpenForm` in normal window mode on a Modal + PopUp form returned in under a second, and `Forms(name).Hwnd` and `CurrentView` answered at once; only `acDialog` window mode blocks, and the tool never uses it. A design open with a second DAO connection holding the file also returned at once, without the exclusive-access notice.
- **`settle_ms`** (default 150, at most 5000) is the pause between opening an object and capturing it.

`ACCESS_VCS_SCREENSHOT_OPEN_TIMEOUT_SEC` (default 10s) bounds the `open_if_closed` open.

### Dialogs and break mode

The screenshot tool never clicks. Two separate tools act on what it reports, and only on Access windows **the server created**: `server_owns` requires a confirmed PID plus creation-time match in the ownership registry, the same proof that decides which windows the server may close. A user's Access is refused with `user_owned_instance`, and the guidance says to ask the user. Both tools are refused when `ACCESS_VCS_DISABLE_WRITES` is set, and both are in `EXEMPT_TOOLS`: the call stuck behind the dialog or the break usually holds the gate, and answering it is what releases the gate.

- **`vcs_dismiss_dialog(dialog_id, button)` clicks only the dialog the agent read.** `dialog_id` is a hash of the window handle, class, title, text, and buttons; if any of them changed, the call is refused with `dialog_changed` and the current `dialogs`. Standard dialogs get a posted `BM_CLICK` (posted, so a busy Access cannot block the caller); Office dialogs get the MSAA default action. The usage log records the title, button, and `dialog_id`, never the dialog text.
- **`vcs_reset_vba()` needs `vbe_state == "break"`.** On its own COM thread with a deadline (Access still answers COM in break mode), it reads the break location from `VBE.ActiveCodePane`, then runs `CommandBars.FindControl(1, 228)` (Reset). Pass the arguments positionally: win32com ignores `Id=` as a keyword and returns the first control ("&File").
- **`vcs_compile_vba` compiles from the editor** (`FindControl(1, 578)`) in `compile_check.py`, not with `acCmdCompileAllModules`. The Access command shows "Cannot Compile Project.", blocks until someone clicks OK, and never moves to the error. The editor command returns at once; about two seconds later the editor shows itself, highlights the error, and opens a standard "Compile error:" dialog. The gate thread watches for that dialog with window calls only, because a COM call issued while it is up waits behind it. Success is asked on a separate COM thread (`IsCompiled`, through a marshaled proxy). Before clicking OK the tool records the message and the module from the code window title; after OK it reads the line, columns, and code from the code pane (`location_source: "editor"`), falling back to the title alone (`"window"`). It clicks only the editor's own compile dialog; any other dialog returns `dialog_shown`. The add-in's `CompileVBA` remains the fallback when the editor has no compile control.

## Logging

The server writes two parallel JSON Lines streams. Both filenames use the `vcs-mcp-` prefix so they don't collide with other tools that share the same logs directory.

### VCS operation logs (written by the add-in, not the server)

Separate from the two streams below, the add-in writes a per-operation log to `{source_dir}/logs/<Base>_<yyyymmdd_hhnnss_fff>.log`, where `<Base>` is `Export`, `Merge`, `Build`, `TestRun`, or `Other`. Note the base name tracks the *operation*, not the tool: `vcs_import_objects` and `vcs_import_object` both produce `Merge_*.log`.

The add-in also writes a `.gitignore` into the source folder containing `logs/` and `*.log`, which means **Cursor's Glob and Grep silently skip these files** — a search returns no matches rather than an error, so an agent can burn several calls before falling back to a shell listing. To avoid that:

- Every operation tool returns `log_path` for the run it just performed. Use it directly with the Read tool.
- On failure, those tools also return `log_excerpt` (tail of the log), so the error is usually available without any follow-up call.
- If the path is no longer at hand, call `vcs_get_log(database_path, log_type=...)` with the base name matching the operation.

The add-in's sync API returns this as camelCase `logPath`; `_addin_json_result` in `tools.py` normalizes it to `log_path` at the boundary and keeps the original key as an alias. Async completion callbacks already use `log_path`. Keep the normalizer even if the add-in changes: a newer server may run against an older add-in build.

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
3. Add the name to `READ_ONLY_TOOLS` or `DESTRUCTIVE_TOOLS` in `tools.py` when it is either; the wrapper sets the MCP `readOnlyHint` / `destructiveHint` annotations from them
4. Keep the description under 4,000 characters (`tests/test_agent_text.py` checks), with what the tool does and its most common mistake first
5. Add tests in `tests/`

## Key Conventions

- All tool names use the `vcs_` prefix
- Tool handlers return `dict[str, Any]` with at least a `success` key
- Error results include `"error"` key (detected by usage logging)
- Async tools (`async def`) are supported by `@vcs_tool` transparently
- `Context` parameters from FastMCP are filtered out of usage logs automatically
