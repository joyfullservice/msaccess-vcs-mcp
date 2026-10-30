# msaccess-vcs-mcp

A lightweight MCP server (Model Context Protocol server) that provides AI agent integration for the [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-integration). This tool acts as a bridge, allowing AI assistants (e.g., Cursor, Claude Code, and other MCP-compatible clients) to work with Microsoft Access databases through version control operations.

## Features

- **Full Export**: Export all database objects (forms, reports, queries, modules, macros, etc.)
- **Fast Save**: Incremental exports (only changed objects)
- **Per-Object Operations**: Export or import individual objects by name and type
- **Merge Build**: Import source changes into existing database
- **Build from Source**: Create fresh database from source files
- **Object Inventory**: List all objects in an Access database
- **Change Tracking**: Compare database against source files
- **SQL Queries**: Execute read-only SELECT queries via the add-in's DAO connection
- **VBA Execution**: Call existing VBA functions or run agent-generated code
- **Async Operations**: Long-running exports/builds report progress via HTTP callbacks
- **Add-in Rebuild**: `vcs_rebuild_addin` streams build callbacks, then watches compile/install status
- **Terminal CLI**: `msaccess-vcs` streams the same MCP progress to stdout
- **Add-in Options**: Read and write VCS add-in settings at runtime
- **Safety Guardrails**: Path validation, permission checks, and write-disable mode

## Architecture

This MCP server is a **lightweight wrapper** around the MSAccess VCS add-in, delegating all database operations to the battle-tested add-in:

```
AI Agent -> MCP Server (Python) -> VCS Add-in (VBA) -> Access Database
                |
          Path validation
          Permission checks
          Result formatting
          Async progress tracking
```

All business logic lives in VBA. The MCP layer validates inputs, manages async lifecycle, and formats responses.

### Access windows are visible

When a tool needs Access, the window is shown rather than kept hidden — expect an Access window to appear if one isn't already open. This is intentional: Access asks questions that only a person can answer (trust prompts, "convert this database?", a VBA error breaking into the debugger), and an operation that stalls behind an invisible dialog looks like a hang with no way to clear it. The server never closes an Access window it did not open. Instances it created stay open across tool calls so merge / test / edit loops reuse the same process; you close the window when you are done. The server closes a window it created only when a rebuild must replace a file it holds, when recycling a stuck owned instance, or when `ACCESS_VCS_LEAVE_ACCESS_OPEN=false`.

Which windows are the server's own is recorded on disk, keyed by process id *and* creation time, so it survives a server restart and cannot be fooled by Windows reusing a process id. Anything it cannot positively confirm is treated as yours and left alone. A server-created instance that ignores the request to close within `ACCESS_VCS_CLOSE_TIMEOUT_SEC` is force-terminated, which only ever applies to a process the server started.

## Prerequisites

- **Python**: 3.10 or higher
- **Microsoft Access**: Installed on Windows (for COM automation)
- **pywin32**: Python COM interface (installed automatically)

If MCP startup fails with a `CLSIDToClassMap` / `win32com.gen_py` error, the server deletes the stale type-library folder under `%TEMP%\gen_py` and retries once automatically. If it still fails after that, Microsoft Access is probably not installed or not registered for COM automation.
- **MSAccess VCS Add-in**: Must be installed ([download latest release](https://github.com/joyfullservice/msaccess-vcs-integration/releases/latest))

## Getting Started

The quickest way to get running is with [uvx](https://docs.astral.sh/uv/guides/tools/) (the tool runner from uv). No cloning or virtual environments needed.

### 1. Install uv

```bash
# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. Install the MSAccess VCS Add-in

This MCP server requires the MSAccess VCS add-in to be installed.

1. Download the [latest release](https://github.com/joyfullservice/msaccess-vcs-integration/releases/latest) from GitHub
2. Extract `Version Control.accda` to a trusted location
3. Open `Version Control.accda` to launch the installer
4. Click **Install Add-In**

The add-in will be installed to `%AppData%\MSAccessVCS\Version Control.accda` by default.

### 3. Register the MCP server

Add the server entry to your MCP client config. Both Cursor and Claude Code use the same `mcpServers` format, just in different files:

| Client | Project-level config | User-level (global) config |
|--------|---------------------|---------------------------|
| Cursor | `.cursor/mcp.json` | `~/.cursor/mcp.json` |
| Claude Code | `.mcp.json` | `~/.claude.json` |

Add this to the appropriate config file:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "uvx",
      "args": ["msaccess-vcs-mcp@latest"]
    }
  }
}
```

The `@latest` suffix ensures `uvx` always pulls the latest version from [PyPI](https://pypi.org/project/msaccess-vcs-mcp/) instead of using a cached copy.

**Alternative for Claude Code** -- you can use the CLI instead of editing JSON:

```bash
claude mcp add msaccess-vcs-mcp -- uvx msaccess-vcs-mcp@latest
```

**Client on another machine (SSE)** -- Access must run on Windows, but the MCP client need not run on the same machine. See [Running the Server for a Client on Another Machine](#running-the-server-for-a-client-on-another-machine).

### 4. Configure your database

Create a `.env` file in your project root:

```bash
# Target database for this project
ACCESS_VCS_DATABASE=C:\Projects\MyApp\database.accdb

# Optional: Disable database writes for safety
# ACCESS_VCS_DISABLE_WRITES=true

# Optional: Custom add-in path
# ACCESS_VCS_ADDIN_PATH=C:\Custom\Path\Version Control.accda
```

### 5. Restart your client

Close and reopen Cursor or Claude Code. The MCP server will be detected and loaded automatically.

### 6. Try it out

Ask the AI assistant to use the VCS tools:

> "List all objects in my Access database using vcs_list_objects"

> "Export my Access database to the src folder using vcs_export_database"

> "Compare my database against the source files using vcs_diff_database"

## Configuration

All configuration is done through environment variables, typically in a `.env` file in your project root. The server loads `.env` automatically at startup.

### Environment Variables

| Variable | Description | Default | Required |
|----------|-------------|---------|----------|
| `ACCESS_VCS_DATABASE` | Target database path (.accdb, .accda, .mdb) | - | Recommended |
| `ACCESS_VCS_PROJECT_DIR` | Explicit path to the project containing your `.env` (overrides auto-discovery) | auto | No |
| `ACCESS_VCS_ADDIN_PATH` | Path to VCS add-in file | `%AppData%\MSAccessVCS\Version Control.accda` | No |
| `ACCESS_VCS_DISABLE_WRITES` | Disable database modifications (`true`/`false`) | `false` | No |
| `ACCESS_VCS_CALLBACK_ENABLED` | Enable HTTP callback server for async progress | `true` | No |
| `ACCESS_VCS_CALLBACK_HOST` | Host address for callback server | `127.0.0.1` | No |
| `ACCESS_VCS_VALIDATE_STARTUP` | Validate Access and add-in availability on startup | `false` | No |
| `ACCESS_VCS_ENABLE_LOGGING` | Enable structured JSONL usage logging | `false` | No |
| `ACCESS_VCS_LOG_DIR` | Custom log directory (overrides auto-detection) | auto | No |
| `ACCESS_VCS_LOG_MAX_SIZE_MB` | Max log file size before rotation | `10` | No |
| `ACCESS_VCS_LOG_BACKUP_COUNT` | Number of rotated backup files to keep | `5` | No |

### Project-Specific Configuration

This tool supports per-project configurations using `.env` files:

#### `.env` (Gitignored, Project-Specific)

Store project-specific database path:

```bash
ACCESS_VCS_DATABASE=C:\Projects\MyApp\database.accdb
```

#### `.env.local` (Optional, Gitignored)

For personal overrides that shouldn't affect the team:

```bash
ACCESS_VCS_DATABASE=C:\Users\YourName\dev\testdb.accdb
```

### Environment Variable Precedence

1. **MCP server `env` section** (highest priority) -- values set in the MCP config file
2. **`.env.local`** -- personal overrides
3. **`.env`** -- project-specific configuration (lowest priority)

### Using msaccess-vcs-mcp from Another Project

When the server is installed once at the user level (in `~/.cursor/mcp.json` rather than per-project), it still needs to find each individual project's `.env` so that settings like `ACCESS_VCS_DATABASE` and `ACCESS_VCS_ENABLE_LOGGING` take effect. The server resolves the project root in the following order:

1. **`ACCESS_VCS_PROJECT_DIR`** -- explicit override, set in the MCP `env` section if needed.
2. **MCP workspace roots** -- when the client (Cursor, VS Code, Claude Code, etc.) supports `roots/list`, the server lazily discovers the workspace on the first tool call and loads `.env` from it. No configuration required.
3. **Upward search from the working directory** -- works automatically when the IDE launches the server with the workspace as `cwd` (the typical case).
4. **Upward search from the installed package location** -- last-resort fallback for development installs.

`.env` files are also hot-reloaded: editing `.env` while the server is running causes the next tool call to pick up the new values, including `ACCESS_VCS_ENABLE_LOGGING`.

The server prints diagnostics to stderr at startup (visible in Cursor's MCP server output pane) showing the resolved project root and which files were loaded -- check there first if a setting isn't taking effect.

### Running the Server for a Client on Another Machine

The default stdio transport needs the MCP client and Access on the same Windows machine. With `--transport sse`, the server runs on the Windows machine that has Access and the client connects to it across the network: Claude Code on a Mac or Linux machine, say, or on another Windows PC. The Windows machine can be a physical PC on the LAN or a VM on any platform (Parallels Desktop, Hyper-V, VMware, Proxmox and so on).

On the Windows machine, in the logged-in session (not as a service, so Access can show its windows), start the server listening on that machine's own IP address:

```bash
msaccess-vcs-mcp --transport sse --host 192.168.1.50 --port 8765
```

On the client machine, use the same address:

```bash
claude mcp add --transport sse msaccess-vcs-mcp http://192.168.1.50:8765/sse
```

The server checks each request's `Host` header (protection against DNS rebinding) and accepts only the address it listens on, plus `localhost`, `127.0.0.1` and `[::1]`. That shapes the setup:

- **One address.** `--host` defaults to `127.0.0.1`, and `0.0.0.0` / `::` are refused: the check needs a specific address to allow.
- **The IP address in the URL, not a name.** A URL such as `http://win11.lan:8765/sse` sends `win11.lan` as the `Host` and is refused with `421 Misdirected Request`.
- **An address that does not change.** Give the Windows machine a static address or a DHCP reservation. If the address changes, restart the server with the new one.
- **Reached directly, not through NAT.** A VM needs an address the client can route to: bridged onto the LAN, or the host-only / shared network of a desktop hypervisor when the client runs on that same host. Through a NAT port forward the client connects to the forwarding host's address, which the server cannot listen on, so it is refused (421).
- **Otherwise, a tunnel.** Where none of that is possible, leave the server on `127.0.0.1` and forward a port over SSH (OpenSSH Server is an optional feature of Windows): `ssh -N -L 8765:127.0.0.1:8765 user@windows-machine` on the client, then `http://localhost:8765/sse`. `localhost` is always accepted, and only SSH is exposed.

The SSE transport has **no authentication**: anyone who can reach the port can drive Access on that machine. Open the port in Windows Firewall to the client machine's address only, not the whole subnet:

```powershell
New-NetFirewallRule -DisplayName "msaccess-vcs-mcp" -Direction Inbound -Protocol TCP -LocalPort 8765 -RemoteAddress 192.168.1.20 -Action Allow
```

If the hypervisor has its own firewall switched on for the VM (Proxmox's, for example), it needs the same opening. Do not expose the port beyond the LAN; reach it over a VPN or an SSH tunnel instead.

## Available Tools

### Database-Level Operations

#### `vcs_export_database(database_path, output_dir, object_types, full_export)`

Export Access database objects to source files via the VCS add-in.

With no `object_types`, exports the whole project (async progress). With
`object_types`, exports only those categories via a scoped sync call. Uses
fast save by default (only changed objects).

**Args:**
- `database_path`: Path to Access database (.accdb, .accda, .mdb)
- `output_dir`: Directory to export source files to
- `object_types`: Optional list of categories (defaults to entire project)
- `full_export`: If True, export all objects in scope (not just changed ones)

**Returns:** `success`, `exported_count`, `export_path`, `objects_by_type`, `log_path`

```python
vcs_export_database("C:\\db.accdb", "C:\\src\\mydb")

vcs_export_database("C:\\db.accdb", "C:\\src\\mydb", object_types=["modules"])
```

#### `vcs_list_objects(database_path)`

List all objects in an Access database by type.

**Args:**
- `database_path`: Path to Access database

**Returns:** `tables`, `queries`, `modules` (and `forms`, `reports` in future)

```python
vcs_list_objects("C:\\db.accdb")
```

#### `vcs_diff_database(database_path, source_dir, show_details)`

Compare database objects against source files to see what has changed.

**Args:**
- `database_path`: Path to Access database
- `source_dir`: Directory containing source files
- `show_details`: If True, show detailed diff

**Returns:** `queries`, `modules`, `new_in_db`, `new_in_source`

```python
vcs_diff_database("C:\\db.accdb", "C:\\src\\mydb")
```

#### `vcs_import_objects(database_path, source_dir, object_types, full_import)`

Import objects from source files into Access database using merge build.

With no `object_types`, runs a full project merge. With `object_types`, merges
only those categories (orphan reconciliation within them; no database backup).
Requires write permission.

**Args:**
- `database_path`: Path to Access database
- `source_dir`: Directory containing source files
- `object_types`: Optional list of categories to merge (e.g. `["queries"]`)
- `full_import`: When scoped: if True, merge all source files in those
  categories (ignore change index); if False (default), only changed files

```python
vcs_import_objects("C:\\db.accdb", "C:\\src\\mydb")
vcs_import_objects("C:\\db.accdb", "C:\\src\\mydb", object_types=["modules"], full_import=True)
```

#### `vcs_rebuild_database(source_dir, output_path, template_path)`

Build a complete Access database from source files.

Creates a fresh database from source files, useful for clean builds and distribution. Requires write permission.

**Args:**
- `source_dir`: Directory containing source files
- `output_path`: Path for the new database file
- `template_path`: Optional template database to use as starting point

```python
vcs_rebuild_database("C:\\src\\mydb", "C:\\output\\fresh.accdb")
```

#### `vcs_rebuild_addin(source_dir, timeout_seconds)`

Rebuild `Version Control.accda` from source and wait until it is installed.

Derives the development copy beside `source_dir` and launches `RebuildAddIn`.
The callback URL and operation ID cross the disconnected worker boundary into
the builder Access process, so the same `Log.Add` and `Log.Progress` HTTP
callbacks used by ordinary builds provide detailed output. After the builder
exits, the tool watches `<source>/logs/rebuild-status.json` for the compile and
install phases until this attempt reaches `complete` or a terminal failure.
This is not `vcs_rebuild_database`, which rebuilds a user project.

```python
vcs_rebuild_addin(r"C:\path\to\msaccess-vcs-addin\Version Control.accda.src")
```

Do not poll the status file yourself unless the tool timed out. `vcs_call_vba(..., ["RebuildAddIn", source])` remains a launch-only escape hatch.

Before launch the server closes Access windows **it created** that hold a file the rebuild replaces. Because every tool call loads the add-in as a library, and that locks the file, this includes server-created instances with an unrelated database open. The close runs inside the Access gate, so another window's tool call cannot reopen the file between the close and the launch. The rebuild still refuses when a **user-owned** `MSACCESS.EXE` holds one of those files — it checks the loaded VBA projects in each one, so an instance with an unrelated database open is left alone. An instance that cannot be asked also blocks it. The add-in never closes another Access process. On refusal, `otherInstances` names each process and which file it holds.

`refused` and `launch-failed` come back immediately and mean nothing was rebuilt. `launch-failed` means the helper script never started; Access is left open and the call is safe to retry.

#### Terminal CLI (`msaccess-vcs`)

Cursor 3.13 often shows only "Running..." for MCP progress. The `msaccess-vcs` command launches this same server over stdio and prints each progress line immediately:

```text
msaccess-vcs export C:\db.accdb C:\src --full
msaccess-vcs merge C:\db.accdb C:\src
msaccess-vcs rebuild-database C:\src C:\out.accdb
msaccess-vcs rebuild-addin C:\path\to\msaccess-vcs-addin\Version Control.accda.src
msaccess-vcs run-tests C:\path\to\msaccess-vcs-addin\Version Control.accda --filter clsTestInstall
```

#### Running the add-in's own tests

Pass the development copy in the add-in's repository as the database path:

```python
vcs_run_tests(r"C:\path\to\msaccess-vcs-addin\Version Control.accda", filter="clsTestInstall")
```

A run needs two projects and they are different files: the installed add-in loads as a library and supplies the runner and `TestAssert`, while the code under test is whatever the current database holds. The runner scans the current VBA project, so the host decides which tests are found — a run aimed at a user database, or at anything in the repository's `Testing` folder, finds that database's tests instead. The server opens the development copy as the current database for you, so there is no manual pre-open step.

The installed add-in is refused as a host (see below), and it has no source tree beside it for the tests that read one. The add-in refuses such a run itself, so the server's refusal is the earlier of two.

#### The installed add-in is never a target

No tool accepts the installed add-in as `database_path`, `output_path`, or `template_path`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. The check runs before the Access gate and before any COM work, so it applies to every tool — export, import, rebuild, `vcs_run_vba`, `vcs_run_tests`, `vcs_call_vba`, and the rest. Refusals carry `error_pattern: installed_addin_refused`.

`vcs_get_version_info()` reports the installed add-in's version without opening it.

The comparison ignores the file extension, since a compiled install is a `.accde` built from the same `.accda`. Folder parameters are not checked — an export folder beside the install is not the install.

With `ACCESS_VCS_ADDIN_PATH` unset, the install path comes from `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`, the same `Install Folder` and `Compile accde` values the add-in's own installer writes. `Install Folder` is absent for a default install, which means `%AppData%\MSAccessVCS`.

Run these tests through the server rather than from the add-in's own window. Assertions are recorded by the runner in whichever project received the call, while `TestAssert` always routes to the *installed* add-in, so a run started inside a development copy discards every assertion and reports `EMPTY` for each test. An all-`EMPTY` result means the harness was bypassed, not that the tests passed.

### Per-Object Operations

#### `vcs_export_object(database_path, object_type, object_name)`

Export a single named database object to its source file. Much faster than a full database export when you only need to refresh one object.

**Args:**
- `database_path`: Path to Access database
- `object_type`: Type of object: `"query"`, `"form"`, `"report"`, `"module"`, `"table"`, `"macro"`
- `object_name`: Name of the object to export

```python
vcs_export_object("C:\\db.accdb", "query", "qryCustomers")
vcs_export_object("C:\\db.accdb", "module", "modUtils")
```

#### `vcs_import_object(database_path, object_type, object_name)`

Import a single named object from source files into the database. The source file must exist in the project's export folder. Requires write permission.

**Args:**
- `database_path`: Path to Access database
- `object_type`: Type of object: `"query"`, `"form"`, `"report"`, `"module"`, `"table"`, `"macro"`
- `object_name`: Name of the object to import

```python
vcs_import_object("C:\\db.accdb", "query", "qryCustomers")
vcs_import_object("C:\\db.accdb", "module", "modUtils")
```

### SQL and VBA Execution

#### `vcs_execute_sql(database_path, sql, max_rows)`

Execute a read-only SELECT query against the database via the add-in's DAO connection.

Only SELECT statements are allowed -- INSERT, UPDATE, DELETE, and DDL are rejected. Uses the same database connection the add-in already holds, avoiding file-locking conflicts.

**Args:**
- `database_path`: Path to Access database
- `sql`: SELECT statement to execute
- `max_rows`: Maximum number of rows to return (default: 100)

**Returns:** `rows`, `rowCount`, `truncated`

```python
vcs_execute_sql("C:\\db.accdb", "SELECT Name, Type FROM MSysObjects WHERE Type=5")
vcs_execute_sql("C:\\db.accdb", "SELECT * FROM Customers", max_rows=50)
```

#### `vcs_call_vba(database_path, function_name, args)`

Call an existing public VBA function by name via `Application.Run`. Lighter weight than `vcs_run_vba` since there is no temp module creation or compilation step.

**Args:**
- `database_path`: Path to Access database
- `function_name`: Fully qualified function name (e.g., `"ModuleName.FunctionName"`)
- `args`: Optional list of string arguments (max 3)

```python
vcs_call_vba("C:\\db.accdb", "MyModule.GetQuerySQL", ["qryCustomers"])
vcs_call_vba("C:\\db.accdb", "Version Control.API", ["GetVCSVersion"])
```

#### `vcs_run_vba(database_path, code)`

Execute agent-generated VBA code in a temporary module.

The add-in handles the full lifecycle: creates a temp module, wraps the code in a function with error handling, compiles the project to validate, executes, captures the result, removes the temp module, and returns structured JSON.

**Requires `McpAllowRunVBA` option to be enabled** (default: off). The user must enable this manually in the VCS Options form — agents cannot set this option programmatically.

**Args:**
- `database_path`: Path to Access database
- `code`: VBA code block to execute

```python
vcs_run_vba("C:\\db.accdb", "MCP_TempFunction = CurrentDb.TableDefs.Count")
```

### VBA Compilation

#### `vcs_check_vba_compiled(database_path)`

Check if VBA code in an Access database is compiled. Returns the compilation state without attempting to compile. Useful for establishing a baseline before making code changes.

**Args:**
- `database_path`: Path to Access database

**Returns:** `success`, `compiled`

```python
result = vcs_check_vba_compiled("C:\\db.accdb")
# result["compiled"] -> True or False
```

#### `vcs_compile_vba(database_path, suppress_warnings)`

Compile all VBA modules in an Access database and return success status.

When compilation fails, MCP cannot report the failing module or line. **Stop
making further code changes.** Ask the user to open the database in Access,
open the Visual Basic Editor, and choose **Debug → Compile** — Access navigates
to the first error line. Ask the user to paste the code around that line (a few
lines above and below). Once you have the snippet, propose a targeted fix; do not
guess or iterate through speculative edits.

The response includes an `agent_guidance` field when compilation fails.

**Args:**
- `database_path`: Path to Access database
- `suppress_warnings`: If True, suppress message boxes during compilation

**Returns:** `success`

```python
result = vcs_compile_vba("C:\\db.accdb", suppress_warnings=True)
```

### Add-in Options

#### `vcs_set_option(database_path, option_name, value)`

Set a VCS add-in option for the current session.

Changes take effect immediately but do not persist to `vcs-options.json` until explicitly saved.

**Args:**
- `database_path`: Path to Access database
- `option_name`: Name of the VCS option property
- `value`: Value to set (string, boolean, or integer)

```python
vcs_set_option("C:\\db.accdb", "ShowDebug", True)
```

#### `vcs_get_option(database_path, option_name)`

Read a VCS add-in option value.

**Args:**
- `database_path`: Path to Access database
- `option_name`: Name of the VCS option property to read

```python
vcs_get_option("C:\\db.accdb", "ShowDebug")
vcs_get_option("C:\\db.accdb", "McpAllowRunVBA")
```

### Diagnostics

#### `vcs_get_version_info()`

Get version information for MCP server, MSAccess VCS add-in, and Access application.

Returns `mcp_version`, `vcs_version`, `access_version`, `bitness`, `target_database`, `addin_path`, `callback_url`, `async_available`, and any `errors` or `warnings`.

```python
vcs_get_version_info()
```

#### `vcs_cancel_operation(operation_id)`

Cancel a running async operation. Requests cancellation of a long-running export, build, or import. The VBA add-in will detect the cancellation during its next DoEvents cycle.

**Args:**
- `operation_id`: UUID of the operation to cancel (returned by async tool calls)

```python
vcs_cancel_operation("a1b2c3d4-5678-90ab-cdef-1234567890ab")
```

#### `vcs_get_log(database_path, log_type)`

Read the most recent operation log file from the source folder's logs directory.

**Args:**
- `database_path`: Path to Access database
- `log_type`: Type of log to read: `"Export"` (default) or `"Build"`

```python
vcs_get_log("C:\\db.accdb")
vcs_get_log("C:\\db.accdb", log_type="Build")
```

## Async Operations

Long-running operations (export, import, rebuild) support async execution with progress reporting. When the callback server is running (enabled by default), these operations:

1. Spawn a detached VBA process via the add-in's `APIAsync` entry point
2. Return immediately with an `operation_id`
3. Receive progress updates via HTTP callbacks from VBA
4. Support cancellation via `vcs_cancel_operation`

VBA progress is per-category and can reset (for example 28/30 queries, then 1/50 modules). MCP requires a strictly increasing `progress` value, so the server sends an incrementing sequence and keeps VBA's `current/total` in the message. Log callbacks are forwarded the same way; MCP logging notifications are deprecated and are not used.

Cursor 3.13 often does not display those notifications. Use `msaccess-vcs` in a terminal when you need to see updates as they arrive.

The server automatically detects when another operation is in progress for the same database and returns a busy response with the active operation details.

If the callback server is not available, operations fall back to synchronous execution.

See [docs/VBA_CALLBACK_API.md](docs/VBA_CALLBACK_API.md) for the full callback protocol specification.

## Export Format

The VCS add-in exports database objects to a comprehensive folder structure with text-based files. See the [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-integration/blob/main/Version%20Control.accda.src/AGENTS.md) guide for complete format documentation.

### Example Structure

```
database.src/
├── queries/           # SQL queries (.sql, .bas)
├── modules/           # VBA modules (.bas, .cls)
├── forms/             # Form definitions (.bas, .cls)
├── reports/           # Report definitions (.bas, .cls)
├── macros/            # Macros (.bas)
├── tables/            # Table data (.txt, .xml)
├── tbldefs/           # Table structure (.sql, .xml)
├── vcs-options.json   # Export options
├── vcs-index.json     # Fast save index
└── Export.log         # Operation log
```

All files use **UTF-8 with BOM** encoding, which is critical for proper import back into Access.

## Workflows

### Version Control Workflow

```python
# 1. Export database to source
vcs_export_database("C:\\db.accdb", "C:\\src\\db")

# 2. Initialize git (if not already done)
# cd C:\src\db && git init && git add . && git commit -m "Initial export"

# 3. Make changes in Access...

# 4. See what changed
vcs_diff_database("C:\\db.accdb", "C:\\src\\db")

# 5. Export changes
vcs_export_database("C:\\db.accdb", "C:\\src\\db")

# 6. Commit changes
# git add . && git commit -m "Updated customer queries"
```

### Per-Object Development Workflow

```python
# 1. Edit a VBA module in source files...

# 2. Import just that module
vcs_import_object("C:\\db.accdb", "module", "modUtils")

# 3. Compile to validate
vcs_compile_vba("C:\\db.accdb")

# 4. Test via VBA
vcs_call_vba("C:\\db.accdb", "modUtils.RunTests")

# 5. Export the module back (picks up any Access-side formatting)
vcs_export_object("C:\\db.accdb", "module", "modUtils")
```

### Integration with db-inspector-mcp

Both tools work together for comprehensive database workflows:

```python
# 1. Analyze database structure (db-inspector-mcp)
db_list_tables(database="legacy")
db_list_views(database="legacy")

# 2. Export to version control (msaccess-vcs-mcp)
vcs_export_database("C:\\legacy.accdb", "C:\\src\\legacy-db")

# 3. Run queries against the database (msaccess-vcs-mcp)
vcs_execute_sql("C:\\legacy.accdb", "SELECT * FROM Customers WHERE Active = True")

# 4. Cross-database comparison (db-inspector-mcp)
db_compare_queries(
    "SELECT * FROM Customers",
    "SELECT * FROM Customers",
    database1="legacy",
    database2="new"
)
```

## MCP Client Setup

### Cursor

**Project-level** -- add `.cursor/mcp.json` to your project root (can be version-controlled for team sharing):

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "uvx",
      "args": ["msaccess-vcs-mcp@latest"]
    }
  }
}
```

**User-level** -- add to `~/.cursor/mcp.json` to make the server available in all projects.

### Claude Code

**Project-level** -- add `.mcp.json` to your project root:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "uvx",
      "args": ["msaccess-vcs-mcp@latest"]
    }
  }
}
```

**CLI alternative** -- register without editing JSON:

```bash
claude mcp add msaccess-vcs-mcp -- uvx msaccess-vcs-mcp@latest
```

### Development Install

For contributing or running from source:

```bash
git clone https://github.com/joyfullservice/msaccess-vcs-mcp.git
cd msaccess-vcs-mcp
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux
pip install -e ".[dev]"
```

For development installs, use `python -m msaccess_vcs_mcp.main` as the command in your MCP config:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "python",
      "args": ["-m", "msaccess_vcs_mcp.main"]
    }
  }
}
```

### Troubleshooting

If the MCP server doesn't load:

1. **Check MCP logs** -- In Cursor, open the Command Palette (`Ctrl+Shift+P`) and look for MCP-related output. In Claude Code, check the terminal output.

2. **Verify the command works** -- run `uvx msaccess-vcs-mcp --help` in your terminal.

3. **Check your `.env` file** -- make sure it's in the project root, the database path is correct, and there are no syntax errors.

## Security Model

### Write Operations

Database write operations (import, rebuild, run VBA) are **enabled by default**. To prevent modifications:

```bash
ACCESS_VCS_DISABLE_WRITES=true
```

### VBA Execution Permissions

VBA execution has two tiers with separate permissions controlled via add-in options:

- **`vcs_call_vba`**: Calls existing named functions via `Application.Run`. Controlled by `McpAllowCallVBA` (default: True).
- **`vcs_run_vba`**: Executes agent-generated code via temp module. Controlled by `McpAllowRunVBA` (default: False). Must be explicitly enabled by the user in the VCS Options form — agents cannot enable this option programmatically via `vcs_set_option`.

### Path Validation

All file paths are validated to prevent:
- Access to system directories
- Invalid file extensions
- Non-existent paths (unless creating)

### Code Execution Audit Trail

When usage logging is enabled (`ACCESS_VCS_ENABLE_LOGGING=true`), every call to `vcs_execute_sql`, `vcs_call_vba`, or `vcs_run_vba` writes a `"code_execution"` event **before** the code runs. This entry preserves the full, untruncated SQL or VBA text alongside the target database path — providing a forensic record even if the process crashes during execution.

```jsonc
// Example log entry (written before execution)
{
  "event": "code_execution",
  "tool": "vcs_execute_sql",
  "database": "C:\\Projects\\mydb.accdb",
  "code_type": "sql",
  "code": "SELECT * FROM Customers WHERE Active = True",
  "timestamp": "2026-04-15T18:30:00+00:00",
  "version": "0.1.0"
}
```

The standard post-execution `"tool_call"` event follows with success/error status and timing.

### Safe Workflow

For safety when working with production databases:
1. Set `ACCESS_VCS_DISABLE_WRITES=true` in production
2. Use export operations to review changes
3. Enable writes only when ready to merge changes

## Development

### Running Tests

Tests must run inside the project virtual environment:

```bash
# Activate the virtual environment first
.\venv\Scripts\Activate.ps1          # Windows PowerShell
# source venv/bin/activate           # macOS/Linux

# Run all tests
pytest

# Run with coverage report
pytest --cov=msaccess_vcs_mcp --cov-report=html

# Skip integration tests (require Access installed)
pytest -m "not integration"
```

### Project Structure

```
msaccess-vcs-mcp/
├── src/
│   └── msaccess_vcs_mcp/
│       ├── __init__.py
│       ├── main.py              # MCP server entry point
│       ├── cli.py               # Terminal client with live progress
│       ├── tools.py             # MCP tool definitions
│       ├── config.py            # Configuration management
│       ├── usage_logging.py     # Structured JSONL usage logging
│       ├── security.py          # Path validation & safety
│       ├── validation.py        # Component validation
│       ├── addin_integration.py # VCS add-in integration
│       ├── operation_manager.py # Async operation tracking
│       ├── rebuild_watcher.py   # Event-driven add-in rebuild wait
│       ├── callback_server.py   # HTTP callback server
│       └── access_com/
│           ├── connection.py    # COM connection management
│           └── dao_helpers.py   # DAO utility functions
├── tests/                       # Test suite
└── docs/                        # Documentation
    ├── AGENT_WORKFLOWS.md       # AI agent usage patterns
    ├── VBA_INTEGRATION.md       # Add-in integration details
    ├── VBA_CALLBACK_API.md      # Async callback protocol
    ├── EXPORT_FORMATS.md        # Export format reference
    └── TESTING.md               # Testing guide
```

## License

MIT License - see LICENSE file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Related Tools

- **[db-inspector-mcp](https://github.com/joyfullservice/db-inspector-mcp)**: Cross-database MCP server for introspection and migration validation
- **[MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-integration)**: The VBA add-in that powers all export/import operations
