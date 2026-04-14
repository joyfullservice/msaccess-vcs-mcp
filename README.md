# msaccess-vcs-mcp

A lightweight MCP server that provides AI agent integration for the [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-integration). This tool acts as a bridge, allowing AI assistants to work with Microsoft Access databases through version control operations.

## Architecture

This MCP tool is a **lightweight wrapper** around the MSAccess VCS add-in, delegating all export/import/build operations to the battle-tested add-in:

```
AI Agent → MCP Tool (Python) → VCS Add-in (VBA) → Access Database
              ↓
       Path validation
       Permission checks
       Result formatting
```

## Features

- **Full Export**: Export all database objects (forms, reports, queries, modules, macros, etc.)
- **Fast Save**: Incremental exports (only changed objects)
- **Merge Build**: Import source changes into existing database
- **Build from Source**: Create fresh database from source files
- **Object Inventory**: List all objects in an Access database
- **Change Tracking**: Compare database against source files
- **Write Operations**: Enabled by default, optional disable for safety
- **Safety Guardrails**: Path validation and permission checks

All export/import operations leverage the comprehensive MSAccess VCS add-in, ensuring consistency and reliability.

## Why This Tool?

While SQL Server, PostgreSQL, and other databases have excellent MCP tools, **Microsoft Access is underserved**. This tool fills that gap by:

- **AI Agent Integration**: Allows AI assistants to work with Access databases
- **Proven Export/Import**: Uses the MSAccess VCS add-in with years of development
- **Complete Feature Set**: All Access object types supported (forms, reports, queries, modules, macros, tables)
- **Version Control Friendly**: Text-based exports work seamlessly with git
- **Collaborative Development**: Multiple developers can work independently and merge changes

## Prerequisites

- **Python**: 3.10 or higher
- **Microsoft Access**: Installed on Windows (for COM automation)
- **pywin32**: Python COM interface (installed automatically)
- **MSAccess VCS Add-in**: Must be installed ([download latest release](https://github.com/joyfullservice/msaccess-vcs-integration/releases/latest))

## Installation

### Basic Installation

```bash
# Clone the repository
git clone https://github.com/joyfullservice/msaccess-vcs-mcp.git
cd msaccess-vcs-mcp

# Create and activate a virtual environment (recommended)
python -m venv venv

# On Windows:
venv\Scripts\activate

# Install in development mode (editable install)
pip install -e ".[dev]"
```

**Note:** The `-e` flag installs the package in "editable" mode, which means changes to the source code are immediately reflected without needing to reinstall.

### Windows PATH Configuration

After installation, you may see warnings about scripts not being on PATH. If you need to run `msaccess-vcs-mcp` from the command line, add the Python Scripts directory to your PATH:

**PowerShell (run as Administrator):**
```powershell
$scriptsPath = "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\Scripts"
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($currentPath -notlike "*$scriptsPath*") {
    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$scriptsPath", "User")
}
```

After updating PATH, restart your terminal (or Cursor) for changes to take effect.

### Install MSAccess VCS Add-in

**Important:** This MCP tool requires the MSAccess VCS add-in to be installed.

1. Download the [latest release](https://github.com/joyfullservice/msaccess-vcs-integration/releases/latest) from GitHub
2. Extract `Version Control.accda` to a trusted location
3. Open `Version Control.accda` to launch the installer
4. Click **Install Add-In**

The add-in will be installed to `%AppData%\MSAccessVCS\Version Control.accda` by default.

### Environment Configuration

Configure your target database by creating a `.env` file:

```bash
# Create .env file in project root
# Add your database path:

ACCESS_VCS_DATABASE=C:\Projects\MyApp\database.accdb

# Optional: Use custom add-in path
# ACCESS_VCS_ADDIN_PATH=C:\Custom\Path\Version Control.accda

# Optional: Disable database writes for safety
# ACCESS_VCS_DISABLE_WRITES=true
```

The tool will work with the database specified in `ACCESS_VCS_DATABASE`. You can also pass database paths directly to tool functions.

## Quick Start

Get up and running with msaccess-vcs-mcp in Cursor in just a few steps:

### Step 1: Install the Package

Make sure the package is installed (if you haven't already):

```bash
pip install -e ".[dev]"
```

Verify the command is available:

```bash
msaccess-vcs-mcp --help
```

### Step 2: Configure Cursor

The MCP configuration file (`.cursor/mcp.json`) is already created in this repository:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "msaccess-vcs-mcp",
      "env": {
        "ACCESS_VCS_DATABASE": "C:\\Projects\\MyApp\\database.accdb"
      }
    }
  }
}
```

**Note:** If you're using this in a different project, create `.cursor/mcp.json` in your project root with the above content.

### Step 3: Set Up Your Configuration

Create a `.env` file in your project root with your database path:

```bash
# Target database for your project
ACCESS_VCS_DATABASE=C:\Projects\MyApp\database.accdb

# Optional: Disable writes for safety
# ACCESS_VCS_DISABLE_WRITES=true

# Optional: Custom add-in path
# ACCESS_VCS_ADDIN_PATH=C:\Custom\Path\Version Control.accda
```

### Step 4: Restart Cursor

After creating the `.env` file:
1. **Close Cursor completely** (not just the window - fully quit the application)
2. **Reopen Cursor** in your project directory
3. Cursor will automatically detect and load the MCP server

### Step 5: Test the Tool

Once Cursor restarts, the MCP tools will be available. You can test them by asking the AI assistant in Cursor to:

- **List database objects:**
  > "Can you list all objects in my Access database at C:\\mydb.accdb?"

- **Export database to source:**
  > "Export my Access database to the src folder"

- **Check what's changed:**
  > "Compare my database against the source files and tell me what changed"

## Configuration

### Environment Variables

| Variable | Description | Default | Required |
|----------|-------------|---------|----------|
| `ACCESS_VCS_DATABASE` | Target database path | - | Recommended |
| `ACCESS_VCS_ADDIN_PATH` | Path to VCS add-in file | `%AppData%\MSAccessVCS\Version Control.accda` | No |
| `ACCESS_VCS_DISABLE_WRITES` | Disable database modifications | `false` | No |

### Project-Specific Configuration

This tool supports per-project configurations using `.env` files:

#### `.cursor/mcp.json` (Version Controlled)

Store non-sensitive settings that can be shared with your team:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "msaccess-vcs-mcp",
      "env": {
        "ACCESS_VCS_DATABASE": "C:\\Projects\\MyApp\\database.accdb"
      }
    }
  }
}
```

#### `.env` (Gitignored, Project-Specific)

Store project-specific database path:

```bash
# Target database for this project
ACCESS_VCS_DATABASE=C:\Projects\MyApp\database.accdb

# Optional: Disable writes for safety
# ACCESS_VCS_DISABLE_WRITES=true
```

#### `.env.local` (Optional, Gitignored)

For personal overrides that shouldn't affect the team:

```bash
# Override for local development
ACCESS_VCS_DATABASE=C:\Users\YourName\dev\testdb.accdb
```

### Environment Variable Precedence

1. **MCP server `env` section** (highest priority) - values in `.cursor/mcp.json`
2. **`.env.local`** - personal overrides
3. **`.env`** - project-specific configuration (lowest priority)

### Default Behavior

- **Database writes**: Enabled by default. The add-in can modify your database.
- **Add-in path**: Auto-detected at `%AppData%\MSAccessVCS\Version Control.accda`
- **Target database**: Must be specified via `ACCESS_VCS_DATABASE` or passed to each tool function

## Available Tools

### `vcs_export_database(database_path, output_dir, object_types)`

Export Access database objects to source files via the VCS add-in.

Supports all Access object types: tables, queries, forms, reports, modules, macros, and more. Uses fast save by default (only exports changed objects).

**Args:**
- `database_path`: Path to Access database (.accdb, .accda, .mdb)
- `output_dir`: Directory to export source files to
- `object_types`: Optional list of types (defaults to all types)

**Returns:**
- `success`: Whether export succeeded
- `exported_count`: Number of objects exported
- `export_path`: Path where files were written
- `objects_by_type`: Breakdown of what was exported
- `log_path`: Path to Export.log file from add-in

**Example:**
```python
# Export entire database
vcs_export_database("C:\\db.accdb", "C:\\src\\mydb")

# Export only VBA modules
vcs_export_database(
    "C:\\db.accdb", 
    "C:\\src\\mydb",
    object_types=["modules"]
)
```

### `vcs_list_objects(database_path)`

List all objects in an Access database.

**Args:**
- `database_path`: Path to Access database

**Returns:**
- `tables`: List of table names
- `queries`: List of query names with types
- `modules`: List of module names
- `forms`: List of form names (future)
- `reports`: List of report names (future)

**Example:**
```python
vcs_list_objects("C:\\db.accdb")
```

### `vcs_diff_database(database_path, source_dir, show_details)`

Compare database objects against source files.

**Args:**
- `database_path`: Path to Access database
- `source_dir`: Directory containing source files
- `show_details`: If True, show detailed diff (future)

**Returns:**
- `queries`: What's different in queries
- `modules`: What's different in modules
- `new_in_db`: Objects in database but not in source
- `new_in_source`: Objects in source but not in database

**Example:**
```python
vcs_diff_database("C:\\db.accdb", "C:\\src\\mydb")
```

### `vcs_import_objects(database_path, source_dir, object_types, overwrite)`

Import objects from source files into Access database using merge build.

Merges source file changes into the existing database without requiring a full rebuild.

**Note:** Database writes are enabled by default. Set `ACCESS_VCS_DISABLE_WRITES=true` to prevent modifications.

### `vcs_rebuild_database(source_dir, output_path, template_path)`

Build a complete Access database from source files.

Creates a fresh database from source files, useful for clean builds and distribution.

**Note:** Database writes are enabled by default. Set `ACCESS_VCS_DISABLE_WRITES=true` to prevent modifications.

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

### Example Files

**Query** (`queries/QueryName.sql`):
```sql
SELECT * FROM Customers WHERE Active = True
```

**VBA Module** (`modules/ModuleName.bas`):
```vba
Attribute VB_Name = "ModuleName"
Option Compare Database
Option Explicit

Public Function MyFunction() As Boolean
    MyFunction = True
End Function
```

All files use **UTF-8 with BOM** encoding, which is critical for proper import back into Access.

## Workflows

### Version Control Workflow

```bash
# 1. Export database to source
vcs_export_database("C:\\db.accdb", "C:\\src\\db")

# 2. Initialize git (if not already done)
cd C:\src\db
git init
git add .
git commit -m "Initial export of database"

# 3. Make changes in Access...

# 4. See what changed
vcs_diff_database("C:\\db.accdb", "C:\\src\\db")

# 5. Export changes
vcs_export_database("C:\\db.accdb", "C:\\src\\db")

# 6. Commit changes
git add .
git commit -m "Updated customer queries"
```

### Integration with db-inspector-mcp

Both tools work together for comprehensive database workflows:

```python
# 1. Analyze database structure (db-inspector-mcp)
db_list_tables(database="legacy")
db_list_views(database="legacy")

# 2. Export to version control (msaccess-vcs-mcp)
vcs_export_database("C:\\legacy.accdb", "C:\\src\\legacy-db")

# 3. Make changes to source files...

# 4. Validate changes (db-inspector-mcp)
db_compare_queries(
    "SELECT * FROM Customers",
    "SELECT * FROM Customers",
    database1="legacy",
    database2="test"
)

# 5. Check what changed (msaccess-vcs-mcp)
vcs_diff_database("C:\\legacy.accdb", "C:\\src\\legacy-db")
```

## Development

### Complete Setup

1. **Clone and navigate to the repository:**
   ```bash
   git clone https://github.com/joyfullservice/msaccess-vcs-mcp.git
   cd msaccess-vcs-mcp
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv venv
   
   # Windows:
   venv\Scripts\activate
   ```

3. **Install the package with development dependencies:**
   ```bash
   pip install -e ".[dev]"
   ```

4. **Run tests:**
   ```bash
   # Run all tests
   pytest
   
   # Run with coverage report
   pytest --cov=msaccess_vcs_mcp --cov-report=html
   
   # Run specific test file
   pytest tests/test_config.py
   ```

### Project Structure

```
msaccess-vcs-mcp/
├── src/
│   └── msaccess_vcs_mcp/
│       ├── __init__.py
│       ├── main.py              # MCP server entry point
│       ├── tools.py             # MCP tool definitions
│       ├── config.py            # Configuration management
│       ├── security.py          # Path validation & safety
│       ├── addin_integration.py # VCS add-in integration
│       └── access_com/
│           ├── connection.py    # COM connection management
│           └── dao_helpers.py   # DAO utility functions
├── tests/                       # Test suite
└── docs/                        # Documentation
    ├── AGENT_WORKFLOWS.md       # AI agent usage patterns
    ├── VBA_INTEGRATION.md       # Add-in integration details
    ├── EXPORT_FORMATS.md        # Export format reference
    └── TESTING.md               # Testing guide
```

## Security Model

### Write Operations

Database write operations (import, rebuild) are **enabled by default**. To prevent modifications:

```bash
ACCESS_VCS_DISABLE_WRITES=true
```

### Path Validation

All file paths are validated to prevent:
- Access to system directories
- Invalid file extensions
- Non-existent paths (unless creating)

### Safe Workflow

For safety when working with production databases:
1. Set `ACCESS_VCS_DISABLE_WRITES=true` in production
2. Use export operations to review changes
3. Enable writes only when ready to merge changes

## Current Status

This MCP tool is **production-ready** as a lightweight wrapper around the MSAccess VCS add-in.

### Completed Features ✓
- [x] VCS add-in integration via COM automation
- [x] Full database export (all object types)
- [x] Fast save (incremental exports)
- [x] Merge build (import changes into existing database)
- [x] Build from source (create fresh database)
- [x] Object inventory and diff operations
- [x] Comprehensive documentation for AI agents
- [x] Unit and integration tests

### Architecture Benefits ✓
- **Leverage proven technology**: Uses the battle-tested MSAccess VCS add-in
- **Complete feature set**: All Access object types supported out of the box
- **Future-proof**: Add-in updates automatically benefit the MCP tool
- **Minimal maintenance**: ~90% less code than reimplementing everything

### Future Enhancements

Potential improvements that could be added:

- **Enhanced diff**: Detailed line-by-line comparison of changes
- **Git integration**: Auto-commit after export, pull before merge
- **Conflict resolution UI**: Help agents resolve merge conflicts
- **VBA code analysis**: Parse and understand VBA for better assistance
- **Schema migrations**: Help agents plan and execute schema changes

These features would build upon the existing add-in integration, not replace it.

## License

MIT License - see LICENSE file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Acknowledgments

- Wraps the excellent [MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-integration) by Adam Waller
- Architecture inspired by [db-inspector-mcp](https://github.com/joyfullservice/db-inspector-mcp)
- Built with [FastMCP](https://github.com/modelcontextprotocol/mcp)
- Fills a critical gap in Access database tooling for AI agents

## Related Tools

- **[db-inspector-mcp](https://github.com/joyfullservice/db-inspector-mcp)**: Database introspection and migration validation
- Works seamlessly with msaccess-vcs-mcp for complete Access workflows
