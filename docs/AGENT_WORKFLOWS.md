# AI Agent Workflows for Microsoft Access Development

This guide documents common workflows for AI agents working with Microsoft Access databases using the msaccess-vcs-mcp tool.

## Overview

The MCP tool enables AI agents to work iteratively with Access databases by:
1. Exporting database objects to text-based source files
2. Modifying source files (queries, VBA modules, etc.)
3. Merging changes back into the database
4. Testing and iterating

## Core Workflow Pattern

All Access development workflows follow this pattern:

```
┌─────────────┐
│  Database   │
└──────┬──────┘
       │ Export
       ↓
┌─────────────┐
│Source Files │ ← AI Agent reads/edits
└──────┬──────┘
       │ Merge Build
       ↓
┌─────────────┐
│  Database   │ ← Testing
└──────┬──────┘
       │ Iterate
       ↓
```

## Common Workflows

### 1. Modify a Query

**Use case:** Update SQL logic, fix bugs, or optimize a query

**Steps:**
```python
# 1. Export the database to source files
vcs_export_database("C:\\mydb.accdb", "C:\\mydb.src")

# 2. Read the query source file
query_sql = read_file("C:\\mydb.src\\queries\\CustomerReport.sql")

# 3. Modify the SQL
# (AI agent makes changes to the SQL)

# 4. Write the updated query
write_file("C:\\mydb.src\\queries\\CustomerReport.sql", updated_sql)

# 5. Merge changes back into database
vcs_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 6. Test the query
# (Open database and run query to verify)
```

**Tips:**
- Query files are in `queries/` folder with `.sql` extension
- Include SQL comments for context
- Test with sample data before committing

### 2. Add or Update VBA Code

**Use case:** Create new functions, fix bugs, or refactor VBA modules

**Steps:**
```python
# 1. Export VBA modules only (faster)
vcs_export_database(
    "C:\\mydb.accdb", 
    "C:\\mydb.src",
    object_types=["modules"]
)

# 2. Read the module source file
module_code = read_file("C:\\mydb.src\\modules\\Utilities.bas")

# 3. Add or modify VBA code
# (AI agent makes changes to the module)

# 4. Write the updated module
write_file("C:\\mydb.src\\modules\\Utilities.bas", updated_code)

# 5. Merge changes back
vcs_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 6. Compile to validate
result = vcs_compile_vba("C:\\mydb.accdb")
if not result["success"]:
    # Stop — ask the user to Debug → Compile in the VBE and paste the
    # code snippet around the highlighted error line. See agent_guidance.
    pass

# 7. Test the functions
# (Open VBE and test the new/updated functions)
```

**Tips:**
- Module files are in `modules/` folder with `.bas` (standard modules) or `.cls` (class modules) extensions
- Preserve the `Attribute VB_Name` header
- Use Option Explicit for type safety
- Include XML doc comments for functions
- If `vcs_compile_vba` fails, stop editing and ask the user to **Debug → Compile** in the VBE, then paste the code snippet around the highlighted line

### 3. Create a New Database Object

**Use case:** Add a new query, module, or other object

**Steps:**
```python
# 1. Export existing database
vcs_export_database("C:\\mydb.accdb", "C:\\mydb.src")

# 2. Create new source file
# For a new query:
new_query = """-- Query: NewCustomerList
-- Type: Select
-- Exported: 2026-01-20T10:30:00

SELECT CustomerID, CompanyName, ContactName
FROM Customers
WHERE Active = True
ORDER BY CompanyName
"""

write_file("C:\\mydb.src\\queries\\NewCustomerList.sql", new_query)

# 3. Merge into database
vcs_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 4. Verify the object was created
objects = vcs_list_objects("C:\\mydb.accdb")
print(objects["queries"])
```

**Tips:**
- Follow existing file naming conventions
- Include metadata headers for queries
- Use descriptive names
- Test immediately after creation

### 4. Bulk Export for Version Control

**Use case:** Initial export to git or periodic full export

**Steps:**
```python
# 1. Full export to source files
result = vcs_export_database("C:\\mydb.accdb", "C:\\mydb.src")

# 2. Review what was exported
print(f"Exported {result['exported_count']} objects")
for obj_type, count in result['objects_by_type'].items():
    print(f"  {obj_type}: {count}")

# 3. Commit to version control (using git tools)
git_add("C:\\mydb.src")
git_commit("Initial database export")
```

**Tips:**
- First export is always a full export
- Subsequent exports use "fast save" (only changed objects)
- Review Export.log for details
- Commit frequently for granular history

### 5. Pull Changes and Merge

**Use case:** Integrate changes from other developers

**Steps:**
```python
# 1. Commit any local changes first
vcs_export_database("C:\\mydb.accdb", "C:\\mydb.src")
git_add("C:\\mydb.src")
git_commit("My changes before pull")

# 2. Pull changes from remote
git_pull()

# 3. Review what changed
diff_result = vcs_diff_database("C:\\mydb.accdb", "C:\\mydb.src")
print("Modified objects:")
print(diff_result)

# 4. Merge source changes into database
vcs_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 5. Test the merged result
# (Verify database functions correctly)

# 6. Export and commit if merge succeeded
vcs_export_database("C:\\mydb.accdb", "C:\\mydb.src")
git_add("C:\\mydb.src")
git_commit("Merged changes from team")
```

**Tips:**
- Always commit before pulling
- Review diffs carefully
- Test thoroughly after merge
- Resolve conflicts in source files, not in Access

### 6. Build Fresh Database from Source

**Use case:** Clean build, deployment, or distribution

**Steps:**
```python
# 1. Build from source files (omit the output path to build to the file
#    the source files name; a different path needs add-in 5.1.1 or later)
result = vcs_rebuild_database(
    "C:\\mydb.src",
    "C:\\builds\\mydb_v1.0.accdb"
)

# 2. Verify build succeeded
if result["success"]:
    print(f"Database built: {result['output_path']}")
    print(f"Build log: {result['log_path']}")
else:
    print(f"Build failed: {result['error']}")

# 3. Test the built database
objects = vcs_list_objects("C:\\builds\\mydb_v1.0.accdb")
print(f"Built database contains {len(objects['queries'])} queries")
```

**Tips:**
- Build from source creates a fresh database
- Use for deployments and releases
- Review Build.log for any issues
- Test thoroughly before distributing

### 6b. Rebuild the VCS add-in from source

**Use case:** You edited add-in source (for example `clsQueryComposer.cls`) and need the running add-in to pick up those changes without waiting for a person.

This is **not** `vcs_rebuild_database`. That tool rebuilds a user project. The add-in rebuilds itself through `vcs_rebuild_addin`.

**Preconditions:** no other `MSACCESS.EXE` in the Windows session may hold a file the rebuild replaces — the installed add-in or the build target. An instance with an unrelated database open does not block it; one that loaded the add-in does, as does one that cannot be asked. The guard reports them but closes nothing. The repo folder must be a trusted location and the helper script enabled.

**Steps:**
```python
result = vcs_rebuild_addin(
    r"C:\path\to\msaccess-vcs-addin\Version Control.accda.src"
)
# Returns when this attempt is complete, refused, or *-failed.
# Do not poll rebuild-status.json yourself unless the tool timed out.
```

The tool derives the development copy beside the source folder. Its callback
URL and operation ID cross the disconnected worker boundary, so the builder
Access process emits the same detailed HTTP `Log.Add` / `Log.Progress` stream
as an export or merge. The status file remains authoritative for compile,
install, terminal failure, and recovery after the builder exits. Rebuilding
the add-in is a repository operation and belongs to the repository's own copy,
which closes itself once the worker handoff is confirmed. Do not open a user
database, anything in the repository's `Testing` folder, or a scratch `.accdb`.
The installed add-in is refused here as it is for every tool — see the note
under 6c.

MCP progress is best-effort in Cursor 3.13. For live terminal output:

```text
msaccess-vcs rebuild-addin "C:\path\to\msaccess-vcs-addin\Version Control.accda.src"
```

Keep the CLI in the foreground so its stream stays in the primary chat. It
exits when the operation reaches terminal status; that process exit is the
completion signal. Do not background it just to wait on a notification, and
do not add a second timer wait, fixed-duration sleep, or
`rebuild-status.json` poll after it has already finished.

**Tips:**
- `refused` and `launch-failed` come back immediately; nothing was rebuilt
- A `refused` result lists each other process in `otherInstances`; close those yourself and call again
- `launch-failed` means the helper script never started — Access stays open and the call is safe to retry
- `compile-failed` leaves Access open on the rebuilt file for Debug > Compile
- After `complete`, later MCP calls load the newly installed add-in
- Builder and silent-installer Access processes request full-power QoS; user-owned Access is left alone
- `vcs_call_vba(..., ["RebuildAddIn", source])` is a launch-only escape hatch
- If the tool times out, recover by reading `<source>/logs/rebuild-status.json` and matching `phaseStarted`

### 6c. Run the add-in's own test suite

**Use case:** You changed add-in source and want its own tests to confirm the rebuild before touching a user database.

**Steps:**
```python
vcs_run_tests(r"C:\path\to\msaccess-vcs-addin\Version Control.accda", filter="clsTestInstall")
```

MCP progress is best-effort in Cursor. For live per-test output:

```text
msaccess-vcs run-tests "C:\path\to\msaccess-vcs-addin\Version Control.accda" --filter clsTestInstall
```

Keep the CLI in the foreground so its stream stays in the primary chat. It
exits when the run finishes; that process exit is the completion signal.
The stream is dots for fast passes, a named line for tests ≥ 1s, and
full FAIL/ERROR/EMPTY lines. The last stdout line is a human summary such as
`Tests passed. 12 subs, 40 assertions in 1.48s`. Headless means no add-in UI
(no web runner, no console form), not a hidden Access window.

**Why the path is the development copy:** a run needs two projects and they are different files. The installed add-in loads as a library and supplies the runner and `TestAssert`; the code under test is whatever the current database holds. The runner scans the current VBA project, so the host decides which tests are found — aim the call at a user database, or anything in the repository's `Testing` folder, and you get that database's tests. Access will not bind a file moniker to an `.accda`, so the server opens the development copy as the current database explicitly; you do not need to open it first.

**The installed add-in is never a target.** No tool accepts it as `database_path`, `output_path`, or `template_path` — not this one, not export, import, rebuild, `vcs_run_vba`, or `vcs_call_vba`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. It also has no source tree beside it for the tests that read one. The check runs before the Access gate and any COM work, and returns `error_pattern: installed_addin_refused`; the add-in refuses such a run itself, so the server's refusal is the earlier of two. `vcs_get_version_info()` reports the installed version without opening anything. The comparison ignores the extension, because a compiled install is a `.accde` built from the same `.accda`.

**Tips:**
- Run through the MCP server, not from the add-in's own window; assertions route to the installed add-in while the runner lives in the calling project, so a development-copy run discards them all
- An all-`EMPTY` result (zero assertions) is a bypassed harness, not a pass
- A rebuild closes server-owned Access that held the files it replaces, so expect a cold start on the next call

### 7. Iterative Development Cycle

**Use case:** Rapid development with frequent testing

**Steps:**
```python
def develop_feature(db_path, src_path, feature_name):
    """Iterative development cycle for a feature."""
    
    while not feature_complete:
        # 1. Export current state
        vcs_export_database(db_path, src_path, object_types=["modules"])
        
        # 2. Make incremental changes
        # (AI agent modifies code)
        
        # 3. Merge changes
        vcs_import_objects(db_path, src_path)
        
        # 4. Test
        test_result = run_tests(db_path)
        
        # 5. Evaluate and iterate
        if test_result.passed:
            # Commit this iteration
            git_commit(f"Progress on {feature_name}")
        else:
            # Debug and retry
            analyze_errors(test_result.errors)
```

**Tips:**
- Export frequently to track progress
- Test each iteration
- Commit working iterations
- Use fast save for speed

## Best Practices

### File Organization

The VCS add-in exports to a structured folder:

```
mydb.src/
├── queries/          # SQL query files (.sql)
├── modules/          # VBA standard modules (.bas)
├── forms/            # Form definitions (.bas, .cls)
├── reports/          # Report definitions (.bas, .cls)
├── macros/           # Macro definitions (.bas)
├── tables/           # Table data (if enabled)
├── tbldefs/          # Table structure (.sql, .xml)
├── vcs-options.json  # Export options
└── vcs-index.json    # Fast save index
```

### Encoding

**Critical:** All source files use UTF-8 with BOM encoding.

- Always preserve UTF-8 BOM when editing files
- The add-in requires BOM for proper import
- Check file encoding before writing changes

### Error Handling

```python
# Always check for errors
result = vcs_export_database(db_path, src_path)

if not result["success"]:
    print(f"Export failed: {result.get('error')}")
    # Check if add-in is installed
    # Check database path
    # Review error message
```

### Testing

Test database changes immediately:
1. Open database in Access
2. Test affected objects
3. Run any VBA tests
4. Verify data integrity

### Version Control

Commit frequently with descriptive messages:
```bash
git commit -m "Add customer search query"
git commit -m "Fix date calculation in Utilities module"
git commit -m "Update invoice report layout"
```

## Troubleshooting

### Add-in Not Found

**Error:** `VCS add-in not found`

**Solution:**
1. Install MSAccess VCS add-in
2. Or set `ACCESS_VCS_ADDIN_PATH` in `.env`

### Import Fails

**Error:** Import/merge build fails

**Solution:**
1. Check Build.log for details
2. Verify source files are valid
3. Ensure UTF-8 BOM encoding
4. Try full rebuild if merge fails

### Objects Not Exporting

**Error:** Some objects missing from export

**Solution:**
1. Check Export.log for errors
2. Ensure objects aren't open in Access
3. Verify object names don't contain invalid characters
4. Check VCS options (vcs-options.json)

### Merge Conflicts

**Error:** Git merge conflicts in source files

**Solution:**
1. Resolve conflicts in source files (not in Access)
2. Use standard git conflict resolution
3. Test merged result in Access
4. Re-export to verify

## Advanced Patterns

### Conditional Logic Updates

When updating complex VBA logic:
1. Export current version
2. Add comprehensive comments
3. Make incremental changes
4. Test each change
5. Commit working versions

### Schema Migrations

When changing table structure:
1. Export table definitions (`tbldefs/`)
2. Modify SQL CREATE TABLE statements
3. Handle data migration separately
4. Test with sample data first
5. Document migration steps

### Form and Report Development

Forms and reports are exported but harder to edit as text:
1. Export for version control
2. Make visual changes in Access
3. Export again to capture changes
4. Commit with descriptive message

## Resources

- [MSAccess VCS Add-in Documentation](https://github.com/joyfullservice/msaccess-vcs-integration/wiki)
- [Export File Format Reference](EXPORT_FORMATS.md)
- [VBA Integration Guide](VBA_INTEGRATION.md)
- [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-integration/blob/main/Version%20Control.accda.src/AGENTS.md) - Comprehensive file structure guide

## Support

For issues or questions:
1. Check the [Wiki](https://github.com/joyfullservice/msaccess-vcs-integration/wiki)
2. Review Export.log or Build.log
3. Open an issue on GitHub
4. Include error messages and context
