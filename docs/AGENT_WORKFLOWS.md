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
access_export_database("C:\\mydb.accdb", "C:\\mydb.src")

# 2. Read the query source file
query_sql = read_file("C:\\mydb.src\\queries\\CustomerReport.sql")

# 3. Modify the SQL
# (AI agent makes changes to the SQL)

# 4. Write the updated query
write_file("C:\\mydb.src\\queries\\CustomerReport.sql", updated_sql)

# 5. Merge changes back into database
access_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

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
access_export_database(
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
access_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 6. Test the functions
# (Open VBE and test the new/updated functions)
```

**Tips:**
- Module files are in `modules/` folder with `.bas` (standard modules) or `.cls` (class modules) extensions
- Preserve the `Attribute VB_Name` header
- Use Option Explicit for type safety
- Include XML doc comments for functions

### 3. Create a New Database Object

**Use case:** Add a new query, module, or other object

**Steps:**
```python
# 1. Export existing database
access_export_database("C:\\mydb.accdb", "C:\\mydb.src")

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
access_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 4. Verify the object was created
objects = access_list_objects("C:\\mydb.accdb")
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
result = access_export_database("C:\\mydb.accdb", "C:\\mydb.src")

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
access_export_database("C:\\mydb.accdb", "C:\\mydb.src")
git_add("C:\\mydb.src")
git_commit("My changes before pull")

# 2. Pull changes from remote
git_pull()

# 3. Review what changed
diff_result = access_diff_database("C:\\mydb.accdb", "C:\\mydb.src")
print("Modified objects:")
print(diff_result)

# 4. Merge source changes into database
access_import_objects("C:\\mydb.accdb", "C:\\mydb.src")

# 5. Test the merged result
# (Verify database functions correctly)

# 6. Export and commit if merge succeeded
access_export_database("C:\\mydb.accdb", "C:\\mydb.src")
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
# 1. Build from source files
result = access_rebuild_database(
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
objects = access_list_objects("C:\\builds\\mydb_v1.0.accdb")
print(f"Built database contains {len(objects['queries'])} queries")
```

**Tips:**
- Build from source creates a fresh database
- Use for deployments and releases
- Review Build.log for any issues
- Test thoroughly before distributing

### 7. Iterative Development Cycle

**Use case:** Rapid development with frequent testing

**Steps:**
```python
def develop_feature(db_path, src_path, feature_name):
    """Iterative development cycle for a feature."""
    
    while not feature_complete:
        # 1. Export current state
        access_export_database(db_path, src_path, object_types=["modules"])
        
        # 2. Make incremental changes
        # (AI agent modifies code)
        
        # 3. Merge changes
        access_import_objects(db_path, src_path)
        
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
result = access_export_database(db_path, src_path)

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
