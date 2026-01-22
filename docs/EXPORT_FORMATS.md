# Export Formats

This document describes the file formats used when exporting Access database objects to source files.

## Overview

The export process converts Access database objects into text-based files that are:
- **Git-friendly**: Easily tracked, diffed, and merged
- **Human-readable**: Can be reviewed and edited in any text editor
- **Deterministic**: Consistent output for consistent inputs
- **Well-structured**: Organized by object type

## Directory Structure

```
output_dir/
├── database.json          # Metadata about the export
├── queries/              # SQL query definitions
│   ├── QueryName1.sql
│   ├── QueryName2.sql
│   └── ...
├── modules/              # VBA module code
│   ├── ModuleName1.bas
│   ├── ModuleName2.bas
│   └── ...
├── tables/               # Table schemas (future)
│   └── ...
├── forms/                # Form definitions (future)
│   └── ...
└── reports/              # Report definitions (future)
    └── ...
```

## Format Details

### Metadata File (`database.json`)

The root metadata file contains information about the export:

```json
{
  "version": "1.0",
  "database_name": "MyDatabase.accdb",
  "database_path": "C:\\path\\to\\database.accdb",
  "exported_date": "2026-01-20T10:30:00.123456",
  "object_counts": {
    "tables": 15,
    "queries": 42,
    "forms": 8,
    "reports": 5,
    "modules": 3,
    "macros": 2
  }
}
```

**Fields:**
- `version`: Format version (for future compatibility)
- `database_name`: Original database filename
- `database_path`: Full path to source database
- `exported_date`: ISO 8601 timestamp of export
- `object_counts`: Count of each object type exported

### Query Files (`.sql`)

Query files contain the SQL definition with metadata header:

```sql
-- Query: CustomersByRegion
-- Type: Select
-- Exported: 2026-01-20T10:30:00

SELECT 
    Customers.CustomerID,
    Customers.CompanyName,
    Customers.Region
FROM Customers
WHERE Customers.Active = True
ORDER BY Customers.Region, Customers.CompanyName;
```

**Header Format:**
- `-- Query:` Object name
- `-- Type:` Query type (Select, Union, PassThrough, etc.)
- `-- Exported:` Export timestamp
- Blank line separates header from SQL

**SQL Format:**
- Original SQL from Access (not reformatted)
- Preserves indentation and spacing
- Ends with newline

**Filename Convention:**
- `QueryName.sql` - sanitized to remove invalid filename characters
- Characters `<>:"/\|?*` are replaced with `_`

### VBA Module Files (`.bas`)

Module files contain VBA code exported using Access's native export format:

```vba
Attribute VB_Name = "ModuleName"
Option Compare Database
Option Explicit

'==================================================================
' Module: ModuleName
' Description: Utility functions for data processing
'==================================================================

Public Function GetCurrentFiscalYear() As Integer
    ' Returns the current fiscal year
    If Month(Date) >= 7 Then
        GetCurrentFiscalYear = Year(Date) + 1
    Else
        GetCurrentFiscalYear = Year(Date)
    End If
End Function
```

**Format:**
- Uses Access's native `.bas` export format
- Includes VBA attributes (`Attribute VB_Name`, etc.)
- Preserves all code formatting and comments
- Includes Option statements

**Filename Convention:**
- `ModuleName.bas`
- Standard and Class modules included
- Form/Report modules in future versions

### Table Schema Files (`.sql`) - Future

Table definition files will contain DDL with metadata:

```sql
-- Table: Customers
-- Rows: 1234
-- Created: 2025-01-15T08:00:00
-- Modified: 2026-01-20T09:45:00

CREATE TABLE [Customers] (
    [CustomerID] COUNTER CONSTRAINT [PK_Customers] PRIMARY KEY,
    [CompanyName] TEXT(255) NOT NULL,
    [ContactName] TEXT(100),
    [Region] TEXT(50),
    [Active] BIT DEFAULT True,
    [CreatedDate] DATETIME DEFAULT Now()
);

-- Indexes
CREATE INDEX [IX_Customers_Region] ON [Customers] ([Region]);
CREATE INDEX [IX_Customers_Active] ON [Customers] ([Active]);
```

### Form/Report Files - Future

Two options will be supported:

**Option 1: Text Format (SaveAsText)**
```
Version =21
VersionRequired =20
Begin Form
    Caption ="CustomerForm"
    ...
End
```

**Option 2: XML Format**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Access>
  <Form Name="CustomerForm">
    ...
  </Form>
</Access>
```

## File Encoding

### Default Encoding: UTF-8 with BOM

Files are exported using **UTF-8 with BOM** (`utf-8-sig`):
- **UTF-8**: Unicode support for all characters
- **BOM** (Byte Order Mark): Ensures Windows tools recognize UTF-8
- Better compatibility with Windows text editors

### Alternative Encodings

Can be configured via `ACCESS_VCS_ENCODING`:
- `utf-8-sig` (default): UTF-8 with BOM
- `utf-8`: UTF-8 without BOM
- `utf-16`: UTF-16 with BOM (Windows standard)
- `cp1252`: Windows Latin-1

## Line Endings

### Windows Line Endings (CRLF)

Files use Windows line endings (`\r\n`):
- Native format for Windows/Access
- Expected by most Windows tools
- Git can normalize these if configured

### Git Configuration

Recommended `.gitattributes` for mixed-platform teams:

```gitattributes
# Text files - normalize to LF in repo, checkout as native
*.sql text
*.bas text
*.json text
*.md text

# Or enforce CRLF everywhere (Windows-only teams)
*.sql text eol=crlf
*.bas text eol=crlf
```

## Filename Sanitization

Object names may contain characters invalid in filenames. These are sanitized:

| Character | Replacement | Reason |
|-----------|-------------|--------|
| `<` | `_` | Invalid filename character |
| `>` | `_` | Invalid filename character |
| `:` | `_` | Drive letter separator |
| `"` | `_` | String delimiter |
| `/` | `_` | Path separator |
| `\` | `_` | Path separator |
| `|` | `_` | Pipe character |
| `?` | `_` | Wildcard |
| `*` | `_` | Wildcard |

**Example:**
- Object: `Customer/Supplier Analysis`
- Filename: `Customer_Supplier Analysis.sql`

## Deterministic Export

Exports are deterministic to enable meaningful diffs:

1. **Consistent Ordering**: Objects exported in alphabetical order
2. **Stable Formatting**: SQL and VBA code not reformatted
3. **Reproducible Metadata**: Timestamps optional (future)

## Future Enhancements

### Metadata Options

Future versions may support:
- **Minimal mode**: Exclude timestamps for cleaner diffs
- **Detailed mode**: Include creation dates, last modified, object properties
- **Hash mode**: Include MD5/SHA256 of object content

### Additional Formats

Planned support for:
- **XML export**: More structured but verbose
- **JSON export**: For programmatic processing
- **Markdown documentation**: Auto-generated from object properties

### Compression

For large databases:
- **Archived exports**: `.tar.gz` or `.zip` of entire export
- **Incremental exports**: Only changed objects
- **Differential format**: Store only differences

## Best Practices

### Git Integration

1. **Add `.gitattributes`** for proper line ending handling
2. **Exclude binary files** in `.gitignore` (`.accdb`, `.mdb`)
3. **Use meaningful commit messages** describing database changes
4. **Review diffs** before committing to catch unintended changes

### Project Structure

```
MyProject/
├── .gitignore
├── .gitattributes
├── database.accdb           # Not in version control
├── src/
│   └── database/           # Version controlled
│       ├── database.json
│       ├── queries/
│       ├── modules/
│       └── ...
└── README.md
```

### Workflow

1. **Export regularly**: After significant changes
2. **Review changes**: Use `git diff` to see what changed
3. **Document changes**: Commit with clear messages
4. **Track branches**: Use git branches for experimental changes

## Troubleshooting

### Encoding Issues

**Problem**: Special characters appear as �
**Solution**: Verify `ACCESS_VCS_ENCODING=utf-8-sig`

### Line Ending Issues

**Problem**: Git shows entire file as changed
**Solution**: Configure `.gitattributes` properly

### Filename Collisions

**Problem**: Two objects with similar names create same filename
**Solution**: Rename objects in Access to be more distinct

### Large Exports

**Problem**: Export takes too long or creates huge files
**Solution**: Export only changed object types using `object_types` parameter
