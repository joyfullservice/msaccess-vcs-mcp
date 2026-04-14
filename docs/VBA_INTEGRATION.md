# MSAccess VCS Add-in Integration

This document explains how the `msaccess-vcs-mcp` tool integrates with the MSAccess VCS add-in through COM automation.

## Overview

The MCP tool is a **lightweight Python wrapper** that delegates all export/import/build operations to the MSAccess VCS add-in. This architecture provides:

1. **Proven reliability**: Leverages the battle-tested VCS add-in
2. **Complete features**: Access to all add-in functionality
3. **Consistent results**: Same export format whether using GUI or MCP
4. **Easy maintenance**: Add-in updates automatically benefit the MCP tool

## Architecture

The MCP tool integrates with the VCS add-in through COM automation:

```
┌─────────────────────────────────────────────────┐
│                  AI Agent                       │
└────────────────────┬────────────────────────────┘
                     │ MCP Protocol
                     ↓
┌─────────────────────────────────────────────────┐
│            msaccess-vcs-mcp (Python)            │
│  • Path validation                              │
│  • Permission checks                            │
│  • Parameter translation                        │
│  • Result formatting                            │
└────────────────────┬────────────────────────────┘
                     │ COM Automation
                     │ (Application.Run)
                     ↓
┌─────────────────────────────────────────────────┐
│         MSAccess VCS Add-in (VBA)               │
│  • Export all object types                      │
│  • Fast save (incremental)                      │
│  • Merge build                                  │
│  • Build from source                            │
│  • Conflict resolution                          │
└────────────────────┬────────────────────────────┘
                     │ DAO/Access APIs
                     ↓
┌─────────────────────────────────────────────────┐
│           Microsoft Access Database             │
└─────────────────────────────────────────────────┘
```

## How It Works

### 1. Loading the Add-in

The MCP tool loads the VCS add-in using COM automation:

```python
# From addin_integration.py
import win32com.client

# Connect to Access
app = win32com.client.Dispatch("Access.Application")

# Load the add-in by calling a simple function
addin_path = r"%AppData%\MSAccessVCS\Version Control"
app.Run(f'"{addin_path}.Preload"')
```

### 2. Calling Add-in Functions

The MCP tool calls add-in functions via `Application.Run`:

```python
# Format: "FullPathWithoutExtension.FunctionName"
addin_lib = r"C:\Users\...\AppData\Roaming\MSAccessVCS\Version Control"

# Call the export function
app.Run(f'"{addin_lib}.HandleRibbonCommand"', "btnExport")

# Call the merge build function
app.Run(f'"{addin_lib}.HandleRibbonCommand"', "btnMergeBuild")
```

### 3. Add-in API

The VCS add-in exposes functions through `HandleRibbonCommand`:

**VBA side (from modAPI.bas):**
```vba
Public Function HandleRibbonCommand(strCommand As String, _
    Optional strArgument As String) As Boolean
    
    ' Trim off control ID prefix when calling command
    ' For example, "btnExport" calls VCS().Export()
    CallByName VCS, Mid(strCommand, 4), VbMethod, strArgument
End Function

Public Function VCS() As clsVersionControl
    ' Returns instance with Export(), Build(), MergeBuild(), etc.
    Set VCS = New clsVersionControl
End Function
```

**Python side (from addin_integration.py):**
```python
def export_source(self, db_path: str, source_folder: str = None) -> dict:
    """Export database via add-in."""
    # Call add-in's export function
    self._call_addin_function("HandleRibbonCommand", "btnExport")
    # Parse results from Export.log
    return self._parse_results()
```

## Integration Points

### Export Operations

**Full Export:**
```python
addin = VCSAddinIntegration(addin_path)
addin.load_addin(app)
result = addin.export_source(db_path, source_folder)
```

Maps to VBA:
```vba
VCS().Export()  ' via HandleRibbonCommand("btnExport")
```

**VBA-Only Export:**
```python
result = addin.export_vba(db_path, source_folder)
```

Maps to VBA:
```vba
VCS().ExportVBA()  ' via HandleRibbonCommand("btnExportVBA")
```

### Import Operations

**Merge Build:**
```python
result = addin.merge_build(db_path, source_folder)
```

Maps to VBA:
```vba
VCS().MergeBuild()  ' via HandleRibbonCommand("btnMergeBuild")
```

**Build from Source:**
```python
result = addin.build_from_source(source_folder, output_path)
```

Maps to VBA:
```vba
VCS().Build(source_folder)  ' via HandleRibbonCommand("btnBuild")
VCS().BuildAs()  ' via HandleRibbonCommand("btnBuildAs")
```

## Result Handling

### Log Files

The add-in writes detailed logs for each operation:

**Export.log:**
```
Beginning Export of Source Files
MyDatabase.accdb
VCS Version 4.1.0
Export Folder: C:\mydb.src
Using Fast Save

Exporting Queries... (15)
Exporting Modules... (8)
Exporting Forms... (12)
...
Export Complete! (3.2 seconds)
```

**Build.log:**
```
Beginning Build from Source
Source Folder: C:\mydb.src
Building Database...
Importing Queries... (15)
Importing Modules... (8)
...
Build Complete! (8.5 seconds)
```

### Python Parsing

The MCP tool parses log files for detailed results:

```python
def parse_log_file(self, log_path: str) -> dict:
    """Parse add-in log file."""
    with open(log_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    return {
        "found": True,
        "content": content,
        "path": log_path
    }
```

## Configuration

### Add-in Path

The MCP tool finds the add-in at:

1. `ACCESS_VCS_ADDIN_PATH` environment variable
2. Default: `%AppData%\MSAccessVCS\Version Control.accda`

**In .env:**
```bash
# Use default installation
# (no configuration needed)

# Or specify custom path
ACCESS_VCS_ADDIN_PATH=C:\CustomPath\Version Control.accda
```

### Add-in Options

The add-in uses options from `vcs-options.json` in the source folder:

```json
{
  "Info": {
    "AddinVersion": "4.1.0"
  },
  "Options": {
    "ExportFolder": "",
    "ShowDebugWindow": false,
    "UseFastSave": true,
    "SavePrintVars": true
  }
}
```

The MCP tool respects these options - they're controlled by the add-in.

## Error Handling

### Add-in Not Found

```python
try:
    addin = VCSAddinIntegration(addin_path)
    if not addin.verify_addin_exists():
        print(f"Add-in not found at: {addin.addin_path}")
        print("Install from: https://github.com/joyfullservice/msaccess-vcs-integration/releases")
except RuntimeError as e:
    print(f"Failed to load add-in: {e}")
```

### Operation Failures

All operations return a result dictionary with `success` flag:

```python
result = addin.export_source(db_path, source_folder)

if not result["success"]:
    print(f"Export failed: {result['message']}")
    # Check log file for details
    if result.get("log_path"):
        print(f"See log: {result['log_path']}")
```

### COM Errors

```python
try:
    addin.load_addin(app)
except RuntimeError as e:
    # Common issues:
    # - Add-in file not found
    # - Add-in not trusted by Access
    # - COM automation blocked
    print(f"Failed to load add-in: {e}")
```

## Export Format

The add-in exports database objects to a comprehensive folder structure:

```
database.src/
├── queries/           # SQL queries (.sql, .bas)
├── modules/           # VBA modules (.bas, .cls)
├── forms/             # Form definitions (.bas, .cls, .json)
├── reports/           # Report definitions (.bas, .cls, .json)
├── macros/            # Macros (.bas)
├── tables/            # Table data (.txt, .xml)
├── tbldefs/           # Table structure (.sql, .xml)
├── relations/         # Table relationships (.json)
├── vcs-options.json   # Export options
├── vcs-index.json     # Fast save index
└── Export.log         # Operation log
```

See [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-integration/blob/main/Version%20Control.accda.src/AGENTS.md) for complete format documentation.

### Encoding

**Critical:** All text files use UTF-8 with BOM encoding.

- The add-in requires BOM for import
- Python must preserve BOM when editing files
- Check encoding before writing changes

### File Formats

**Queries (.sql):**
```sql
-- Query: CustomerList
-- Type: Select
-- Exported: 2026-01-20T10:30:00

SELECT * FROM Customers
WHERE Active = True
```

**Modules (.bas):**
```vba
Attribute VB_Name = "Utilities"
Option Compare Database
Option Explicit

Public Function GetVersion() As String
    GetVersion = "1.0"
End Function
```

**Forms (.bas, .cls):**
- `.bas` - Form layout and properties
- `.cls` - Code-behind (if "Split Layout from VBA" enabled)

## VBA Add-in Development

If creating a new VBA add-in to work with this tool:

### Recommended Functions

```vba
' Export all objects
Public Function ExportDatabase( _
    dbPath As String, _
    outputDir As String, _
    Optional objectTypes As String = "all" _
) As Boolean
    ' Implementation
End Function

' Import objects
Public Function ImportObjects( _
    dbPath As String, _
    sourceDir As String, _
    Optional overwrite As Boolean = False _
) As Boolean
    ' Implementation
End Function

' List objects
Public Function ListObjects( _
    dbPath As String _
) As String  ' JSON string
    ' Implementation
End Function
```

### Error Handling

```vba
On Error GoTo ErrorHandler
    ' Operation
    ExportDatabase = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    ExportDatabase = False
End Function
```

### Logging

```vba
' Write log file for Python to read
Public Sub WriteLog(message As String)
    Dim logFile As String
    logFile = Environ("TEMP") & "\vba_addin_log.txt"
    
    Open logFile For Append As #1
    Print #1, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    Close #1
End Sub
```

## Testing Integration

### Test Checklist

- [ ] Both tools export same queries identically
- [ ] Both tools export same modules identically
- [ ] File formats are compatible
- [ ] Metadata is consistent
- [ ] Error handling works
- [ ] Performance is acceptable

### Test Script

```python
# test_integration.py
import os
import filecmp
from msaccess_vcs_mcp.tools import vcs_export_database

# Export with Python tool
vcs_export_database("test.accdb", "python_export")

# Export with VBA tool (manual or automated)
# compare_exports.vbs

# Compare outputs
def compare_directories(dir1, dir2):
    dircmp = filecmp.dircmp(dir1, dir2)
    if dircmp.diff_files:
        print("Differences found:", dircmp.diff_files)
        return False
    return True

assert compare_directories("python_export", "vba_export")
```

## Best Practices

### Naming Conventions

Use consistent naming:
- **Queries**: `QueryName.sql` (both tools)
- **Modules**: `ModuleName.bas` (both tools)
- **Metadata**: `database.json` (Python) or `metadata.xml` (VBA)

### Version Control

Track which tool exported:
```json
{
  "tool": "msaccess-vcs-mcp",
  "tool_version": "0.1.0",
  "exported_date": "2026-01-20T10:30:00"
}
```

### Documentation

Document your integration:
```markdown
# Our Integration Setup

## Tools
- VBA Add-in: v2.5 (forms/reports)
- msaccess-vcs-mcp: v0.1 (queries/modules)

## Workflow
1. Export queries/modules with Python
2. Export forms/reports with VBA
3. Commit all to git

## Future
- Migrate forms/reports to Python when available
```

## Troubleshooting

### Different Output

**Problem**: Tools produce different output for same object
**Solution**: 
1. Compare formats side-by-side
2. Adjust Python export to match VBA
3. Or adjust VBA to match Python

### Performance Issues

**Problem**: Calling VBA from Python is slow
**Solution**:
1. Use native Python for fast operations
2. Call VBA only for complex operations
3. Consider migrating more to Python

### COM Errors

**Problem**: COM automation fails randomly
**Solution**:
1. Use file-based communication instead
2. Implement retry logic
3. Better error handling

## Future Enhancements

Planned improvements for VBA integration:

1. **Auto-detection**: Detect VBA add-in automatically
2. **Unified config**: Single config for both tools
3. **Seamless calling**: Python→VBA calls transparent
4. **Format converter**: Convert between formats
5. **Migration assistant**: Help migrate from VBA to Python

## Resources

- [Win32com Documentation](https://github.com/mhammond/pywin32)
- [Access Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/access)
- [COM Interop Best Practices](https://docs.microsoft.com/en-us/dotnet/standard/native-interop/com-interop)

## Support

For integration questions:
1. Open an issue on GitHub
2. Include both tool versions
3. Describe your current workflow
4. Share sample outputs (sanitized)
