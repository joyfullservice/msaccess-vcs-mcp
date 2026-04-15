# Testing Guide for msaccess-vcs-mcp

This guide covers testing procedures for the MCP tool, including unit tests, integration tests, and end-to-end workflow testing.

## Virtual Environment

All test commands assume the project virtual environment is activated. Activate it before running any tests:

```powershell
cd C:\Repos\msaccess-vcs-mcp
.\venv\Scripts\Activate.ps1
```

If the venv doesn't exist yet, create and install:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

## Running Unit Tests

Unit tests cover individual components without requiring Access or the VCS add-in:

```bash
# Run all unit tests
pytest

# Run with coverage report
pytest --cov=msaccess_vcs_mcp --cov-report=html

# Run specific test file
pytest tests/test_usage_logging.py -v

# Run specific test
pytest tests/test_addin_integration.py::TestVCSAddinIntegration::test_load_addin_success -v
```

## Integration Tests

Integration tests require:
- Microsoft Access installed
- MSAccess VCS add-in installed
- Test database available

### Prerequisites

1. **Install MSAccess VCS Add-in:**
   - Download from [releases](https://github.com/joyfullservice/msaccess-vcs-integration/releases/latest)
   - Install to default location: `%AppData%\MSAccessVCS\`

2. **Create Test Database:**
   ```python
   # Create a simple test database
   import win32com.client
   
   app = win32com.client.Dispatch("Access.Application")
   app.NewCurrentDatabase("C:\\test\\TestDB.accdb")
   
   # Add a simple query
   db = app.CurrentDb()
   qd = db.CreateQueryDef("TestQuery")
   qd.SQL = "SELECT 1 AS TestValue"
   
   app.CloseCurrentDatabase()
   app.Quit()
   ```

### Running Integration Tests

```bash
# Run integration tests (marked with @pytest.mark.integration)
pytest -m integration

# Skip integration tests (when Access not available)
pytest -m "not integration"
```

## End-to-End Workflow Testing

### Test Setup

1. **Create test directory structure:**
   ```
   C:\test\
   ├── TestDB.accdb          # Test database
   ├── TestDB.src\           # Source files (will be created)
   └── TestDB_built.accdb    # Rebuilt database (will be created)
   ```

2. **Activate virtual environment** (see [Virtual Environment](#virtual-environment) above)

### Manual End-to-End Test

This test verifies the complete workflow from export through modify to merge build.

#### Step 1: Export Database

```python
from msaccess_vcs_mcp.tools import vcs_export_database

# Export test database
result = vcs_export_database(
    "C:\\test\\TestDB.accdb",
    "C:\\test\\TestDB.src"
)

print(f"Success: {result['success']}")
print(f"Exported: {result['exported_count']} objects")
print(f"Path: {result['export_path']}")

# Verify source files were created
import os
assert os.path.exists("C:\\test\\TestDB.src")
assert os.path.exists("C:\\test\\TestDB.src\\vcs-options.json")
print("✓ Export successful")
```

#### Step 2: Modify Source Files

```python
# Read a query
query_path = "C:\\test\\TestDB.src\\queries\\TestQuery.sql"
with open(query_path, 'r', encoding='utf-8-sig') as f:
    original_sql = f.read()

# Modify the query
modified_sql = original_sql.replace(
    "SELECT 1 AS TestValue",
    "SELECT 2 AS TestValue, 'Modified' AS Status"
)

# Write back with UTF-8 BOM
with open(query_path, 'w', encoding='utf-8-sig') as f:
    f.write(modified_sql)

print("✓ Query modified")
```

#### Step 3: Merge Changes

```python
from msaccess_vcs_mcp.tools import vcs_import_objects

# Merge source changes into database
result = vcs_import_objects(
    "C:\\test\\TestDB.accdb",
    "C:\\test\\TestDB.src"
)

print(f"Success: {result['success']}")
if result.get('log_path'):
    print(f"Log: {result['log_path']}")

print("✓ Merge successful")
```

#### Step 4: Verify Changes

```python
from msaccess_vcs_mcp.tools import vcs_list_objects
from msaccess_vcs_mcp.access_com.connection import AccessConnection

# List objects to verify query exists
objects = vcs_list_objects("C:\\test\\TestDB.accdb")
print(f"Queries: {[q['name'] for q in objects['queries']]}")

# Read the modified query SQL directly
with AccessConnection("C:\\test\\TestDB.accdb") as conn:
    app, db = conn.connect()
    query = db.QueryDefs("TestQuery")
    print(f"Modified SQL: {query.SQL}")
    
    # Should contain our modifications
    assert "Status" in query.SQL
    assert "2" in query.SQL

print("✓ Changes verified in database")
```

#### Step 5: Build from Source

```python
from msaccess_vcs_mcp.tools import vcs_rebuild_database

# Build fresh database from source
result = vcs_rebuild_database(
    "C:\\test\\TestDB.src",
    "C:\\test\\TestDB_built.accdb"
)

print(f"Success: {result['success']}")
print(f"Output: {result['output_path']}")

# Verify built database
import os
assert os.path.exists("C:\\test\\TestDB_built.accdb")
print("✓ Build from source successful")
```

#### Step 6: Compare Databases

```python
# Export both databases and compare
vcs_export_database(
    "C:\\test\\TestDB.accdb",
    "C:\\test\\TestDB_original.src"
)

vcs_export_database(
    "C:\\test\\TestDB_built.accdb",
    "C:\\test\\TestDB_rebuilt.src"
)

# Compare source files
import filecmp

dcmp = filecmp.dircmp(
    "C:\\test\\TestDB_original.src",
    "C:\\test\\TestDB_rebuilt.src"
)

print(f"Identical files: {len(dcmp.same_files)}")
print(f"Different files: {len(dcmp.diff_files)}")
print(f"Different: {dcmp.diff_files}")

# Some files may differ (timestamps, GUIDs, etc.)
# But core files should match
print("✓ Database comparison complete")
```

### Automated End-to-End Test Script

Save this as `tests/test_e2e_workflow.py`:

```python
import os
import pytest
from pathlib import Path
import win32com.client

from msaccess_vcs_mcp.tools import (
    vcs_export_database,
    vcs_import_objects,
    vcs_rebuild_database,
    vcs_list_objects
)


@pytest.mark.integration
class TestEndToEndWorkflow:
    """End-to-end workflow tests requiring Access and VCS add-in."""
    
    @pytest.fixture
    def test_db_path(self, tmp_path):
        """Create a test database."""
        db_path = tmp_path / "TestDB.accdb"
        
        # Create database with Access
        app = win32com.client.Dispatch("Access.Application")
        app.NewCurrentDatabase(str(db_path))
        
        # Add a test query
        db = app.CurrentDb()
        qd = db.CreateQueryDef("TestQuery")
        qd.SQL = "SELECT 1 AS TestValue"
        
        app.CloseCurrentDatabase()
        app.Quit()
        
        return str(db_path)
    
    def test_complete_workflow(self, test_db_path, tmp_path):
        """Test complete export -> modify -> merge -> build workflow."""
        src_path = tmp_path / "TestDB.src"
        
        # 1. Export
        result = vcs_export_database(
            test_db_path,
            str(src_path)
        )
        assert result["success"]
        assert src_path.exists()
        
        # 2. Modify source
        query_file = src_path / "queries" / "TestQuery.sql"
        assert query_file.exists()
        
        content = query_file.read_text(encoding='utf-8-sig')
        modified = content.replace(
            "SELECT 1 AS TestValue",
            "SELECT 2 AS TestValue"
        )
        query_file.write_text(modified, encoding='utf-8-sig')
        
        # 3. Merge changes
        result = vcs_import_objects(
            test_db_path,
            str(src_path)
        )
        assert result["success"]
        
        # 4. Verify changes
        objects = vcs_list_objects(test_db_path)
        assert any(q["name"] == "TestQuery" for q in objects["queries"])
        
        # 5. Build from source
        built_db = tmp_path / "TestDB_built.accdb"
        result = vcs_rebuild_database(
            str(src_path),
            str(built_db)
        )
        assert result["success"]
        assert built_db.exists()


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-m", "integration"])
```

## Common Test Scenarios

### Test Add-in Not Installed

```python
# Temporarily move or rename add-in
import os
from msaccess_vcs_mcp.config import get_config

config = get_config()
addin_path = config["ACCESS_VCS_ADDIN_PATH"]

# Should fail gracefully
try:
    result = vcs_export_database("test.accdb", "test.src")
    assert not result["success"]
    assert "not found" in result.get("error", "").lower()
except RuntimeError as e:
    assert "not found" in str(e).lower()
```

### Test Permission Denied

```python
from msaccess_vcs_mcp.tools import vcs_import_objects
import os

# Without write permission
os.environ["ACCESS_VCS_ALLOW_WRITES"] = "false"

result = vcs_import_objects("test.accdb", "test.src")
assert not result["success"]
assert "permission" in result["error"].lower()
```

### Test Invalid Database Path

```python
result = vcs_export_database(
    "C:\\NonExistent\\Database.accdb",
    "C:\\test\\output"
)
assert not result["success"]
```

### Test Fast Save (Incremental Export)

```python
# First export
result1 = vcs_export_database("test.accdb", "test.src")
count1 = result1["exported_count"]

# Second export without changes (should be fast)
result2 = vcs_export_database("test.accdb", "test.src")
count2 = result2["exported_count"]

# Fast save should export fewer objects
assert count2 <= count1
```

## Performance Testing

### Measure Export Time

```python
import time

start = time.time()
result = vcs_export_database("large.accdb", "large.src")
duration = time.time() - start

print(f"Exported {result['exported_count']} objects in {duration:.2f}s")
print(f"Average: {duration/result['exported_count']:.3f}s per object")
```

### Measure Merge Build Time

```python
start = time.time()
result = vcs_import_objects("large.accdb", "large.src")
duration = time.time() - start

print(f"Merge build completed in {duration:.2f}s")
```

## Troubleshooting Tests

### Enable Debug Mode

Set debug environment variables:
```bash
set ACCESS_VCS_DEBUG=1
pytest -v -s
```

### View COM Errors

```python
import win32com.client
import pythoncom

# Enable COM error details
pythoncom.CoInitialize()
```

### Check Log Files

After operations, check log files:
```python
# Export log
log_path = "C:\\test\\TestDB.src\\Export.log"
if os.path.exists(log_path):
    with open(log_path, 'r') as f:
        print(f.read())

# Build log  
log_path = "C:\\test\\TestDB.src\\Build.log"
if os.path.exists(log_path):
    with open(log_path, 'r') as f:
        print(f.read())
```

## Continuous Integration

For CI/CD environments:

1. **Skip integration tests** (no Access available):
   ```yaml
   # .github/workflows/test.yml
   - name: Run tests
     run: pytest -m "not integration"
   ```

2. **Run integration tests on Windows runners** (with Access):
   ```yaml
   - name: Install Access
     # Install Access runtime or full version
   
   - name: Install VCS Add-in
     # Download and install add-in
   
   - name: Run integration tests
     run: pytest -m integration
   ```

## Test Coverage

Check test coverage:

```bash
# Generate coverage report
pytest --cov=msaccess_vcs_mcp --cov-report=html

# View in browser
start htmlcov/index.html
```

Target coverage goals:
- Unit tests: >80%
- Integration tests: All major workflows
- End-to-end: Complete export/import cycle

## Resources

- [pytest Documentation](https://docs.pytest.org/)
- [pytest-cov Plugin](https://pytest-cov.readthedocs.io/)
- [MSAccess VCS Add-in Testing](https://github.com/joyfullservice/msaccess-vcs-integration/wiki)
