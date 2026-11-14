"""vcs_list_objects lists the modules of the database's own VBA project.

Once the add-in has been loaded, ``VBE.ActiveVBProject`` can be the
add-in's project rather than the database's (issue #5).
"""

import asyncio
from pathlib import Path
from unittest.mock import MagicMock, PropertyMock, patch

import pytest

DB_PATH = r"C:\data\Example.accdb"
ADDIN_PATH = r"C:\Users\me\AppData\Roaming\MSAccessVCS\Version Control.accda"
TABLES = [{"name": "Customers"}]
QUERIES = [{"name": "qCustomers", "type": "Select"}]


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


def _project(name, file_name, modules):
    project = MagicMock()
    project.Name = name
    if isinstance(file_name, Exception):
        type(project).FileName = PropertyMock(side_effect=file_name)
    else:
        project.FileName = file_name
    components = []
    for module in modules:
        component = MagicMock()
        component.Name = module
        component.Type = 1  # vbext_ct_StdModule
        components.append(component)
    project.VBComponents = components
    return project


def _fake_app(projects, active):
    app = MagicMock()
    app.CurrentProject.FullName = DB_PATH
    app.VBE.VBProjects = projects
    app.VBE.ActiveVBProject = active
    return app


def _list_objects(app):
    from msaccess_vcs_mcp.tools import vcs_list_objects

    conn = MagicMock()
    conn.__enter__ = MagicMock(return_value=conn)
    conn.__exit__ = MagicMock(return_value=False)
    conn.connect.return_value = (app, MagicMock())
    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=conn),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=Path(DB_PATH)),
        patch("msaccess_vcs_mcp.tools.list_table_defs", return_value=TABLES),
        patch("msaccess_vcs_mcp.tools.list_query_defs", return_value=QUERIES),
    ):
        return asyncio.run(_unwrap(vcs_list_objects)(DB_PATH))


def test_lists_database_modules_when_database_project_is_active():
    db_project = _project("Example", DB_PATH, ["modExample"])
    addin_project = _project("MSAccessVCS", ADDIN_PATH, ["clsVersionControl", "modAPI"])
    app = _fake_app([db_project, addin_project], active=db_project)

    result = _list_objects(app)

    assert result["success"] is True
    assert result["modules"] == ["modExample"]


def test_lists_database_modules_when_addin_project_is_active():
    db_project = _project("Example", DB_PATH, ["modExample"])
    addin_project = _project("MSAccessVCS", ADDIN_PATH, ["clsVersionControl", "modAPI"])
    app = _fake_app([db_project, addin_project], active=addin_project)

    result = _list_objects(app)

    assert result["success"] is True
    assert result["modules"] == ["modExample"]


@pytest.mark.parametrize(
    "db_file_name",
    [DB_PATH.upper(), DB_PATH.replace("\\", "/")],
    ids=["different-case", "forward-slashes"],
)
def test_matches_database_project_by_normalised_path(db_file_name):
    addin_project = _project("MSAccessVCS", ADDIN_PATH, ["modAPI"])
    db_project = _project("Example", db_file_name, ["modExample"])
    app = _fake_app([addin_project, db_project], active=addin_project)

    assert _list_objects(app)["modules"] == ["modExample"]


def test_skips_project_whose_file_name_raises():
    broken = _project("Broken", RuntimeError("FileName unavailable"), ["modBroken"])
    db_project = _project("Example", DB_PATH, ["modExample"])
    app = _fake_app([broken, db_project], active=broken)

    assert _list_objects(app)["modules"] == ["modExample"]


def test_no_matching_project_is_an_error_and_keeps_tables_and_queries():
    addin_project = _project("MSAccessVCS", ADDIN_PATH, ["clsVersionControl", "modAPI"])
    app = _fake_app([addin_project], active=addin_project)

    result = _list_objects(app)

    assert result["success"] is False
    assert "VBA project" in result["error"]
    assert result["modules"] == []
    assert result["tables"] == ["Customers"]
    assert result["queries"] == QUERIES


def test_current_project_full_name_raising_lists_no_modules():
    # Known limit: this falls into the existing module-listing handler,
    # which reports success with no modules rather than an error.
    db_project = _project("Example", DB_PATH, ["modExample"])
    app = _fake_app([db_project], active=db_project)
    type(app.CurrentProject).FullName = PropertyMock(side_effect=RuntimeError("no current project"))

    result = _list_objects(app)

    assert result["success"] is True
    assert result["modules"] == []
    assert result["tables"] == ["Customers"]
