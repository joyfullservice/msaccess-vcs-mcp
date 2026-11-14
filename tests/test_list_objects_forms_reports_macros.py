"""vcs_list_objects lists forms, reports and macros by name.

Names come from ``CurrentProject.AllForms``, ``AllReports`` and ``AllMacros``,
which hold ``AccessObject`` entries: nothing is opened to read them.
"""

import asyncio
from pathlib import Path
from unittest.mock import MagicMock, PropertyMock, patch

import pytest

DB_PATH = r"C:\data\Example.accdb"
QUERIES = [{"name": "qryItems", "type": "Select"}]
COLLECTIONS = pytest.mark.parametrize(
    "key, collection",
    [("forms", "AllForms"), ("reports", "AllReports"), ("macros", "AllMacros")],
)


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


def _access_objects(names):
    objects = []
    for name in names:
        obj = MagicMock()
        obj.Name = name
        objects.append(obj)
    return objects


def _fake_app(forms=(), reports=(), macros=()):
    app = MagicMock()
    app.CurrentProject.FullName = DB_PATH
    app.CurrentProject.AllForms = _access_objects(forms)
    app.CurrentProject.AllReports = _access_objects(reports)
    app.CurrentProject.AllMacros = _access_objects(macros)
    module = MagicMock()
    module.Name = "modExample"
    module.Type = 1  # vbext_ct_StdModule
    project = MagicMock()
    project.FileName = DB_PATH
    project.VBComponents = [module]
    app.VBE.VBProjects = [project]
    app.VBE.ActiveVBProject = project
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
        patch("msaccess_vcs_mcp.tools.list_table_defs", return_value=[{"name": "Items"}]),
        patch("msaccess_vcs_mcp.tools.list_query_defs", return_value=QUERIES),
    ):
        return asyncio.run(_unwrap(vcs_list_objects)(DB_PATH))


def _assert_call_failed(result, key, collection, message):
    assert result["success"] is False
    assert key in result["error"]
    assert collection in result["error"]
    assert message in result["error"]
    for listed in ("forms", "reports", "macros"):
        assert listed not in result
    assert result["tables"] == ["Items"]
    assert result["queries"] == QUERIES
    assert result["modules"] == ["modExample"]


def test_lists_form_report_and_macro_names_sorted_by_name():
    app = _fake_app(
        forms=["frmBravo", "frmCharlie", "frmAlpha"],
        reports=["rptBravo", "rptAlpha"],
        macros=["mcrBravo", "mcrAlpha"],
    )

    result = _list_objects(app)

    assert result["success"] is True
    assert result["forms"] == ["frmAlpha", "frmBravo", "frmCharlie"]
    assert result["reports"] == ["rptAlpha", "rptBravo"]
    assert result["macros"] == ["mcrAlpha", "mcrBravo"]
    assert result["modules"] == ["modExample"]
    assert result["tables"] == ["Items"]
    assert result["queries"] == QUERIES


def test_sort_ignores_case():
    app = _fake_app(forms=["frmbeta", "frmAlpha", "FrmGamma"])

    assert _list_objects(app)["forms"] == ["frmAlpha", "frmbeta", "FrmGamma"]


def test_listing_opens_nothing():
    app = _fake_app(forms=["frmAlpha"], reports=["rptAlpha"], macros=["mcrAlpha"])

    result = _list_objects(app)

    assert result["success"] is True
    assert result["forms"] == ["frmAlpha"]
    assert result["reports"] == ["rptAlpha"]
    assert result["macros"] == ["mcrAlpha"]
    assert app.DoCmd.mock_calls == []


def test_empty_database_lists_no_forms_reports_or_macros():
    result = _list_objects(_fake_app())

    assert result["success"] is True
    assert result["forms"] == []
    assert result["reports"] == []
    assert result["macros"] == []


@COLLECTIONS
def test_collection_that_raises_fails_the_call(key, collection):
    app = _fake_app(forms=["frmAlpha"], reports=["rptAlpha"], macros=["mcrAlpha"])
    setattr(
        type(app.CurrentProject),
        collection,
        PropertyMock(side_effect=RuntimeError("collection unavailable")),
    )

    _assert_call_failed(_list_objects(app), key, collection, "collection unavailable")


@COLLECTIONS
def test_collection_whose_iteration_raises_fails_the_call(key, collection):
    app = _fake_app(forms=["frmAlpha"], reports=["rptAlpha"], macros=["mcrAlpha"])
    broken = MagicMock()
    broken.__iter__.side_effect = RuntimeError("iteration failed")
    setattr(app.CurrentProject, collection, broken)

    _assert_call_failed(_list_objects(app), key, collection, "iteration failed")


@COLLECTIONS
def test_item_whose_name_raises_fails_the_call(key, collection):
    app = _fake_app(forms=["frmAlpha"], reports=["rptAlpha"], macros=["mcrAlpha"])
    item = MagicMock()
    type(item).Name = PropertyMock(side_effect=RuntimeError("name unavailable"))
    setattr(app.CurrentProject, collection, [item])

    _assert_call_failed(_list_objects(app), key, collection, "name unavailable")
