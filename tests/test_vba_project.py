"""The current database's VBA project is found by file name, not by the editor's focus."""

import asyncio
import inspect
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_com.vba_project import current_vb_project

DB = "C:\\work\\db.accdb"
ADDIN = "C:\\Users\\me\\AppData\\Roaming\\MSAccessVCS\\Version Control.accda"


class _Projects:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def __call__(self, index):
        return self._items[index - 1]


def _project(file_name, modules):
    components = [SimpleNamespace(Name=name, Type=1) for name in modules]
    return SimpleNamespace(FileName=file_name, VBComponents=components)


def _app_with_addin_active():
    addin = _project(ADDIN, ["modAddInOnly"])
    db = _project(DB, ["modDatabase"])
    vbe = SimpleNamespace(VBProjects=_Projects([addin, db]), ActiveVBProject=addin)
    return SimpleNamespace(VBE=vbe, CurrentProject=SimpleNamespace(FullName=DB)), db


def test_current_vb_project_skips_active_addin():
    app, db = _app_with_addin_active()
    assert current_vb_project(app) is db


def test_current_vb_project_falls_back_to_active_project():
    addin = _project(ADDIN, [])
    vbe = SimpleNamespace(VBProjects=_Projects([addin]), ActiveVBProject=addin)
    app = SimpleNamespace(VBE=vbe, CurrentProject=SimpleNamespace(FullName=DB))
    assert current_vb_project(app) is addin


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


def test_list_objects_lists_database_modules_when_addin_is_active():
    app, _ = _app_with_addin_active()
    conn = MagicMock()
    conn.__enter__.return_value = conn
    conn.connect.return_value = (app, MagicMock())

    fn = tools.vcs_list_objects
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__

    with (
        patch.object(tools, "validate_database_path", return_value=Path(DB)),
        patch.object(tools, "AccessConnection", return_value=conn),
        patch.object(tools, "list_table_defs", return_value=[]),
        patch.object(tools, "list_query_defs", return_value=[]),
    ):
        result = fn(DB)
        if inspect.iscoroutine(result):
            result = asyncio.run(result)

    assert result["success"] is True
    assert result["modules"] == ["modDatabase"]
