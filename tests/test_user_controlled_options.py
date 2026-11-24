"""The McpAllow* permissions are the user's: no tool may change them."""

from pathlib import Path
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools


PERMISSIONS = ["McpAllowRunVBA", "McpAllowCallVBA", "McpAllowExecuteSQL", "McpAllowImport"]


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


@pytest.fixture
def no_access(tmp_path):
    addin = tmp_path / "Version Control.accda"
    db = tmp_path / "db.accdb"
    with (
        patch.object(tools, "get_config", return_value={"ACCESS_VCS_ADDIN_PATH": str(addin)}),
        patch.object(tools, "validate_database_path", return_value=Path(db)),
        patch.object(tools, "log_code_execution"),
        patch.object(tools, "AccessConnection") as connection,
    ):
        yield str(db), connection


def _set_option(db_path, name, value):
    fn = tools.vcs_set_option
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn(db_path, name, value)


@pytest.mark.parametrize("name", PERMISSIONS + ["mcpallowimport", " McpAllowRunVBA "])
def test_set_option_refuses_permission(no_access, name):
    db_path, connection = no_access
    result = _set_option(db_path, name, True)
    assert result["success"] is False
    assert result["error_pattern"] == "user_controlled_option"
    assert "Ask the user" in result["error"]
    connection.assert_not_called()


@pytest.mark.parametrize("name", PERMISSIONS)
def test_call_vba_set_option_refuses_permission(no_access, name):
    db_path, connection = no_access
    result = tools._execute_call_vba(db_path, "VCS.API", ["SetOption", name, "True"])
    assert result["success"] is False
    assert result["error_pattern"] == "user_controlled_option"
    connection.assert_not_called()


def test_call_vba_other_option_is_not_refused(no_access):
    db_path, connection = no_access
    tools._execute_call_vba(db_path, "VCS.API", ["SetOption", "ShowDebug", "True"])
    connection.assert_called_once()


def test_call_vba_non_addin_function_is_not_refused(no_access):
    db_path, connection = no_access
    tools._execute_call_vba(db_path, "MyModule.API", ["SetOption", "McpAllowRunVBA", "True"])
    connection.assert_called_once()
