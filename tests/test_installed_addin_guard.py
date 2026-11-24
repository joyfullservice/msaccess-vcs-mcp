"""Tests for the refusal that keeps the installed add-in out of every tool.

The install exists to be loaded as a library. Opening it as a database, or writing
into it, resets a VBA project while it is executing -- so the refusal sits in the
``vcs_tool`` wrapper, ahead of the gate and of any COM work, and covers every tool
rather than the handful that had grown their own guards.

Paths here are synthetic. They only have to look like real installs and real
repositories to the comparison under test; nothing is read from disk.
"""

from __future__ import annotations

import asyncio
from pathlib import Path
from unittest.mock import MagicMock

import pytest

import msaccess_vcs_mcp.tools as tools_module


INSTALLED_ADDIN = r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.accda"
INSTALLED_ADDIN_COMPILED = (
    r"C:\Users\Example\AppData\Roaming\MSAccessVCS\Version Control.accde"
)
REPO_ADDIN = r"C:\Projects\msaccess-vcs-addin\Version Control.accda"
SOURCE_DIR = r"C:\Projects\msaccess-vcs-addin\Version Control.accda.src"
USER_DB = r"C:\Projects\example\Sales.accdb"


@pytest.fixture(autouse=True)
def _installed_addin_configured(monkeypatch):
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": INSTALLED_ADDIN},
    )
    monkeypatch.setattr(tools_module, "validate_database_path", lambda p: Path(p))


def _refusal(tool: str, func, *args, **kwargs):
    return tools_module._refuse_installed_addin_target(tool, func, args, kwargs)


def _unwrap(tool_callable):
    fn = tool_callable
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.mark.parametrize(
    "tool_name",
    [
        "vcs_run_tests",
        "vcs_export_database",
        "vcs_import_objects",
        "vcs_run_vba",
        "vcs_call_vba",
        "vcs_compile_vba",
        "vcs_execute_sql",
        "vcs_export_object",
        "vcs_import_object",
        "vcs_diff_database",
        "vcs_list_objects",
        "vcs_capture_screenshot",
        "vcs_reset_vba",
    ],
)
def test_every_tool_refuses_the_install_as_database_path(tool_name):
    """The refusal is the wrapper's, so no tool can be the one that forgot it."""
    func = _unwrap(getattr(tools_module, tool_name))

    result = _refusal(tool_name, func, INSTALLED_ADDIN)

    assert result is not None
    assert result["success"] is False
    assert result["error_pattern"] == "installed_addin_refused"
    assert tool_name in result["error"]
    assert "database_path" in result["error"]
    # A refusal that does not name the host that works just restates the problem.
    assert "development copy" in result["error"]


def test_dismiss_dialog_refuses_the_install_as_database_path():
    func = _unwrap(tools_module.vcs_dismiss_dialog)

    result = _refusal("vcs_dismiss_dialog", func, "abc", "OK", database_path=INSTALLED_ADDIN)

    assert result is not None
    assert result["error_pattern"] == "installed_addin_refused"


def test_rebuild_output_path_refused():
    """vcs_rebuild_database takes no database_path -- the install would be its output."""
    func = _unwrap(tools_module.vcs_rebuild_database)

    result = _refusal("vcs_rebuild_database", func, SOURCE_DIR, INSTALLED_ADDIN)

    assert result is not None
    assert result["error_pattern"] == "installed_addin_refused"
    assert "output_path" in result["error"]


def test_compiled_install_also_refused():
    """An install configured for the compiled add-in is a .accde built from the .accda.

    Only one of the two is ever named in ACCESS_VCS_ADDIN_PATH, so a full-name
    comparison would let the other variant through. The add-in makes the same
    allowance in modInstall.PathsMatchIgnoringExtension.
    """
    func = _unwrap(tools_module.vcs_run_tests)

    result = _refusal("vcs_run_tests", func, INSTALLED_ADDIN_COMPILED)

    assert result is not None
    assert result["error_pattern"] == "installed_addin_refused"


def test_development_copy_and_user_databases_pass():
    """The development copy shares the install's file name but is a different file."""
    func = _unwrap(tools_module.vcs_run_tests)

    assert _refusal("vcs_run_tests", func, REPO_ADDIN) is None
    assert _refusal("vcs_run_tests", func, USER_DB) is None


def test_keyword_arguments_are_checked_too():
    func = _unwrap(tools_module.vcs_export_database)

    result = _refusal(
        "vcs_export_database", func, database_path=INSTALLED_ADDIN, output_dir=SOURCE_DIR
    )

    assert result is not None
    assert result["error_pattern"] == "installed_addin_refused"


def test_export_folder_beside_the_install_is_not_the_install():
    """Only the add-in file is off limits; a folder next to it is not that file."""
    func = _unwrap(tools_module.vcs_export_database)

    result = _refusal(
        "vcs_export_database",
        func,
        REPO_ADDIN,
        INSTALLED_ADDIN + ".src",
    )

    assert result is None


def test_refusal_happens_before_the_gate_and_any_com(monkeypatch):
    """End to end through the registered wrapper, with COM and the gate booby-trapped."""
    monkeypatch.setattr(tools_module, "_ensure_env_loaded", _noop_async)
    monkeypatch.setattr(tools_module, "load_config", lambda: {})

    def _explode():
        raise AssertionError("the gate must not be reached")

    monkeypatch.setattr(tools_module, "get_access_gate", _explode)
    connect_cls = MagicMock()
    monkeypatch.setattr(tools_module, "AccessConnection", connect_cls)

    result = asyncio.run(tools_module.vcs_run_tests(INSTALLED_ADDIN))

    assert result["error_pattern"] == "installed_addin_refused"
    connect_cls.assert_not_called()


async def _noop_async(*_args, **_kwargs):
    return None
