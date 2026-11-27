"""Copies of a project refuse work whose export folder escapes their working tree."""

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.worktree import (
    build_source_refusal,
    export_folder_refusal,
    working_tree_root,
)


def _layout(tmp_path: Path, export_folder: str) -> tuple[Path, Path]:
    """A main checkout and a linked worktree, each with a src folder."""
    main = tmp_path / "project"
    wt = tmp_path / "project-wt"
    admin = main / ".git" / "worktrees" / "project-wt"
    admin.mkdir(parents=True)
    (admin / "commondir").write_text("../..\n", encoding="utf-8")
    (wt).mkdir()
    (wt / ".git").write_text(f"gitdir: {admin}\n", encoding="utf-8")
    for checkout in (main, wt):
        src = checkout / "src"
        src.mkdir()
        (src / "vbe-project.json").write_text(
            json.dumps({"Items": {"FileName": "App.accdb"}}), encoding="utf-8")
        (src / "vcs-options.json").write_text(
            json.dumps({"Options": {"ExportFolder": export_folder}}), encoding="utf-8")
    return main, wt


def test_working_tree_root_for_clone_and_worktree(tmp_path):
    main, wt = _layout(tmp_path, "")
    assert working_tree_root(main / "src") == main.resolve()
    assert working_tree_root(wt / "src") == wt.resolve()
    assert working_tree_root(tmp_path) is None


def test_worktree_build_with_absolute_export_folder_is_refused(tmp_path):
    main, wt = _layout(tmp_path, str(tmp_path / "project" / "src") + "\\")

    refusal = build_source_refusal(wt / "src")

    assert refusal["error_pattern"] == "export_folder_outside_working_tree"
    assert build_source_refusal(main / "src") is None


def test_relative_and_blank_export_folders_build_anywhere(tmp_path):
    _, wt = _layout(tmp_path, "\\src\\")
    assert build_source_refusal(wt / "src") is None
    (wt / "src" / "vcs-options.json").write_text(json.dumps({"Options": {}}), encoding="utf-8")
    assert build_source_refusal(wt / "src") is None


def test_absolute_export_folder_inside_the_same_tree_is_allowed(tmp_path):
    _, wt = _layout(tmp_path, "")
    (wt / "src" / "vcs-options.json").write_text(
        json.dumps({"Options": {"ExportFolder": str(wt / "src") + "\\"}}), encoding="utf-8")
    assert build_source_refusal(wt / "src") is None


def test_db_name_placeholder_expands(tmp_path):
    main, wt = _layout(tmp_path, str(tmp_path / "project" / "%dbName%.src") + "\\")
    (main / "App.accdb.src").mkdir()
    assert build_source_refusal(wt / "src")["error_pattern"] == "export_folder_outside_working_tree"


def test_export_folder_refusal_for_database_outside_any_tree(tmp_path):
    assert export_folder_refusal(tmp_path / "App.accdb", tmp_path / "elsewhere") is None


def test_export_folder_refusal_names_both_folders(tmp_path):
    main, wt = _layout(tmp_path, "")
    refusal = export_folder_refusal(wt / "App.accdb", main / "src")
    assert refusal["working_tree"] == str(wt.resolve())
    assert refusal["export_path"] == str(main / "src")
    assert export_folder_refusal(wt / "App.accdb", wt / "App.accdb.src") is None


def _unwrap(fn):
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


def test_set_option_refuses_export_folder():
    result = _unwrap(tools.vcs_set_option)(r"C:\x\App.accdb", "exportFolder", r"C:\y")
    assert result["error_pattern"] == "export_folder_not_settable"
    assert "output_dir" in result["agent_guidance"]


def test_source_folder_mismatch(tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    a.mkdir(), b.mkdir()
    assert tools._source_folder_mismatch(a, Path(str(a).upper())) is None
    assert tools._source_folder_mismatch(a, b)["error_pattern"] == "source_folder_mismatch"
