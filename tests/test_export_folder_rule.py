"""The relative/absolute ExportFolder rule matches the add-in's."""

import json

import pytest

from msaccess_vcs_mcp.worktree import (
    configured_export_folder,
    is_inside,
    is_relative_export_folder,
    same_folder,
)


@pytest.mark.parametrize("value, relative", [
    (None, True),
    ("", True),
    ("\\src\\", True),
    ("\\..\\src\\", True),
    ("\\\\server\\share\\src\\", False),
    ("C:\\Repos\\example\\src\\", False),
    ("%USERPROFILE%\\example\\", False),
    ("src\\", False),
])
def test_relative_rule(value, relative):
    assert is_relative_export_folder(value) is relative


def test_configured_export_folder_reads_vcs_options(tmp_path):
    (tmp_path / "vcs-options.json").write_text(
        "\ufeff" + json.dumps({"Info": {}, "Options": {"ExportFolder": "C:\\x\\src\\"}}),
        encoding="utf-8",
    )
    assert configured_export_folder(tmp_path) == "C:\\x\\src\\"


def test_configured_export_folder_blank_and_missing(tmp_path):
    assert configured_export_folder(tmp_path) is None
    (tmp_path / "vcs-options.json").write_text(json.dumps({"Options": {}}), encoding="utf-8")
    assert configured_export_folder(tmp_path) == ""


def test_folder_comparisons(tmp_path):
    inner = tmp_path / "wt" / "src"
    inner.mkdir(parents=True)
    assert same_folder(str(inner).upper() + "\\", inner)
    assert is_inside(inner, tmp_path / "wt")
    assert is_inside(tmp_path / "wt", tmp_path / "wt")
    assert not is_inside(tmp_path / "wt-other", tmp_path / "wt")
    assert not is_inside(tmp_path / "wt" / ".." / "elsewhere", tmp_path / "wt")
