"""output_dir exports to another folder, on a new enough add-in, under the same guards."""

import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from msaccess_vcs_mcp import tools


def _addin(configured: Path, version: str = "5.2.0") -> MagicMock:
    addin = MagicMock()
    addin.call_sync.side_effect = lambda command, *_a: {
        "GetExportFolder": str(configured) + "\\",
        "GetVCSVersion": version,
    }.get(command)
    return addin


def test_no_output_dir_uses_the_configured_folder(tmp_path):
    configured = tmp_path / "App.accdb.src"
    path, override, refusal = tools._resolve_export_target(_addin(configured), tmp_path / "App.accdb", None)
    assert (path, override, refusal) == (configured, None, None)


def test_output_dir_equal_to_configured_needs_no_override(tmp_path):
    configured = tmp_path / "App.accdb.src"
    path, override, refusal = tools._resolve_export_target(
        _addin(configured), tmp_path / "App.accdb", str(configured).upper() + "\\")
    assert override is None and refusal is None and path == configured


def test_other_folder_on_new_addin_is_an_override(tmp_path):
    target = tmp_path / "review"
    path, override, refusal = tools._resolve_export_target(
        _addin(tmp_path / "App.accdb.src"), tmp_path / "App.accdb", str(target))
    assert refusal is None
    assert path == target and override == str(target)


def test_other_folder_on_old_addin_is_refused(tmp_path):
    configured = tmp_path / "App.accdb.src"
    path, override, refusal = tools._resolve_export_target(
        _addin(configured, "5.1.1"), tmp_path / "App.accdb", str(tmp_path / "review"))
    assert refusal["error_pattern"] == "export_folder_mismatch"
    assert override is None and path == configured


def test_folder_with_an_index_is_refused(tmp_path):
    target = tmp_path / "other.accdb.src"
    target.mkdir()
    (target / "vcs-index.idx").write_bytes(b"")
    _, override, refusal = tools._resolve_export_target(
        _addin(tmp_path / "App.accdb.src"), tmp_path / "App.accdb", str(target))
    assert refusal["error_pattern"] == "export_folder_has_index"
    assert override is None


def test_worktree_output_dir_must_stay_in_the_worktree(tmp_path):
    wt = tmp_path / "wt"
    wt.mkdir()
    (wt / ".git").mkdir()
    _, _, refusal = tools._resolve_export_target(
        _addin(wt / "App.accdb.src"), wt / "App.accdb", str(tmp_path / "outside"))
    assert refusal["error_pattern"] == "export_folder_outside_working_tree"


def test_one_folder_one_export_at_a_time(tmp_path):
    folder = tmp_path / "out"
    assert tools._claim_export_folder(folder, "vcs_export_database") is None
    try:
        busy = tools._claim_export_folder(str(folder).upper() + "\\", "vcs_export_object")
        assert busy["error_pattern"] == "server_busy"
        assert busy["busy_with"] == {"tool": "vcs_export_database", "database": None, "reason": "same_folder"}
    finally:
        tools._release_export_folder(folder)
    assert tools._claim_export_folder(folder, "vcs_export_object") is None
    tools._release_export_folder(folder)


def _source(folder: Path, export_folder: str, build_path: str | None = None) -> Path:
    folder.mkdir(parents=True)
    (folder / "vbe-project.json").write_text(json.dumps({"Items": {"FileName": "App.accdb"}}), encoding="utf-8")
    (folder / "vcs-options.json").write_text(
        json.dumps({"Options": {"ExportFolder": export_folder}}), encoding="utf-8")
    if build_path:
        (folder / "proj-properties.json").write_text(
            json.dumps({"Items": {"VCS Build Path": build_path}}), encoding="utf-8")
    return folder


def test_build_target_follows_the_addin(tmp_path):
    configured = str(tmp_path / "orig" / "src") + "\\"
    apps = str(tmp_path / "apps")
    orig = _source(tmp_path / "orig" / "src", configured, apps)
    copy = _source(tmp_path / "copy" / "src", configured, apps)
    blank = _source(tmp_path / "blank" / "App.accdb.src", "")

    assert tools._build_target(orig) == str(Path(apps) / "App.accdb")
    assert tools._build_target(copy) == str(tmp_path / "copy" / "App.accdb")
    assert tools._build_target(blank) == str(tmp_path / "blank" / "App.accdb")
    assert tools._build_target(copy, r"C:\out\x.accdb") == r"C:\out\x.accdb"
