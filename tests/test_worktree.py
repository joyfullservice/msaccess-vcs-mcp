"""Worktree detection and the warnings attached to rebuild results."""

import json
from pathlib import Path

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.worktree import (
    data_source_warnings,
    find_worktree,
    worktree_install_fields,
)


def _make_repo(tmp_path: Path) -> tuple[Path, Path]:
    """A main checkout and one linked worktree, laid out the way git does."""
    main = tmp_path / "project"
    worktree = tmp_path / "project-wt"
    admin = main / ".git" / "worktrees" / "project-wt"
    admin.mkdir(parents=True)
    (admin / "commondir").write_text("../..\n", encoding="utf-8")
    for checkout in (main, worktree):
        (checkout / "App.accdb.src").mkdir(parents=True)
    (worktree / ".git").write_text(f"gitdir: {admin}\n", encoding="utf-8")
    return main, worktree


def test_worktree_is_found_with_its_main_checkout(tmp_path):
    main, worktree = _make_repo(tmp_path)

    found = find_worktree(worktree / "App.accdb.src")

    assert found is not None
    assert found.root == worktree.resolve()
    assert found.main_checkout == main.resolve()


def test_a_file_path_in_a_worktree_is_found(tmp_path):
    _, worktree = _make_repo(tmp_path)
    assert find_worktree(worktree / "App.accdb") is not None


def test_main_checkout_is_not_a_worktree(tmp_path):
    main, _ = _make_repo(tmp_path)
    assert find_worktree(main / "App.accdb.src") is None


def test_submodule_is_not_a_worktree(tmp_path):
    admin = tmp_path / "super" / ".git" / "modules" / "lib"
    admin.mkdir(parents=True)
    sub = tmp_path / "super" / "lib"
    sub.mkdir()
    (sub / ".git").write_text(f"gitdir: {admin}\n", encoding="utf-8")
    assert find_worktree(sub) is None


def test_relative_gitdir_is_resolved(tmp_path):
    main, worktree = _make_repo(tmp_path)
    (worktree / ".git").write_text(
        "gitdir: ../project/.git/worktrees/project-wt\n", encoding="utf-8"
    )
    found = find_worktree(worktree)
    assert found is not None and found.main_checkout == main.resolve()


def test_outside_any_repository(tmp_path):
    assert find_worktree(tmp_path) is None


def test_install_from_worktree_names_the_main_source(tmp_path):
    main, worktree = _make_repo(tmp_path)

    fields = worktree_install_fields(worktree / "App.accdb.src")

    assert fields["installed_from_worktree"] is True
    assert fields["main_checkout"] == str(main.resolve())
    assert str(main.resolve() / "App.accdb.src") in fields["agent_guidance"]
    assert fields["warnings"]


def test_install_from_main_checkout_adds_nothing(tmp_path):
    main, _ = _make_repo(tmp_path)
    assert worktree_install_fields(main / "App.accdb.src") == {}


SOURCES = [
    {"type": "table", "kind": "project", "source": "C:\\wt\\be.accdb", "linked": 3, "notLinked": 0},
    {"type": "table", "kind": "network", "source": "\\\\server\\share\\data.accdb", "linked": 2, "notLinked": 0},
    {"type": "table", "kind": "server", "source": "ODBC:DSN=APPDB", "linked": 4, "notLinked": 0},
    {"type": "library", "kind": "missing", "source": "C:\\wt\\lib.accda", "linked": 0, "notLinked": 1},
]


def test_worktree_build_warns_about_shared_and_missing_sources(tmp_path):
    main, worktree = _make_repo(tmp_path)

    warnings = data_source_warnings(str(worktree / "App.accdb"), SOURCES)

    assert len(warnings) == 2
    missing, shared = warnings
    assert "lib.accda" in missing and str(main.resolve()) in missing
    assert "ODBC:DSN=APPDB" in shared and "share" in shared
    assert "be.accdb" not in shared


def test_main_checkout_build_warns_only_about_missing_sources(tmp_path):
    main, _ = _make_repo(tmp_path)

    warnings = data_source_warnings(str(main / "App.accdb"), SOURCES)

    assert len(warnings) == 1
    assert "missing" in warnings[0]


def test_no_sources_no_warnings(tmp_path):
    _, worktree = _make_repo(tmp_path)
    assert data_source_warnings(str(worktree / "App.accdb"), []) == []
    assert data_source_warnings(str(worktree / "App.accdb"), None) == []


def test_build_result_carries_data_sources_and_warnings(tmp_path):
    _, worktree = _make_repo(tmp_path)
    database = worktree / "App.accdb"
    raw = json.dumps({
        "success": True,
        "logPath": str(worktree / "App.accdb.src" / "logs" / "Build.log"),
        "databasePath": str(database),
        "dataSources": SOURCES,
    })

    result = tools._build_headless_result(
        {"success": True, "result": raw}, "VCS.API", worktree / "App.accdb.src", None
    )

    assert result["data_sources"] == SOURCES
    assert "dataSources" not in result
    assert len(result["warnings"]) == 2


def test_build_result_from_older_addin_has_no_data_sources(tmp_path):
    raw = json.dumps({"success": True, "databasePath": str(tmp_path / "App.accdb")})

    result = tools._build_headless_result(
        {"success": True, "result": raw}, "VCS.API", tmp_path, None
    )

    assert "data_sources" not in result
    assert "warnings" not in result
