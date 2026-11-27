"""Tell whether a path lies in a git worktree, and what that means for a result.

A worktree builds its own database but shares everything outside its folder
with the main checkout: the installed add-in, and any back end the add-in's
build reports as ``outside``, ``network``, or ``server``. These helpers find
the worktree from the ``.git`` file at its root and word the warnings that
``vcs_rebuild_addin`` and ``vcs_rebuild_database`` attach.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

SHARED_KINDS = frozenset({"outside", "network", "server"})
_LISTED_SOURCES = 5


@dataclass(frozen=True)
class Worktree:
    root: Path
    # None for a worktree of a bare repository, which has no main checkout.
    main_checkout: Path | None


def is_relative_export_folder(value: str | None) -> bool:
    """True when an ExportFolder value resolves from the database's own folder.

    Mirrors the add-in's ``clsOptions.IsRelativeExportFolder``: blank (the
    default ``<file name>.src``) or a single leading backslash. Everything
    else, including ``%VAR%`` placeholders and a bare ``src\\``, names the same
    folder for every copy of the project.
    """
    if not value:
        return True
    return value.startswith("\\") and not value.startswith("\\\\")


def configured_export_folder(source_dir: str | Path) -> str | None:
    """The ExportFolder value in a source folder's vcs-options.json, or None.

    None means the file is missing or unreadable; a project without one uses
    the default, which is relative.
    """
    try:
        text = (Path(source_dir) / "vcs-options.json").read_text(encoding="utf-8-sig")
        options = json.loads(text).get("Options") or {}
    except (OSError, ValueError, AttributeError):
        return None
    value = options.get("ExportFolder")
    return value if isinstance(value, str) else ""


def same_folder(a: str | Path, b: str | Path) -> bool:
    def norm(path: str | Path) -> str:
        return os.path.normcase(os.path.realpath(os.path.abspath(str(path).rstrip("\\/"))))
    return norm(a) == norm(b)


def is_inside(path: str | Path, folder: str | Path) -> bool:
    """True when ``path`` is ``folder`` or lies under it (after resolving ``..``)."""
    def norm(p: str | Path) -> str:
        return os.path.normcase(os.path.realpath(os.path.abspath(str(p).rstrip("\\/"))))
    child, parent = norm(path), norm(folder)
    return child == parent or child.startswith(parent.rstrip("\\/") + os.sep)


def _git_root(path: Path) -> Path | None:
    for folder in [path, *path.parents]:
        if (folder / ".git").exists():
            return folder
    return None


def _resolve(base: Path, value: str) -> Path:
    target = Path(value.strip())
    return target if target.is_absolute() else (base / target).resolve()


def find_worktree(path: str | Path) -> Worktree | None:
    """The linked worktree containing ``path``, or None.

    A main checkout has a ``.git`` folder. A linked worktree has a ``.git``
    file naming its admin folder, which holds a ``commondir`` file; a
    submodule also has a ``.git`` file, but no ``commondir``.
    """
    try:
        start = Path(path).resolve()
        root = _git_root(start if start.is_dir() else start.parent)
        if root is None or not (root / ".git").is_file():
            return None
        match = re.match(r"\s*gitdir:\s*(.+)", (root / ".git").read_text(encoding="utf-8"))
        if not match:
            return None
        admin = _resolve(root, match.group(1))
        commondir = admin / "commondir"
        if not commondir.is_file():
            return None
        common = _resolve(admin, commondir.read_text(encoding="utf-8"))
    except OSError:
        return None
    main = common.parent if common.name == ".git" else None
    return Worktree(root=root, main_checkout=main)


def working_tree_root(path: str | Path) -> Path | None:
    """The git working tree (main checkout, clone, or linked worktree) holding ``path``."""
    try:
        start = Path(path).resolve()
    except OSError:
        return None
    return _git_root(start if start.is_dir() else start.parent)


def source_db_file_name(source_dir: str | Path) -> str:
    """The database file name recorded in a source folder's vbe-project.json."""
    try:
        text = (Path(source_dir) / "vbe-project.json").read_text(encoding="utf-8-sig")
        name = str((json.loads(text).get("Items") or {}).get("FileName") or "")
    except (OSError, ValueError, AttributeError):
        return ""
    return name[4:] if name.startswith("rel:") else name


def expand_export_folder(value: str, db_file_name: str) -> str:
    """Expand environment variables and ``%dbName%`` the way the add-in does."""
    expanded = os.path.expandvars(value)
    return re.sub("%dbName%", lambda _m: db_file_name, expanded, flags=re.IGNORECASE)


_REFUSAL_GUIDANCE = (
    "Ask the user to set the project's ExportFolder to blank (the default "
    "<file name>.src beside the database) or to a folder below the database such as "
    "\\src\\, or work in the main checkout instead. vcs_tips(\"worktrees\") explains why."
)


def export_folder_refusal(path: str | Path, export_folder: str | Path | None) -> dict[str, Any] | None:
    """A refusal when ``path`` is in a git working tree and its export folder is not.

    A copy of the project (a linked worktree or a clone) whose export folder is
    outside its own working tree would read or write another checkout's source.
    A database outside any working tree is not affected.
    """
    if export_folder is None:
        return None
    root = working_tree_root(path)
    if root is None or is_inside(export_folder, root):
        return None
    return {
        "success": False,
        "error": (
            f"{path} is in the git working tree {root}, but its export folder "
            f"{export_folder} is outside it. Building, exporting, or importing here "
            "would read or write another checkout's source files, so nothing was done."
        ),
        "error_pattern": "export_folder_outside_working_tree",
        "export_path": str(export_folder),
        "working_tree": str(root),
        "agent_guidance": _REFUSAL_GUIDANCE,
    }


def build_source_refusal(source_dir: str | Path) -> dict[str, Any] | None:
    """A refusal for building ``source_dir`` when its absolute ExportFolder escapes its tree.

    Building the configured folder itself is allowed; a copy elsewhere in the
    same working tree is too, since it builds beside itself.
    """
    configured = configured_export_folder(source_dir)
    if is_relative_export_folder(configured):
        return None
    folder = expand_export_folder(configured, source_db_file_name(source_dir))
    if same_folder(folder, source_dir):
        return None
    return export_folder_refusal(source_dir, folder)


def _main_equivalent(path: Path, worktree: Worktree) -> Path | None:
    if worktree.main_checkout is None:
        return None
    try:
        return worktree.main_checkout / Path(path).resolve().relative_to(worktree.root)
    except ValueError:
        return None


def worktree_install_fields(source_dir: str | Path) -> dict[str, Any]:
    """Fields for a completed ``vcs_rebuild_addin`` whose source is in a worktree."""
    worktree = find_worktree(source_dir)
    if worktree is None:
        return {}
    main_source = _main_equivalent(Path(source_dir), worktree)
    restore = (
        f'vcs_rebuild_addin(r"{main_source}")' if main_source
        else "a rebuild from the main checkout"
    )
    return {
        "installed_from_worktree": True,
        "worktree": str(worktree.root),
        "main_checkout": str(worktree.main_checkout) if worktree.main_checkout else None,
        "warnings": [
            f"The installed add-in now comes from the worktree {worktree.root}. Every "
            "Access session on this computer loads it, including the user's own work "
            "and other agents' test runs."
        ],
        "agent_guidance": (
            "Tell the user the installed add-in now comes from this worktree. When the "
            f"branch is merged or dropped, put it back with {restore}."
        ),
    }


def _names(sources: list[dict[str, Any]]) -> str:
    names = [str(s.get("source")) for s in sources[:_LISTED_SOURCES]]
    extra = len(sources) - len(names)
    return ", ".join(names) + (f", and {extra} more" if extra > 0 else "")


def data_source_warnings(database_path: str | None, sources: Any) -> list[str]:
    """Warnings for a build's ``data_sources``: missing ones, and shared ones in a worktree."""
    if not isinstance(sources, list):
        return []
    entries = [s for s in sources if isinstance(s, dict)]
    worktree = find_worktree(database_path) if database_path else None
    warnings = []

    missing = [s for s in entries if s.get("kind") == "missing"]
    if missing:
        where = (
            f" Copy them from the main checkout ({worktree.main_checkout}) into this worktree,"
            " or add the .env keys, then build again."
            if worktree and worktree.main_checkout else ""
        )
        warnings.append(
            f"{len(missing)} linked data source(s) are missing, so the tables or "
            f"references on them were not created: {_names(missing)}.{where}"
        )

    shared = [s for s in entries if s.get("kind") in SHARED_KINDS]
    if worktree and shared:
        warnings.append(
            f"This database is in a git worktree, but {len(shared)} linked data "
            "source(s) are shared with the main checkout and every other copy: "
            f"{_names(shared)}. Check what any code you run writes to, and ask the "
            "user before writing to a source that may hold live data."
        )
    return warnings
