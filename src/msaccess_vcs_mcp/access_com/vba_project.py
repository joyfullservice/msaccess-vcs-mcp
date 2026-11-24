"""Find the current database's VBA project.

``VBE.ActiveVBProject`` is whichever project the editor last touched. Once the
add-in is loaded as a library, that is often the add-in, so code that reads
modules or compiles must look the project up by file name instead.
"""

from __future__ import annotations

from typing import Any

from .instance_registry import paths_match


def current_vb_project(app: Any) -> Any:
    """The current database's VBA project, not a loaded add-in's."""
    vbe = app.VBE
    target = str(app.CurrentProject.FullName)
    for index in range(1, int(vbe.VBProjects.Count) + 1):
        project = vbe.VBProjects(index)
        try:
            if paths_match(str(project.FileName), target):
                return project
        except Exception:
            continue
    return vbe.ActiveVBProject
