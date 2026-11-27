# Working in a git worktree, a second copy of a project, or alongside other agents

A worktree (or any second clone) gets only the files tracked in git. It builds its own database, but everything gitignored stays behind, and everything outside its folder is shared with the original and every other copy.

## Before building

- Pass the worktree's own paths in every call. Paths in docs, transcripts, and the original's `.env` name the original's files.
- The project's `ExportFolder` option must be blank or relative (`\src\`). An absolute one names the original's folder, so build, export, and import in a worktree or clone are refused with `export_folder_outside_working_tree`. Ask the user to change the option in the original; do not change it yourself. Never use `\..\src\`: a build cannot find the database from it.
- Copy the gitignored inputs from the original folder: `.env` and `.env.*` in the export folder (connection strings for `env:conn_*` links), and any back-end or library database files that sit beside the database. Without them the build links nothing, or stops on a connection prompt.
- A project can do this when the worktree is created. Cursor runs `.cursor/worktrees.json`, for example:

      {"setup-worktree-windows": ["Copy-Item \"$env:ROOT_WORKTREE_PATH\\App.accdb.src\\.env\" App.accdb.src\\.env"]}

  The add-in repository's `tools/setup-worktree.ps1` does all of it: it copies the `.env` files and points their paths into the worktree, and with `-CopyDatabase` it also copies the built database, its `vcs-index.idx`, and the `rel:` back ends and libraries, skipping a database that is open. A copied database needs no build, and its first export is incremental.
- Otherwise build with `vcs_rebuild_database(source_dir="<worktree>\\App.accdb.src")`. The database lands beside that folder, inside the worktree, and keeps no path properties from the original.
- The folder must be an Access trusted location for the database's code to run without a prompt. Cursor's worktree folder (`%USERPROFILE%\.cursor\worktrees`) usually is not one; ask the user.

## What is shared

`vcs_rebuild_database` returns `data_sources`, one entry per linked back end or library reference, also listed at the end of the build log:

- `project`: a file inside the database's folder. Each copy needs its own; it was copied or built.
- `missing`: a file that does not exist. It was not copied, or the path is wrong. Tables on it were not linked (`notLinked`).
- `outside`, `network`: a file elsewhere on this computer or the network. Every copy reads and writes the same file.
- `server`: an ODBC or SharePoint connection. Every copy uses the same server.

Before running code that writes data (event procedures, action queries, `vcs_run_vba`), check what it writes to. If a shared source might hold live data, ask the user first. The result carries a warning when the build is inside a worktree and a source is shared.

Also shared: the installed add-in. `vcs_rebuild_addin` installs for every Access session on the machine, and closes every Access window this server created that had the add-in loaded.

## Several copies of one project at once

Copies with the same file name and the same VBA project name (`App.accdb` in each worktree) can be open at the same time, each in its own Access window. The server tells them apart by full path, so:

- Name a copy only by its full `database_path`, or by `pid`. Never pick a window by its title: a database that sets its own application title shows the same title in every copy, and a screenshot that can only match by name returns `ambiguous_instance` with the candidates.
- One copy per agent. Two agents on the same file share one Access window and take turns, and each one's imports overwrite the other's changes. Give each agent its own worktree instead.
- Leave the user's copy alone. If the user has the original open, work in a worktree; an import into a database someone else has open can be lost when they save, and the import result warns when that is the case.
- Several copies reading and writing one shared back end is ordinary multi-user Access. Changing that back end's design, or compacting it, needs it to yourself and fails while any copy has it open.
- `vcs_export_database` and `vcs_export_object` write to each copy's own export folder (`export_path` in the result). `output_dir` exports somewhere else instead (add-in 5.2.0 or later): always a full export, leaving the configured folder, its index, and the options untouched. In a worktree it must be inside the worktree, and a folder holding another project's `vcs-index.idx` is refused.

## Parallel agents

Several agents (for example one per worktree in an arena run) can share this server. Calls on the same database run one at a time; calls on different databases run in parallel, up to `ACCESS_VCS_MAX_PARALLEL` (default 3). Each database gets its own Access window.

- Work only on your own database. Never answer another agent's dialog or reset its VBA, even though the server created that window.
- `server_busy` means your database, or an exclusive operation, is busy. Retry.
- Do not call `vcs_rebuild_addin` during a parallel run; it closes every agent's Access window.
- Access runs only on this Windows machine, so the agents must be local, not cloud agents.
- A coordinating agent should create the worktrees, copy the gitignored files, and build each database before it starts the others, so the slow builds do not compete.
- A coordinator never sees a subagent's inline image. Capture with `vcs_capture_screenshot(..., save_image=True)` and report `image.saved_path` and `image.markdown`. If `image.retention` says nothing was saved, say so; do not describe a picture the coordinator cannot check.
