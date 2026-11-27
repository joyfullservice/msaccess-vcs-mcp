# Rebuilding the add-in and running its tests

How `vcs_rebuild_addin` and `vcs_run_tests` work on the VCS add-in's own repository, and why the installed add-in is never a target. Read this before changing either tool, `rebuild_watcher.py`, `_refuse_installed_addin_target`, or `get_default_addin_path()`. How agents use these tools is in [AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md) (sections 6b and 6c) and the [addin-development](../src/msaccess_vcs_mcp/tips/addin-development.md) tip.

## Rebuilding the VCS add-in

To rebuild `Version Control.accda` from source after editing add-in files, do **not** use `vcs_rebuild_database` (that rebuilds a user project). Call:

```python
vcs_rebuild_addin(r"C:\path\to\msaccess-vcs-addin\Version Control.accda.src")
```

The tool derives the development copy beside that folder, launches `RebuildAddIn`, and passes its HTTP callback identity through the disconnected worker to the builder Access process. Detailed `Log.Add` / `Log.Progress` messages stream during the build; the status file covers compile, install, and durable terminal recovery. Do **not** poll it yourself in the normal workflow. `vcs_call_vba(..., ["RebuildAddIn", source])` remains launch-only.

MCP progress notifications are best-effort in Cursor 3.13 (often only "Running..." until the tool returns). For guaranteed live output, run `msaccess-vcs rebuild-addin <source>` from a terminal and keep that command in the foreground so the stream stays in the primary chat. The CLI exits when the operation reaches terminal status; that process exit is the completion signal. Do not background the CLI just to wait on a notification, and do not add a second timer wait, sleep, or `rebuild-status.json` poll after it has already finished.

`refused` and `launch-failed` are returned immediately and mean nothing was rebuilt: before launch, and inside the Access gate so nothing can reopen the file in between, the server closes Access windows **it created** that hold a file the rebuild replaces — including instances whose only claim on the add-in is having loaded it as a library. `refused` when a **user-owned** `MSACCESS.EXE` still holds one of those files or cannot be asked which files it holds (`otherInstances` names what to close; the add-in never closes another process), and `launch-failed` when the helper script never started, which leaves Access open and is safe to retry.

If a client times out (`-32001`) or the tool returns `rebuild_stalled` / `timeout`, recover by reading `<source>/logs/rebuild-status.json` and matching `phaseStarted` against `rebuild_phase_started`. A `complete` whose `phaseStarted` predates the call is an earlier run's record. After any client timeout, call `vcs_get_recent_calls()` before inferring from the status file. A live rebuild always has `MSACCESS.EXE` or `wscript.exe`; neither, with a non-terminal status, means the run died.

## Running the add-in's own tests

Pass the **development copy** in the add-in's repository as `database_path`:

```python
vcs_run_tests(r"C:\path\to\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

A run needs two projects and they are different files: the installed add-in loads as a library and supplies the runner and `TestAssert`, while the code under test is whatever the current database holds. The runner scans the current VBA project, so the host decides which tests are found — aim a run at a user database, or anything in the repo's `Testing` folder, and you get that database's tests reported as a clean pass.

The installed add-in is refused as a host (see below); it also has no source tree beside it for the tests that read one. The add-in refuses such a run itself in `ExecuteTests` via `modInstall.CurrentDbIsInstalledAddIn`, so the server's refusal is the earlier of two.

`AccessConnection` opens the development copy itself (Access refuses to bind a file moniker to an `.accda`, so `GetObject` fails and the explicit `OpenCurrentDatabase` fallback in `_open_as_current_database` takes over), so no manual pre-open is needed.

Run these tests **through this server**, never from the add-in's own window. `modTestAssert.TestAssert` routes through `Application.Run` to the *installed* add-in path, while the runner singleton that records assertions lives in whichever project received `RunTests`. Invoking them from a development copy puts those in different projects: assertions are discarded and every test reports `EMPTY`. Treat an all-`EMPTY` result as a broken harness, not a pass.

MCP progress notifications are best-effort in Cursor. For live per-test output, run `msaccess-vcs run-tests <database>` from a terminal and keep that command in the foreground, the same way as `rebuild-addin`. The CLI prints pytest-style dots for fast passes, names tests that take a second or more, then a human completion line; it does not dump the full `tests` map. Headless means no add-in UI, not a hidden Access window.

## The installed add-in is never a target

No tool accepts the installed add-in as `database_path`, `output_path`, or `template_path`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper, ahead of the gate and of any COM work, so the rule holds for every tool rather than the handful that grew their own guards — export, import, rebuild, `run_vba`, `run_tests`, `call_vba`, and the rest. Refusals carry `error_pattern: installed_addin_refused`.

`vcs_get_version_info()` reports the installed add-in's version without opening it, which is the one thing you might legitimately want from that file.

The comparison ignores the file extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`: a compiled install is a `.accde` built from the same `.accda`, and only one of the two is ever named in `ACCESS_VCS_ADDIN_PATH`. Folder parameters (`source_dir`, `output_dir`) are not checked — an export folder beside the install is not the install.

**Resolving the install path.** With `ACCESS_VCS_ADDIN_PATH` unset, `get_default_addin_path()` reads `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`, which is the only place to read it from — the add-in's own `GetInstalledAddInFileName` is built from exactly these two values. `Install Folder` is present only for a folder the user chose (the installer deletes it when the folder is the default, so absence means `%AppData%\MSAccessVCS`, not "not installed"), and `Compile accde` decides the extension. Do not reconstruct the path from `%AppData%` alone or assume `.accda`. The result is cached for the life of the process; `reset_addin_path_cache()` clears it.
