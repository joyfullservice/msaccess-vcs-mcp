# Developing the VCS add-in itself

Use this when the database you are working on is the msaccess-vcs-addin repository: `Version Control.accda` and its source folder `Version Control.accda.src`.

## Two copies of the add-in

- The **development copy** is the `Version Control.accda` beside `Version Control.accda.src` in the repository. Rebuilds and test runs use it.
- The **installed add-in** lives under `%APPDATA%\MSAccessVCS`. Every tool call loads it as a library. No tool accepts it as `database_path`, `output_path`, or `template_path`; such calls return `error_pattern: installed_addin_refused`. `vcs_get_version_info()` reports its version without opening it.

## Rebuilding after editing source

Edit the files in `Version Control.accda.src`, then call:

    vcs_rebuild_addin("<repo>\\Version Control.accda.src")

It builds the development copy, installs it, and returns when the status is `complete`, `cancelled`, or a `*-failed` status. Do not host the rebuild on a user database, the repository's Testing folder, or a scratch database. `vcs_rebuild_database` is a different tool: it builds a user project.

- The install is machine-wide: every Access session on this computer, including the user's own work and other agents' test runs, loads the new build. The rebuild also closes every Access window this server created that had the add-in loaded. From a git worktree the result carries `installed_from_worktree` and `main_checkout`; tell the user, and when the branch is done, rebuild from the main checkout. See `vcs_tips("worktrees")`.

- `refused`: nothing was rebuilt. An Access window the server did not create holds a file the rebuild must replace. `otherInstances` names each process and the file it holds. Ask the user to close those windows, then call again. Never close them yourself.
- `launch-failed`: the helper script never started. Access is left open and the call is safe to retry.
- If the client times out (`-32001`), call `vcs_get_recent_calls()`, then read `<source>/logs/rebuild-status.json`. A status whose `phaseStarted` predates your call belongs to an earlier run.
- For live output, run `msaccess-vcs rebuild-addin <source>` in a terminal and keep it in the foreground; the process exit is the completion signal.

Never patch the loaded add-in in place (`VBComponents.Import`, `LoadFromText`, or `vcs_import_object` aimed at `Version Control.accda`). It resets a VBA project that is running and can end the session.

## Running the add-in's own tests

    vcs_run_tests("<repo>\\Version Control.accda", "clsTestOptions")

Pass the development copy as `database_path`. The runner scans the current database's VBA project, so a user database or a Testing database finds that database's tests instead. The installed add-in supplies the runner and `TestAssert`, so run tests through this server, not from the add-in's own window. A result where every test is `EMPTY` (no assertions) means the harness was bypassed, not that the tests passed.

For live per-test output, run `msaccess-vcs run-tests <database>` in a terminal.

## Calling the add-in's API

Use `vcs_call_vba(db, "VCS.API", ["<Method>", ...])`. The server rewrites `VCS.API` to the add-in's full path. `"Version Control.API"` does not resolve, because `Application.Run` matches a loaded VBA project name, not a file name. Do not call the API from inside `vcs_run_vba`; that code is already delivered through the API, and the re-entrant call is refused.
