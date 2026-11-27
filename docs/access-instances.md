# Access instances: visibility, ownership, and closing

Which Access windows the server shows, which ones it may close, and how it treats a database someone else has open. Read this before changing `access_com/connection.py`, `access_com/instance_registry.py`, `access_com/process_qos.py`, or `shared_open.py`. The rule that agents never close Access windows themselves is in [AGENTS.md](../AGENTS.md#invariants).

## Access window visibility

Any Access instance holding a database the server works with is left **visible**, whether the server created it or attached to one already running. A hidden instance strands the user: an error dialog, a VBA break, or a trust prompt blocks every later call with nothing on screen to explain why, and nobody can dismiss what they cannot see. `COM automation` normally starts Access hidden, so this is deliberate, not incidental.

Two rules follow from that, both enforced in `access_com/connection.py`:

- Open databases through `open_current_database(app, path)`, never a bare `app.OpenCurrentDatabase(...)`. It lowers `Application.UserControl` across the open — `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's own `AutoRun` opens its installer form when that flag says a person is watching, stranding the instance the server is about to drive — and it shows the window afterwards.
- Show the window *after* the database opens, via `ensure_access_visible(app)`. Making a window visible can itself set `UserControl`, which is why the order is not interchangeable.
- After that, instances the server created also get `UserControl = True` so the window is a normal interactive Access app. `Visible` alone is not enough on a COM-created process.

The one deliberate exception is `validate_access_installation()` in `config.py`: it opens no database and quits immediately, so a window would only flash on screen with nothing to act on.

## Full-power cores

Access is single-threaded. MCP-launched Access processes disable EcoQoS and run at Above Normal so Windows prefers a performance core. User-owned Access the server attaches to is left alone. Do not pin CPU affinity.

## Leaving Access open

With `ACCESS_VCS_LEAVE_ACCESS_OPEN` true (the default), a COM-created Access instance is not quit on disconnect, so a later attach can reuse the boosted process, and the user closes the window. The server still closes instances it created when a rebuild must replace a file they hold, or when it recycles a stuck owned instance after a failed recovery probe. User-owned Access is never closed.

## Which Access windows the server may close

Persistence only pays off if the server can still get a file back when it needs to replace one, so it has to know which windows are its own. That is recorded on disk in the ownership registry, keyed by PID *and* process creation time — a PID alone is reused by Windows, and a reused PID would authorize closing a stranger's process.

The rules, all enforced in `access_com/instance_registry.py` and `access_com/connection.py`:

- **Only a confirmed match is owned.** If the creation stamp cannot be read, or the record does not carry one, the answer is "not ours". Every ambiguity resolves toward leaving the window alone.
- **A failed process query is not an empty one.** `list_access_pids_or_none()` returns `None` when `tasklist` itself fails, and the registry keeps its records rather than forgetting every window it created.
- **Closing re-verifies identity.** Both ways of resolving an instance from a path can land on a different process than the one recorded, and a moniker bind can even start a new one. `pid` and creation time are re-checked against the record before anything is closed.
- **A hung owned instance is terminated.** The graceful `CloseCurrentDatabase` + `Quit` runs in a worker thread with `ACCESS_VCS_CLOSE_TIMEOUT_SEC`; a blocking call here would hang the Access gate. If the process survives, it is force-terminated. This is reachable only for a PID proven to be server-created, where unsaved state is acceptable loss.
- **A live process keeps its claim.** A record is dropped only once the process is confirmed gone. Forgetting a live owned instance would make it permanently uncloseable.
- **A loaded add-in counts as holding a file.** Every tool call loads the add-in as a library, which locks it no matter which database that instance has open. `load_addin` records this, and a rebuild closes those instances too — matching only on the open database would make rebuilds refuse almost every time.

The same proof decides which dialogs and break-mode sessions the server may answer ([screenshots.md](screenshots.md#dialogs-and-break-mode)).

## Imports into a database held elsewhere

Design and VBA changes merged into a database that another session has open can fail to save, or be dropped when that session saves. `vcs_import_objects` and `vcs_import_object` warn but never refuse, because the loss cannot be confirmed in advance. `shared_open.py` lists other local holders from the Restart Manager (minus the instance being driven and the server itself) and other computers from the `.laccdb` / `.ldb` lock file, whose entries outlive their sessions and so are labelled "may be stale". Each warning is one sentence: who holds it, the consequence, and the action (ask the user to close it, then verify with `vcs_diff_database`). `warnings` sits ahead of the per-object detail.
