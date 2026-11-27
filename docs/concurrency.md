# Lanes, timeouts, and cancellation

How one server process schedules Access work for many agents, how a hung call is abandoned, and how a long operation is cancelled. Read this before changing `access_gate.py`, `vba_worker_manager.py`, `com_recovery.py`, `callback_server.py`, `operation_manager.py`, or how a tool reaches Access. Then run the concurrency suite ([AGENTS.md](../AGENTS.md#development-workflow)).

## One lane per database

One MCP server process is shared across all Cursor windows and every agent in them. `access_gate.py` gives each database a lane: a one-at-a-time slot and its own COM apartment thread. Calls on the same database run in order; calls on different databases, each in its own Access instance, run in parallel, up to `ACCESS_VCS_MAX_PARALLEL`. A call that names no database, and `vcs_rebuild_addin`'s launch (`run_exclusive_all`), wait for every lane and hold new lane calls off, because the rebuild closes server-created instances that loaded the add-in.

A tool whose database is not `database_path` names it in `_lane_target` (today `vcs_rebuild_database`). Its target comes from `_build_target`, which mirrors where the add-in will actually write: `output_path`, or else the file `vbe-project.json` names, placed at the `VCS Build Path` when the source folder is the one an absolute `ExportFolder` names, and beside the source folder otherwise. The instance-closing step of `vcs_rebuild_database` uses the same target.

A call that waits too long gets `error_pattern: server_busy` with `busy_with` naming the in-flight tool and a `reason`: `same_database`, `exclusive`, `exclusive_waiting`, or `max_parallel`. Lanes keep one database's calls in order, but two databases can export into one folder (an `output_dir`, or a shared absolute export folder), so `_claim_export_folder` returns `server_busy` with `reason: same_folder` for the second. The result says to retry rather than wait for a client `-32001` timeout. After any client timeout, call `vcs_get_recent_calls()` to see what actually executed.

- **Starting an instance is one step at a time across lanes.** `EnsureDispatch` returns whichever Access registered itself as active, so `_get_access_app` holds `_INSTANCE_CREATION_LOCK` from finding or starting an instance through opening its database. Without it, two lanes adopt the same fresh instance.
- **Module state must be per database or locked.** Recovery state (`com_recovery.py`), screenshot cleanup, and the ownership registry are; anything new that a tool keeps between calls has to be too.
- **The add-in keeps per-operation registry data per Access window** (`Timer <hwnd>`, `State <hwnd>`); a shared section let parallel instances pick up each other's operations.

## A hung `vcs_run_vba` call

`vcs_run_vba` executes Access COM work in a daemon worker thread with a hard timeout. If a snippet hangs because Access is in break mode, blocked on a modal dialog, or otherwise unresponsive, the MCP server abandons that thread and returns a recoverable timeout. It does **not** kill `MSACCESS.EXE` or close user-owned Access windows; after Access becomes responsive, the next call runs an automatic probe and resumes normal operation.

## Cancelling an operation

`vcs_cancel_operation(operation_id?, database_path?)` stops a running export, import, database rebuild, add-in rebuild, or test run; with no arguments it targets the only one running. It is gate-exempt and runs on the event loop, so every long tool must keep the loop free while Access works: whole-project export, merge, and test runs use `APIAsync` timer continuations, and the synchronous `ExportByType`, `ImportByType`, and `BuildHeadless` go through `_run_api_with_callbacks` on a worker thread. Never call a long add-in method with `call_sync` from an async tool.

- **The add-in polls; the server only answers.** `Operation.Pulse` asks `/cancel-status/{id}` about twice a second, and only when the callback info carries `cancel_poll: true`. The callback server is a `ThreadingHTTPServer` so that poll is answered while another callback is in flight. An add-in built before this support never polls, so its operations cannot be cancelled from MCP.
- **A wait that ends early abandons, never just unregisters.** On a wait timeout or client cancellation, `abandon_operation` marks the operation cancelled and keeps a ten-minute tombstone, so Access stops at its next poll instead of running on unobserved. A terminal callback for the id clears it.
- **Results say so.** A canceled run returns `cancelled: true` and `error_pattern: "operation_cancelled"`; a canceled add-in rebuild ends with status `cancelled` and installs nothing.

The callback protocol itself, including the cancellation poll, is in [VBA_CALLBACK_API.md](VBA_CALLBACK_API.md).
