<!-- BEGIN HEADER -->
# Decision Log

A reverse-chronological journal of architectural and strategic decisions.
Maintained by AI coding agents (and human developers) at the end of working
sessions. Each entry captures what was decided, what alternatives were
considered, and why — so future contributors never revisit dead ends or lose
context on trade-offs already evaluated.

Agents: read this file before working on any module referenced here.

### When to log

Log decisions that constrain future design, involved genuine alternatives,
or would be non-obvious to a future contributor. A good litmus test: does
the "What this rules out" section have something meaningful to say?

Do NOT log: bug fixes with obvious solutions, test-only refactors,
documentation updates, or minor config tweaks that don't affect
architecture.

### Entry format

Insert new entries directly below this header, newest first. Do not modify
or reorder existing entries except to add supersession notes (see below).
If a session produced multiple independent decisions, create a separate
entry for each.

```
---

## YYYY-MM-DD — [Short descriptive title]

**Trigger**: What problem, requirement, or situation prompted this work.

**Options explored**:
- For each option, name the approach, its strengths, and why it was or
  wasn't chosen. Include options that were tried and reverted.

**Decision**: What was chosen and the core trade-off.

**What this rules out**: Future directions now constrained or foreclosed.
What would trigger revisiting this decision.

**Relevant files**: Key files created or modified.
```

### Guidelines

- Focus on **why**, not what. The diff shows what changed; this log
  explains the reasoning.
- Capture rejected alternatives with equal care. Future agents need to
  know what was already tried.
- Be specific — name libraries, files, config choices, error messages.
- Aim for 10–50 lines per entry. Reference document, not narrative.
- Plain language. No jargon, no editorializing, no padding.

### Superseded entries

When a new decision invalidates, corrects, or replaces guidance in an older
entry, add a blockquote annotation to the affected older entry — do not
rewrite or delete its original text. Place the note immediately after the
entry's heading or after the paragraph containing the superseded claim.

> **⚠ Superseded** (YYYY-MM-DD): [Brief explanation of what changed and
> why.] See "[title of newer entry]" above.

Use **⚠ Partially superseded** when only specific claims are affected, and
**⚠ Superseded** when the entire entry's premise or decision has been
overturned. Always scan older entries for claims that conflict with the new
decision — agents reading the log linearly will otherwise encounter
contradictory guidance.
<!-- END HEADER -->

---

## 2026-11-27 — Saved screenshots: opt-in per call, expire by age

**Trigger**: Agents could see a `vcs_capture_screenshot` image but not show
it: the PNG only appeared inside the collapsed tool result. Cursor's chat
renders a local image only from a Markdown link written `/C:/...` (forward
slashes, leading slash). `C:\...`, `C:/...`, and `file:///` do not render.
That needs a file on disk, which the tool had never written, because
captures can show client data.

**Options explored**:
- Newest-N count (first draft, 50): rejected. A busy day, a test run, or a
  second Cursor window evicts images that today's chat still shows, and files
  from a quiet week stay forever.
- A size or count cap on top of age: rejected. It brings back eviction of
  recent images to guard against bursts that opt-in saving makes unlikely.
- No pruning: rejected; client pixels would stay on disk with no end date.
- Delete at session end: rejected. The server restarts a median of every
  6.8 minutes, so images would break mid-conversation.
- `%TEMP%`: rejected; Storage Sense deletes on its own schedule.
- `%LOCALAPPDATA%`: rejected. It splits server state across two profile
  roots; a roaming profile is handled by `ACCESS_VCS_SCREENSHOT_DIR`.
- A per-workspace or agent-chosen folder: rejected. Either is one `git add`
  from a commit, and the server cannot reliably tell which workspace a call
  came from.
- Prune at startup in `main.py`: rejected. It runs before the workspace
  `.env` loads, so it could delete files a user had set to keep.
- Refuse `include_image=False` with `save_image=True`: rejected. The user
  asked for posting without image tokens; the docstring says the agent has
  not seen those pixels.
- Drop `readOnlyHint`: rejected. Every read-only tool already writes the
  usage log, the tool writes only the server's own folder, and dropping the
  hint could add confirmation prompts to routine stuck-Access checks.
- **Chosen**: `save_image=True` writes `vcs-shot-YYYYMMDD-HHMMSS-<target>-<8
  hex>.png` to `~/.msaccess-vcs-mcp/screenshots/` and returns `saved_path`,
  `markdown`, and the `target` the pixels show (`application` for a
  fallback). Files older than `ACCESS_VCS_SCREENSHOT_RETENTION_DAYS` (30, `0`
  keeps them) are deleted after each save and once per process after the
  first `load_config()`, off the event loop. Only `vcs-shot-` names past the
  limit are ever deleted, so concurrent pruners and shared folders are safe.

**Decision**: Opt-in per call, kept for a fixed age. The pixels repeat data
already on the machine in the database, so the risks worth designing for are
a commit and a sync, which the folder location covers; 30 days keeps last
month's chats intact. The cost is that a chat older than the retention
period loses its images. Checked in Cursor (2026-11-27): a link under a folder
named with spaces and parentheses, percent-encoded, rendered; the same link
posted after its file was deleted showed nothing, so Cursor reads the file
when it displays the chat rather than keeping a copy.

**What this rules out**: Object or database names in file names; writing
into a repository or workspace by default; a count cap. Revisit the default
if Cursor starts copying images into the chat, which would let files expire
much sooner.

**Relevant files**: `src/msaccess_vcs_mcp/screenshot_store.py`,
`src/msaccess_vcs_mcp/screenshot.py` (`_finish`, `capture_screenshot`),
`src/msaccess_vcs_mcp/tools.py` (`vcs_capture_screenshot`, `vcs_tool`),
`tests/conftest.py`, `tests/test_screenshot_tool.py`.

---

## 2026-11-24 — Rules in the instructions, next steps in results, depth in vcs_tips

**Trigger**: The server instructions had grown to 10,205 characters. Claude
Code cuts instructions at 4,096 (2,048 before v2.1.296), so its users lost
everything after the middle of the installed-add-in section: compile
failures, "Access seems stuck", the tool reference, and the logs. The OpenAI
Agents SDK and LangChain adapters do not document passing instructions to
the model at all. `vcs_run_vba`'s description (4,134) also crossed the limit
for a tool loaded upfront.

**Options explored**:
- Trim the instructions in place: rejected alone. It fits Claude Code but
  still leaves clients that drop instructions with nothing at failure time.
- Move everything into tool descriptions: rejected. Descriptions load late in
  Claude Code (through tool search) and the same rule would repeat in many.
- A `hint` field per failure site, as before: rejected. Only three sites had
  one, each worded differently, and none named the call to make.
- **Chosen**: three layers. The instructions keep the rules every call needs
  (installed add-in, no parallel calls, the stuck-Access call, `vcs_tips`
  triggers) in the first 2,000 characters and stay under 3,500; they survive
  context compaction in long sessions. Failed results carry one
  `agent_guidance` sentence with the exact next call, added centrally in the
  `vcs_tool` wrapper by `error_pattern` unless the result already has its
  own. `vcs_tips(topic)` serves packaged markdown topics on demand; it is
  gate-exempt so `stuck` answers while Access is blocked, and it is the only
  tool marked `anthropic/alwaysLoad` so Claude Code agents see it without a
  search.

**Decision**: Durable rules in the instructions, failure guidance in results,
situational depth in `vcs_tips`. The cost is three places to keep in step;
`tests/test_agent_text.py` checks lengths, placement, and that every tool
name and topic mentioned exists.

**What this rules out**: Growing the instructions back into a manual, and a
second guidance field (`hint`, `tips_topic`). Guidance never goes in `error`,
which the usage log records. Revisit the limits if Claude Code changes them.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`SERVER_INSTRUCTIONS`,
`vcs_tips`, `vcs_tool`), `src/msaccess_vcs_mcp/guidance.py`,
`src/msaccess_vcs_mcp/tips.py`, `src/msaccess_vcs_mcp/tips/`,
`tests/test_guidance.py`, `tests/test_agent_text.py`.

---

## 2026-11-24 — No tool can change the McpAllow* permissions

**Trigger**: `vcs_set_option` refused `McpAllowRunVBA`, but
`vcs_call_vba(db, "VCS.API", ["SetOption", "McpAllowRunVBA", "True"])`
reached the add-in's `SetOption` directly and turned it on. The add-in also
merged `McpAllow*` keys from session override files in `mcp/`, which an agent
can write with ordinary file tools. The other three permissions
(`McpAllowCallVBA`, `McpAllowExecuteSQL`, `McpAllowImport`) were not
protected at all.

**Options explored**:
- Keep the Python-only list and add a matching check in `vcs_call_vba`:
  rejected as the only fence. Any caller of the add-in API, and any override
  file, would still get through.
- Detect "called from MCP" in the add-in and refuse only then: rejected. VBA
  cannot tell its caller apart reliably, and the user never needs an API
  route to these options; the MCP options form is where they are set.
- **Chosen**: the add-in refuses all four in `SetOption` and ignores them in
  override files (`clsOptions.IsUserControlledOption`). The server refuses
  them too, in `vcs_set_option` and in `vcs_call_vba` calls to `SetOption`,
  so an add-in build older than this check is still covered for the direct
  routes. Every refusal returns `error_pattern: user_controlled_option` and
  tells the agent to ask the user.

**Decision**: The permissions are the user's decision, enforced in the
add-in, with the server as a second check.

**What this rules out**: Agents enabling a permission for a session, even
temporarily. The fence is not complete: the project's `vcs-options.json`
holds these options and is an ordinary file in the export folder, and
importing VBA then calling it through `vcs_call_vba` is code execution
whatever `McpAllowRunVBA` says. `ACCESS_VCS_DISABLE_WRITES` refuses import
and rebuild, which closes that route, but it does not gate `vcs_call_vba` or
`vcs_run_vba`; the README spells out both limits. Revisit if the options file
moves out of agent reach.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`USER_CONTROLLED_OPTIONS`),
`tests/test_user_controlled_options.py`. Add-in side: `clsOptions.cls`,
`clsVersionControl.cls` (`SetOption`).

---

## 2026-11-24 — Answer dialogs and reset break mode only on server-created Access

**Trigger**: Screenshots made dialogs and break mode visible to agents, but
an agent still had to stop and ask the user to click OK or press Reset before
any autonomous work could continue.

**Options explored**:
- Let `vcs_capture_screenshot` click buttons: rejected. It would no longer be
  a read-only tool, and an agent asking "what is Access doing?" could change
  what Access does next.
- Allow dismissal on user-owned Access behind an opt-in setting: rejected
  for now. A user's window may hold unsaved work or a deliberate debugging
  session, and a process list cannot tell an agent whose window it is.
- Click by title or button caption alone: rejected. The dialog may have
  been answered and replaced between the screenshot and the click; a `Yes`
  meant for one prompt would land on the next.
- **Chosen**: separate `vcs_dismiss_dialog` and `vcs_reset_vba`, limited to
  instances with a confirmed PID plus creation-time match in the ownership
  registry, refused when writes are disabled, and gate-exempt because the
  stuck call holds the gate. A dismiss must quote the `dialog_id` (a hash of
  the window handle, title, text, and buttons) it read.

**Decision**: Owned instances only, with a content fingerprint as the click
guard. The cost is that a user's own Access still needs the user.

**What this rules out**: Clicking dialogs in a window the server did not
create. Revisit if users ask for unattended work in their own window; that
would need an explicit per-session opt-in, not a default.

**Relevant files**: `src/msaccess_vcs_mcp/vba_control.py`,
`src/msaccess_vcs_mcp/accessibility.py`, `src/msaccess_vcs_mcp/access_gate.py`

---

## 2026-11-24 — Compile from the VBA editor to report the failing line

**Trigger**: `vcs_compile_vba` returned only success or failure, and agents
were told to ask the user to compile by hand and paste the error line.

**Options explored**:
- Keep the add-in's `CompileVBA` (`acCmdCompileAllModules`): it shows
  Office's "Cannot Compile Project." dialog, blocks the caller until someone
  clicks OK, and never moves to the error, so there is no location to read.
- Read the location from a screenshot alone: the code window title gives
  the module, and the image shows the highlighted line, but window calls
  cannot read a line number from the custom-drawn code pane.
- **Chosen**: run the editor's own Compile control (578) over COM. It
  returns at once, then the editor highlights the error and shows a
  standard "Compile error:" dialog. The gate thread waits for that dialog
  with window calls only (a COM call made while it is up waits behind it),
  records the message and the module from the window title, clicks OK, and
  reads the line and code from `ActiveCodePane` over COM. Success is
  detected by polling `IsCompiled` on a separate COM thread.

**Decision**: Editor-side compile with two sources of evidence: window
evidence while the dialog is up, and the exact selection afterwards. The
add-in route stays as a fallback when the control is missing.

**What this rules out**: Compiling through `acCmdCompileAllModules` from
the server. The editor is left open on the error, which is visible to the
user; revisit if that proves disruptive.

**Relevant files**: `src/msaccess_vcs_mcp/compile_check.py`,
`src/msaccess_vcs_mcp/tools.py`

---

## 2026-11-20 — Cancel by add-in poll; every long tool keeps the event loop free

**Trigger**: An operation started by automation could not be stopped: the
add-in's progress window refused to cancel in unattended mode, and
`vcs_cancel_operation` only set a flag the add-in never read (its COM call was
commented out). The add-in side now funnels every cancel into
`Operation.RequestCancel`; see the add-in's DECISIONS.md entry of the same date.

**Options explored**:
- *Call the add-in's `Cancel` API over COM from `vcs_cancel_operation`.*
  Rejected. Access is single-threaded and busy running the operation, so the
  call queues behind it; the server could also hang on a modal dialog.
- *The add-in polls `/cancel-status/{id}` from `Operation.Pulse` (chosen).*
  About twice a second, with sub-second `ServerXMLHTTP` timeouts. The server
  had to become a `ThreadingHTTPServer`: on the single-threaded one a poll
  queued behind a log callback, and every poll timed out (seen as a stalled
  add-in rebuild). The add-in polls only when the callback info carries
  `cancel_poll: true`, so an older server is never polled.
- *Keep `call_sync` for scoped export/import and `BuildHeadless`.* Rejected.
  `vcs_cancel_operation` is gate-exempt and runs on the event loop, and those
  calls blocked it, so a cancel could not arrive until the work was done. They
  now go through `APIAsync` with callback info on a worker thread
  (`_run_api_with_callbacks`); an older add-in falls through to `API` there.
- *Unregister an operation when its wait times out.* Rejected. Access keeps
  running with nobody watching. `abandon_operation` marks it cancelled and
  keeps a tombstone for ten minutes, or until a terminal callback arrives.

**Decision**: Cancellation is cooperative and add-in driven. The server
records the request; the add-in asks for it. `vcs_cancel_operation` takes an
optional id or database path, and with neither cancels the only running
operation. A timed-out or abandoned wait cancels as well. Canceled results
carry `cancelled: true` and `error_pattern: "operation_cancelled"`, and a
canceled add-in rebuild ends with status `cancelled`.

**What this rules out**: Calling a long add-in method with `call_sync` from an
async tool. Killing `MSACCESS.EXE` to cancel. Revisit if Access gains a way to
interrupt a running VBA call from outside.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_cancel_operation`,
`_run_api_with_callbacks`, `_run_scoped_api`, `_mark_cancelled`),
`src/msaccess_vcs_mcp/operation_manager.py` (`abandon_operation`, tombstones,
`normalize_database_path`), `src/msaccess_vcs_mcp/callback_server.py`,
`src/msaccess_vcs_mcp/rebuild_watcher.py`, `tests/test_cancel_operation.py`,
`docs/VBA_CALLBACK_API.md`. Add-in side: `clsOperation.cls`, `clsMCP.cls`,
`modAPI.bas`, `docs/automation-contract.md`.

---

## 2026-10-30 — Unattended vcs_rebuild_database through BuildHeadless

**Trigger**: A field report showed `vcs_rebuild_database` failing with "The
application called an interface that was marshalled for a different thread"
before reaching the add-in, and `vcs_import_objects` waiting out its timeout on
a database with no index. Fixing the thread error exposed two more failures
that would have come next: `BuildAs` shows a folder picker and a Save As
dialog and ignores `output_path`, and `Application.Run` into the add-in with
no database open crashes Access ("The server threw an exception", then RPC
server unavailable).

**Options explored**:
- *Keep the async `Build` / `BuildAs` path and suppress its dialogs.* Rejected.
  `BuildAs` exists to ask a person where to build; teaching it to stay quiet
  duplicates `BuildHeadless`, which already forces unattended mode and returns
  a JSON outcome.
- *`BuildHeadless` with no database open, as the add-in's CI example did.*
  Rejected after testing: Access needs a current database to run a library
  procedure. The tool opens a scratch database in a temp folder, calls the
  build from there, and deletes it. The add-in's Wiki example was corrected too.
- *Always pass `output_path` to `BuildHeadless`.* Rejected. An extra argument to
  an older add-in raises inside the library, and Access shows an error dialog
  that blocks the call. The tool reads `GetVCSVersion` first: 5.1.0 is needed
  for `BuildHeadless`, 5.1.1 for its output path. `output_path` is now optional
  so a user on 5.1.0 can still build to the file the source files name.
- *Pre-check the index from the server before `MergeBuild`.* Not done. The
  add-in now posts an `error` callback with `code: merge_not_available` and
  returns `False`, which covers every caller of the API, not only this server.
  Add-ins older than 5.1.1 still wait out the timeout.
- *Attach with `ensure_dispatch`, as the old code did.* Rejected. It can return
  a window the user already has open; the old code then quit that window, and
  opening a scratch database there would also close whatever it held. Found in
  testing with a development copy left open by a test run ("You already have
  the database open"). Every build gets its own process from `DispatchEx`.

**Decision**: `vcs_rebuild_database` starts its own Access process, loads the
add-in from a scratch host database, and runs `BuildHeadless(source[, output])`
on a worker thread, polled from the event loop with a heartbeat progress
message every 15 seconds and bounded by `ACCESS_VCS_BUILD_TIMEOUT_SEC` (default 600). On timeout Access is
left open, since the build may still be running there. When no database path
can be looked up in the Running Object Table, worker threads get the caller's
proxy through `CoMarshalInterThreadInterfaceInStream` /
`CoGetInterfaceAndReleaseStream` (`VCSAddinIntegration.worker_app_source`).
Access is out of process, so the unmarshalled proxy calls Access directly; a
probe against real Access answered while the main thread sat blocked in
`join`, so the timeout still holds. The marshal has to happen on the thread
that owns the proxy, which is why the call starts on the COM executor and the
wait runs on the event loop instead of using `asyncio.to_thread`.

**What this rules out**: Per-object build progress in `vcs_rebuild_database`;
`BuildHeadless` returns only when the build is over. `template_path` is
accepted and ignored. Revisit if the add-in grows a headless build that
streams callbacks.

> **⚠ Partially superseded** (2026-11-20): `BuildHeadless` now runs through
> `APIAsync` with callback info, so the build log streams as progress and the
> build can be canceled; the heartbeat remains for quiet stretches, and
> `_wait_for_build` is replaced by `_run_api_with_callbacks`. See "Cancel by
> add-in poll; every long tool keeps the event loop free" above.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_rebuild_database`,
`_create_build_instance`, `_open_build_host`, `_wait_for_build`,
`_start_application_call`, `_merge_fallback_failure`), `src/msaccess_vcs_mcp/addin_integration.py`
(`worker_app_source`, `merge_build`, `MERGE_NOT_AVAILABLE_MESSAGE`),
`src/msaccess_vcs_mcp/vba_worker_manager.py` (`get_build_timeout`),
`tests/test_rebuild_database.py`. Add-in side: `clsVersionControl.cls`
(`MergeBuild`, `BuildHeadless`), `docs/automation-contract.md`.

---

## 2026-09-10 — Persist server-created Access; own it by PID

**Trigger**: Merge / test / edit loops were spawning a new `MSACCESS.EXE` per
tool call. Attach costs ~3 ms; spawn+quit costs ~3 s median. Ninety-seven
`vcs_run_vba` / import / test calls failed with "Cannot find Access instance"
because the worker only looks in the ROT. `_owns_app` is recomputed per
connection, so leaving a window open reclassified it as user-owned.

**Options explored**:
- *Keep quit-per-call, document `ACCESS_VCS_LEAVE_ACCESS_OPEN`.* Rejected.
  The flag already existed and was unused in this workspace; the reliability
  failures are the default path.
- *In-memory owned-PID set.* Rejected. The server restarts a median of every
  6.8 minutes; an in-memory set would orphan windows several times an hour.
- *Never Quit when UserControl is True.* Still rejected: owned instances
  deliberately set `UserControl` so they look interactive.
- *Persist by default + disk PID registry* (chosen). GetObject reattaches, so
  the bound is one window per distinct database, not an accumulating leak.
  Rebuild pre-flight closes registry-owned holders; user-owned still refuse.
  Stuck owned instances are recycled after a failed recovery probe.

**Decision**: Default `ACCESS_VCS_LEAVE_ACCESS_OPEN` to true. Record
`(pid, create_time, database_path)` in `~/.msaccess-vcs-mcp/owned-instances.json`.
`_owns_app` now means "the server created this process and may close it when
it must," not "quit at end of call."

**What this rules out**: Quitting a leftover owned instance at the end of
every tool call. Revisit if long-lived Access shows a correctness problem
(stale VBA project, memory growth) that in-process reset cannot fix.

**Relevant files**: `access_com/instance_registry.py`,
`access_com/connection.py`, `vba_worker_manager.py`, `tools.py`, `main.py`.

---

## 2026-09-10 — Ownership needs proof; closing needs identity

**Trigger**: Review of the persistence work above. Ownership had become the
authority to close somebody's Access process, but the code answered "is this
ours?" with heuristics that failed toward yes: `is_owned` accepted a record
with no creation stamp, a failed `tasklist` looked identical to "no Access
running" and silently wiped the registry, and closure resolved an instance by
*path* and quit whatever came back. Two rebuild-blocking bugs sat alongside
it: `GetObject(path)` launches Access when nothing holds the file, and that
process was never registered; and every tool call loads the add-in as a
library, which locks it, while the pre-flight matched only on open database.

**Options explored**:
- *Treat unconfirmed as owned, since the server usually is the one running
  Access.* Rejected. The cost of a false positive is closing a user's window
  with unsaved work; the cost of a false negative is a rebuild that refuses
  and tells the user which window to close.
- *Verify ownership by `UserControl` or window title instead of a registry.*
  Rejected again (as in the entry above): owned instances deliberately set
  `UserControl`, so it cannot distinguish them.
- *Leave a hung owned instance alone and let the rebuild refuse.* Rejected on
  the user's call: the instance is server-created, so unsaved state is
  acceptable loss, and refusing leaves a window nobody can act on.
- *Match the add-in by path at each of the fifteen `load_addin` call sites.*
  Rejected. Recording it once inside `load_addin` covers every path through
  the code, including ones added later.

**Decision**: Ownership requires a confirmed `(pid, create_time)` match;
every ambiguity resolves to "not ours". `list_access_pids_or_none()`
distinguishes a failed process query from an empty one, and the registry
keeps its records when liveness is unknowable. Closure re-verifies pid and
creation time against the record before touching an instance, runs the
graceful `Quit` in a worker thread under `ACCESS_VCS_CLOSE_TIMEOUT_SEC`
(a blocking COM call here would hang the Access gate), then force-terminates
a survivor. A record is dropped only once the process is confirmed gone.
`load_addin` records the loaded add-in so rebuild pre-flight can close
library holders, and that pre-flight moved inside the gate.

**What this rules out**: Any ownership signal that is not the durable
registry, and any close path that authorizes on a path match. Revisit the
terminate policy if owned instances ever hold state a user would miss —
today they do not, because the user's own windows are never server-created.

**Relevant files**: `access_com/instance_registry.py`,
`access_com/connection.py`, `access_com/process_qos.py`,
`addin_integration.py`, `tools.py`, `vba_worker_manager.py`.

---

## 2026-08-28 — Optional leave-open for COM-created Access

> **⚠ Superseded** (2026-09-10): Persist-by-default plus a durable PID
> registry replaced the opt-in flag as the default. The flag remains as an
> explicit `false` to restore quit-per-call. The 2026-08-28 leak concern
> assumed accumulation; GetObject reattach bounds the count to one window
> per database. See "Persist server-created Access; own it by PID" above.


**Trigger**: Sequential `msaccess-vcs run-tests` always Quit the Access the
CLI child created, so there was no way to time a boosted process and then
reattach to the same PID. QoS is process-lifetime; attach to a leftover
boosted instance is the comparison that matters.

**Options explored**:
- *Never Quit when UserControl is True.* Rejected as the default. Owned
  instances would leak Access processes after every CLI or MCP session.
- *ACCESS_VCS_LEAVE_ACCESS_OPEN* (chosen). When set, `close()` still releases
  COM but skips `CloseCurrentDatabase`/`Quit`. Attach later finds the same
  PID via GetObject. Default remains Quit for owned instances.

**Decision**: Opt-in env flag only. User-owned Access is unchanged.

**What this rules out**: Leaving every MCP-created Access running by default.

**Relevant files**: `access_com/connection.py`, `tests/test_access_visibility.py`.

---

## 2026-08-28 — CLI test output is compact pytest-style dots

**Trigger**: `msaccess-vcs run-tests` printed every MCP log fragment (Access
same-line console layout) and then the full per-test JSON. A 468-test run
took ~3 minutes, most of it HTTP. The last line was not a human summary.

**Options explored**:
- *Leave the stream as raw `Log.Add`.* Rejected after the first full run.
- *Live current-test name via `\\r` on stderr.* Rejected. Cursor's capture is
  not a TTY, so overwrite is unreadable.
- *Infer a dot from each start-of-test `(n/m)` progress.* Rejected. That is
  468 HTTP posts. The add-in now coalesces fast passes into log lines.
- *Print `.` from those batched log payloads; ignore `(n/m)`; compact JSON
  plus a human completion line* (chosen). Wrap dots at 80. Names remain the
  add-in's ≥ 1s PASS and FAIL/ERROR/EMPTY lines. `vcs_run_tests` as an MCP
  tool still returns the `tests` map.

**Decision**: The CLI is for watching. The MCP tool is for programmatic
reruns. Completion is a sentence such as `Tests passed. 468 subs, 2281
assertions, 3 empty in 14.42s`, not a sentinel.

**What this rules out**: Dumping the `tests` map to CLI stdout. Printing
`(n/m)` progress during `run-tests` (rebuild/export still do). Using progress
callbacks as the pytest dot stream. A `VCS_CLI_EXIT` token.

**Relevant files**: `cli.py` (`ProgressPrinter`, `compact_result_payload`,
`completion_message`), add-in `clsLog.cls`, `clsTestRunner.cls`.

---

## 2026-08-28 — Agent test runs stream through APIAsync and the CLI

**Trigger**: Agent-initiated test runs returned one JSON blob after the suite
finished. Cursor shows only "Running..." for MCP progress, so a long run looked
stuck. Builds already streamed: `APIAsync` + HTTP `Log.Add`/`Log.Progress` +
`msaccess-vcs rebuild-addin` printing each notification. Tests already posted
those callbacks when `MCP.IsActive`, but `vcs_run_tests` used blocking
`call_sync("RunFilteredTests")`, which never registered a callback.

**Options explored**:
- *Keep the sync COM call and tail `TestRun_*.log`.* Rejected. Duplicates the
  callback path builds already use, and the log is gitignored so agents miss it.
- *Register the callback then block in `call_sync`.* Rejected. The tool handler
  would hold COM for the whole suite and could not emit MCP progress until
  return.
- *Same async path as export* (chosen). `RunFilteredTests` joins the
  `APIAsync` timer list. MCP waits on callbacks. `msaccess-vcs run-tests`
  prints each line. Old add-ins still return `{sync: true, result: ...}`.

Failed tests complete as `eorFailed`, which posts type `error`. The results
file rides on every terminal callback as `results_path` so MCP can still
return per-test JSON for reruns. What the CLI prints is the compact stream
above, not every Access-console fragment.

**Decision**: `vcs_run_tests` follows the export pattern. The CLI is the live
output path; the MCP tool remains for short or programmatic runs. Tool
`success` still comes from the runner summary, not from `Operation.Result`.

**What this rules out**: Treating a finished-with-failures suite as a lost
result. Adding a second log-tailer for tests. Exempting `vcs_run_tests` from
the Access gate (the suite runs in the connected instance).

**Relevant files**: `tools.py` (`vcs_run_tests`), `cli.py`,
`operation_manager.py`, add-in `modAPI.bas`, `clsOperation.cls`,
`clsTestRunner.cls`.

---

## 2026-08-28 — Owned Access instances get UserControl after the database opens

**Trigger**: Headless MCP test runs created a new Access for
`Version Control.accda` and set `Visible = True`, but the window often never
appeared as a normal interactive app. Automation-created Access with
`UserControl` still False stays off the desktop even when Visible is True.
The common `EnsureDispatch` new-instance path never set `UserControl`; only
the isolated `DispatchEx` path restored it after open.

**Options explored**:
- *Set UserControl before OpenCurrentDatabase.* Rejected. AutoExec/`AutoRun`
  would see a person watching and open the installer form.
- *Visible only, as in the 2026-08-12 rule.* Insufficient for COM-created
  instances; that is the gap this closes.
- *UserControl = True after open, on owned instances only* (chosen). AutoExec
  already ran with the flag down. Attached user instances are left alone.

**Decision**: `_get_access_app` still shows the window after the database is
open, then sets `UserControl = True` when `_owns_app`.
`validate_access_installation()` stays hidden (no database, quit immediately).

**What this rules out**: Hiding MCP-created Access for tidiness. Setting
`UserControl` before a database opens. Forcing `UserControl` on instances the
server attached to.

**Relevant files**: `access_com/connection.py`
(`_ensure_owned_instance_interactive`), `tests/test_access_visibility.py`.

---

## 2026-08-28 — MCP-launched Access prefers a full-power core

**Trigger**: The add-in's agentic rebuild was slower than a ribbon rebuild of the
same source because Windows scheduled the COM-launched, windowless Access
process onto an LP-E core. Access is single-threaded, so that one core is the
operation.

**Options explored**:
- *Pin CPU affinity to P-cores.* Rejected. Topology differs by SKU, and pinning
  is the wrong default on a machine with only performance cores.
- *Promote every Access process the server sees.* Rejected. The server often
  attaches to a database the user already has open.
- *EcoQoS off + Above Normal for MCP-launched processes only* (chosen).
  `AccessConnection` applies it when `_owns_app` is true. `vcs_rebuild_database`
  applies it to an empty `EnsureDispatch` instance, not one that already has a
  database open. `vcs_rebuild_addin` snapshots `MSACCESS.EXE` PIDs before launch
  and promotes processes that appear afterward (builder, silent installer).
  The probe in `validate_access_installation` is not promoted: it quits
  immediately. Failures are swallowed.

**Decision**: Disable `PROCESS_POWER_THROTTLING_EXECUTION_SPEED` and set Above
Normal. Do not set an affinity mask. User-owned Access is unchanged.

**What this rules out**: Pinning cores. Changing QoS on an Access instance the
user launched. Assuming this closes the entire agent/ribbon gap.

**Relevant files**: `access_com/process_qos.py`, `access_com/connection.py`,
`rebuild_watcher.py`, `tools.py`.

---

## 2026-08-27 — Carry the existing HTTP stream through add-in self-rebuild

**Trigger**: The first `vcs_rebuild_addin` smoke test completed in Access but
showed only "Rebuild launched" and then hung. Even without the hang, watching
the coarse status file could report only `building`, `compiling`, and
`installing`, while the build log contained the useful per-object detail.

**Options explored**:
- *Tail `Build_*.log` from Python.* Viable fallback, but duplicates the stream
  already emitted by `Log.Add` and `Log.Progress`, requires discovering a
  timestamped file, and cannot preserve native progress values.
- *Have Worker.vbs implement HTTP posting.* Rejected. That would be a second
  callback implementation and a second formatter.
- *Pass the existing callback identity to builder Access* (chosen). The parent
  callback server outlives the launching Access instance, and the builder can
  invoke the existing `APIAsync(..., "Build", source)` path.

**Decision**: `vcs_rebuild_addin` registers an operation before launch and
passes its callback JSON as the second `RebuildAddIn` argument. The add-in
extracts the shell-safe URL and operation ID, carries them through Worker.vbs,
and reconstructs the callback JSON in builder Access. Build `log`, `progress`,
and terminal callbacks feed the same monotonic reporter as status phases. A
build `complete` callback ends only the callback pump; the status watcher
remains authoritative for compile/install and the final result. The original
`phaseStarted` now also crosses the worker boundary instead of being replaced.
`ReadDirectoryChangesW` uses overlapped I/O and `CancelIo`, avoiding the
deadlock caused by closing a handle with a synchronous read on another thread.

**What this rules out**: Tailing the build log as the primary live channel.
Reimplementing HTTP in Worker.vbs. Treating the build-phase `complete` callback
as proof that compile/install finished. Replacing durable terminal status with
an ephemeral callback.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`,
`operation_manager.py`, `rebuild_watcher.py`, `cli.py`,
`tests/test_rebuild_watcher.py`; add-in `clsVersionControl.cls` and
`clsWorker.cls`.

---

## 2026-08-27 — Real-time operation feedback: monotonic MCP progress, rebuild watcher, CLI

> **⚠ Partially superseded** (2026-08-27): builder HTTP callbacks now cross
> the disconnected worker boundary. The status watcher remains for
> compile/install and durable recovery, but is no longer the only live source.
> See "Carry the existing HTTP stream through add-in self-rebuild" above.

**Trigger**: Long Access operations looked idle in Cursor. VBA progress resets per category, so forwarding `current/total` as MCP `progress` violated the strictly-increasing rule. Full merge and database rebuild also dropped the MCP context (`ctx=None`). Add-in self-rebuild still required the agent to poll `rebuild-status.json` after `vcs_call_vba` returned. Cursor 3.13 often shows only "Running..." for progress notifications.

**Options explored**:
- *Forward VBA `current/total` as MCP progress.* Rejected. The spec requires each notification's `progress` to increase; category resets (28/30 queries, then 1/50 modules) are invalid and clients drop them.
- *MCP logging notifications (`ctx.log`).* Rejected. Deprecated 2026-07-28 / SEP-2577.
- *MCP Tasks.* Rejected. Extra protocol surface; does not fix Cursor display.
- *Have the add-in worker POST HTTP callbacks after Access quits.* Deferred. Would need a new worker contract; `rebuild-status.json` is already the durable source of truth.
- *New `vcs_rebuild_addin` that blocks on COM.* Rejected again. The host Access instance must quit so files can be replaced.
- *Agent-side Read polling of the status file.* Rejected as the normal workflow. Agents sat on timers after the work had finished.
- *Dedicated CLI that reimplements COM.* Rejected. A second operation path would drift from the MCP tools.

**Decision**: A shared `MonotonicProgressReporter` owns an incrementing sequence and keeps VBA counts in the message, with `total=None`. Full export, merge, and database rebuild all pass `ctx`. `vcs_rebuild_addin(source_dir)` launches via existing `RebuildAddIn` (gate held only for launch), then watches `rebuild-status.json` with `ReadDirectoryChangesW` (poll fallback), correlating `phaseStarted`. Cancel abandons the wait only — it does not kill `MSACCESS.EXE` or `wscript`. `vcs_call_vba` stays launch-only. `msaccess-vcs` is a stdio MCP client for the same tools, printing progress with immediate flush. Status-file polling remains recovery after a client timeout or stall.

> **⚠ Partially superseded** (2026-11-20): an abandoned rebuild wait now also
> asks the add-in to stop, through a tombstoned operation its poll still sees,
> and a canceled rebuild ends with status `cancelled`. Nothing is killed. See
> "Cancel by add-in poll; every long tool keeps the event loop free" above.

**What this rules out**: Treating Cursor chat progress as a keepalive or as guaranteed UI. Killing Access/wscript when a client stops waiting. Making generic `vcs_call_vba` wait minutes for install. Adopting deprecated MCP logging notifications. A second COM implementation beside the MCP server.

**Relevant files**: `src/msaccess_vcs_mcp/operation_manager.py`, `rebuild_watcher.py`, `cli.py`, `tools.py` (`vcs_rebuild_addin`), `access_gate.py`, `tests/test_operation_manager.py`, `tests/test_rebuild_watcher.py`, `tests/test_cli.py`.

---

## 2026-08-21 — Self-heal corrupted pywin32 gen_py cache

**Trigger**: MCP startup died with `module 'win32com.gen_py.4AFFC9A0-...' has no attribute 'CLSIDToClassMap'` when `%TEMP%\gen_py` held a half-built Access type-library folder (only `__pycache__`, no wrapper `.py` files). Deleting the folder manually and retrying `EnsureDispatch` regenerated the wrappers and succeeded. The failure had recurred several times.

**Options explored**:
- **Do nothing** — users must manually delete `%TEMP%\gen_py`. Simple, but fatal: `validate_access_installation()` exits and Cursor reports `MCP error -32000: Connection closed`.
- **Wipe all of `gen_py` on every startup** — always safe, but forces a full makepy rebuild for every COM library the process touches.
- **Fall back to late-bound `Dispatch`** — avoids the cache, but breaks early-binding semantics (`Application.Run` return shape).
- **Targeted folder delete + one retry** — parse the type-library folder from the `AttributeError`, `rmtree` only that folder, drop `sys.modules` entries, `gencache.Rebuild()`, retry `EnsureDispatch` once.

**Decision**: Centralize early-bound dispatch in `ensure_dispatch()` (`access_com/connection.py`) with targeted purge and a single retry. Route every `EnsureDispatch("Access.Application")` call site through it. Log `gen_py_cache_rebuilt` to the diagnostic stream.

**What this rules out**: Startup wipe-all of `gen_py`. Late-bind fallback for Access Application dispatch. Folding this into `com_recovery.py` (that module is for live RPC disconnects, not local makepy cache repair). More than one automatic retry per call.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`, `config.py`, `validation.py`, `tools.py`, `usage_logging.py` (`gen_py_cache` error pattern), `tests/test_gen_py_cache.py`.

---

## 2026-08-21 — The installed add-in is never a target, for any tool

**Trigger**: Agents repeatedly passed the installed add-in as `database_path` — to
run the add-in's own tests, and to host `RebuildAddIn`. That file exists to be
loaded as a library: opening it as a database, or writing into it, resets a VBA
project while it is executing. A test run makes the point sharply, needing two
roles at once from two different files (the install supplies the runner and
`TestAssert`; the current database holds the code under test), so collapsing them
makes the runner scan the library's own components and has
`InstallTestAssertModule` write into the executing project. The server's own
instructions had drifted into recommending the wrong host, naming "a user database
Access already has open" for rebuilds — which sent agents hunting for a database to
borrow, and into the add-in repo's `Testing` folder when they could not find one.

**Options explored**:
- *Rely on the add-in's VBA guard alone.* Rejected. `ExecuteTests` does refuse a
  run there, but only once the install is already the current database — and
  opening it as a database is the thing being prevented, not just the run that
  follows. It also covers only test runs.
- *Per-tool guards.* Tried and withdrawn within a day. Two were written, for
  `vcs_run_tests` and for `RebuildAddIn` via `vcs_call_vba`, and they still left
  `vcs_export_database`, `vcs_run_vba`, and the rebuild's own `output_path`
  uncovered. Tailoring each message to name the right alternative host was the
  argument for them, and it turned out there is only one answer to name.
- *Compare full file names.* Rejected. An install configured for the compiled
  add-in is a `.accde` built from the same `.accda`, and only one of the two
  appears in `ACCESS_VCS_ADDIN_PATH`, so the other variant would slip past.
- *Check folder parameters too* (`source_dir`, `output_dir`). Rejected. An export
  folder beside the install is not the install, and refusing it would block reading
  the install's own exported source.

**Decision**: `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper,
ahead of the gate and of any COM work, and rejects `database_path`, `output_path`,
or `template_path` naming the install with `error_pattern:
installed_addin_refused`. The comparison, `_is_installed_addin_path`, ignores the
extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`. The
message names the development copy in the add-in's repository and rules out the
substitutes agents actually reached for: a user database, the repo's `Testing`
folder, a scratch `.accdb`. `vcs_get_version_info()` is the sanctioned way to learn
the installed version, and needs no path.

**What this rules out**: Any tool reaching the installed file, including ones not
yet written — a new tool inherits the refusal from the decorator. Hosting a rebuild
or a test run on a borrowed database. Asking a person to nominate a host. A
tool-specific exemption would have to be argued as an exemption, since there is no
longer a per-tool guard to quietly omit.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`
(`_refuse_installed_addin_target`, `_is_installed_addin_path`,
`_installed_addin_target_refusal`, `vcs_tool`), `AGENTS.md`, `README.md`,
`docs/AGENT_WORKFLOWS.md`, `tests/test_installed_addin_guard.py`. The add-in side is
`modInstall.CurrentDbIsInstalledAddIn` and the guard in
`clsVersionControl.ExecuteTests`.

---

## 2026-08-21 — The install path comes from the add-in's own settings key

**Trigger**: `get_default_addin_path()` built `%AppData%\MSAccessVCS\Version
Control.accda` from constants. Both halves can be wrong: the installer lets the
user choose a folder, and a compiled install is a `.accde` — the installer deletes
the `.accda` in that case, so the guessed path names a file that does not exist.
Refusing the installed add-in as a target made this load-bearing rather than
merely untidy, since a wrong path means the refusal protects nothing.

**Options explored**:
- *Access Menu Add-Ins `Library` value.* Rejected. It does hold the full path with
  the correct extension in one read, but it lives under
  `Office\{Application.Version}\...`, and nothing records which Office version ran
  the installer — an external reader has to enumerate versions and guess. Reading
  two values from one fixed key is less machinery and no less certain.
- *Ask a running Access instance for `GetInstalledAddInFileName`.* Rejected.
  Requires COM and a loaded add-in to answer a question needed before either.
- *Probe the disk for whichever of the two extensions exists.* Rejected as the
  primary source: it infers from a side effect instead of reading the recorded
  choice, and reports nothing when neither file is present.

**Decision**: Read `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`
— `Install Folder` and `Compile accde` — which is exactly what the add-in's
`modInstall.GetInstalledAddInFileName` joins. `Install Folder` is absent for a
default install because the installer deletes the value rather than writing the
default, so absence means `%AppData%\MSAccessVCS`, not "not installed".
`Compile accde` holds a VBA integer, so True arrives as `-1`. Cached for the life
of the process, with `reset_addin_path_cache()` for tests and reinstalls.
`ACCESS_VCS_ADDIN_PATH` still overrides everything.

**What this rules out**: Reading the install path from anywhere else — the Menu
Add-Ins registration, the trusted-location entry (folder only), or a reconstructed
`%AppData%` path. Treating a missing `Install Folder` as a missing install.
Assuming the extension.

**Relevant files**: `src/msaccess_vcs_mcp/config.py`
(`get_default_addin_path`, `_read_install_setting`, `reset_addin_path_cache`),
`tests/test_config.py`. The add-in side is `modInstall.GetInstallSettings` /
`GetInstalledAddInFileName`.

---

## 2026-08-21 — The vcs_call_vba timeout worker owns its COM apartment

**Trigger**: Giving `vcs_call_vba` a hard timeout meant dispatching
`Application.Run` on a daemon thread so the parent could stop waiting. Handing that
thread the caller's `app` proxy failed before reaching Access at all —
"CoInitialize has not been called" while the thread had no apartment,
`RPC_E_WRONG_THREAD` (0x8001010E) once it did. Which of the two surfaced looked
random, so the failure read as an add-in problem. `RebuildAddIn` was unreachable
through this tool for as long as the timeout existed.

**Options explored**:
- *Marshal the caller's pointer into the worker* (`CoMarshalInterThreadInterfaceInStream`
  / `CoGetInterfaceAndReleaseStream`). Rejected. A marshalled STA pointer serializes
  the call back onto the apartment that created it, so the calling thread blocks
  anyway and the timeout this function exists to impose never fires.

> **⚠ Partially superseded** (2026-10-30): For out-of-process Access the
> unmarshalled proxy calls Access directly, and the timeout holds; marshalling
> is now the fallback when there is no database path to look up in the ROT.
> The ROT lookup remains the first choice. See "Unattended vcs_rebuild_database
> through BuildHeadless" above.
- *Run the call on the main thread and time out around it.* Rejected. There is
  nothing to time out around — a blocking COM call on the calling thread cannot be
  abandoned.
- *Re-acquire Access from the Running Object Table inside the worker* (chosen).
  Yields an apartment-local proxy, and `VCSAddinIntegration._find_access_in_rot`
  already existed for exactly this reason in the add-in probe.

**Decision**: `_run_application_call_with_timeout` takes `db_path` and the worker
calls `pythoncom.CoInitialize()`, then re-acquires the instance from the ROT. Without
`db_path` there is nothing to look up, so the caller's proxy is used as before — no
worse than it was, and the timeout may simply not fire.

**What this rules out**: Sharing a COM proxy across threads anywhere in this server.
Testing this path by stubbing `AccessConnection` alone — a test that does so silently
reaches the real ROT and passes or fails depending on whether Access happens to be
open, which is how `test_repo_addin_path_not_self_host_refused` became
environment-dependent. Stub `_find_access_in_rot` too; `_fake_access` in
`tests/test_call_vba_guards.py` wires both.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`
(`_run_application_call_with_timeout`), `src/msaccess_vcs_mcp/addin_integration.py`
(`_find_access_in_rot`), `tests/test_call_vba_guards.py`.

---

## 2026-08-12 — Access instances holding a database stay visible

> **⚠ Partially superseded** (2026-08-28): `Visible = True` is necessary but
> not sufficient for instances the server creates. Those also need
> `UserControl = True` *after* the database is open, or the window often never
> appears as a normal interactive app. See "Owned Access instances get
> UserControl after the database opens" above.

**Trigger**: COM automation starts Access hidden, and nothing in the server
ever showed the window. Access still asks questions only a person can answer —
trust prompts, conversion prompts, a VBA error breaking into the debugger — and
behind an invisible window those look like a hang. Every later call then blocks
on an instance the user cannot see, may not know exists, and has no way to
clear. The `.accda` work made this sharper by opening databases in instances
the server creates itself.

**Options explored**:
- *Leave instances hidden and detect the stall instead.* Rejected. The probe
  timeouts already report "Access is likely in break mode or blocked on a modal
  dialog", which is the right diagnosis and still leaves the user with no
  window to act on. Detection is not recovery.
- *Show only instances the server creates.* Rejected. An attached instance can
  raise the same dialog, and moniker binding launches a hidden Access when
  nothing had the file open, so "attached" does not imply "a person can see
  it".
- *Make visibility configurable.* Deferred. An opt-out re-creates the
  unresolvable-dialog failure for whoever sets it; wait for a concrete need
  (unattended CI is the plausible one).

**Decision**: Any instance the server drives a database through ends up
visible. Two helpers in `access_com/connection.py` carry the rule so it cannot
be half-applied: `ensure_access_visible(app)` (best-effort, never fails an
operation) and `open_current_database(app, path)`, which replaces every bare
`app.OpenCurrentDatabase(...)` call. Visibility comes *after* the open, because
showing a window can set `UserControl` and the add-in's `AutoRun` reads that
flag to decide whether a person is watching. `validate_access_installation()`
is the one exception: no database, quit immediately, so a window would only
flash.

**What this rules out**: Bare `OpenCurrentDatabase` calls — the AutoExec and
visibility rules travel with the helper, and a new call site that skips it
silently loses both. Hiding instances for speed or tidiness. Treating a probe
timeout as sufficient handling for a blocked dialog.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`
(`ensure_access_visible`, `open_current_database`), `validation.py`,
`vba_worker_manager.py`, `tools.py` (`vcs_rebuild_database`), `config.py`,
`tests/test_access_visibility.py`.

---

## 2026-08-12 — Opening .accda as the current database for add-in self-tests

> **⚠ Partially superseded** (2026-08-21): the `.accda` this opens is the
> *development copy* in the add-in's repository. The installed copy is now refused
> as a target by every tool, before anything opens it. See "The installed add-in is
> never a target, for any tool" above.

**Trigger**: The add-in's own tests only run when the add-in is the current
database, because the runner walks `CurrentVBProject`. Every `vcs_run_tests`
call against `Version Control.accda` failed with "Cannot find Access instance".
`GetObject(path)` binds a file moniker that resolves through the COM
registration for the extension; that registration opens .accdb and .mdb as the
current database but treats .accda as an add-in, so the bind fails. The
fallback instance then had no current database, and tier 2 of
`_find_access_in_rot` matches on `CurrentDb().Name`, so nothing was found.

**Options explored**:
- *Have a person pre-open the file* (the PowerShell `OpenCurrentDatabase`
  recipe in the add-in repo's `docs/agentic-rebuild.md`). Rejected: a human
  step per iteration, which is what this whole path exists to remove. That
  section of the add-in docs is stale as of this entry.
- *Register a file moniker for .accda.* Rejected: machine-wide COM
  registration change to fix one client.
- *Call `OpenCurrentDatabase` when the moniker bind fails* (chosen). Same
  approach `validation.py` already takes, and no new configuration.

**Decision**: `_open_as_current_database` runs only on the GetObject-failure
branch, and opens only on an instance we own, so a user's database is never
displaced. Two constraints surfaced in review. `UserControl` is lowered across
the open: `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's
`AutoRun` opens its installer form when `Application.UserControl` says a person
is watching — which would strand the instance we are about to automate, and did
so whenever `_create_isolated_instance` supplied the process. A failed open is
also swallowed rather than raised, so `_get_current_db` can still reach its DAO
strategies for read-only callers.

**What this rules out**: Documenting a manual pre-open step for add-in tests.
Setting `UserControl` before a database opens — `_create_isolated_instance`
still needs it for teardown survival, so it has to be restored after AutoExec,
not before. Letting an `OpenCurrentDatabase` failure escape `_get_access_app`,
which would bypass the DAO fallback chain that predates this path.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`
(`_open_as_current_database`, `open_current_database`),
`tests/test_accda_current_database.py`.

---

## 2026-08-12 — Agentic add-in rebuild via existing vcs_call_vba

> **⚠ Partially superseded** (2026-08-27): the normal agent path is now
> `vcs_rebuild_addin`, which watches `rebuild-status.json` after launch. The
> rejection of "a tool that waits on COM" still holds; the new tool waits on
> the status file with the Access gate released. `vcs_call_vba` remains the
> launch-only escape hatch. Agent-side Read polling is recovery, not the
> primary workflow. See "Real-time operation feedback" above.
>
> **⚠ Partially superseded** (2026-08-21): the host is the development copy of
> the add-in in its repository, not an arbitrary open database — see "The
> installed add-in is never a target, for any tool" above. `vcs_call_vba` also has a timeout now
> (`ACCESS_VCS_CALL_VBA_TIMEOUT_SEC`), which closed the follow-up noted below and
> brought its own COM constraint; see "The vcs_call_vba timeout worker owns its
> COM apartment" above.

**Trigger**: Agents iterating on the VCS add-in source could not rebuild
`Version Control.accda` without a person, because server instructions told them
the add-in cannot be rebuilt via MCP (closing every Access instance would close
the user's other databases).

**Options explored**:
- *New `vcs_rebuild_addin` tool that waits on COM.* Rejected. The Access instance
  the MCP is talking to is deliberately quit; a blocking COM wait would hang.
  `vcs_call_vba` already dispatches `RebuildAddIn` through `CallByName`.
- *Have the server poll the status file.* Rejected. The agent already has a Read
  tool, and the status path is known from the source folder. No new MCP surface.

**Decision**: Document the existing `vcs_call_vba` → `VCS.API` → `RebuildAddIn`
path. The add-in writes `<source>/logs/rebuild-status.json` and refuses when
another Access instance holds a file the rebuild replaces. A COM error after
launch is expected.
`vcs_call_vba` still has no timeout; the worker sleeps before quit so the JSON
can return. Adding a timeout remains a follow-up.

**What this rules out**: Treating `vcs_rebuild_database` as the add-in rebuild
path. A dedicated rebuild-add-in tool unless `vcs_call_vba` grows a timeout.
Closing other Access windows from the MCP server stays out; the add-in reports
the offenders instead, for reasons recorded in the add-in repo's own decision
log. A second Access instance held by this server only blocks the rebuild if the
add-in is loaded in it, which happens as soon as any `vcs_*` call routes through
the add-in's API.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (instructions, `vcs_call_vba`
example), `README.md`, `AGENTS.md`, `docs/AGENT_WORKFLOWS.md`.

---

## 2026-08-07 — Scoped object_types via ImportByType / ExportByType

**Trigger**: `vcs_import_objects` and `vcs_export_database` accepted `object_types` but largely ignored them. Import always ran a full `MergeBuild`. Export only special-cased a modules-only list into `ExportVBA` and otherwise exported everything. Agents passed `object_types=["modules"]` expecting a partial merge and got a whole-project one with no warning. The documented `overwrite` flag on import never mapped to any add-in behavior.

**Options explored**:
- *Keep ignoring object_types / document the lie*: rejected. Agents already rely on the parameter.
- *Add per-type MCP wrappers*: rejected. The add-in already exposes category-scoped APIs.
- *Route to ImportByType / ExportByType via call_sync* (chosen): no add-in rebuild; `modAPI.API` reaches any public `clsVersionControl` method through `CallByName`.

**Decision**: When `object_types` is set, call `ImportByType` / `ExportByType` synchronously and attach `log_path` like other sync results. When unset, keep the existing async full-project path (`MergeBuild` / `Export`/`FullExport`). Replace `overwrite` with `full_import`, which maps to `blnFullImport` (reload all files in the named categories vs only index-marked changes). Retire the modules-only `ExportVBA` shortcut so every scoped export uses one rule and reconciles deletions.

Scoped calls are sync-only because those methods are not in `APIAsync`'s command list. That is acceptable for category-sized work; revisit by adding them to the async list if blocking becomes painful. Single-type lists are passed as a bare string to avoid COM array-marshalling edge cases; multi-type lists remain Python lists.

**What this rules out**: Documenting or reintroducing an `overwrite` "skip existing" mode the add-in does not have. Treating a scoped merge as a safe partial without stating orphan deletion and no backup. Preferring `ExportVBA` for modules-only from these tools. Mapping "every category + full_import=True" as the recommended full rebuild path — use `vcs_rebuild_database` instead.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_import_objects`, `vcs_export_database`, `_scoped_types_arg`); add-in `clsVersionControl.ExportByType` / `ImportByType`, `modBuild.MergeScoped`.

---

## 2026-07-30 — Reaching the add-in API from MCP: vcs_call_vba, not vcs_run_vba

**Trigger**: An agent spent a long session trying to run the add-in's round-trip test harness (`VCS.RunRoundtripTests`) through `vcs_run_vba`. Every attempt returned an empty string. That looked like a broken add-in, then a stuck Access instance, and the workaround attempted next (`HandleRibbonCommand`) corrupted the host VBA project with error 2517 and required closing and reopening the database. The session ended by telling the user to paste a command into the Immediate window — for a capability the tools already had.

The cause is structural. `vcs_run_vba` is itself delivered through `modAPI.API`, which has a `Static IsRunning` re-entrancy guard. Submitted code therefore runs *inside* an API call, and anything it calls back into the API is nested by construction and refused. The guard returned `Empty` silently, which is indistinguishable from a method that legitimately returned nothing.

Three separate defects turned a one-line answer into a multi-hour dead end:

1. `vcs_call_vba` — the tool that does work — documented `"Version Control.API"` as its example. That never resolves. `Application.Run` matches a loaded VBA *project* name (`MSAccessVCS`), not the file name (`Version Control`). `"MSAccessVCS.API"` works but only once something has already loaded the add-in; only the full path is correct from a cold start.
2. `vcs_call_vba` did not unwrap the result. Early-bound `Run` returns a 31-element tuple — the return value followed by `Run`'s own 30 `Arg` slots, unused ones showing `DISP_E_PARAMNOTFOUND` (`-2147352572`). `VCSAddinIntegration.call_api_function` already handled this; the generic tool did not, so callers saw the payload buried in noise.
3. Nothing in the tool descriptions said which tool reaches the API, or that `vcs_run_vba` structurally cannot.

**Options explored**:
- *Document the full path and move on*: rejected. It puts an install-specific absolute path in every call and still leaves the silent-`Empty` trap for the next agent.
- *Relax the re-entrancy guard to allow nesting*: rejected. The guard protects `Operation` state owned by the outer call. The nested call is genuinely unserviceable; the defect is that it was refused silently, not that it was refused.
- *Resolve an alias qualifier server-side* (chosen): the server already knows `ACCESS_VCS_ADDIN_PATH`.

**Decision**: `vcs_call_vba` now accepts `"VCS.API"` (also `"Version Control.API"`, `"MSAccessVCS.API"`, case-insensitive) and rewrites the qualifier to the configured add-in's full path, which loads it on demand. Any other qualifier passes through untouched, so an explicit path or a user's own module still works. The result tuple is unwrapped to its first element, matching `call_api_function`. A "cannot find the procedure" failure now explains the project-name-versus-file-name rule. Server instructions name `vcs_call_vba` as the route to API methods and state that `vcs_run_vba` cannot be.

Paired with an add-in change: `modAPI.API` and `APIAsync` now return a message naming the refused method and pointing at `vcs_call_vba`, instead of returning `Empty`. `API` prefixes it with `API_REFUSED_PREFIX` (`"VCS_API_REFUSED: "`); `APIAsync` embeds it in its JSON. `vcs_call_vba` matches the prefix and reports `success: False` so a refusal is never mistaken for data.

That started as an `Err.Raise` and had to be changed after testing. An error raised inside a library database does not propagate across `Application.Run` into the calling project's handler — even with `On Error GoTo` active in the caller, Access shows a modal "Run-time error" dialog and blocks until a human dismisses it. Since this guard only trips on a nested call, which is by definition the case that crosses that boundary, raising was guaranteed to hit it. A blocking dialog is worse for automation than the silence it replaced, so the refusal travels as a marked return value instead.

**What this rules out**: Calling add-in API methods from inside `vcs_run_vba` — use `vcs_call_vba`. Adding tools that wrap individual API methods (`vcs_run_roundtrip_tests` and the like); the generic route now works and does not need per-method surface. Treating a `Run` result as a scalar anywhere else without unwrapping the tuple first. One accepted risk: a user module genuinely named `VCS` would have its qualifier rewritten — revisit if that ever surfaces.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_call_vba`, `_resolve_addin_function_name`, `_describe_run_failure`, server instructions); add-in `modules/API/modAPI.bas` (`RefuseReentrantCall`, `ERR_API_REENTRANT`).

---

## 2026-04-30 — EnsureDispatch ownership fix (DispatchEx fallback)

**Trigger**: Disabling the MCP tool in Cursor killed the user's Access window (with their open database). `AccessConnection._get_access_app()` called `EnsureDispatch("Access.Application")` when `GetObject(db_path)` failed, but `EnsureDispatch` can silently attach to an already-running user-owned Access instance instead of creating a new one. The code set `_owns_app = True` unconditionally, so `close()` called `_app.Quit()` on the user's session. This is the exact bug db-inspector-mcp fixed in their "DispatchEx fallback for COM instance conflicts" decision.

**Decision**: After `EnsureDispatch`, check `app.CurrentDb()` to detect whether the returned instance already has a database open. If it has *our* database, reuse and set `_owns_app = False`. If it has a *different* database, fall back to `DispatchEx("Access.Application")` which always spawns an isolated COM server process. If no database is open, it's a genuinely fresh instance and `_owns_app = True` is correct. `close()` only calls `Quit()` on instances the server actually created.

**What this rules out**: Setting `_owns_app = True` for any `EnsureDispatch` result without checking `CurrentDb()` first. Any future code path that creates an Access COM reference must verify ownership before storing it. If `DispatchEx` proves unreliable on specific Access/Windows configurations, revisit with the same guard pattern.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py` (`_get_access_app`, `_create_or_reuse_instance`, `_create_isolated_instance`).

---

## 2026-04-30 — Isolate `vcs_run_vba` behind a timeout-controlled worker

> **⚠ Partially superseded** (2026-04-30): Subprocess isolation was tried and reverted the same day. The child process had to cold-start Python, import the package, initialize COM, and re-acquire Access via the ROT — adding seconds of overhead per call. Worse, `subprocess.Popen.communicate(timeout=45)` blocked the MCP event loop for the entire duration, preventing the server from processing other requests (including `ListToolsRequest`), which caused `BrokenResourceError` crashes when Cursor timed out. Reverted to a daemon-thread + `thread.join(timeout)` approach modelled on this project's existing `_probe_with_timeout` and db-inspector-mcp's `_run_dao_with_timeout`. The thread creates its own COM apartment via `pythoncom.CoInitialize()` and re-acquires Access through the ROT, so `thread.join(timeout)` fires even when the COM call blocks. A class-level `_active_worker` guard prevents zombie thread pile-up (same pattern as db-inspector-mcp). The recovery state machine, error classification, logging events, and timeout env vars are unchanged. `vba_worker.py` (the subprocess entry point) has been deleted.

**Trigger**: A long-running `vcs_run_vba` call left Cursor's MCP connection closed, followed by reconnect attempts timing out and later calls failing with `Not connected`. The existing add-in probe had `ACCESS_VCS_PROBE_TIMEOUT_SEC`, but the actual `Application.Run(..., "RunVBA", code)` call still happened synchronously inside the stdio MCP process. If Access entered VBA break mode, showed a modal dialog, or never returned from the submitted snippet, the server process could hang or die before it could return a structured error.

**Options explored**:
- **Keep synchronous COM and rely on the existing add-in probe**. Rejected: the probe only proves `GetVCSVersion` responds before dispatch; it does not bound the later arbitrary `RunVBA` call where the observed failure occurred.
- **Launch a short-lived Python COM worker subprocess per call**. Tried and reverted: cold-start overhead (Python import + COM init + ROT lookup) was too high, and `subprocess.communicate(timeout)` blocked the async event loop causing `BrokenResourceError` crashes.
- **Daemon thread + `thread.join(timeout)` (chosen)**. Same pattern as the existing add-in probe and db-inspector-mcp's DAO timeout. Worker thread creates its own COM apartment, re-acquires Access via ROT, and runs VBA. Main thread returns `TimeoutError` if the deadline expires; the daemon thread finishes naturally. No true cancellation, but no event-loop blocking either.
- **Use the async callback path (`APIAsync` + `OperationManager`)**. Not applicable to ad-hoc VBA snippets which don't have a detached async add-in contract.
- **Kill `MSACCESS.EXE` automatically**. Rejected: Access may be user-owned with unsaved work.

**Decision**: `vcs_run_vba` runs arbitrary VBA in a daemon worker thread with a hard timeout (`ACCESS_VCS_RUN_VBA_TIMEOUT_SEC`, default 45s, overridable per call via `timeout_seconds`). A per-database `COMRecoveryManager` classifies COM transport failures and runs a short `ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC` probe before later calls. Pre-dispatch failures are retried once after a successful probe; failures during `run_vba` are not auto-retried because the snippet may already have started.

**What this rules out**: Running `RunVBA` directly on the main thread without a timeout boundary. Do not solve this class of failure by killing `MSACCESS.EXE` unless ownership tracking proves the server created an isolated Access instance. The daemon-thread approach accepts that a timed-out worker thread stays alive until Access responds (or the process exits); this is the same trade-off db-inspector-mcp makes for DAO queries.

**Relevant files**: `src/msaccess_vcs_mcp/vba_worker_manager.py` (thread-based timeout, probe, retry policy, `_active_worker` guard), `src/msaccess_vcs_mcp/com_recovery.py` (classification and per-database state), `src/msaccess_vcs_mcp/tools.py` (`vcs_run_vba` delegates to the worker manager), `src/msaccess_vcs_mcp/usage_logging.py` (worker/recovery events), `src/msaccess_vcs_mcp/main.py` (startup/shutdown/fatal diagnostics), `.env.example`, `AGENTS.md`, `tests/test_vba_worker_manager.py`, `tests/test_usage_logging.py`.

---

## 2026-04-27 — Always-on diagnostic stream + tiered usage-log defaults

**Trigger**: When the server is launched from a user-level Cursor MCP config (`~/.cursor/mcp.json`), the working directory becomes the user's home folder, the upward `.env` walk in `_find_project_root` resolves to `C:\Users\<user>` (because of the `.cursor` directory there), no `.env` is found, and `ACCESS_VCS_ENABLE_LOGGING` defaults to `false` -- so the server falls silent at exactly the moment we most need its self-diagnostics. The lazy MCP-roots `.env` discovery in `_ensure_env_loaded` was added to fix this, but its only debugging output went to stderr (which Cursor's user-level MCP pipes to the server-output pane that the agent cannot read). We had no way to verify that `list_roots()` was even being called, let alone what it returned. Compounding this: the existing single `ACCESS_VCS_ENABLE_LOGGING` switch made auditability all-or-nothing, which conflicts with this server's central capability -- it makes destructive changes to live databases and a forensic record is exactly what an operator needs by default, *but* code-execution bodies (SQL/VBA fragments) frequently embed business data, table names, or PII that the same operator may not want persisted to disk.

**Options explored**:
- **Single switch, leave default off**. Status quo. Rejected: provides no audit trail by default for a tool that mutates live databases, *and* still doesn't solve the observability gap that prompted this work.
- **Single switch, flip default on, log code bodies in full**. Maximum forensic fidelity. Rejected: lumping audit metadata and SQL/VBA bodies behind one switch forces privacy-conscious users to choose between auditability and confidentiality.
- **Hash code bodies (SHA-256) instead of redacting them**. Provides "did this exact code run before?" lookup without storing plaintext. Rejected: agents often generate semantically identical SQL with trivial whitespace/parameter differences, so hash equivalence is too brittle to be useful, and a hash still leaks more about *which* code ran than `code_length` alone.
- **Per-tool granular logging knobs** (`LOG_SQL_BODIES` separate from `LOG_VBA_BODIES`, etc.). Rejected: complexity exceeds value; one body switch is enough until we have evidence of actual divergent needs.
- **Mirror tool calls into the diagnostic file**. Rejected: bloats lifecycle debugging with high-frequency event noise. The two streams have genuinely different lifetimes (lifecycle: rare, small) and review patterns (audit: frequent, large), and conflating them defeats the whole point of a separate always-on stream.
- **Two-stream design with tiered default-on usage logging (chosen)**. (1) Always-on diagnostic stream at `~/.msaccess-vcs-mcp/logs/vcs-mcp-diagnostic.jsonl`, gated only on `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`, capturing server lifecycle events (`server_start`, `startup_env_load`, `lazy_env_load`, `lazy_init_started`, `lazy_init_skipped`, `list_roots_failed`, `list_roots_response`, `lazy_init_loaded`, `lazy_init_no_env_in_roots`, `usage_log_status`). (2) Usage stream defaults `ENABLE_LOGGING=true` so audit metadata is captured by default. (3) Code-execution bodies are redacted to `code_length` only; restore the old behavior with `ACCESS_VCS_LOG_CODE_CONTENT=true`. (4) Parameter keys matching `password|secret|token|api[_-]?key|connection[_-]?string` (case-insensitive) are auto-masked to `"<redacted>"` regardless of any switch -- defense in depth for accidentally-named-bad params.

**Naming sub-decision**: All log filenames carry the `vcs-mcp-` prefix (e.g. `vcs-mcp-usage.jsonl`, `vcs-mcp-diagnostic.jsonl`). Existing setups commonly point multiple MCP servers at a single shared `logs/` directory, where a generic name like `usage.jsonl` from one server collides with the same name from another. The prefix self-identifies the source without needing to inspect file contents.

**Discoverability sub-decision**: `vcs_get_version_info` returns `usage_log_path`, `diagnostic_log_path`, and `log_code_content` so an agent can `Read` either file directly and know whether code bodies are being captured -- without the agent having to inspect env vars itself.

**What this rules out**: Future entries should not add new "is logging on?" switches without first considering whether a *tier* of an existing stream covers the use case. Any new sensitive field that appears in logged data must either go through `_sanitize_parameters` (so the secret-key auto-mask catches it) or be added to the tiered model with its own explicit opt-in. The diagnostic stream is intentionally not a place to mirror tool-call events; future contributors who need richer audit data should extend the *usage* stream, not the diagnostic one. Reverting `ENABLE_LOGGING` to default-off would re-introduce the original "no audit trail by default" hazard and must be justified explicitly. The `code_length`-only default for code bodies is deliberate; do not "improve" it with a preview/first-N-chars representation -- length-only was chosen specifically because partial-content redactions leak just enough to enable per-keyword fishing while still failing to round-trip the actual code.

**Relevant files**: `src/msaccess_vcs_mcp/usage_logging.py` (new `_initialize_diagnostic_logging`, `log_diagnostic_event`, `get_diagnostic_log_path`, `is_diagnostic_logging_enabled`; `_write_log_entry` parameterized over handler; usage filename renamed to `vcs-mcp-usage.jsonl`; `log_code_execution` redaction; `_sanitize_parameters` secret-key masking; `_get_logging_config` adds `log_code_content` and flips `enabled` default), `src/msaccess_vcs_mcp/main.py` (`server_start` and `usage_log_status` events; diagnostic-log status print), `src/msaccess_vcs_mcp/config.py` (`startup_env_load` event in `_load_env_files`, `lazy_env_load` event in `_load_env_from_directory`; `ACCESS_VCS_ENABLE_LOGGING` default flipped, `ACCESS_VCS_LOG_CODE_CONTENT` added), `src/msaccess_vcs_mcp/tools.py` (`_ensure_env_loaded` rewired to emit diagnostic events on every branch; `vcs_get_version_info` exposes `usage_log_path`, `diagnostic_log_path`, `log_code_content`; FastMCP `instructions` text rewritten), `.env.example`, `AGENTS.md`, `tests/test_usage_logging.py` (new diagnostic-stream, body-redaction, and secret-key-mask tests), `tests/test_lazy_workspace_init.py` (extended to assert `lazy_init_*` events).

---

## 2026-04-25 — Honor the add-in lifecycle gate at every call site, with a hard probe timeout

**Trigger**: Commit 6342387 ("fix: surface add-in failures and tighten lifecycle checks") added a strict gate to `VCSAddinIntegration._call_addin_function` requiring both `self._app` AND `self._addin_loaded` to be truthy. The `_addin_loaded` flag is set only by `load_addin()`, which probes the add-in via `Application.Run("…\Version Control.API", "GetVCSVersion")`. However, every call site historically bypassed `load_addin()` and bare-assigned `addin._app = app` to skip a second COM round-trip. The gate refused every call, breaking all 13 user-facing tools plus the `_cleanup_session` atexit handler and `validate_components`. A separate concern surfaced during the fix: if Access is unresponsive (most commonly a developer left VBA paused in the VBE, but also modal dialogs and true hangs), even a "fast" probe will block the MCP server forever on the first tool call.

**Options explored**:
- **Loosen the gate to require only `_app`, restoring the pre-6342387 contract**. Smallest diff. But re-buries lifecycle errors -- a `vcs_run_vba` failure that's actually an add-in load problem still masquerades as a VBA bug, exactly the troubleshooting smell 6342387 was trying to fix.
- **Hybrid: keep the gate, auto-load inside `_call_addin_function` when `_addin_loaded` is False**. Lazy-load is convenient but conflates lifecycle with per-call dispatch and hides the probe cost in unpredictable places. Leaves no clean point to attach a hard timeout.
- **Honor the gate; switch every call site to `addin.load_addin(app, db_path=...)` (chosen)**. Pays one extra `Application.Run` per tool call (typically <10ms once Access has the add-in resident, ~50-200ms on the very first call while it loads). In exchange: lifecycle errors surface at the lifecycle boundary with a single, actionable message, and we have a natural place to attach the hang-protection timeout.

**Hang-protection sub-decision**: For the timeout on `GetVCSVersion`, evaluated `CoCancelCall` (requires server-side cooperation Jet/ACE doesn't implement), killing `MSACCESS.EXE` (loses unsaved work), subprocess isolation (overkill for a sub-second probe), and the disposable-worker-thread pattern that sibling project [`db-inspector-mcp`](C:/Repos/db-inspector-mcp/DECISIONS.md) already evaluated and shipped for DAO query timeouts. The worker-thread pattern is the only practical solution for in-process Access COM. Lifted with light edits: daemon thread + `thread.join(timeout)` + class-level `_active_probe_thread` guard against zombie pile-up + `pythoncom.CoInitialize()` per worker + ROT-based proxy re-acquisition (so the worker has its own apartment-local proxy instead of marshaling back to the main thread, which would defeat the timeout). Configurable via `ACCESS_VCS_PROBE_TIMEOUT_SEC` (default 10s -- short enough to feel responsive, long enough to absorb genuine first-load latency).

> **⚠ Partially superseded** (2026-10-30): Marshalling does not defeat the
> timeout for out-of-process Access, and a probe with no `db_path` now
> marshals the caller's proxy rather than sharing it. See "Unattended
> vcs_rebuild_database through BuildHeadless" above.

**Pre-flight sub-decision**: When `db_path` is supplied to `load_addin()`, run an `os.path.isfile(db_path)` check before any COM activity. Catches stale or typo paths in ~1ms instead of burning the full 10s timeout in ROT lookup for a database that simply isn't there. Defense in depth: today only `tools.py` call sites validate paths via `validate_database_path()`; `validation.py`, `_cleanup_session`, and the internal `get_version_info` did not.

**Instrumentation sub-decision**: Added `log_addin_probe(addin_path, duration_ms, success, timed_out, error)` to `usage_logging.py`, emitting one `"addin_probe"` JSONL event per probe in the same `usage.jsonl` stream as `tool_call` and `code_execution`. Lets us empirically verify the assumed "fast and cheap" cost (`rg '"event":"addin_probe"' logs/usage.jsonl`) and revisit the trade-off if subsequent probes turn out to be expensive in practice. Distinguishes timeouts from generic COM errors (`timed_out` boolean) so analytics can isolate true hangs.

**What this rules out**: Reverting to bare `addin._app = app` at any call site is now an anti-pattern -- it would re-break the gate and lose hang protection. A future contributor who sees the extra `Application.Run` per call should not "optimize" it away without first reading this entry. The probe timeout assumes a well-behaved Access instance can answer `GetVCSVersion` in <10s; if a legitimate use case exists where it cannot, raise `ACCESS_VCS_PROBE_TIMEOUT_SEC` rather than removing the timeout. The class-level `_active_probe_thread` guard is intentionally not cleared on timeout -- a lingering probe thread blocks all subsequent probes with a clear "previous probe still pending" error until Access responds (the daemon thread terminates) or the process exits. This is the desired behavior in production; tests use an autouse fixture to reset it. ROT-based re-acquisition only fires when `db_path` is provided; the `vcs_rebuild_database` and `get_version_info` paths pass `None` and fall back to sharing the main thread's proxy (best-effort, may not respect the timeout as reliably, but acceptable for paths that don't have a database in scope).

**Relevant files**: `src/msaccess_vcs_mcp/addin_integration.py` (new `load_addin(app, db_path=None)` signature with idempotent early return, pre-flight, instrumented probe; new `_probe_with_timeout` and `_find_access_in_rot` helpers; class-level `_active_probe_thread`; `get_version_info` now routes through `load_addin`), `src/msaccess_vcs_mcp/tools.py` (13 call sites updated; redundant inline `GetVCSVersion` pre-flights in `vcs_export_database` and `vcs_import_objects` removed in favor of `load_addin()` with the same friendly error wrapping), `src/msaccess_vcs_mcp/validation.py` (`validate_components` line 255), `src/msaccess_vcs_mcp/main.py` (`_cleanup_session` line 95), `src/msaccess_vcs_mcp/usage_logging.py` (`log_addin_probe`), `.env.example` (`ACCESS_VCS_PROBE_TIMEOUT_SEC` documented), `tests/test_addin_integration.py` (5 new tests: timeout, active-worker guard, db_path pre-flight, idempotency, invalid timeout fallback), `tests/test_usage_logging.py` (4 new tests for `log_addin_probe`).

---

## 2026-04-25 — Drop silent `Application.Eval` fallback in `_call_addin_function`

**Trigger**: Investigating failing add-in integration tests surfaced a long-standing usability and troubleshooting problem: when `Application.Run("…\Version Control.API", funcName)` failed twice in a row, `_call_addin_function` would silently fall through to `self._app.Eval("CallVcsApi(\"" & funcName & "\")")`, using a hard-coded default wrapper name. In any environment where the user had not implemented a `CallVcsApi` VBA function in the open database, this branch would still succeed against unit-test mocks and against any COM object whose `Eval` happened to return a value, producing false-success results for `vcs_export_*` and friends. The original `Run` error was buried in a three-layer retry/fallback message that obscured root cause during troubleshooting.

**Options explored**:
- **Keep current behavior, update tests to mock `Eval` raising**. Lowest-friction, but bakes the silent-success failure mode into production indefinitely.
- **Make the fallback opt-in via `ACCESS_VCS_API_WRAPPER`**. Removes the silent-success default but adds a configuration knob plus a code path almost no one will exercise; documenting *when* to set it requires explaining a real Access COM bug that most users will never hit.
- **Remove the fallback entirely (chosen)**. Single behavioral contract: if `Application.Run` fails twice, a `RuntimeError` is raised with the underlying COM error verbatim. Simpler call graph, fewer places to look during incident triage.

**Decision**: Removed the `Application.Eval` fallback and the `ACCESS_VCS_API_WRAPPER` env-var hook. The single `Run` retry is preserved -- it covers the documented Access first-call add-in load behavior -- but a second failure now surfaces directly as `RuntimeError("Failed to call add-in function '<name>': <error>")`. The decision was driven by simplification of code and of troubleshooting, not by performance or correctness in any narrow sense.

**What this rules out**: The MCP server will no longer auto-recover from genuine `Application.Run`-from-COM bugs by routing through a per-database VBA wrapper. Users who actually need that workaround must reintroduce it deliberately -- ideally as an explicit opt-in setting with a clear error message when it isn't configured -- not as a hidden default. Re-adding any silent fallback that swallows a documented error path should be rejected on review.

**Relevant files**: `src/msaccess_vcs_mcp/addin_integration.py` (`_call_addin_function`).

---

## 2026-04-25 — Multi-strategy `.env` discovery with workspace-roots lazy init

**Trigger**: A user reported that `ACCESS_VCS_ENABLE_LOGGING=true` in a client project's `.env` had no effect when `msaccess-vcs-mcp` was used from another project. Root cause: `_find_project_root()` only walked up from CWD or the installed package location. When the server was launched from a user-level `mcp.json` (`~/.cursor/mcp.json`) and CWD wasn't the project root, the upward walk failed, and the package-location fallback could match `msaccess-vcs-mcp`'s own `pyproject.toml` instead of the user's project. Compounding this, `_load_env_files()` always called `load_dotenv(..., override=False)`, so even a successful reload would silently fail to apply edited values. The sibling `db-inspector-mcp` project had already solved this with a layered resolution strategy.

**Options explored**:
- **Single env-var override (`ACCESS_VCS_PROJECT_DIR`) only**. Simple, works, but requires per-project user configuration. Doesn't help users who configure the server once at the user level and expect it to "just work" across projects.
- **MCP workspace-roots discovery only** (`ctx.session.list_roots()`). Automatic, no user config required. But only works for tools that accept a `Context` parameter and run async — most existing tools are sync without `ctx`.
- **Layered resolution: env-var → workspace roots → CWD walk → package walk → fallback (chosen)**. Mirrors `db-inspector-mcp`'s proven order. Each strategy compensates for the others' blind spots: explicit env-var for power users, workspace roots for IDE-launched servers, CWD walk for terminal-launched servers, package walk as a last-resort dev-install fallback.
- **Eager workspace-roots probe at startup**. Cleaner than lazy init, but FastMCP's `Context` is not available before the first tool call -- the MCP protocol handshake hasn't completed yet.

**Decision**: Ported `db-inspector-mcp`'s discovery machinery in full, with `ACCESS_VCS_` prefix substitution and three project-specific adaptations:

1. **Resolution-method tracking**. `_find_project_root()` and `initialize_from_workspace()` set a module-level `_project_root_method` (one of `RESOLUTION_PROJECT_DIR_ENV`, `RESOLUTION_WORKSPACE_ROOTS`, `RESOLUTION_CWD_ENV`, `RESOLUTION_CWD_MARKER`, `RESOLUTION_PACKAGE_ENV`, `RESOLUTION_PACKAGE_MARKER`, `RESOLUTION_CWD_FALLBACK`). Surfaced via `get_project_root_info()`, printed to stderr ("Resolved project root: X (via Y)"), and embedded in the `logging_initialized` JSONL event so users can audit *which* mechanism actually populated their config in any given session.
2. **mtime-based hot-reload with `override=True` on reload**. `_check_env_reload()` compares stored `.env`/`.env.local` mtimes against current values; when changed, the next `load_config()` triggers a reload using `override=True` so edited values actually replace old ones. Logging is reset only when a reload was detected (was previously reset on every `load_config()` call -- wasteful).
3. **Lazy init wired through the `vcs_tool` decorator, not individual tools**. The decorator inspects the wrapped handler's signature; for async handlers that accept `ctx: Context`, it calls `await _ensure_env_loaded(ctx)` before `load_config()`. Converted `vcs_get_version_info`, `vcs_list_objects`, `vcs_diff_database`, `vcs_import_objects`, and `vcs_rebuild_database` to async with optional `ctx` so the workspace-roots path triggers on the agent's first realistic call regardless of which tool that is.

**What this rules out**: Sync tools without `ctx` cannot trigger workspace-roots lazy init -- but this is fine because once any async-with-ctx tool runs, the project env is loaded for all subsequent calls (sync included). If a future agent goes straight to a sync tool first (e.g., `vcs_execute_sql` before any read tool), workspace-roots discovery won't fire and the user must rely on `ACCESS_VCS_PROJECT_DIR` or CWD-based discovery; converting more tools to async is the escape hatch. The `RESOLUTION_*` constants are part of an implicit public contract -- renaming them would silently break any log-analysis tooling that filters on `project_root_resolution`. The package-walk fallback can still match `msaccess-vcs-mcp`'s own dev tree when the CWD walk finds nothing; this is intentional for development installs and is the lowest-priority strategy.

**Relevant files**: `src/msaccess_vcs_mcp/config.py` (resolution-method tracking, `_check_env_reload`, `_load_env_from_directory`, `initialize_from_workspace`, `get_project_root_info`), `src/msaccess_vcs_mcp/tools.py` (`_lazy_init_attempted`, `_file_uri_to_path`, `_ensure_env_loaded`, `vcs_tool` decorator, 5 tool signatures), `src/msaccess_vcs_mcp/usage_logging.py` (`logging_initialized` event includes `project_root` + `project_root_resolution`), `.env.example`, `README.md` (new "Using msaccess-vcs-mcp from Another Project" section), `tests/test_config_env_loading.py` (17 tests), `tests/test_lazy_workspace_init.py` (9 tests).

---

## 2026-04-15 — Pre-execution audit logging for code execution tools

**Trigger**: The existing usage logging captures tool parameters but truncates all strings to 500 characters and only writes after execution completes. For the three tools that execute arbitrary code against databases (`vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), this leaves two gaps: (1) a complex SQL query or VBA code block may be truncated beyond usefulness in a forensic review, and (2) if the process crashes during execution, no record of what was attempted exists.

**Options explored**:
- **Raise the truncation limit globally**. Simple, but inflates every log entry (file paths, option names, etc.) unnecessarily. Rotation would trigger sooner.
- **Exempt specific parameter names from truncation** (e.g., `sql`, `code`). Mixes audit concerns into the general sanitization logic. Hard to extend cleanly.
- **Dedicated `log_code_execution()` function with a separate event type (chosen)**. Writes a `"code_execution"` event *before* execution begins, with the full untruncated code/SQL and the target database path. The existing `with_logging` decorator continues to write the post-execution `"tool_call"` event with truncated parameters, success/error, and timing. Two complementary records: the audit trail (what was attempted) and the outcome (what happened).

**Decision**: Added `log_code_execution(tool_name, database_path, code, code_type)` to `usage_logging.py`. Called from `vcs_execute_sql` (code_type=`"sql"`), `vcs_run_vba` (code_type=`"vba"`), and `vcs_call_vba` (code_type=`"vba_call"`) immediately after path validation but before any COM/database interaction. The `code` field is never truncated. The event goes to the same `usage.jsonl` file — no separate audit file — distinguished by `"event": "code_execution"`.

**What this rules out**: Code execution entries have no upper size limit on the `code` field. In practice, VBA code blocks and SQL queries are small (under 10 KB). If an agent somehow generates megabyte-scale code strings, log rotation handles it, but this is not a realistic concern. If a separate audit file is ever wanted (e.g., for compliance), the `log_code_execution` function could be retargeted without changing call sites.

**Relevant files**: `usage_logging.py` (`log_code_execution`), `tools.py` (call sites in `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), `tests/test_usage_logging.py` (5 new tests).

---

## 2026-04-15 — Agents cannot enable McpAllowRunVBA programmatically

> **⚠ Superseded** (2026-11-24): `vcs_call_vba` could reach the add-in's
> `SetOption` and bypass this server-only list. The add-in now refuses all four
> `McpAllow*` options itself, without needing to know its caller. See "No tool
> can change the McpAllow* permissions" above.

**Trigger**: The `vcs_set_option` tool allowed agents to set any VCS option, including `McpAllowRunVBA` which gates arbitrary VBA code execution via `vcs_run_vba`. An agent could autonomously enable this option and then run arbitrary code without user awareness, undermining the security boundary that `McpAllowRunVBA` was designed to provide.

**Options explored**:
- **No guard, rely on default-off**. `McpAllowRunVBA` defaults to False, but nothing prevented an agent from calling `vcs_set_option("db.accdb", "McpAllowRunVBA", True)` as its first action. The docstrings even showed this as an example.
- **Server-side blocklist in `vcs_set_option` (chosen)**. A case-insensitive check against a set of protected option names. Returns a descriptive error directing the user to enable the option manually via the VCS Options form.
- **VBA-side enforcement**. Have the add-in's `SetOption` method refuse `McpAllowRunVBA` when called from MCP. Harder to implement since VBA doesn't know the calling context, and the error would be less clear.

**Decision**: `vcs_set_option` blocks setting `McpAllowRunVBA` with a clear error message. The option requires explicit user consent via the VCS Options form in Access. Docstrings and README updated to stop suggesting agents can self-enable this option.

**What this rules out**: Agents cannot autonomously escalate to arbitrary VBA execution. If future protected options emerge (e.g., a hypothetical `McpAllowDDL`), add them to the `PROTECTED_OPTIONS` set in `vcs_set_option`.

**Relevant files**: `tools.py` (`vcs_set_option`), `README.md` (security section).

---

## 2026-04-15 — Object type normalization lives in VBA, not Python

**Trigger**: `vcs_export_object` and `vcs_import_object` only supported 6 core Access object types (query, form, report, module, table, macro). The add-in's `eDatabaseComponentType` enum defines 24+ types (relations, IMEX specs, VBE project, themes, etc.) that couldn't be exported individually. Additionally, the MCP tools used plural strings (`"queries"`) in `vcs_export_database` but singular (`"query"`) in `vcs_export_object`, creating inconsistency that confused AI agents.

**Options explored**:
- **Python-side normalization map**. A `normalize_object_type()` helper in the MCP server that maps plural/alias forms to canonical singular before passing to VBA. Only benefits MCP callers. Creates a second type map to maintain alongside VBA.
- **VBA-side normalization via `ResolveComponentType` (chosen)**. A `Select Case` function in `modContainers.bas` that accepts singular, plural, and alias forms (50+ strings) and maps to `eDatabaseComponentType`. Benefits all callers — MCP tools, direct `Application.Run` API calls, any future integration. Python becomes a transparent pass-through.
- **Accept both in Python AND VBA**. Redundant and creates maintenance burden keeping two maps in sync.

**Decision**: Type normalization lives entirely in VBA's `ResolveComponentType`. Python passes `object_type` strings through to VBA without validation. VBA returns structured error JSON for unrecognized types. `ExportObject` and `ImportObject` on `clsVersionControl` were extended to handle all 24 component types: core AccessObject types use the existing `ExportSingleObject` path; non-core types use `GetComponentClass` + `GetAllFromDB`. Single-file types (like `vbe_project`) don't require an `object_name` parameter.

**What this rules out**: Adding new component types requires updating VBA's `ResolveComponentType` — the Python MCP layer does not need changes. If a Python-only consumer needs early validation without a COM roundtrip, they would need to maintain their own type list, but this is unlikely since the VBA error response is fast and descriptive.

**Relevant files**: `modContainers.bas` (`ResolveComponentType`), `clsVersionControl.cls` (`ExportObject`, `ImportObject` rewritten), `tools.py` (docstrings updated, `object_name` made optional), `addin_integration.py` (no type-related changes).

---

## 2026-04-15 — Structured JSONL usage logging via composite decorator

> **⚠ Partially superseded** (2026-04-25): The "What this rules out" note about adopting `db-inspector-mcp`'s mtime-based hot-reload pattern "if performance becomes a concern" is now fact -- adopted for correctness rather than performance (the original `override=False` reload was silently failing to apply edits). The `logging_initialized` event also now includes `project_root` and `project_root_resolution` fields. See "Multi-strategy `.env` discovery with workspace-roots lazy init" above.

**Trigger**: Need to troubleshoot and evaluate how AI agents use the MCP tools in practice — which tools are called, with what parameters, how often they fail, and how long they take. The `db-inspector-mcp` sibling project already has a proven logging implementation that was requested as the reference pattern.

**Options explored**:
- **Python stdlib `logging` module**. Standard approach, but produces unstructured text. Not suitable for programmatic analysis of tool call patterns.
- **FastMCP middleware / hooks**. FastMCP doesn't expose a per-tool middleware layer. Would require monkey-patching internals.
- **Composite decorator with JSONL file logging (chosen)**. Follows the exact pattern from `db-inspector-mcp`: a `with_logging(name)` decorator that wraps each tool, writing one JSON object per line to a rotating file. A `vcs_tool("name")` composite decorator chains config reload → usage logging → `mcp.tool()` registration, replacing bare `@mcp.tool()` on all 17 tools. Controlled by `ACCESS_VCS_ENABLE_LOGGING` env var (default: off).

**Decision**: Adopted the `db-inspector-mcp` pattern with project-specific adaptations:
- Env var prefix changed from `DB_MCP_` to `ACCESS_VCS_` for consistency with existing config.
- Removed `database`/`dialect` fields from log entries (VCS tools pass `database_path` as a regular parameter, so it's captured in `parameters` automatically).
- Error pattern categories tailored to VCS-specific errors (COM errors, add-in errors, VBA compile errors, write-disabled, database busy) instead of SQL-specific patterns.
- `Context` objects from FastMCP are filtered out of logged parameters (not serializable).
- Lazy initialization: disabled state is not cached (`_logging_enabled` stays `None`) so the first tool call after `load_config()` populates the env can still enable logging. Failure state _is_ cached to avoid retry spam.

**What this rules out**: Log entries do not include tool return values — only parameters, success/failure, errors, and timing. If result logging is needed later, the `with_logging` decorator already receives the result for serialization checking and could be extended. The `vcs_tool` decorator calls `load_config()` on every tool invocation (one `stat()` call); if this becomes a performance concern, the hot-reload pattern from `db-inspector-mcp` (mtime-based gating) could be adopted.

**Relevant files**: `src/msaccess_vcs_mcp/usage_logging.py` (new), `src/msaccess_vcs_mcp/tools.py` (`vcs_tool` decorator, all 17 tools migrated), `src/msaccess_vcs_mcp/config.py` (logging env vars + `reset_logging` on reload), `src/msaccess_vcs_mcp/main.py` (startup status), `.env.example`, `tests/test_usage_logging.py` (36 tests), `AGENTS.md` (new), `.gitignore` (`logs/`).

---

## 2026-04-14 — Tool naming: `vcs_*` prefix

**Trigger**: Tools were originally named `access_*` (e.g., `access_export_database`). A separate MCP server for Access (`MCP-Access` by bclothier) also uses `access_*` prefixed tools. Both servers might be loaded in the same agent session. Additionally, these tools control the VCS add-in, not the Access application itself, so `access_*` was a misnomer.

**Options explored**:
- **Keep `access_*`**. Familiar, but inaccurate and collides with MCP-Access.
- **Use `vcs_*` (chosen)**. Accurately reflects these tools control the VCS add-in. No namespace collision.
- **Use `msaccess_vcs_*`**. Unambiguous but verbose — wastes tokens on every tool call.

**Decision**: All 9 existing tools renamed from `access_*` to `vcs_*`. All 8 new tools use `vcs_*`. Updated across `tools.py`, `README.md`, `validation.py`, and all docs.

**What this rules out**: Any external references to `access_export_database` etc. break. Acceptable since the server is pre-release with no external consumers.

**Relevant files**: `tools.py`, `README.md`, `validation.py`, `docs/*.md`.

---

## 2026-04-14 — Eight new tools for per-object development workflow

> **⚠ Partially superseded** (2026-04-15): `vcs_export_object` and `vcs_import_object` now support all 24 component types (not just the original 6 core types). Type normalization moved to VBA. See "Object type normalization lives in VBA, not Python" above.

**Trigger**: All existing tools operated at the whole-database level. Agents had no way to export/import a single object, execute SQL, run VBA, or control add-in options — capabilities essential for the tight edit-import-compile-test loop needed during add-in development and general database development.

**Options explored**:
- **Extend existing tools with filters** (e.g., `vcs_export_database` with `object_name` parameter). Conflates bulk and single-object semantics. The existing tools have async/callback infrastructure not needed for quick per-object calls.
- **Build intelligence into the MCP server** (Python-side VBE manipulation, SQL execution via separate ODBC connection). Creates tight coupling to Access internals in Python, duplicates logic better handled in VBA, and opens a second database connection causing file-locking conflicts.
- **Thin MCP tools that delegate to add-in API methods (chosen)**. Each new tool calls a corresponding method on `clsVersionControl` via the existing `API()` dispatcher. The add-in handles all Access interaction. The MCP layer just validates paths, parses JSON results, and formats responses.

**Decision**: 8 new tools added: `vcs_export_object`, `vcs_import_object`, `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`, `vcs_set_option`, `vcs_get_option`, `vcs_get_log`. Total: 17 tools. Each delegates to a public method on `clsVersionControl` in the add-in. The MCP server remains a lightweight wrapper — all business logic lives in VBA.

**What this rules out**: The MCP server does not do database introspection, schema analysis, or complex SQL. Those capabilities stay in `db-inspector-mcp`. If the VCS MCP needs to support operations not expressible through `clsVersionControl` API methods, the add-in must be extended first. This is intentional — it keeps the MCP layer thin and ensures all consumers of the add-in API get the same capabilities.

**Relevant files**: `tools.py` (8 new tool definitions), `addin_integration.py` (`call_sync` used by all new tools).

---

## 2026-04-14 — SQL execution via add-in DAO connection, not separate ODBC

**Trigger**: Usage logs from `db-inspector-mcp` showed 67% of all calls were just running SELECT queries (`db_count_query_results`, `db_preview`). Agents frequently need to inspect `MSysObjects`, `MSysQueries`, table data, and query results. Requiring a second MCP server for this basic need adds configuration overhead and creates a second connection to the same Access file (risking file-locking conflicts).

**Options explored**:
- **Keep SQL execution in db-inspector-mcp only**. Clean separation, but requires agents to have both MCPs configured. Two connections to the same `.accdb` file can cause locking issues. Extra overhead for the dominant use case.
- **Add ODBC connection in the VCS MCP server** (Python-side). Avoids VBA roundtrip but opens a second connection. Would need to handle Access SQL dialect quirks in Python.
- **Route through add-in's existing DAO connection (chosen)**. `ExecuteSQL` method on `clsVersionControl` uses `CurrentDb.OpenRecordset` — the same connection the add-in already holds. No file-locking conflict. Access SQL dialect handled natively. Read-only (SELECT only, enforced in VBA).

**Decision**: `vcs_execute_sql` tool calls the add-in's `ExecuteSQL` API method, which runs the query via `CurrentDb.OpenRecordset`, serializes results as JSON, and returns them. Non-SELECT statements are rejected. Results capped at `max_rows` (default 100). The db-inspector MCP remains available for cross-database comparison and heavy analytical work.

**What this rules out**: No write queries (INSERT/UPDATE/DELETE/DDL) through this tool. If agents need to modify data, they use `vcs_run_vba` or `vcs_call_vba` with appropriate VBA code. The SQL validation is simple (checks for `SELECT` prefix) — a determined agent could bypass it via `vcs_run_vba`, which is why `McpAllowRunVBA` defaults to off.

**Relevant files**: `tools.py` (`vcs_execute_sql`), `clsVersionControl.cls` (`ExecuteSQL`).

---

## 2026-04-14 — Two VBA execution tools with distinct roles

**Trigger**: Agents need to execute VBA code for testing and debugging. Two distinct use cases emerged: calling existing functions by name (safe, predictable) and executing arbitrary agent-generated code (powerful, risky). These have fundamentally different security profiles.

**Options explored**:
- **Single `vcs_run_vba` tool for both**. Agent passes either a function name or a code block, tool detects which. Blurs the security boundary — how do you gate "arbitrary code" while allowing "call existing function"?
- **Two tools with distinct roles (chosen)**. `vcs_call_vba` calls existing named functions via `Application.Run` — no temp module, no compilation, lower risk, separate permission (`McpAllowCallVBA`, default: True). `vcs_run_vba` executes agent-generated code via temp module lifecycle — compilation check, error capture, cleanup, higher risk, separate permission (`McpAllowRunVBA`, default: False).

**Decision**: `vcs_call_vba(database, function_name, args)` for existing functions; `vcs_run_vba(database, code)` for ad-hoc code. The add-in handles `RunVBA`'s full lifecycle (create, compile, execute, capture, cleanup). Error capture in `RunVBA` uses module-level variables with accessor functions rather than embedded JSON string construction in generated code — cleaner and avoids VBA quote-escaping nightmares.

**What this rules out**: `vcs_call_vba` is limited to public functions callable via `Application.Run` (max 3 args in current implementation). Private functions or functions requiring object parameters can't be called directly — use `vcs_run_vba` for those. If the 3-arg limit becomes a problem, the `InvokeTypes` + `pythoncom.Missing` padding pattern (used by MCP-Access) would support up to 30 args.

**Relevant files**: `tools.py` (`vcs_call_vba`, `vcs_run_vba`), `clsVersionControl.cls` (`RunVBA`).

---

## 2026-04-15 — Session-scoped option overrides for MCP/API callers

**Trigger**: `vcs_set_option` changes were silently discarded because the add-in reloads options from `vcs-options.json` at the start of every operation. The agent's overrides never persisted past the first subsequent export/build.

**Decision**: The MCP server generates a session ID at startup (`uuid4().hex[:8]`), registers it with the add-in via `RegisterSession`, and `vcs_set_option` now writes overrides to a session-scoped file (`mcp/options-{session_id}.json`) in the export folder. The add-in's operation entry points load these overrides after `LoadProjectOptions` when `Operation.Source` is API/MCP. On shutdown, `atexit` calls `EndSession` to clean up the file. A `vcs_end_session` tool is also available for explicit mid-session cleanup.

**What this rules out**: Session IDs don't persist across server restarts — the agent must re-set options if the server restarts. Stale override files are auto-cleaned after 30 days on the add-in side.

**Relevant files**: `tools.py` (`vcs_set_option`, `vcs_end_session`), `main.py` (session ID, atexit), `config.py` (`get_session_id`). Add-in side: see `DECISIONS.md` in `msaccess-vcs-addin`.
