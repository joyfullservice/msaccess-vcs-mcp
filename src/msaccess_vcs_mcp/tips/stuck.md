# Access is stuck, a call timed out, or the server is busy

## First, look

    vcs_capture_screenshot(database_path, include_image=False)

This works while Access is blocked and while another call holds the server. It returns, without image tokens:

- `dialogs`: the title, text, buttons, and `dialog_id` of every dialog Access is showing (message boxes, compile errors, prompts).
- `vbe_state`: `break` when VBA stopped at a line, `running`, or `design`.
- `owned_by_server`: whether this server created that Access window.

Call it again with the image when the layout matters. `target="vbe"` shows the VBA editor with the module and highlighted line.

## Then act

- **The server created the window** (`owned_by_server: true`): answer a dialog with `vcs_dismiss_dialog(dialog_id, button)`, quoting the `dialog_id` you just read; if the dialog changed, the call is refused with the current list. End break mode with `vcs_reset_vba()`. Both work while another call is stuck behind the dialog.
- **The user opened the window**: ask the user to answer the dialog or press Reset. Do not click it.
- Never stop, kill, or quit `MSACCESS.EXE`. Another agent or the user may own that window, and a process list cannot tell you whose it is. If a window holds a file you need, name its PID and ask the user.

## Result patterns

- `error_pattern: server_busy`: another call (perhaps from another Cursor window or agent) holds this database, or an exclusive operation such as `vcs_rebuild_addin` holds every database; `busy_with` names it. Retry shortly. Calls on one database run one at a time, so do not issue them in parallel.
- `error_pattern: timeout`: the server stopped waiting. It does not close Access. Look first (above); the next call probes Access and resumes once it responds. `access_state`, when present, already carries what the screenshot would show.
- `error_pattern: access_unresponsive`: Access is not answering automation, usually because of a dialog or break mode. Same steps.
- Client error `-32001`: the client gave up and no result reached you. Call `vcs_get_recent_calls()` to see what actually ran, and whether it finished, before retrying.

## Stopping a long operation

`vcs_cancel_operation()` stops a running export, import, rebuild, or test run; with no arguments it targets the only one running, and when several are running it lists them so you can pass the `operation_id` or `database_path` of yours. A canceled run returns `cancelled: true` and `error_pattern: operation_cancelled`.
