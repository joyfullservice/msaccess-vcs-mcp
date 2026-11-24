# Running VBA

Two tools run VBA, each behind its own permission. The user sets both in the add-in's MCP options form; no tool can change them.

- `vcs_call_vba(database_path, function_name, args)` calls an existing public function through `Application.Run` (`McpAllowCallVBA`, on by default). Up to 3 string arguments. Use `"VCS.API"` as the function name to reach the add-in's API.
- `vcs_run_vba(database_path, code)` runs code you write in a temporary module (`McpAllowRunVBA`, off by default).

There are no other code tools; names such as `vcs_eval` or `vcs_execute_code` do not exist.

## Returning a value from vcs_run_vba

Your code becomes the body of a function named `MCP_TempFunction`. Assign to it to return a value:

    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs("qryExample")
    MCP_TempFunction = qd.SQL

Send statements, not a bare expression. Return a string or number; build JSON yourself when you need structure.

## Errors and line numbers

The wrapper runs with `On Error Resume Next` and reports the last runtime error as `errorNumber`, `error`, and `errorLine`. Lines are numbered for you: `errorLine: 7` is line 7 of the `code` you sent, counting blank lines, comments, and continuations.

To collect every failing line in one call, use your own handler. `Erl` matches `errorLine`:

    Dim col As New Collection
    On Error GoTo H
    CurrentDb.Execute "DELETE * FROM tblExample"
    CurrentDb.Execute "UPDATE tblOther SET x = 1"
    MCP_TempFunction = "errors=" & col.Count
    Exit Function
    H: col.Add Erl & ": " & Err.Number & " " & Err.Description
    Resume Next

## When the result reports a problem

- `compileError` saying the host project does not compile: the error is in the database's existing VBA. See `vcs_tips("compile")`.
- `temp_module_sweep_failed`, `temp_module_unresolvable`, or `temp_module_cleanup_failed`: an `MCP_Temp_*` module could not be created or removed. Stop and tell the user rather than calling again.
- `error_pattern: timeout`: the server stopped waiting; Access is left running. See `vcs_tips("stuck")`.

## Timeouts

`vcs_run_vba` and `vcs_call_vba` default to 45 seconds (`ACCESS_VCS_RUN_VBA_TIMEOUT_SEC`, `ACCESS_VCS_CALL_VBA_TIMEOUT_SEC`). Pass `timeout_seconds` for longer work, but keep it below the client's own request timeout or the client returns `-32001` with no result.
