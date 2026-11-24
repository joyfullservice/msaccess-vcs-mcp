# VBA compile errors

## Compile and find the error

    vcs_compile_vba(database_path)

It compiles from the VBA editor. On failure it returns `compile_error` with the message, module, line, columns, and code of the first error. Fix that line in source, import it (`vcs_import_object`), and compile again. `vcs_check_vba_compiled(database_path)` only reports whether the project is compiled.

- `compile_error` has no line: call `vcs_capture_screenshot(database_path, target="vbe")`; the editor highlights the error.
- Neither shows the error: stop editing and ask the user to compile in the VBA editor (Debug, Compile) and paste the code around the highlighted line. Do not iterate through speculative edits.
- `compile_timeout`: the editor never finished or never showed its error. Look with `vcs_capture_screenshot(database_path, include_image=False)`.
- `dialog_shown`: a different dialog appeared. Read it from the screenshot metadata before doing anything else.

## Compile errors from vcs_run_vba

When a `vcs_run_vba` result has `compileError` saying the host project does not compile on its own, the error is in the database's existing VBA, not in the code you sent. Run `vcs_compile_vba` to find it rather than editing your snippet.

## After fixing source

An import does not compile the project. Compile again after each fix; one error hides the next.
