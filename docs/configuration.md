# Configuration and user settings

How the server reads its configuration, and how the few settings that need a user's decision are asked for and stored. Read this before adding an environment variable or a setting that asks the user.

[`.env.example`](../.env.example) is the one list of environment variables, with each default and what it does. A new variable gets its entry there, not a second list here.

## Where values come from

All configuration comes from environment variables. `config.py` finds the project root (the order is at the top of `.env.example`), then loads `.env` and `.env.local` from it:

- On the first load, a value already in the environment, such as one from the `env` section of `mcp.json`, beats `.env` (`override=False`). `.env.local` beats both.
- `load_config()` runs at the start of every tool call, inside the `vcs_tool` wrapper, and reloads both files when either one's modification time changes. A reload uses `override=True`, so an edited value replaces the old one.
- When the client reports workspace roots, the `.env` found there is loaded with `override=False` and its `.env.local` with `override=True`.

Settings in the user settings registry below add one more layer: `settings.json`, read after the environment and before the default.

## User settings

A few settings get an explicit decision from the user instead of a silent default; `DECISION_SETTINGS` in `user_settings.py` lists them (today only screenshot retention). Decisions live per user in `settings.json`, because one server process serves every window and a per-project file would ask in every repository.

- **Read through `get_setting(key)`, never `os.environ` alone.** The order is environment (`mcp.json`, `.env`, `.env.local`), then `settings.json`, then the default. Nothing is written into `os.environ`: the workspace-roots `.env` load uses `override=False`, so a pre-filled value would beat the project's `.env`. The file is re-read when its modification time changes, so a change from the CLI or another server process applies on the next read.
- **Writes merge under a lock file** (`settings.json.lock`, cleared after 30 seconds) and replace the file atomically. Unknown keys are kept. An unreadable file is renamed to `settings.json.corrupt-*` and treated as empty.
- **Only a deliberate answer is a decision.** Each entry records `decided_by` (`form`, `agent`, or `cli`). `{"decision": "default"}` follows the built-in default even if a later release changes it. A form that is declined, cancelled, or times out saves nothing. A key with no entry is pending, which is also how a setting added in a later release is detected.
- **Never ask inside an unrelated call.** `vcs_configure` (gate-exempt, `settings_form.py`) is the only place the server elicits. Check `elicitation.form` directly, because the SDK's `check_client_capability` accepts a URL-only client. Choice fields need `str` with an `enum` in `json_schema_extra`, because `Literal` fails the SDK's schema check. Without form support, it returns `agent_guidance` telling the agent to ask the user and pass the answer.
- **Surface a pending setting where it matters.** `vcs_capture_screenshot` adds `image.retention` (days, `expires_at`, `decided`) to every saved image, plus `agent_guidance` while retention is pending. Under `never` it saves nothing, warns, and deletes earlier `vcs-shot-` files. A new decision setting needs the same hook in the tool whose results depend on it, not a notice on every call. Retention's values and cleanup are in [screenshots.md](screenshots.md#returned-and-saved-images).
- **Keep audit and security switches out of the registry.** An agent answering for the user could otherwise turn off `ACCESS_VCS_ENABLE_LOGGING`.
- `msaccess-vcs configure` (`--show`, `--set NAME=VALUE`, `--reset NAME`) writes the file directly without starting a server, and prompts only when stdin is a terminal.
