# Contributing to msaccess-vcs-mcp

For the architecture, conventions, and the rules every change must keep, read
[AGENTS.md](AGENTS.md) first. This file covers the development environment,
dependencies, and publishing a release.

## Development environment

Day-to-day development uses the editable install in `venv/` described in
[AGENTS.md](AGENTS.md#development-workflow):

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e ".[dev]"
pytest -m "not integration"
```

CI and the publish workflow use [uv](https://docs.astral.sh/uv/) with the
committed `uv.lock`, so every run gets the same dependency versions. To
reproduce a CI run locally:

```powershell
uv sync --frozen --extra dev
uv run pytest tests -m "not integration"
```

## Dependencies

When you change a dependency in `pyproject.toml`, regenerate and commit the
lock, or CI fails on `uv sync --frozen`:

```powershell
uv lock              # re-resolve after editing pyproject.toml
uv lock --upgrade    # refresh every resolution to the newest allowed version
```

Resolution skips releases published in the last 3 days (`exclude-newer` in
`pyproject.toml`), matching the Dependabot cooldown, so a compromised upload
has time to be found and pulled before it is locked. The cooldown is for
routine updates. When `uv audit` reports a vulnerability whose fix is younger
than 3 days, exempt that one package with a fixed cutoff just after the fix
was published, then run `uv lock --upgrade-package <name>`:

```toml
[tool.uv]
exclude-newer = "3 days"
exclude-newer-package = { pyjwt = "2026-10-05T00:00:00Z" }
```

Remove the exemption once the fixed version is more than 3 days old.

## Publishing a release

Publishing a GitHub Release runs `.github/workflows/publish.yml`, which tests,
builds, and uploads the package to
[PyPI](https://pypi.org/project/msaccess-vcs-mcp/) through trusted publishing
(no API token is stored anywhere). Users install with
`uvx msaccess-vcs-mcp@latest`.

### One-time PyPI setup

1. On PyPI, open **Your account → Publishing** and add a **pending
   publisher** (the project does not exist on PyPI until the first upload):
   - PyPI Project Name: `msaccess-vcs-mcp`
   - Owner: `joyfullservice`
   - Repository name: `msaccess-vcs-mcp`
   - Workflow name: `publish.yml`
   - Environment name: `pypi`
2. On GitHub, open the repository's **Settings → Environments**, create an
   environment named `pypi`, and add yourself as a required reviewer so each
   upload waits for approval.

A pending publisher does not reserve the name. Publish the first release soon
after adding it.

### Release checklist

1. **Check the add-in it needs.** Features gated on an add-in version (see the
   `_MIN_VERSION` constants in `tools.py`) must be in a published add-in
   release, or the README must say which add-in version they need.
2. **Update the version** in both places; they must match:
   - `pyproject.toml`, the `version` field
   - `src/msaccess_vcs_mcp/__init__.py`, `__version__`
3. **Run the quality gates locally**, including the integration tests, which
   need Access and the add-in installed and do not run in CI:

   ```powershell
   uv run pytest tests -m "not integration"
   uv run pytest tests -m integration
   uv build
   uv run twine check dist/*
   ```

4. **Commit and push** the version bump to `main`, and check that CI passes.
5. **Create a GitHub Release** with a new tag matching the version (`v0.2.0`
   for version `0.2.0`; the workflow refuses a mismatch) and release notes.
6. **Approve the upload** in the Actions tab when the `pypi` environment asks.
7. **Check the published package:**

   ```powershell
   uvx msaccess-vcs-mcp@latest --version
   ```
