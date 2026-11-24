"""Screenshot captures against a real Access window (Windows desktop required).

These attach read-only to whatever Access is already running and skip when
there is none. They never start, focus, restore, or close Access, so they
are safe to run while a user is working in it.
"""

from __future__ import annotations

import asyncio
import json
import sys

import pytest

pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(sys.platform != "win32", reason="Win32 capture"),
]


def _running_access() -> dict[int, dict]:
    from msaccess_vcs_mcp import screenshot

    return screenshot.survey_access()


@pytest.fixture
def access_pid() -> int:
    survey = {pid: s for pid, s in _running_access().items() if s["title"]}
    if not survey:
        pytest.skip("No Access window is open.")
    return next(iter(survey))


def _meta(result) -> dict:
    return json.loads(result.content[0].text)


def _capture(**kwargs):
    from msaccess_vcs_mcp import screenshot

    args = dict(database_path=None, target="application", object_type=None, object_name=None,
                open_if_closed=False, pid=None, max_dimension=800, include_image=True,
                get_gate=lambda: pytest.fail("window captures must not use the gate"))
    args.update(kwargs)
    return asyncio.run(screenshot.capture_screenshot(**args))


def test_application_capture(access_pid):
    result = _capture(pid=access_pid)
    meta = _meta(result)
    if meta.get("error_pattern") in ("window_minimized", "window_hidden"):
        pytest.skip(meta["error"])
    assert meta["success"] is True, meta
    assert meta["window"]["class"] == "OMain"
    assert max(meta["image"]["width"], meta["image"]["height"]) <= 800
    image = result.content[1]
    assert image.type == "image" and image.mimeType == "image/png"


def test_metadata_only_capture(access_pid):
    meta = _meta(_capture(pid=access_pid, include_image=False))
    assert "image" not in meta
    assert isinstance(meta.get("dialogs", []), list)


def test_vbe_capture_or_not_open(access_pid):
    meta = _meta(_capture(pid=access_pid, target="vbe"))
    assert meta["success"] is True or meta["error_pattern"] in (
        "vbe_not_open", "window_minimized", "window_hidden"
    ), meta


def test_restart_manager_finds_holders_without_com(access_pid):
    """The Restart Manager answers for any file; a missing one has no holders."""
    from msaccess_vcs_mcp import screen_capture

    assert screen_capture.pids_holding_file(r"C:\path\to\missing.accdb") in (set(), None)
