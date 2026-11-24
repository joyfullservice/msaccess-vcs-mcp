"""Failed results carry agent_guidance with the next call."""

import asyncio

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import EXEMPT_TOOLS
from msaccess_vcs_mcp.guidance import insert_after, with_guidance
from msaccess_vcs_mcp.tips import get_tips, topic_names


@pytest.mark.parametrize(
    "pattern, expected",
    [
        ("timeout", "vcs_capture_screenshot"),
        ("access_unresponsive", "vcs_capture_screenshot"),
        ("server_busy", "one at a time"),
        ("installed_addin_refused", "development copy"),
    ],
)
def test_guidance_by_error_pattern(pattern, expected):
    result = with_guidance("vcs_export_database", {
        "success": False, "error": "x", "error_pattern": pattern, "recoverable": True,
    })
    assert expected in result["agent_guidance"]
    assert list(result) == ["success", "error", "error_pattern", "agent_guidance", "recoverable"]


def test_rebuild_refusal_gets_guidance():
    result = with_guidance("vcs_rebuild_addin", {"success": False, "status": "refused"})
    assert "otherInstances" in result["agent_guidance"]


def test_existing_guidance_is_kept():
    result = with_guidance("vcs_run_vba", {
        "success": False, "error_pattern": "timeout", "agent_guidance": "specific",
    })
    assert result["agent_guidance"] == "specific"


@pytest.mark.parametrize("result", [
    {"success": True, "error_pattern": "timeout"},
    {"success": False, "error_pattern": "unknown"},
    "not a dict",
])
def test_no_guidance_added(result):
    assert with_guidance("vcs_run_vba", result) == result


def test_error_text_is_unchanged():
    original = {"success": False, "error": "x", "error_pattern": "timeout"}
    assert with_guidance("vcs_run_vba", original)["error"] == "x"


def test_insert_after_missing_anchor_appends():
    assert list(insert_after({"a": 1}, "missing", "b", 2)) == ["a", "b"]


def test_tips_lists_topics_without_argument():
    result = get_tips(None)
    names = [entry["topic"] for entry in result["topics"]]
    assert names == topic_names()
    assert {"addin-development", "stuck", "compile", "logs", "run-vba"} <= set(names)
    assert all(entry["about"] for entry in result["topics"])


def test_tips_returns_named_topic():
    result = get_tips("Stuck")
    assert result["topic"] == "stuck"
    assert "vcs_capture_screenshot" in result["content"]


def test_tips_unknown_topic_searches_text():
    assert get_tips("MCP_TempFunction")["topic"] == "run-vba"


def test_tips_no_match_lists_topics():
    result = get_tips("zzz-nothing")
    assert result["success"] is False
    assert result["topics"]


def test_tips_tool_is_exempt_and_always_loaded():
    assert "vcs_tips" in EXEMPT_TOOLS
    tool = next(t for t in tools.mcp._tool_manager.list_tools() if t.name == "vcs_tips")
    assert tool.meta == {"anthropic/alwaysLoad": True}
    others = [t.name for t in tools.mcp._tool_manager.list_tools()
              if t.name != "vcs_tips" and (t.meta or {}).get("anthropic/alwaysLoad")]
    assert others == []


def test_tips_tool_runs_through_wrapper(monkeypatch, tmp_path):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path))
    result = asyncio.run(tools.vcs_tips("logs"))
    assert result["topic"] == "logs"
