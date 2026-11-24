"""What agents read stays inside client limits and names only things that exist.

Claude Code cuts server instructions, and any tool description it loads
upfront, at 4,096 characters. The limits below leave headroom.
"""

import ast
import re
from pathlib import Path

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.tips import topic_names

INSTRUCTIONS_LIMIT = 3500
DESCRIPTION_LIMIT = 4000
CRITICAL_WINDOW = 2000

PACKAGE = Path(tools.__file__).parent
GUIDANCE_SOURCES = [
    "tools.py", "guidance.py", "access_state.py", "vba_worker_manager.py",
    "screenshot.py", "vba_control.py", "compile_check.py",
]
# vcs_-prefixed words that are not tool names: tools run-vba.md says do not
# exist, and a result field.
NOT_TOOL_NAMES = {"vcs_eval", "vcs_execute_code", "vcs_version"}

TOOL_NAME = re.compile(r"(?<![\w-])vcs_[a-z_]+")
TIPS_CALL = re.compile(r"vcs_tips\(\s*\\?[\"']([a-z0-9-]+)\\?[\"']\s*\)")


def _tools():
    return tools.mcp._tool_manager.list_tools()


def _string_literals(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    return [
        node.value for node in ast.walk(tree)
        if isinstance(node, ast.Constant) and isinstance(node.value, str)
    ]


def _agent_texts() -> dict[str, str]:
    texts = {"instructions": tools.SERVER_INSTRUCTIONS}
    texts.update({f"description of {t.name}": t.description or "" for t in _tools()})
    texts.update({
        f"tips/{p.name}": p.read_text(encoding="utf-8")
        for p in (PACKAGE / "tips").glob("*.md")
    })
    for name in GUIDANCE_SOURCES:
        texts[name] = "\n".join(_string_literals(PACKAGE / name))
    return texts


def test_instructions_fit():
    assert len(tools.SERVER_INSTRUCTIONS) <= INSTRUCTIONS_LIMIT


@pytest.mark.parametrize("tool", _tools(), ids=lambda t: t.name)
def test_description_fits(tool):
    assert len(tool.description or "") <= DESCRIPTION_LIMIT


@pytest.mark.parametrize("phrase", [
    "installed add-in",
    "not in parallel",
    "vcs_capture_screenshot(database_path, include_image=False)",
    'vcs_tips("addin-development")',
    'vcs_tips("stuck")',
])
def test_critical_rules_come_first(phrase):
    position = tools.SERVER_INSTRUCTIONS.find(phrase)
    assert 0 <= position
    assert position + len(phrase) <= CRITICAL_WINDOW


def test_every_tool_name_mentioned_is_registered():
    registered = {t.name for t in _tools()}
    unknown = {
        f"{name} in {where}"
        for where, text in _agent_texts().items()
        for name in TOOL_NAME.findall(text)
        if name not in registered and name not in NOT_TOOL_NAMES
    }
    assert not unknown


def test_every_tips_topic_mentioned_has_a_file():
    topics = set(topic_names())
    missing = {
        f"{topic} in {where}"
        for where, text in _agent_texts().items()
        for topic in TIPS_CALL.findall(text)
        if topic not in topics
    }
    assert not missing


def test_tips_topics_are_mentioned_where_agents_look():
    text = tools.SERVER_INSTRUCTIONS + next(
        t.description for t in _tools() if t.name == "vcs_tips"
    )
    assert all(topic in text for topic in topic_names())


def test_topic_names_are_unique_ignoring_case():
    names = topic_names()
    assert len({n.lower() for n in names}) == len(names)
