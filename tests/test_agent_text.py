"""What agents read stays inside client limits and names only things that exist.

Claude Code cuts server instructions, and any tool description it loads
upfront, at 4,096 characters. The limits below leave headroom.

AGENTS.md loads on every turn for agents changing the server, so its content
and its routing table have separate budgets, and every reference doc must be
reachable from it: an unlinked doc is rarely opened.
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
    "screenshot.py", "vba_control.py", "compile_check.py", "settings_form.py",
    "worktree.py",
]
# vcs_-prefixed words that are not tool names: tools run-vba.md says do not
# exist, and a result field.
NOT_TOOL_NAMES = {"vcs_eval", "vcs_execute_code", "vcs_version"}

TOOL_NAME = re.compile(r"(?<![\w-])vcs_[a-z_]+")
TIPS_CALL = re.compile(r"vcs_tips\(\s*\\?[\"']([a-z0-9-]+)\\?[\"']\s*\)")

REPO = Path(__file__).resolve().parent.parent
AGENTS_FILE = REPO / "AGENTS.md"
ROUTING_HEADING = "## Where to read next"
AGENTS_CONTENT_LIMIT = 6000
ROUTING_LIMIT = 2400
ROUTING_ROWS_LIMIT = 20
LOCAL_LINK = re.compile(r"\]\((?!https?:|#)([^)#\s]+)")


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


def _agents_md() -> str:
    return AGENTS_FILE.read_text(encoding="utf-8").replace("\r\n", "\n")


def _routing_section(text: str) -> str:
    start = text.find(ROUTING_HEADING)
    assert start >= 0, f"AGENTS.md has no {ROUTING_HEADING!r} section"
    end = text.find("\n## ", start + len(ROUTING_HEADING))
    return text[start:] if end < 0 else text[start:end]


def _cost(text: str) -> int:
    """Characters an agent reads, ignoring line breaks so rewrapping costs nothing."""
    return len(text.replace("\n", ""))


def test_agents_md_content_fits():
    text = _agents_md()
    content = _cost(text) - _cost(_routing_section(text))
    assert content <= AGENTS_CONTENT_LIMIT, "move depth into docs/ and add a routing row"


def test_agents_md_routing_fits():
    routing = _routing_section(_agents_md())
    table = [line for line in routing.splitlines() if line.lstrip().startswith("|")]
    header, separator, *rows = table
    assert _cost(routing) <= ROUTING_LIMIT
    assert len(rows) <= ROUTING_ROWS_LIMIT


def test_agents_md_links_resolve():
    missing = [t for t in LOCAL_LINK.findall(_agents_md()) if not (REPO / t).exists()]
    assert not missing


def test_every_doc_is_linked_from_agents_md():
    linked = {(REPO / t).resolve() for t in LOCAL_LINK.findall(_agents_md())}
    unlinked = [p.name for p in sorted((REPO / "docs").glob("*.md")) if p.resolve() not in linked]
    assert not unlinked, "add a routing row for each doc in AGENTS.md"
