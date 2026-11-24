"""Tests for picking an Office dialog's message and buttons out of its MSAA tree."""

from __future__ import annotations

from msaccess_vcs_mcp import accessibility as acc
from msaccess_vcs_mcp.accessibility import Node

DIALOG = 0x12
WINDOW = 0x09


def tree() -> list[Node]:
    """The shape the "Cannot Compile Project." dialog reports."""
    return [
        Node((), WINDOW, "Microsoft Access"),
        Node((WINDOW,), acc.ROLE_TITLEBAR, ""),
        Node((WINDOW, acc.ROLE_TITLEBAR), acc.ROLE_PUSHBUTTON, "Close"),
        Node((WINDOW,), DIALOG, "Microsoft Access"),
        Node((WINDOW, DIALOG), acc.ROLE_STATICTEXT, ""),
        Node((WINDOW, DIALOG), acc.ROLE_STATICTEXT, "Microsoft Access"),
        Node((WINDOW, DIALOG), acc.ROLE_STATICTEXT, "Cannot Compile Project."),
        Node((WINDOW, DIALOG), acc.ROLE_PUSHBUTTON, ""),
        Node((WINDOW, DIALOG), acc.ROLE_PUSHBUTTON, "&OK", handle=7, child_id=0),
        Node((WINDOW, DIALOG), acc.ROLE_PUSHBUTTON, "Hidden", invisible=True),
        Node((WINDOW,), acc.ROLE_SCROLLBAR, "Vertical"),
        Node((WINDOW, acc.ROLE_SCROLLBAR), acc.ROLE_PUSHBUTTON, "Line up"),
    ]


def test_dialog_content_skips_title_bar_scroll_bars_and_title():
    lines, buttons = acc.dialog_content(tree(), "Microsoft Access")
    assert lines == ["Cannot Compile Project."]
    assert buttons == ["OK"]


def test_find_button_ignores_accelerators_and_case():
    node = acc.find_button(tree(), "ok")
    assert node is not None and node.handle == 7


def test_find_button_never_matches_non_client_or_hidden_buttons():
    assert acc.find_button(tree(), "Close") is None
    assert acc.find_button(tree(), "Line up") is None
    assert acc.find_button(tree(), "Hidden") is None
