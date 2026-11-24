"""Read-only and destructive tools say so to the client."""

from msaccess_vcs_mcp import tools


def _by_name():
    return {t.name: t for t in tools.mcp._tool_manager.list_tools()}


def test_annotation_sets_name_registered_tools():
    registered = set(_by_name())
    assert tools.READ_ONLY_TOOLS <= registered
    assert tools.DESTRUCTIVE_TOOLS <= registered
    assert not tools.READ_ONLY_TOOLS & tools.DESTRUCTIVE_TOOLS


def test_read_only_tools():
    for name in tools.READ_ONLY_TOOLS:
        assert _by_name()[name].annotations.readOnlyHint is True, name


def test_destructive_tools():
    for name in tools.DESTRUCTIVE_TOOLS:
        annotations = _by_name()[name].annotations
        assert annotations.readOnlyHint is False, name
        assert annotations.destructiveHint is True, name


def test_code_execution_is_not_marked_read_only():
    for name in ("vcs_run_vba", "vcs_call_vba", "vcs_run_tests", "vcs_dismiss_dialog"):
        annotations = _by_name()[name].annotations
        assert annotations is None or not annotations.readOnlyHint, name
