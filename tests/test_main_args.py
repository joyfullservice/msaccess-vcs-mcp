"""Command-line flags for the msaccess-vcs-mcp entry point.

The README tells users to check an install with ``--help`` and the release
checklist checks ``--version``. Both must exit before the server starts, which
would otherwise sit waiting on stdin.
"""

import pytest

from msaccess_vcs_mcp import __version__
from msaccess_vcs_mcp.main import _parse_args


@pytest.mark.parametrize("flag", ["--help", "--version"])
def test_flag_exits_before_server_starts(flag, capsys):
    with pytest.raises(SystemExit) as exc:
        _parse_args([flag])
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert "msaccess-vcs-mcp" in out


def test_version_reports_package_version(capsys):
    with pytest.raises(SystemExit):
        _parse_args(["--version"])
    assert capsys.readouterr().out.strip() == f"msaccess-vcs-mcp {__version__}"


def test_transport_defaults_and_unknown_arguments_ignored():
    args = _parse_args(["--some-client-flag"])
    assert args.transport == "stdio"
    assert args.port == 8000


def test_sse_transport():
    args = _parse_args(["--transport", "sse", "--port", "9000"])
    assert args.transport == "sse"
    assert args.port == 9000
