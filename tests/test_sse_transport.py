"""Tests for ``--transport sse``: where it listens, and which requests it accepts.

``FastMCP.run()`` takes no host or port (mcp 1.28 to 1.30): they are settings, and
passing them to ``run()`` raised a TypeError at start-up. The server object is built
for 127.0.0.1, which also turns on the Host header check for localhost only, so a
client on another machine (the reason for SSE here: an MCP client on one machine
reaching Access on a Windows machine across the LAN, physical or a VM) was refused
with 421. ``_configure_sse`` sets the host and port, and keeps
the check on with the one address listened on allowed as well.
"""

import pytest
from mcp.server.fastmcp import FastMCP
from starlette.testclient import TestClient

from msaccess_vcs_mcp.main import _configure_sse

# The Windows machine's address: on a LAN, and on a desktop hypervisor's own network.
# Nothing in the configuration depends on which kind it is.
LAN = "192.168.1.50"
PARALLELS = "10.211.55.3"
ADDRESSES = [LAN, PARALLELS]


def _post(app, host, origin=None):
    """POST to the SSE message endpoint with the given Host (and Origin)."""
    headers = {"Host": host, "Content-Type": "application/json"}
    if origin:
        headers["Origin"] = origin
    return TestClient(app).post("/messages/", headers=headers, content=b"{}")


def _configured(host):
    mcp = FastMCP("test")
    _configure_sse(mcp, host, 8765)
    return mcp


def test_run_takes_no_host_or_port():
    """The call this replaces fails before anything listens."""
    with pytest.raises(TypeError):
        FastMCP("test").run(transport="sse", host=LAN, port=8765)


def test_settings_carry_host_and_port():
    mcp = _configured(LAN)
    assert mcp.settings.host == LAN
    assert mcp.settings.port == 8765


@pytest.mark.parametrize("address", ADDRESSES)
def test_unconfigured_server_refuses_another_machine(address):
    """Why the configuration is needed: the default check allows localhost only."""
    assert _post(FastMCP("test").sse_app(), f"{address}:8765").status_code == 421


@pytest.mark.parametrize("address", ADDRESSES)
@pytest.mark.parametrize("host_header", ["{address}:8765", "127.0.0.1:8765", "localhost:8765"])
def test_configured_server_accepts_its_own_address_and_localhost(address, host_header):
    """Its own address for a direct connection; localhost for an SSH tunnel."""
    # 400 "session_id is required" means the request passed the Host check
    response = _post(_configured(address).sse_app(), host_header.format(address=address))
    assert response.status_code == 400
    assert "session_id" in response.text


def test_configured_server_still_refuses_other_hosts():
    assert _post(_configured(LAN).sse_app(), "evil.example:8765").status_code == 421


def test_a_name_for_the_machine_is_refused():
    """The client must use the IP address it listens on: a name is another Host."""
    assert _post(_configured(LAN).sse_app(), "win11.lan:8765").status_code == 421


def test_configured_server_still_refuses_other_origins():
    response = _post(_configured(LAN).sse_app(), f"{LAN}:8765", origin="http://evil.example")
    assert response.status_code == 403


def test_ipv6_address_is_bracketed():
    app = _configured("fe80::1").sse_app()
    assert _post(app, "[fe80::1]:8765").status_code == 400
    assert _post(app, "evil.example:8765").status_code == 421
