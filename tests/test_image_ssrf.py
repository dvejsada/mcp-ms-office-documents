"""Tests for SSRF protection on caller-supplied image URLs.

Covers:
- Non-public literals (loopback, RFC 1918, link-local/cloud metadata) blocked
- Alternative IP encodings (decimal, octal, IPv4-mapped IPv6) blocked
- Public addresses allowed
- Unresolvable hosts rejected rather than passed through
- Redirects to a non-public address blocked
- SSRF_ALLOW_PRIVATE_ADDRESSES escape hatch

Hostname resolution is stubbed where a test would otherwise need DNS; the
alternative-encoding cases deliberately use the real resolver because
normalising them is the behaviour under test (libc parses them locally, so no
network access is involved).
"""

import socket
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

import requests

from pptx_tools.image_utils import (
    REQUEST_TIMEOUT,
    SSRFProtectionError,
    _get_following_redirects,
    assert_url_is_public,
    download_image,
)


@pytest.fixture(autouse=True)
def _protection_enabled():
    """Run every test with protection on unless it opts out."""
    with patch("pptx_tools.image_utils.get_config") as get_config:
        get_config.return_value.allow_private_image_addresses = False
        yield get_config


def _resolver(hosts):
    """Stub getaddrinfo from a {hostname: [address, ...]} mapping.

    Hosts not in the mapping resolve to themselves, so an IP literal in a
    redirect target still behaves like the real resolver.
    """
    def _getaddrinfo(host, *a, **kw):
        return [(None, None, None, "", (addr, 0)) for addr in hosts.get(host, [host])]
    return _getaddrinfo


@pytest.mark.parametrize("url", [
    "http://127.0.0.1/image.png",                      # loopback
    "http://[::1]/image.png",                          # IPv6 loopback
    "http://10.0.0.1/image.png",                       # RFC 1918
    "http://192.168.1.1/image.png",                    # RFC 1918
    "http://172.16.0.1/image.png",                     # RFC 1918
    "http://169.254.169.254/latest/meta-data/",        # cloud metadata
    "http://100.64.0.1/image.png",                     # CGNAT
    "http://0.0.0.0/image.png",                        # unspecified
    "http://user:pass@127.0.0.1/image.png",            # credentials do not bypass
    "http://127.0.0.1:8080/image.png",                 # port does not bypass
])
def test_non_public_addresses_are_blocked(url):
    with pytest.raises(SSRFProtectionError):
        assert_url_is_public(url)


@pytest.mark.parametrize("url", [
    "http://2130706433/image.png",         # 127.0.0.1 in decimal
    "http://0177.0.0.1/image.png",         # 127.0.0.1 with an octal first octet
    "http://[::ffff:127.0.0.1]/image.png",  # IPv4-mapped IPv6
])
def test_alternative_encodings_of_loopback_are_blocked(url):
    """These forms bypass a naive ipaddress.ip_address() blocklist check."""
    with pytest.raises(SSRFProtectionError):
        assert_url_is_public(url)


def test_public_address_is_allowed():
    with patch.object(socket, "getaddrinfo", _resolver({"example.com": ["93.184.216.34"]})):
        assert_url_is_public("https://example.com/image.png")  # does not raise


def test_all_resolved_addresses_must_be_public():
    """A host with one public and one private record is rejected."""
    resolver = _resolver({"split-horizon.example": ["93.184.216.34", "10.0.0.5"]})
    with patch.object(socket, "getaddrinfo", resolver):
        with pytest.raises(SSRFProtectionError, match="10.0.0.5"):
            assert_url_is_public("https://split-horizon.example/image.png")


def test_unresolvable_host_is_rejected():
    def _fail(*a, **kw):
        raise socket.gaierror("Name or service not known")

    with patch.object(socket, "getaddrinfo", _fail):
        with pytest.raises(SSRFProtectionError, match="Cannot resolve"):
            assert_url_is_public("https://no-such-host.example/image.png")


def test_url_without_hostname_is_rejected():
    with pytest.raises(SSRFProtectionError, match="no hostname"):
        assert_url_is_public("file:///etc/passwd")


def test_redirect_to_private_address_is_blocked():
    """A public URL must not be able to bounce the request onto loopback."""
    with patch.object(socket, "getaddrinfo", _resolver({"example.com": ["93.184.216.34"]})):
        with patch("pptx_tools.image_utils.requests.get") as get:
            get.return_value.is_redirect = True
            get.return_value.headers = {"Location": "http://127.0.0.1/secret.png"}
            with pytest.raises(SSRFProtectionError):
                download_image("https://example.com/image.png")


def test_escape_hatch_allows_private_addresses(_protection_enabled):
    _protection_enabled.return_value.allow_private_image_addresses = True
    assert_url_is_public("http://127.0.0.1/image.png")  # does not raise


def _redirect_to(location):
    return MagicMock(is_redirect=True, headers={"Location": location})


def test_intermediate_redirect_responses_are_closed():
    """A 3xx body is never read, so its response is released explicitly."""
    redirect = _redirect_to("https://example.com/final.png")
    final = MagicMock(is_redirect=False, headers={"Content-Type": "image/png"})
    final.iter_content.return_value = [b"image-bytes"]

    with patch.object(socket, "getaddrinfo", _resolver({"example.com": ["93.184.216.34"]})):
        with patch("pptx_tools.image_utils.requests.get", side_effect=[redirect, final]):
            download_image("https://example.com/image.png")

    redirect.close.assert_called_once()


def test_redirect_chain_shares_one_timeout_budget():
    """The whole chain is bounded by REQUEST_TIMEOUT, not each hop separately."""
    redirect = _redirect_to("https://example.com/next.png")
    # monotonic(): deadline base, then one reading per loop iteration.
    clock = iter([0.0, 0.0, REQUEST_TIMEOUT + 1])

    with patch.object(socket, "getaddrinfo", _resolver({"example.com": ["93.184.216.34"]})):
        with patch("pptx_tools.image_utils.time.monotonic", lambda: next(clock)):
            with patch("pptx_tools.image_utils.requests.get", return_value=redirect):
                with pytest.raises(requests.exceptions.Timeout):
                    _get_following_redirects("https://example.com/image.png")


def test_each_hop_gets_only_the_remaining_budget():
    """A hop's timeout shrinks as the shared budget is consumed."""
    redirect = _redirect_to("https://example.com/next.png")
    final = MagicMock(is_redirect=False, headers={})
    clock = iter([0.0, 0.0, 10.0])
    timeouts = []

    def _get(url, **kwargs):
        timeouts.append(kwargs["timeout"])
        return redirect if len(timeouts) == 1 else final

    with patch.object(socket, "getaddrinfo", _resolver({"example.com": ["93.184.216.34"]})):
        with patch("pptx_tools.image_utils.time.monotonic", lambda: next(clock)):
            with patch("pptx_tools.image_utils.requests.get", _get):
                _get_following_redirects("https://example.com/image.png")

    assert timeouts == [REQUEST_TIMEOUT, REQUEST_TIMEOUT - 10.0]
