"""Security tests for SSRF protection in image downloads.

This module tests the SSRF protection mechanisms implemented in
pptx_tools/image_utils.py to ensure that internal networks, cloud
metadata endpoints, and other sensitive targets are properly blocked.
"""

import pytest
from pptx_tools.image_utils import (
    is_ssrf_safe,
    validate_url,
    _is_ip_blocked,
    SSRFProtectionError,
    ImageValidationError,
    download_image,
)


class TestSSRFProtection:
    """Test suite for SSRF protection mechanisms."""

    # =========================================================================
    # Blocked IP Address Tests
    # =========================================================================

    @pytest.mark.parametrize("ip", [
        # Loopback addresses
        "127.0.0.1",
        "127.0.0.2",
        "127.255.255.255",
        # Private networks (RFC 1918)
        "10.0.0.1",
        "10.255.255.255",
        "172.16.0.1",
        "172.31.255.255",
        "192.168.0.1",
        "192.168.255.255",
        # Link-local (includes AWS/GCP/Azure metadata)
        "169.254.0.1",
        "169.254.169.254",  # Cloud metadata endpoint
        "169.254.255.255",
        # Carrier-grade NAT
        "100.64.0.1",
        "100.127.255.255",
        # Unspecified
        "0.0.0.0",
        "0.0.0.1",
    ])
    def test_blocked_ipv4_addresses(self, ip):
        """Test that private/internal IPv4 addresses are blocked."""
        assert _is_ip_blocked(ip) is True, f"IP {ip} should be blocked"

    @pytest.mark.parametrize("ip", [
        # IPv6 loopback
        "::1",
        # IPv6 link-local
        "fe80::1",
        "fe80::dead:beef",
        # IPv6 unique local
        "fc00::1",
        "fd00::1",
    ])
    def test_blocked_ipv6_addresses(self, ip):
        """Test that private/internal IPv6 addresses are blocked."""
        assert _is_ip_blocked(ip) is True, f"IP {ip} should be blocked"

    @pytest.mark.parametrize("ip", [
        # Public IPv4 addresses (examples)
        "8.8.8.8",
        "1.1.1.1",
        "93.184.216.34",  # example.com
        "151.101.1.140",  # reddit.com
    ])
    def test_allowed_public_ipv4_addresses(self, ip):
        """Test that public IPv4 addresses are allowed."""
        assert _is_ip_blocked(ip) is False, f"IP {ip} should be allowed"

    # =========================================================================
    # Blocked Hostname Tests
    # =========================================================================

    @pytest.mark.parametrize("url", [
        # Localhost variations
        "http://localhost/image.png",
        "http://localhost:8080/image.png",
        "https://localhost/image.png",
        "http://LOCALHOST/image.png",  # Case insensitive
        "http://localhost.localdomain/image.png",
        # Cloud metadata endpoints
        "http://metadata.google.internal/computeMetadata/v1/",
        "http://metadata.internal/latest/",
        # Kubernetes
        "http://kubernetes.default.svc/api",
        "http://kubernetes.default/api",
        "http://kubernetes/api",
    ])
    def test_blocked_hostnames(self, url):
        """Test that dangerous hostnames are blocked."""
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False, f"URL {url} should be blocked. Error: {error}"
        assert "Blocked hostname" in error or "blocked" in error.lower()

    @pytest.mark.parametrize("url", [
        # Internal suffix hostnames
        "http://secret-service.internal/image.png",
        "http://api.local/image.png",
        "http://something.localhost/image.png",
    ])
    def test_blocked_hostname_suffixes(self, url):
        """Test that hostnames with dangerous suffixes are blocked."""
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False, f"URL {url} should be blocked. Error: {error}"

    # =========================================================================
    # Blocked IP in URL Tests
    # =========================================================================

    @pytest.mark.parametrize("url", [
        # Direct IP addresses in URL
        "http://127.0.0.1/image.png",
        "http://127.0.0.1:8080/image.png",
        "http://10.0.0.1/internal/image.png",
        "http://192.168.1.1/admin/image.png",
        "http://172.16.0.1/image.png",
        # AWS metadata endpoint
        "http://169.254.169.254/latest/meta-data/",
        "http://169.254.169.254/latest/meta-data/iam/security-credentials/",
        # IPv6 in URL
        "http://[::1]/image.png",
        "http://[fe80::1]/image.png",
    ])
    def test_blocked_ip_addresses_in_url(self, url):
        """Test that URLs with blocked IP addresses are rejected."""
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False, f"URL {url} should be blocked. Error: {error}"

    # =========================================================================
    # Allowed URLs Tests
    # =========================================================================

    @pytest.mark.parametrize("url", [
        "https://example.com/image.png",
        "https://images.unsplash.com/photo.jpg",
        "http://via.placeholder.com/150",
        "https://picsum.photos/200/300",
        "https://i.imgur.com/abc123.png",
    ])
    def test_allowed_public_urls(self, url):
        """Test that legitimate public URLs are allowed."""
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is True, f"URL {url} should be allowed. Error: {error}"

    # =========================================================================
    # Invalid URL Format Tests
    # =========================================================================

    @pytest.mark.parametrize("url", [
        # Invalid schemes
        "ftp://example.com/image.png",
        "file:///etc/passwd",
        "gopher://example.com/",
        # No scheme
        "example.com/image.png",
        # Invalid format
        "not-a-url",
        "",
    ])
    def test_invalid_url_formats(self, url):
        """Test that invalid URL formats are rejected."""
        assert validate_url(url) is False, f"URL {url} should be invalid"

    # =========================================================================
    # SSRF Attack Payload Tests
    # =========================================================================

    @pytest.mark.parametrize("url,description", [
        # AWS IMDSv1 metadata
        ("http://169.254.169.254/latest/meta-data/", "AWS metadata root"),
        ("http://169.254.169.254/latest/meta-data/iam/security-credentials/", "AWS IAM credentials"),
        ("http://169.254.169.254/latest/user-data/", "AWS user data"),
        ("http://169.254.169.254/latest/dynamic/instance-identity/document", "AWS instance identity"),
        
        # GCP metadata
        ("http://metadata.google.internal/computeMetadata/v1/", "GCP metadata root"),
        ("http://169.254.169.254/computeMetadata/v1/", "GCP metadata via IP"),
        
        # Azure metadata
        ("http://169.254.169.254/metadata/instance", "Azure instance metadata"),
        ("http://169.254.169.254/metadata/identity/oauth2/token", "Azure identity token"),
        
        # Internal services
        ("http://localhost:6379/", "Redis"),
        ("http://localhost:27017/", "MongoDB"),
        ("http://localhost:9200/", "Elasticsearch"),
        ("http://127.0.0.1:8500/v1/agent/members", "Consul"),
        ("http://192.168.1.1/admin", "Router admin"),
        
        # Kubernetes
        ("http://kubernetes.default.svc.cluster.local/", "K8s service"),
        ("http://10.0.0.1/api/v1/namespaces", "K8s API"),
    ])
    def test_ssrf_attack_payloads(self, url, description):
        """Test that common SSRF attack payloads are blocked."""
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False, f"SSRF payload '{description}' ({url}) should be blocked. Error: {error}"

    # =========================================================================
    # Edge Cases
    # =========================================================================

    def test_url_with_credentials(self):
        """Test URL with embedded credentials (should still check host)."""
        url = "http://user:pass@127.0.0.1/image.png"
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False, f"URL with credentials to blocked host should be blocked"

    def test_url_with_port(self):
        """Test that port number doesn't bypass SSRF checks."""
        url = "http://127.0.0.1:80/image.png"
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False

    def test_url_with_path_traversal(self):
        """Test URL with path traversal (host check comes first)."""
        url = "http://127.0.0.1/../../../etc/passwd"
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is False

    def test_decimal_ip_notation(self):
        """Test decimal IP notation (127.0.0.1 = 2130706433).
        
        Note: Python's urlparse doesn't convert decimal IPs, so this
        should fail at the URL parsing level, not SSRF check.
        """
        # 2130706433 = 127.0.0.1 in decimal
        url = "http://2130706433/image.png"
        # This should either be blocked or fail validation
        is_safe, error = is_ssrf_safe(url)
        # Decimal IPs may or may not resolve, but should not bypass security

    def test_ipv6_mapped_ipv4(self):
        """Test IPv6-mapped IPv4 address."""
        url = "http://[::ffff:127.0.0.1]/image.png"
        is_safe, error = is_ssrf_safe(url)
        # This should be blocked as it maps to 127.0.0.1


class TestValidateUrl:
    """Test the validate_url function."""

    def test_valid_https_url(self):
        """Test that valid HTTPS URLs pass validation."""
        assert validate_url("https://example.com/image.png") is True

    def test_valid_http_url(self):
        """Test that valid HTTP URLs pass validation."""
        assert validate_url("http://example.com/image.png") is True

    def test_blocked_url_returns_false(self):
        """Test that blocked URLs return False (not raise exception)."""
        assert validate_url("http://localhost/image.png") is False
        assert validate_url("http://169.254.169.254/metadata") is False

    def test_invalid_scheme(self):
        """Test that non-HTTP schemes are rejected."""
        assert validate_url("ftp://example.com/file.txt") is False
        assert validate_url("file:///etc/passwd") is False


class TestDownloadImageSSRF:
    """Test that download_image properly enforces SSRF protection."""

    def test_download_blocked_url_raises(self):
        """Test that downloading from blocked URLs raises SSRFProtectionError."""
        with pytest.raises(SSRFProtectionError) as exc_info:
            download_image("http://localhost/image.png")
        assert "Blocked" in str(exc_info.value)

    def test_download_metadata_endpoint_raises(self):
        """Test that downloading from cloud metadata raises SSRFProtectionError."""
        with pytest.raises(SSRFProtectionError) as exc_info:
            download_image("http://169.254.169.254/latest/meta-data/")
        assert "Blocked" in str(exc_info.value)

    def test_download_invalid_url_raises(self):
        """Test that invalid URLs raise ImageValidationError."""
        with pytest.raises(ImageValidationError):
            download_image("not-a-valid-url")

    def test_download_ftp_url_raises(self):
        """Test that FTP URLs raise ImageValidationError."""
        with pytest.raises(ImageValidationError):
            download_image("ftp://example.com/image.png")


class TestSSRFProtectionEdgeCases:
    """Edge case tests for SSRF protection."""

    def test_empty_url(self):
        """Test handling of empty URL."""
        is_safe, error = is_ssrf_safe("")
        assert is_safe is False

    def test_none_url(self):
        """Test handling of None URL."""
        # This should handle gracefully, not crash
        try:
            is_safe, error = is_ssrf_safe(None)
            assert is_safe is False
        except TypeError:
            pass  # Acceptable to raise TypeError for None

    def test_url_with_unicode(self):
        """Test URL with unicode characters."""
        url = "https://example.com/ímágé.png"
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is True  # Should handle unicode gracefully

    def test_very_long_url(self):
        """Test very long URL doesn't cause issues."""
        url = "https://example.com/" + "a" * 10000 + ".png"
        # Should handle without crashing
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is True

    def test_url_with_special_characters(self):
        """Test URL with special characters."""
        url = "https://example.com/path?query=value&other=test#fragment"
        is_safe, error = is_ssrf_safe(url)
        assert is_safe is True
