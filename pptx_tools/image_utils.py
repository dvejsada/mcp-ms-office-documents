"""Utility functions for handling images in PowerPoint presentations.

This module provides functionality to download and validate images from URLs
for embedding in PowerPoint slides.

Downloads are restricted to publicly routable addresses (see
:func:`assert_url_is_public`) so that a caller-supplied image URL cannot be
used to reach loopback, private-network or cloud-metadata services.
"""

import io
import ipaddress
import logging
import socket
import time
from typing import Tuple
from urllib.parse import urljoin, urlparse

import requests

from config import get_config

logger = logging.getLogger(__name__)

# Allowed image MIME types
ALLOWED_MIME_TYPES = {
    'image/png',
    'image/jpeg',
    'image/jpg',
    'image/gif',
    'image/bmp',
    'image/webp',
    'image/tiff',
}

# Maximum image size in bytes (10 MB)
MAX_IMAGE_SIZE = 10 * 1024 * 1024

# Request timeout in seconds
REQUEST_TIMEOUT = 30

# Maximum number of HTTP redirects to follow. Every hop is re-validated, so a
# public URL cannot bounce the request onto an internal address.
MAX_REDIRECTS = 5


class ImageDownloadError(Exception):
    """Exception raised when image download fails."""
    pass


class ImageValidationError(Exception):
    """Exception raised when image validation fails."""
    pass


class SSRFProtectionError(ImageValidationError):
    """Exception raised when a URL resolves to a non-public address.

    Subclasses :class:`ImageValidationError` so callers that already handle
    validation failures treat a blocked URL the same way.
    """
    pass


def validate_url(url: str) -> bool:
    """Validate that a URL is well-formed and uses http/https.

    Args:
        url: URL string to validate.

    Returns:
        True if URL is valid, False otherwise.
    """
    try:
        parsed = urlparse(url)
        return parsed.scheme in ('http', 'https') and bool(parsed.netloc)
    except Exception:
        return False


def _addr_is_public(ip) -> bool:
    """Return True only for globally routable unicast addresses.

    ``is_global`` already excludes loopback, RFC 1918, link-local (including
    the 169.254.169.254 cloud-metadata address), CGNAT, documentation and
    reserved ranges for both IPv4 and IPv6; multicast is rejected separately
    because ``is_global`` does not cover it.
    """
    # Collapse IPv4-mapped IPv6 (::ffff:127.0.0.1) to the IPv4 address it targets.
    ip = getattr(ip, 'ipv4_mapped', None) or ip
    return ip.is_global and not (
        ip.is_private or ip.is_loopback or ip.is_link_local
        or ip.is_reserved or ip.is_multicast
    )


def assert_url_is_public(url: str) -> None:
    """Raise unless every address the URL's host resolves to is public.

    Resolution goes through ``getaddrinfo`` for *all* hosts, including bare IP
    literals, so that alternative encodings (``2130706433``, ``0177.0.0.1``)
    are normalised the same way the HTTP client will normalise them.

    Known limitation: the hostname is resolved again when the connection is
    made, so a DNS entry with a very short TTL can answer differently for the
    two lookups (DNS rebinding). Closing that requires pinning the validated
    address into the connection; it is not attempted here.

    Args:
        url: URL to check.

    Raises:
        SSRFProtectionError: If the host is missing, cannot be resolved, or
            resolves to a non-public address.
    """
    if get_config().allow_private_image_addresses:
        return

    hostname = urlparse(url).hostname
    if not hostname:
        raise SSRFProtectionError(f"URL has no hostname: {url}")

    try:
        addr_infos = socket.getaddrinfo(hostname, None, type=socket.SOCK_STREAM)
    except socket.gaierror as e:
        raise SSRFProtectionError(f"Cannot resolve host '{hostname}': {e}")

    for info in addr_infos:
        address = ipaddress.ip_address(info[4][0])
        if not _addr_is_public(address):
            raise SSRFProtectionError(
                f"Host '{hostname}' resolves to non-public address {address}"
            )


def _get_following_redirects(url: str) -> requests.Response:
    """GET *url*, validating the initial URL and every redirect hop.

    Redirects are followed manually because ``requests`` would otherwise
    follow them without giving us a chance to check where they lead.

    ``REQUEST_TIMEOUT`` budgets the whole chain rather than each hop, so a
    server cannot hold a worker thread for a multiple of it by chaining slow
    redirects. Tool handlers run on a small bounded thread pool (see
    RUN_BLOCKING_MAX_WORKERS), so that multiple matters.
    """
    assert_url_is_public(url)
    deadline = time.monotonic() + REQUEST_TIMEOUT

    for _ in range(MAX_REDIRECTS + 1):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise requests.exceptions.Timeout(
                f"Timed out following redirects for {url}"
            )

        response = requests.get(
            url,
            timeout=remaining,
            stream=True,
            allow_redirects=False,
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerPoint-MCP/1.0'
            }
        )
        if not response.is_redirect:
            response.raise_for_status()
            return response

        location = response.headers.get('Location')
        # Nothing reads a 3xx body. requests.get() closes its own session, so
        # this is belt-and-braces rather than a fix for an observed leak, but
        # it keeps the intent explicit if this ever moves to a shared Session.
        response.close()
        if not location:
            raise ImageDownloadError(f"Redirect without Location header from {url}")
        url = urljoin(url, location)
        assert_url_is_public(url)

    raise ImageDownloadError(f"Too many redirects (>{MAX_REDIRECTS}) downloading image")


def download_image(url: str) -> Tuple[io.BytesIO, str]:
    """Download an image from a URL and return it as a BytesIO object.

    Args:
        url: HTTP(S) URL of the image to download.

    Returns:
        Tuple of (BytesIO object containing image data, detected file extension).

    Raises:
        ImageDownloadError: If download fails.
        ImageValidationError: If image validation fails.
        SSRFProtectionError: If the URL, or any redirect it follows, resolves
            to a non-public address.
    """
    if not validate_url(url):
        raise ImageValidationError(f"Invalid URL format: {url}")

    logger.info(f"Downloading image from: {url}")

    try:
        response = _get_following_redirects(url)

    except requests.exceptions.Timeout:
        raise ImageDownloadError(f"Timeout downloading image from {url}")
    except requests.exceptions.ConnectionError:
        raise ImageDownloadError(f"Connection error downloading image from {url}")
    except requests.exceptions.HTTPError as e:
        raise ImageDownloadError(f"HTTP error {e.response.status_code} downloading image from {url}")
    except requests.exceptions.RequestException as e:
        raise ImageDownloadError(f"Error downloading image from {url}: {str(e)}")

    # Check content type
    content_type = response.headers.get('Content-Type', '').split(';')[0].strip().lower()
    if content_type and content_type not in ALLOWED_MIME_TYPES:
        raise ImageValidationError(
            f"Invalid image type: {content_type}. Allowed types: {', '.join(ALLOWED_MIME_TYPES)}"
        )

    # Check content length if provided
    content_length = response.headers.get('Content-Length')
    if content_length:
        try:
            size = int(content_length)
            if size > MAX_IMAGE_SIZE:
                raise ImageValidationError(
                    f"Image too large: {size / (1024*1024):.1f}MB. Maximum size: {MAX_IMAGE_SIZE / (1024*1024):.0f}MB"
                )
        except ValueError:
            pass  # Invalid Content-Length header, continue with download

    # Download image data
    image_data = io.BytesIO()
    total_size = 0

    for chunk in response.iter_content(chunk_size=8192):
        total_size += len(chunk)
        if total_size > MAX_IMAGE_SIZE:
            raise ImageValidationError(
                f"Image too large. Maximum size: {MAX_IMAGE_SIZE / (1024*1024):.0f}MB"
            )
        image_data.write(chunk)

    image_data.seek(0)

    # Determine file extension from content type or URL
    extension = get_image_extension(content_type, url)

    logger.info(f"Successfully downloaded image: {total_size / 1024:.1f}KB, type: {extension}")

    return image_data, extension


def get_image_extension(content_type: str, url: str) -> str:
    """Determine image file extension from content type or URL.

    Args:
        content_type: MIME type of the image.
        url: Original URL of the image.

    Returns:
        File extension (e.g., 'png', 'jpg').
    """
    # Try to get from content type
    type_to_ext = {
        'image/png': 'png',
        'image/jpeg': 'jpg',
        'image/jpg': 'jpg',
        'image/gif': 'gif',
        'image/bmp': 'bmp',
        'image/webp': 'webp',
        'image/tiff': 'tiff',
    }

    if content_type in type_to_ext:
        return type_to_ext[content_type]

    # Try to get from URL
    parsed = urlparse(url)
    path = parsed.path.lower()
    for ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tiff'):
        if path.endswith(f'.{ext}'):
            return 'jpg' if ext == 'jpeg' else ext

    # Default to png
    return 'png'




