"""Utility functions for handling images in PowerPoint presentations.

This module provides functionality to download and validate images from URLs
for embedding in PowerPoint slides.
"""

import io
import logging
from typing import Optional, Tuple
from urllib.parse import urlparse

import requests

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


class ImageDownloadError(Exception):
    """Exception raised when image download fails."""
    pass


class ImageValidationError(Exception):
    """Exception raised when image validation fails."""
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


def download_image(url: str) -> Tuple[io.BytesIO, str]:
    """Download an image from a URL and return it as a BytesIO object.

    Args:
        url: HTTP(S) URL of the image to download.

    Returns:
        Tuple of (BytesIO object containing image data, detected file extension).

    Raises:
        ImageDownloadError: If download fails.
        ImageValidationError: If image validation fails.
    """
    if not validate_url(url):
        raise ImageValidationError(f"Invalid URL format: {url}")

    logger.info(f"Downloading image from: {url}")

    try:
        response = requests.get(
            url,
            timeout=REQUEST_TIMEOUT,
            stream=True,
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerPoint-MCP/1.0'
            }
        )
        response.raise_for_status()

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


def get_image_dimensions(image_data: io.BytesIO) -> Optional[Tuple[int, int]]:
    """Try to get image dimensions without external dependencies.

    This is a basic implementation that reads PNG and JPEG headers.

    Args:
        image_data: BytesIO object containing image data.

    Returns:
        Tuple of (width, height) or None if cannot be determined.
    """
    try:
        image_data.seek(0)
        header = image_data.read(32)
        image_data.seek(0)

        # PNG
        if header[:8] == b'\x89PNG\r\n\x1a\n':
            width = int.from_bytes(header[16:20], 'big')
            height = int.from_bytes(header[20:24], 'big')
            return (width, height)

        # JPEG
        if header[:2] == b'\xff\xd8':
            image_data.seek(0)
            data = image_data.read()
            image_data.seek(0)

            # Find SOF marker
            i = 2
            while i < len(data) - 9:
                if data[i] == 0xFF:
                    marker = data[i + 1]
                    # SOF0, SOF1, SOF2 markers
                    if marker in (0xC0, 0xC1, 0xC2):
                        height = int.from_bytes(data[i + 5:i + 7], 'big')
                        width = int.from_bytes(data[i + 7:i + 9], 'big')
                        return (width, height)
                    elif marker == 0xD9:  # EOI
                        break
                    elif marker in (0xD0, 0xD1, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0x01):
                        i += 2
                    else:
                        length = int.from_bytes(data[i + 2:i + 4], 'big')
                        i += 2 + length
                else:
                    i += 1

        return None

    except Exception as e:
        logger.debug(f"Could not determine image dimensions: {e}")
        return None

