"""Utility functions for file upload operations.

This module provides filename sanitization and generation utilities
for safe file handling across different storage backends.

Security features:
- Path traversal prevention
- Unsafe character removal
- Filename length limits
"""

import re
import uuid


def sanitize_filename(name: str) -> str:
    """Sanitize a human-readable name into a safe filename component.

    Strips unsafe characters, replaces whitespace with underscores, and truncates
    to a reasonable length. Prevents path traversal attacks.

    Security measures:
    - Removes path separators (/, \\)
    - Removes path traversal sequences (.., .)
    - Removes null bytes and control characters
    - Removes characters unsafe for filenames/URLs
    - Truncates to 100 characters
    - Provides fallback for empty results

    :param name: Raw filename or title string
    :return: Sanitized string safe for use in object/blob names
    """
    if not name:
        return "document"
    
    # Remove null bytes and control characters first
    name = re.sub(r'[\x00-\x1f\x7f]', '', name)
    
    # Remove path separators to prevent path traversal
    name = name.replace('/', '').replace('\\', '')
    
    # Remove dot sequences that could be used for path traversal
    # This handles .., ..., etc.
    name = re.sub(r'\.{2,}', '', name)
    
    # Remove leading/trailing dots (hidden files, path tricks)
    name = name.strip('.')
    
    # Remove characters that are unsafe for filenames/URLs
    # Allow: word chars (a-z, A-Z, 0-9, _), whitespace, hyphen, single dot
    name = re.sub(r'[^\w\s\-.]', '', name)
    
    # Collapse whitespace to single underscores
    name = re.sub(r'\s+', '_', name.strip())
    
    # Collapse multiple underscores
    name = re.sub(r'_+', '_', name)
    
    # Collapse multiple dots
    name = re.sub(r'\.+', '.', name)
    
    # Remove leading/trailing underscores and dots
    name = name.strip('_.') 
    
    # Truncate to 100 chars to avoid overly long names
    name = name[:100]
    
    # Final safety check - if result is empty or just dots/underscores, use fallback
    if not name or name in ('.', '..', '_'):
        return "document"
    
    return name


def generate_unique_object_name(suffix: str) -> str:
    """Generate a unique object name using UUID and preserve the file extension.
    
    :param suffix: File extension without dot (e.g., 'docx', 'xlsx')
    :return: Unique filename like 'a1b2c3d4-e5f6-7890-abcd-ef1234567890.docx'
    """
    unique_id = str(uuid.uuid4())
    # Sanitize suffix to prevent injection
    safe_suffix = re.sub(r'[^a-zA-Z0-9]', '', suffix)[:10]
    return f"{unique_id}.{safe_suffix}"


def generate_named_object_name(
    filename: str, 
    suffix: str, 
    add_unique_prefix: bool = False
) -> str:
    """Generate an object name using a human-readable filename.
    
    :param filename: Human-readable filename (will be sanitized)
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :param add_unique_prefix: If True, adds 8-char UUID prefix for uniqueness.
        Default False - LibreChat handles uniqueness with its own UUID prefix.
    :return: Object name like 'My_Report.docx' or 'a1b2c3d4_My_Report.docx'
    """
    safe_name = sanitize_filename(filename)
    # Sanitize suffix to prevent injection
    safe_suffix = re.sub(r'[^a-zA-Z0-9]', '', suffix)[:10]
    
    if add_unique_prefix:
        short_id = uuid.uuid4().hex[:8]
        return f"{short_id}_{safe_name}.{safe_suffix}"
    return f"{safe_name}.{safe_suffix}"


def get_content_type(file_name: str) -> str:
    """Determine content type based on file extension.

    :param file_name: Name of the file
    :return: MIME type string
    :raises ValueError: If file type is unknown
    """
    file_name_lower = file_name.lower()
    
    if "pptx" in file_name_lower:
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif "docx" in file_name_lower:
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif "xlsx" in file_name_lower:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif "eml" in file_name_lower:
        return "application/octet-stream"
    elif "xml" in file_name_lower:
        return "application/xml"
    else:
        raise ValueError("Unknown file type")
