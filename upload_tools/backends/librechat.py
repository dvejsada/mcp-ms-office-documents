"""LibreChat file service upload backend.

This module handles uploading generated documents to LibreChat's service
endpoint, which stores files and associates them with conversations.
The uploaded files appear as attachments in the chat UI.
"""

import io
import logging
import mimetypes
from typing import Optional

import httpx

from config import LibreChatSettings

logger = logging.getLogger(__name__)

# MIME type mapping for document extensions
MIME_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
    "eml": "message/rfc822",
    "xml": "application/xml",
    "txt": "text/plain",
    "md": "text/markdown",
    "json": "application/json",
    "csv": "text/csv",
}


def get_mime_type(filename: str) -> str:
    """Get MIME type for a filename.

    Args:
        filename: Filename with extension

    Returns:
        MIME type string
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext in MIME_TYPES:
        return MIME_TYPES[ext]
    # Fallback to mimetypes module
    mime_type, _ = mimetypes.guess_type(filename)
    return mime_type or "application/octet-stream"


async def upload_to_librechat(
    file_object: io.BytesIO,
    filename: str,
    user_context: dict,
    config: LibreChatSettings,
    timeout: float = 60.0,
) -> dict:
    """Upload a file to LibreChat's service endpoint.

    This function uploads a file to LibreChat and returns metadata that can
    be used to construct an MCP file artifact response.

    Args:
        file_object: BytesIO object containing the file data
        filename: Name of the file (with extension)
        user_context: Dict containing user information from request headers:
            - user_id: Required - LibreChat user ID (from X-User-Id header)
            - user_email: Optional - User email (from X-User-Email header)
            - conversation_id: Optional - Conversation ID (from X-Conversation-Id header)
        config: LibreChatSettings with service_url and service_token
        timeout: Request timeout in seconds (default 60)

    Returns:
        Dict with file metadata:
        {
            "file_id": "uuid-string",
            "filename": "document.docx",
            "filepath": "/path/to/file",
            "type": "application/vnd.openxmlformats-...",
            "bytes": 12345,
            "source": "local",
            "download_url": "/api/files/download/{user_id}/{file_id}"
        }

    Raises:
        ValueError: If user_context is missing required user_id
        RuntimeError: If upload fails
    """
    # Validate user context
    user_id = user_context.get("user_id")
    if not user_id:
        raise ValueError(
            "LibreChat upload requires user_id in user_context. "
            "Ensure X-User-Id header is passed from LibreChat."
        )

    # Prepare headers
    headers = {
        "X-Service-Token": config.service_token,
        "X-User-Id": user_id,
    }

    # Add optional headers
    if user_context.get("user_email"):
        headers["X-User-Email"] = user_context["user_email"]
    if user_context.get("conversation_id"):
        headers["X-Conversation-Id"] = user_context["conversation_id"]

    # Determine MIME type
    mime_type = get_mime_type(filename)

    logger.info(
        "Uploading file to LibreChat: %s (%s) for user %s",
        filename,
        mime_type,
        user_id,
    )

    # Ensure file object is at the beginning
    file_object.seek(0)
    file_size = len(file_object.getvalue())

    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            # Prepare multipart form data
            files = {
                "file": (filename, file_object, mime_type),
            }
            data = {
                "filename": filename,
            }

            response = await client.post(
                config.service_url,
                headers=headers,
                files=files,
                data=data,
            )

            # Check for errors
            if response.status_code == 401:
                raise RuntimeError(
                    "LibreChat upload failed: Invalid service token. "
                    "Verify LIBRECHAT_SERVICE_TOKEN matches MCP_SERVICE_TOKEN in LibreChat."
                )
            elif response.status_code == 400:
                error_detail = response.json().get("message", response.text)
                raise RuntimeError(f"LibreChat upload failed: {error_detail}")

            response.raise_for_status()

            result = response.json()

    except httpx.TimeoutException as e:
        logger.error("LibreChat upload timeout: %s", e)
        raise RuntimeError(
            f"LibreChat upload timed out after {timeout}s. "
            "Try increasing timeout or check network connectivity."
        ) from e
    except httpx.HTTPStatusError as e:
        logger.error("LibreChat upload HTTP error: %s", e)
        raise RuntimeError(f"LibreChat upload failed: {e}") from e
    except httpx.RequestError as e:
        logger.error("LibreChat upload request error: %s", e)
        raise RuntimeError(
            f"LibreChat upload failed: {e}. "
            f"Check LIBRECHAT_SERVICE_URL ({config.service_url}) is accessible."
        ) from e

    # Extract file info from response
    if not result.get("success", False):
        error_msg = result.get("error") or result.get("message") or "Unknown error"
        raise RuntimeError(f"LibreChat upload failed: {error_msg}")

    file_info = result.get("file", {})
    if not file_info.get("file_id"):
        raise RuntimeError(
            "LibreChat upload succeeded but response missing file_id. "
            f"Response: {result}"
        )

    logger.info(
        "File uploaded successfully to LibreChat: file_id=%s, filename=%s",
        file_info.get("file_id"),
        file_info.get("filename", filename),
    )

    # Construct download URL for LLM to use in responses
    file_id = file_info["file_id"]
    download_url = f"/api/files/download/{user_id}/{file_id}"

    # Return normalized file metadata
    return {
        "file_id": file_id,
        "filename": file_info.get("filename", filename),
        "filepath": file_info.get("filepath"),
        "type": file_info.get("type", mime_type),
        "bytes": file_info.get("bytes", file_size),
        "source": file_info.get("source", "local"),
        "download_url": download_url,
    }


def format_file_artifact(file_info: dict, text_message: Optional[str] = None):
    """Format file info as an MCP tool response with file artifact.

    Returns a ToolResult with structured_content containing the file artifact.
    LibreChat frontend expects this format for rendering file attachments.

    The structured_content contains:
    {
        "content": [
            {"type": "text", "text": "Success message"},
            {"type": "file", "file_id": "...", "filename": "...", "mimeType": "..."}
        ]
    }

    Args:
        file_info: Dict returned by upload_to_librechat()
        text_message: Optional text message to include with the file

    Returns:
        ToolResult with structured content for LibreChat
    """
    from fastmcp.server.server import ToolResult
    from mcp import types

    content_items = []

    # Add text content if provided
    if text_message:
        content_items.append({
            "type": "text",
            "text": text_message,
        })

    # Add file artifact - this format is expected by LibreChat frontend
    content_items.append({
        "type": "file",
        "file_id": file_info["file_id"],
        "filename": file_info["filename"],
        "mimeType": file_info.get("type", "application/octet-stream"),
    })

    # Return as ToolResult with structured_content
    # This bypasses FastMCP's content_and_artifact detection
    return ToolResult(
        content=[types.TextContent(type="text", text=text_message or "File created.")],
        structured_content={"content": content_items},
    )
