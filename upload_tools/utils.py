import uuid


def generate_unique_object_name(suffix: str) -> str:
    """Generate a unique object name using UUID and preserve the file extension."""
    unique_id = str(uuid.uuid4())
    return f"{unique_id}.{suffix}"


def get_content_type(file_name: str) -> str:
    """Determine content type based on file extension.

    :param file_name: Name of the file
    :return: MIME type string
    :raises ValueError: If file type is unknown
    """
    if "pptx" in file_name:
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif "docx" in file_name:
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif "xlsx" in file_name:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif "eml" in file_name:
        return "application/octet-stream"
    else:
        raise ValueError("Unknown file type")

