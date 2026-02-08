"""XML file creation tool.

This module provides functionality to validate XML content and save it as a file.
"""

import io
import logging
import xml.etree.ElementTree as ET
from typing import Tuple

from upload_tools import upload_file

logger = logging.getLogger(__name__)


class XMLValidationError(Exception):
    """Raised when XML content is invalid or incomplete."""
    pass


def validate_xml(xml_content: str) -> Tuple[bool, str]:
    """Validate that the provided string is well-formed XML.

    Args:
        xml_content: The XML content to validate.

    Returns:
        A tuple of (is_valid, error_message). If valid, error_message is empty.
    """
    try:
        # Try to parse the XML content
        ET.fromstring(xml_content)
        return True, ""
    except ET.ParseError as e:
        return False, f"XML parsing error: {str(e)}"
    except Exception as e:
        return False, f"Unexpected error during XML validation: {str(e)}"


def create_xml_file(xml_content: str) -> str:
    """Create an XML file from the provided XML content.

    Validates that the content is well-formed XML before saving.

    Args:
        xml_content: Complete, valid XML content string.

    Returns:
        A status message with the download URL or file path.

    Raises:
        XMLValidationError: If the XML content is invalid.
    """
    logger.info("Starting XML file creation")

    # Strip leading/trailing whitespace
    xml_content = xml_content.strip()

    # Validate the XML content
    is_valid, error_message = validate_xml(xml_content)
    if not is_valid:
        logger.error(f"XML validation failed: {error_message}")
        raise XMLValidationError(error_message)

    logger.debug("XML content validated successfully")

    # Ensure the content starts with XML declaration if not present
    if not xml_content.startswith('<?xml'):
        xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_content
        logger.debug("Added XML declaration to content")

    try:
        # Create a file-like object from the XML content
        xml_bytes = xml_content.encode('utf-8')
        file_object = io.BytesIO(xml_bytes)

        # Upload the file
        result = upload_file(file_object, "xml")
        logger.info("XML file uploaded successfully")
        return result

    except Exception as e:
        logger.error(f"Error creating XML file: {str(e)}", exc_info=True)
        return f"Error creating XML file: {str(e)}"

