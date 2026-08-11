import io
import logging
from typing import List, Dict, Any, Optional

from upload_tools import upload_file
from .slide_builder import PowerpointPresentation

logger = logging.getLogger(__name__)


def _create_presentation_buffer(
    slides: List[Dict[str, Any]],
    format: str = "4:3",
    author: Optional[str] = None,
    footer_text: Optional[str] = None,
    show_slide_numbers: bool = False,
) -> io.BytesIO:
    """Create a PowerPoint presentation and return as BytesIO buffer.

    This function is useful when the caller needs to handle upload separately,
    such as for LibreChat file artifact uploads.

    :param slides: List of slide dicts with keys based on slide_type
    :param format: "4:3" or "16:9"
    :param author: Author name for document properties
    :param footer_text: Optional footer text displayed on all slides
    :param show_slide_numbers: Whether to show slide numbers
    :return: BytesIO buffer containing the PowerPoint file (position at start)
    """
    if not slides:
        raise ValueError("No slides provided")

    logger.info(f"Starting _create_presentation_buffer: slides={len(slides)}, format={format}")

    presentation = PowerpointPresentation(
        slides, format,
        author=author,
        footer_text=footer_text,
        show_slide_numbers=show_slide_numbers,
    )
    return presentation.save()


def create_presentation(
    slides: List[Dict[str, Any]],
    format: str = "4:3",
    file_name: Optional[str] = None,
    author: Optional[str] = None,
    footer_text: Optional[str] = None,
    show_slide_numbers: bool = False,
) -> str:
    """Create a PowerPoint presentation from structured slides and upload it.

    :param slides: List of slide dicts with keys based on slide_type
    :param format: "4:3" or "16:9"
    :param file_name: Optional custom filename (without extension)
    :param author: Author name for document properties
    :param footer_text: Optional footer text displayed on all slides
    :param show_slide_numbers: Whether to show slide numbers
    :return: Upload status or URL text
    """
    file_object = _create_presentation_buffer(
        slides, format,
        author=author,
        footer_text=footer_text,
        show_slide_numbers=show_slide_numbers,
    )

    try:
        text = upload_file(file_object, "pptx", filename=file_name)
    finally:
        file_object.close()

    logger.info("PowerPoint upload completed")
    return text

