import io
import logging
from typing import List, Dict, Any, Tuple, Optional
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from template_utils import find_pptx_templates

# Slide layout constants
TITLE_LAYOUT = 2
SECTION_LAYOUT = 7
CONTENT_LAYOUT = 4
BLANK_LAYOUT = 6  # Blank layout for table slides

logger = logging.getLogger(__name__)


def load_templates() -> Tuple[Optional[str], Optional[str]]:
    """Resolve presentation templates (4:3, 16:9) from custom/default template dirs.

    Returns: tuple[str|None, str|None] -> (path_4_3, path_16_9)
    """
    t43, t169 = find_pptx_templates()
    if not t43 or not t169:
        logger.info("One or more PPT templates missing; will fall back to PowerPoint defaults where needed")
    return t43, t169


def parse_table(table_data: List[List[str]]) -> List[List[str]]:
    """Parse table data and return cleaned data without separator rows.

    Args:
        table_data: List of rows, where each row is a list of cell strings.

    Returns:
        Cleaned table data with separator rows removed.
    """
    if not table_data:
        return []

    cleaned_data = []
    for row in table_data:
        # Skip separator lines (contain dashes like ---, :-:, :--, --:)
        if row and all(
            '---' in cell or ':-:' in cell or ':--' in cell or '--:' in cell or cell.strip() == ''
            for cell in row
        ):
            continue
        cleaned_data.append(row)

    return cleaned_data


def parse_markdown_table(lines: List[str]) -> List[List[str]]:
    """Parse markdown table lines and return table data.

    Args:
        lines: List of markdown table lines (each line starts and ends with |).

    Returns:
        List of rows, where each row is a list of cell strings.
    """
    if not lines:
        return []

    table_data = []
    for line in lines:
        line = line.strip()
        if not line.startswith('|') or not line.endswith('|'):
            continue

        # Skip separator line (contains dashes)
        if '---' in line or ':-:' in line or ':--' in line or '--:' in line:
            continue

        # Split by | and clean up
        cells = [cell.strip() for cell in line.split('|')[1:-1]]  # Remove empty first/last
        table_data.append(cells)

    return table_data


class PowerpointPresentation:
    """Helper to build a PPTX presentation from structured slide dictionaries."""

    def __init__(self, slides: List[Dict[str, Any]], format: str):
        """Initialize PowerPoint presentation with slides and format."""
        logger.info(f"Initializing PowerPoint: slides={len(slides)}, format={format}")

        # Validate input
        if not slides:
            raise ValueError("At least one slide is required")

        # Load templates
        self.template_regular, self.template_wide = load_templates()
        logger.debug(f"Selected templates -> 4:3={self.template_regular}, 16:9={self.template_wide}")

        # Create presentation based on the format used
        try:
            if format == "4:3":
                if self.template_regular:
                    self.presentation = Presentation(self.template_regular)
                else:
                    self.presentation = Presentation()  # Use default template
                    logger.warning("No 4:3 template found, using PowerPoint default template")
            elif format == "16:9":
                if self.template_wide:
                    self.presentation = Presentation(self.template_wide)
                else:
                    self.presentation = Presentation()  # Use default template
                    logger.warning("No 16:9 template found, using PowerPoint default template")
            else:
                logger.warning(f"Unknown format '{format}', defaulting to 4:3")
                if self.template_regular:
                    self.presentation = Presentation(self.template_regular)
                else:
                    self.presentation = Presentation()  # Use default template
        except Exception as e:
            logger.error(f"Failed to load template: {e}")
            logger.info("Falling back to default PowerPoint template")
            self.presentation = Presentation()  # Fallback to default template

        # Remove default slide if it exists (some templates add one automatically)
        if len(self.presentation.slides) > 0:
            try:
                logger.debug("Removing default first slide from new presentation")
                slide_to_remove = self.presentation.slides[0]
                # Use underlying element removal to clear the initial slide
                self.presentation.slides.element.remove(slide_to_remove.element)
            except Exception as e:
                logger.debug(f"Could not remove default slide (non-fatal): {e}")

        # Create slides
        self._create_slides(slides)

    def _create_slides(self, slides: List[Dict[str, Any]]):
        """Create all slides from the slides data."""
        logger.info(f"Creating {len(slides)} slides")
        for i, slide in enumerate(slides):
            try:
                slide_type = slide.get("slide_type")
                logger.debug(f"Creating slide index={i}, type={slide_type}, title={slide.get('slide_title','')}")

                if slide_type == "content":
                    self.create_content_slide(slide)
                elif slide_type == "section":
                    self.create_section_slide(slide)
                elif slide_type == "title":
                    self.create_title_slide(slide)
                elif slide_type == "table":
                    self.create_table_slide(slide)
                else:
                    logger.warning(f"Unknown slide type '{slide_type}' for slide {i}, skipping")

            except Exception as e:
                logger.error(f"Failed to create slide {i}: {e}")
                raise ValueError(f"Error creating slide {i}: {str(e)}")

    def create_title_slide(self, slide: Dict[str, Any]):
        """Create a title slide."""
        try:
            title_layout = self.presentation.slide_layouts[TITLE_LAYOUT]
            title_slide = self.presentation.slides.add_slide(title_layout)
            logger.debug("Added title slide")

            # Set title
            if len(title_slide.placeholders) > 0:
                title_text = slide.get("slide_title", "")
                title_slide.placeholders[0].text = title_text
                logger.debug(f"Title slide title set: {title_text!r}")

            # Set author
            if len(title_slide.placeholders) > 1:
                author_text = slide.get("author", "")
                title_slide.placeholders[1].text = author_text
                logger.debug(f"Title slide author set: {author_text!r}")

        except Exception as e:
            logger.error(f"Failed to create title slide: {e}")
            raise

    def create_section_slide(self, slide: Dict[str, Any]):
        """Create a section slide."""
        try:
            section_layout = self.presentation.slide_layouts[SECTION_LAYOUT]
            section_slide = self.presentation.slides.add_slide(section_layout)
            logger.debug("Added section slide")

            # Set title
            if len(section_slide.placeholders) > 0:
                title_text = slide.get("slide_title", "")
                section_slide.placeholders[0].text = title_text
                logger.debug(f"Section slide title set: {title_text!r}")

        except Exception as e:
            logger.error(f"Failed to create section slide: {e}")
            raise

    def create_content_slide(self, slide: Dict[str, Any]):
        """Create a content slide with bullet points."""
        try:
            content_layout = self.presentation.slide_layouts[CONTENT_LAYOUT]
            content_slide = self.presentation.slides.add_slide(content_layout)
            logger.debug("Added content slide")

            # Set title
            if len(content_slide.placeholders) > 0:
                title_text = slide.get("slide_title", "")
                content_slide.placeholders[0].text = title_text
                logger.debug(f"Content slide title set: {title_text!r}")

            # Add content
            slide_text = slide.get("slide_text", [])
            if slide_text and len(content_slide.placeholders) > 1:
                logger.debug(f"Adding {len(slide_text)} bullet items to content slide")
                # Clear existing text
                content_slide.placeholders[1].text = ""

                # Add first paragraph
                first_item = slide_text[0]
                first_text = first_item.get("text", "")
                first_level = max(0, int(first_item.get("indentation_level", 1)) - 1)
                content_slide.placeholders[1].text_frame.paragraphs[0].text = first_text
                content_slide.placeholders[1].text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                content_slide.placeholders[1].text_frame.paragraphs[0].level = first_level
                logger.debug(f"Bullet[0]: level={first_level} text={first_text!r}")

                # Add remaining paragraphs
                for idx, paragraph_data in enumerate(slide_text[1:], start=1):
                    p = content_slide.placeholders[1].text_frame.add_paragraph()
                    p.text = paragraph_data.get("text", "")
                    p.alignment = PP_ALIGN.LEFT
                    level = max(0, int(paragraph_data.get("indentation_level", 1)) - 1)
                    p.level = level
                    logger.debug(f"Bullet[{idx}]: level={level} text={p.text!r}")

        except Exception as e:
            logger.error(f"Failed to create content slide: {e}")
            raise

    def create_table_slide(self, slide: Dict[str, Any]):
        """Create a table slide with a title and a table."""
        try:
            # Use blank layout for table slides
            blank_layout = self.presentation.slide_layouts[BLANK_LAYOUT]
            table_slide = self.presentation.slides.add_slide(blank_layout)
            logger.debug("Added table slide")

            # Get slide dimensions for positioning
            slide_width = self.presentation.slide_width
            slide_height = self.presentation.slide_height

            # Add title as a text box at the top
            title_text = slide.get("slide_title", "")
            if title_text:
                title_left = Inches(0.5)
                title_top = Inches(0.3)
                title_width = slide_width - Inches(1)
                title_height = Inches(0.8)

                title_shape = table_slide.shapes.add_textbox(
                    title_left, title_top, title_width, title_height
                )
                title_frame = title_shape.text_frame
                title_para = title_frame.paragraphs[0]
                title_para.text = title_text
                title_para.font.size = Pt(32)
                title_para.font.bold = True
                title_para.alignment = PP_ALIGN.LEFT
                logger.debug(f"Table slide title set: {title_text!r}")

            # Get table data
            table_data = slide.get("table_data", [])
            if not table_data:
                logger.warning("No table data provided for table slide")
                return

            # Clean the table data (remove separator rows if any)
            table_data = parse_table(table_data)
            if not table_data:
                logger.warning("Table data is empty after parsing")
                return

            # Calculate table dimensions
            num_rows = len(table_data)
            num_cols = max(len(row) for row in table_data) if table_data else 0

            if num_rows == 0 or num_cols == 0:
                logger.warning("Invalid table dimensions")
                return

            # Position table below title
            table_left = Inches(0.5)
            table_top = Inches(1.3)
            table_width = slide_width - Inches(1)
            table_height = slide_height - Inches(1.8)

            # Create the table
            shape = table_slide.shapes.add_table(
                num_rows, num_cols, table_left, table_top, table_width, table_height
            )
            table = shape.table
            logger.debug(f"Created table with {num_rows} rows and {num_cols} columns")

            # Populate table cells
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_text in enumerate(row_data):
                    if col_idx < num_cols:
                        cell = table.cell(row_idx, col_idx)
                        cell.text = str(cell_text) if cell_text else ""

                        # Style first row as header
                        if row_idx == 0:
                            cell.text_frame.paragraphs[0].font.bold = True

            logger.debug(f"Table populated with {num_rows} rows")

        except Exception as e:
            logger.error(f"Failed to create table slide: {e}")
            raise

    def save(self) -> io.BytesIO:
        """Save presentation to BytesIO object."""
        try:
            logger.info("Saving PowerPoint to memory buffer")
            file_like_object = io.BytesIO()
            self.presentation.save(file_like_object)
            file_like_object.seek(0)
            return file_like_object
        except Exception as e:
            logger.error(f"Failed to save presentation: {e}")
            raise

