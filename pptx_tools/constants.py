"""PowerPoint presentation constants.

This module contains layout indices, typography settings, colors, and margins
used throughout the PowerPoint generation.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# =============================================================================
# Presentation Formats
# =============================================================================
# Single source of truth for the accepted aspect ratios. Previously the tool
# handler defaulted to "16:9" while _create_presentation_buffer() defaulted to
# "4:3", so a caller that omitted the argument got a different deck depending on
# which entry point it used.

SLIDE_FORMAT_4_3 = "4:3"
SLIDE_FORMAT_16_9 = "16:9"
VALID_SLIDE_FORMATS = (SLIDE_FORMAT_4_3, SLIDE_FORMAT_16_9)
DEFAULT_SLIDE_FORMAT = SLIDE_FORMAT_16_9


# =============================================================================
# Slide Layout Indices (PowerPoint default template)
# =============================================================================

TITLE_LAYOUT = 0           # Title Slide
CONTENT_LAYOUT = 1         # Title and Content
SECTION_LAYOUT = 2         # Section Header
TWO_COLUMN_LAYOUT = 3      # Two Content (no subheaders)
TWO_COLUMN_TEXT_LAYOUT = 4 # Comparison (with subheaders)
TITLE_ONLY_LAYOUT = 5      # Title Only
BLANK_LAYOUT = 6           # Blank


# =============================================================================
# Typography
# =============================================================================

DEFAULT_SUBTITLE_FONT_SIZE = Pt(20)
DEFAULT_BODY_FONT_SIZE = Pt(18)
DEFAULT_CAPTION_FONT_SIZE = Pt(14)
DEFAULT_QUOTE_FONT_SIZE = Pt(28)


# =============================================================================
# Bullet Indentation
# =============================================================================
# PowerPoint supports paragraph levels 0-8; the slide schema documents 1-3.
# Values outside 1..MAX_INDENT_LEVEL are clamped rather than rejected, because a
# model occasionally emits a deeper level than the prompt asked for.

MAX_INDENT_LEVEL = 5


# =============================================================================
# Table Colors
# =============================================================================

TABLE_HEADER_FILL = RGBColor(0x41, 0x72, 0xC4)
TABLE_HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ALT_ROW_FILL = RGBColor(0xE9, 0xEC, 0xEF)


# =============================================================================
# Margins and Dimensions
# =============================================================================

MARGIN_LEFT = Inches(0.5)

