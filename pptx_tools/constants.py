"""PowerPoint presentation constants.

This module contains layout indices, typography settings, colors, and margins
used throughout the PowerPoint generation.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


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

DEFAULT_TITLE_FONT_SIZE = Pt(32)
DEFAULT_SUBTITLE_FONT_SIZE = Pt(20)
DEFAULT_BODY_FONT_SIZE = Pt(18)
DEFAULT_CAPTION_FONT_SIZE = Pt(14)
DEFAULT_QUOTE_FONT_SIZE = Pt(28)


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
MARGIN_TOP = Inches(0.3)
MARGIN_BOTTOM = Inches(0.5)
TITLE_HEIGHT = Inches(0.8)
CONTENT_TOP = Inches(1.3)

