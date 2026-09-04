"""PowerPoint slide builder class.

Builds a deck from the typed slide models in :mod:`pptx_tools.schema`, using
the SlideHelpers mixin for text, tables, images and charts.

Anything the builder had to work around — an image that would not download, a
slide whose text will not fit, a layout the template does not provide — is
recorded on ``self.warnings`` and returned to the caller alongside the file.
Those situations used to be logged server-side only, so the model was told the
deck was fine and had no way to correct itself.
"""

import io
import copy
import logging
from typing import Any, Dict, List, Optional, Sequence

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

from template_utils import find_pptx_templates
from .constants import (
    TITLE_LAYOUT, SECTION_LAYOUT, CONTENT_LAYOUT,
    TWO_COLUMN_LAYOUT, TWO_COLUMN_TEXT_LAYOUT,
    DEFAULT_BODY_FONT_SIZE, DEFAULT_SUBTITLE_FONT_SIZE,
    DEFAULT_CAPTION_FONT_SIZE, DEFAULT_QUOTE_FONT_SIZE,
    TABLE_HEADER_FILL,
    DEFAULT_SLIDE_FORMAT, SLIDE_FORMAT_16_9, VALID_SLIDE_FORMATS,
)
from .helpers import (
    SlideHelpers,
    apply_autofit, body_to_bullets, estimate_text_fill, parse_table_data, resolve_fill, set_runs_language, table_overflows,
)
from .inline_formatting import needs_inline_processing, apply_inline_formatting
from .chart_utils import (
    add_chart_to_slide, add_scatter_to_slide, configure_data_labels,
    set_axis_titles, ChartDataError,
)
from .schema import coerce_slides

logger = logging.getLogger(__name__)


# Cache for loaded template paths (resolved once at first use)
_template_cache: Dict[str, Any] = {}


def _get_templates():
    """Get presentation templates for 4:3 and 16:9 formats (cached).

    Returns:
        Tuple of (path_4_3, path_16_9) template paths.
    """
    if "resolved" not in _template_cache:
        t43, t169 = find_pptx_templates()
        if not t43 or not t169:
            logger.info("One or more PPT templates missing; using PowerPoint defaults")
        _template_cache["4:3"] = t43
        _template_cache["16:9"] = t169
        _template_cache["resolved"] = True
    return _template_cache.get("4:3"), _template_cache.get("16:9")


class PowerpointPresentation(SlideHelpers):
    """Builder class for creating PowerPoint presentations from structured data."""

    def __init__(self, slides: Sequence[Any], format: str,
                 author: Optional[str] = None,
                 footer_text: Optional[str] = None,
                 show_slide_numbers: bool = False,
                 language: Optional[str] = None):
        """Initialize and build presentation.

        Args:
            slides: Slide models, or plain dicts in either the current or the
                previous key spelling — both are validated through
                :func:`~pptx_tools.schema.coerce_slides`.
            format: Presentation format ("4:3" or "16:9").
            author: Author name stored in document metadata/properties.
            footer_text: Optional footer text displayed on all slides.
            show_slide_numbers: Whether to show slide numbers on all slides.
            language: BCP-47 tag (e.g. "cs-CZ") stamped on every text run so
                the deck is proof-read in the right language.
        """
        if not slides:
            raise ValueError("At least one slide is required")

        self.warnings: List[str] = []
        self.slides = coerce_slides(slides)

        logger.info(
            "Initializing PowerPoint: slides=%d, format=%s", len(self.slides), format
        )

        self.presentation = self._create_presentation(format)
        self._footer_text = footer_text
        self._show_slide_numbers = show_slide_numbers
        self._language = language
        self._remove_template_slides()
        self._build_slides(self.slides)
        if footer_text or show_slide_numbers:
            self._apply_footer_and_slide_numbers()
        if language:
            self._apply_language(language)
        if author:
            self.presentation.core_properties.author = author

    # -------------------------------------------------------------------------
    # Setup
    # -------------------------------------------------------------------------

    def _warn(self, slide_index: int, message: str) -> None:
        """Record a caller-visible warning about one slide."""
        entry = f"slide {slide_index}: {message}"
        self.warnings.append(entry)
        logger.warning("[pptx] %s", entry)

    def _create_presentation(self, format: str) -> Presentation:
        """Create presentation with appropriate template."""
        if format not in VALID_SLIDE_FORMATS:
            logger.warning(
                "Unknown presentation format %r; using %s. Valid formats: %s",
                format, DEFAULT_SLIDE_FORMAT, ", ".join(VALID_SLIDE_FORMATS),
            )
            self.warnings.append(
                f"Unknown format {format!r}; used {DEFAULT_SLIDE_FORMAT}."
            )
            format = DEFAULT_SLIDE_FORMAT

        template_4_3, template_16_9 = _get_templates()
        template = template_16_9 if format == SLIDE_FORMAT_16_9 else template_4_3

        if template:
            try:
                return Presentation(template)
            except Exception as e:
                logger.error(f"Failed to load template: {e}")
                self.warnings.append(
                    f"Template could not be opened ({e}); used the built-in PowerPoint theme."
                )

        logger.warning(f"Using default PowerPoint template for {format}")
        return Presentation()

    def _remove_template_slides(self) -> None:
        """Remove every slide the template ships with, parts included.

        Dropping only the ``<p:sldId>`` entry leaves the slide's relationship in
        place, and python-pptx serialises parts by walking relationships — so a
        template carrying a sample slide shipped that slide's XML, text and
        images inside every generated file even though PowerPoint showed the
        right slide count. Dropping the relationship removes the part too.
        """
        sldIdLst = self.presentation.slides._sldIdLst
        prs_part = self.presentation.part

        removed = 0
        for sldId in list(sldIdLst):
            rId = sldId.rId
            try:
                sldIdLst.remove(sldId)
                prs_part.drop_rel(rId)
                removed += 1
            except Exception as e:
                logger.warning("Could not fully remove template slide %s: %s", rId, e)

        if removed:
            logger.debug("Removed %d slide(s) carried by the template", removed)

    def _build_slides(self, slides: Sequence[Any]) -> None:
        """Build all slides from validated models."""
        builders = {
            "title": self._build_title_slide,
            "section": self._build_section_slide,
            "content": self._build_content_slide,
            "table": self._build_table_slide,
            "image": self._build_image_slide,
            "two_column": self._build_two_column_slide,
            "chart": self._build_chart_slide,
            "scatter": self._build_scatter_slide,
            "quote": self._build_quote_slide,
        }

        logger.info("Building %d slides", len(slides))

        for i, slide in enumerate(slides):
            builder = builders.get(slide.type)
            if builder is None:  # pragma: no cover - schema forbids it
                raise ValueError(f"No builder for slide type {slide.type!r} at slide {i}")

            # `layout` is accepted now and honoured once template layouts are
            # resolved by name; say so rather than ignoring it silently.
            if slide.layout:
                self._warn(
                    i,
                    f"layout {slide.layout!r} was ignored: named layouts are not "
                    "resolved yet, the default layout for this slide type was used.",
                )

            try:
                logger.debug("Building slide %d: type=%s", i, slide.type)
                builder(slide, i)
            except Exception as e:
                logger.error("Failed to create slide %d: %s", i, e)
                raise ValueError(f"Error creating slide {i} ({slide.type}): {e}")

    # -------------------------------------------------------------------------
    # Slide Builders
    # -------------------------------------------------------------------------

    def _build_title_slide(self, slide_data, index: int) -> None:
        """Build a title slide with title and optional subtitle."""
        layout = self.presentation.slide_layouts[TITLE_LAYOUT]
        slide = self.presentation.slides.add_slide(layout)

        if len(slide.placeholders) > 0:
            slide.placeholders[0].text = slide_data.title or ""
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = slide_data.subtitle or ""

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_section_slide(self, slide_data, index: int) -> None:
        """Build a section divider slide."""
        layout = self.presentation.slide_layouts[SECTION_LAYOUT]
        slide = self.presentation.slides.add_slide(layout)

        if len(slide.placeholders) > 0:
            slide.placeholders[0].text = slide_data.title or ""

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_content_slide(self, slide_data, index: int) -> None:
        """Build a content slide with bullet points."""
        layout = self.presentation.slide_layouts[CONTENT_LAYOUT]
        slide = self.presentation.slides.add_slide(layout)

        if len(slide.placeholders) > 0:
            slide.placeholders[0].text = slide_data.title or ""

        bullets = body_to_bullets(slide_data.body)
        if bullets and len(slide.placeholders) > 1:
            placeholder = slide.placeholders[1]
            placeholder.text = ""
            self._fill_bullets(placeholder.text_frame, bullets)
            self._fit_text(placeholder, bullets, index)

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_table_slide(self, slide_data, index: int) -> None:
        """Build a table slide with a styled table."""
        slide, left, top, width, height = self._add_title_content_slide(slide_data.title or "")

        rows, col_alignments = parse_table_data(slide_data.rows)
        if not rows:
            self._warn(index, "table has no rows; the slide is empty.")
            return

        # An explicit `align` beats an inline markdown separator row.
        if slide_data.align:
            col_alignments = [
                {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[a]
                for a in slide_data.align
            ]

        header_fill = resolve_fill(slide_data.header_color, TABLE_HEADER_FILL)

        _, points = self._create_styled_table(
            slide,
            rows,
            left=left,
            top=top,
            width=width,
            height=height,
            header_color=header_fill,
            alternate_rows=slide_data.zebra,
            column_alignments=col_alignments,
            font_size=slide_data.font_size,
        )

        if points and table_overflows(len(rows), height, points):
            self._warn(
                index,
                f"table of {len(rows)} rows will not fit the content area even at "
                f"{points}pt; split it across slides.",
            )
        elif slide_data.font_size is None and points and points < int(DEFAULT_BODY_FONT_SIZE.pt):
            self._warn(
                index,
                f"table font reduced to {points}pt to fit {len(rows)} rows.",
            )

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_image_slide(self, slide_data, index: int) -> None:
        """Build a slide with an image from a URL or inline data URI."""
        slide, left, top, width, height = self._add_title_content_slide(slide_data.title or "")

        caption = slide_data.caption
        max_height = height - (Inches(0.6) if caption else 0)

        picture, error = self._add_image(
            slide, slide_data.source,
            left=left, top=top, max_width=width, max_height=max_height,
        )

        if picture and caption:
            self._add_text_box(
                slide, caption,
                left=left,
                top=picture.top + picture.height + Inches(0.1),
                width=width,
                height=Inches(0.5),
                font_size=DEFAULT_CAPTION_FONT_SIZE,
                italic=True,
                alignment=PP_ALIGN.CENTER,
            )
        elif not picture:
            self._add_image_placeholder(
                slide, "Image could not be loaded", left, top + Inches(1), width
            )
            self._warn(index, f"image could not be loaded ({error}); a placeholder was drawn.")

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_two_column_slide(self, slide_data, index: int) -> None:
        """Build a slide with two text columns using built-in PowerPoint layouts.

        Uses the Comparison layout when either column has a heading, otherwise
        Two Content.

        Placeholder indices:
        - Two Content (3): idx 0=Title, 1=Left content, 2=Right content
        - Comparison (4): idx 0=Title, 1=Left heading, 2=Left content,
          3=Right heading, 4=Right content
        """
        left_col, right_col = slide_data.left, slide_data.right
        has_headings = bool(left_col.heading or right_col.heading)

        layout_index = TWO_COLUMN_TEXT_LAYOUT if has_headings else TWO_COLUMN_LAYOUT
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[layout_index])

        left_bullets = body_to_bullets(left_col.body)
        right_bullets = body_to_bullets(right_col.body)

        if has_headings:
            content_slots = {2: left_bullets, 4: right_bullets}
            heading_slots = {1: left_col.heading, 3: right_col.heading}
        else:
            content_slots = {1: left_bullets, 2: right_bullets}
            heading_slots = {}

        for shape in slide.placeholders:
            idx = shape.placeholder_format.idx

            if idx == 0:
                if slide_data.title:
                    shape.text = slide_data.title
            elif idx in heading_slots:
                if heading_slots[idx]:
                    shape.text = heading_slots[idx]
            elif idx in content_slots:
                bullets = content_slots[idx]
                if bullets:
                    self._fill_bullets(shape.text_frame, bullets)
                    self._fit_text(shape, bullets, index)

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_chart_slide(self, slide_data, index: int) -> None:
        """Build a slide with a category chart."""
        slide, left, top, width, height = self._add_title_content_slide(slide_data.title or "")

        chart_data = {
            "categories": slide_data.categories,
            "series": [
                {"name": s.name, "values": s.values} for s in slide_data.series
            ],
        }

        try:
            chart = add_chart_to_slide(
                slide,
                chart_type=slide_data.chart_type,
                chart_data=chart_data,
                left=left, top=top, width=width, height=height,
                has_legend=slide_data.legend != "none",
                legend_position=slide_data.legend,
                title=slide_data.chart_title,
            )
            configure_data_labels(chart, slide_data.data_labels, slide_data.number_format)
            set_axis_titles(chart, slide_data.x_title, slide_data.y_title)
        except ChartDataError as e:
            logger.error(f"Chart error: {e}")
            self._add_text_box(
                slide, f"[Chart error: {e}]",
                left, top, width, Inches(1), alignment=PP_ALIGN.CENTER,
            )
            self._warn(index, f"chart could not be built ({e}).")
            self._add_speaker_notes(slide, slide_data.notes)
            return

        # A series whose length disagrees with the categories is accepted by
        # python-pptx but renders with gaps, so say so rather than let it pass.
        for series in slide_data.series:
            if len(series.values) != len(slide_data.categories):
                self._warn(
                    index,
                    f"series {series.name!r} has {len(series.values)} values for "
                    f"{len(slide_data.categories)} categories.",
                )

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_scatter_slide(self, slide_data, index: int) -> None:
        """Build a slide with an XY (scatter) chart."""
        slide, left, top, width, height = self._add_title_content_slide(slide_data.title or "")

        try:
            add_scatter_to_slide(
                slide,
                series=slide_data.series,
                left=left, top=top, width=width, height=height,
                legend=slide_data.legend,
                title=slide_data.chart_title,
                x_title=slide_data.x_title,
                y_title=slide_data.y_title,
            )
        except ChartDataError as e:
            logger.error(f"Scatter chart error: {e}")
            self._add_text_box(
                slide, f"[Chart error: {e}]",
                left, top, width, Inches(1), alignment=PP_ALIGN.CENTER,
            )
            self._warn(index, f"scatter chart could not be built ({e}).")

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_quote_slide(self, slide_data, index: int) -> None:
        """Build a quote/citation slide."""
        slide, left, top, width, height = self._add_title_content_slide(slide_data.title or "")

        quote_box = slide.shapes.add_textbox(left, top, width, height)
        tf = quote_box.text_frame
        tf.word_wrap = True

        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        formatted_quote = f'"{slide_data.text}"'
        if needs_inline_processing(formatted_quote):
            apply_inline_formatting(para, formatted_quote,
                                    font_size=DEFAULT_QUOTE_FONT_SIZE, italic=True)
        else:
            para.text = formatted_quote
            para.font.size = DEFAULT_QUOTE_FONT_SIZE
            para.font.italic = True

        if slide_data.attribution:
            author_para = tf.add_paragraph()
            author_para.text = f"— {slide_data.attribution}"
            author_para.font.size = DEFAULT_SUBTITLE_FONT_SIZE
            author_para.font.bold = True
            author_para.alignment = PP_ALIGN.CENTER
            author_para.space_before = Pt(24)

        self._add_speaker_notes(slide, slide_data.notes)

    # -------------------------------------------------------------------------
    # Fit, language, footer
    # -------------------------------------------------------------------------

    def _fit_text(self, placeholder, bullets, index: int) -> None:
        """Ask PowerPoint to shrink overfull body text, and warn when it is far gone."""
        fill = estimate_text_fill(
            bullets, placeholder.width, placeholder.height,
            font_size_pt=float(DEFAULT_BODY_FONT_SIZE.pt),
        )
        apply_autofit(placeholder.text_frame, scale=(1.0 / fill) if fill > 1.0 else None)

        # Below the shrink floor the slide is genuinely overfull; shrinking
        # further would produce text nobody can read from a room.
        if fill > 1.0 / 0.6:
            self._warn(
                index,
                f"body text is about {fill:.1f}x the space available and will be "
                "shrunk to fit; consider splitting it across slides.",
            )

    def _apply_language(self, language: str) -> None:
        """Stamp the proofing language on every run in the deck."""
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    set_runs_language(shape.text_frame, language)
                elif getattr(shape, "has_table", False):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            set_runs_language(cell.text_frame, language)
            if slide.has_notes_slide:
                set_runs_language(slide.notes_slide.notes_text_frame, language)

        logger.debug("Applied language %s to all runs", language)

    def _apply_footer_and_slide_numbers(self) -> None:
        """Apply footer text and/or slide numbers to all slides.

        Clones the footer/slide-number placeholder shapes from each slide's
        layout into the slide itself so they become visible. Assigns unique
        shape IDs to avoid PPTX corruption.
        """
        from xml.sax.saxutils import escape as xml_escape

        missing_footer = 0

        for slide in self.presentation.slides:
            layout = slide.slide_layout
            spTree = slide.shapes._spTree

            # Determine next available shape ID on this slide by collecting the
            # id attributes of the tree's direct children.
            existing_ids = set()
            for sp in spTree:
                cNvPr = sp.find('.//' + qn('p:cNvPr'))
                if cNvPr is None:
                    cNvPr = sp.find('.//' + qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                if cNvPr is not None and cNvPr.get('id'):
                    existing_ids.add(int(cNvPr.get('id')))
            next_id = max(existing_ids, default=0) + 1

            layout_indices = {ph.placeholder_format.idx for ph in layout.placeholders}
            if self._footer_text and 11 not in layout_indices:
                missing_footer += 1

            for ph in layout.placeholders:
                idx = ph.placeholder_format.idx

                if idx == 11 and self._footer_text:  # FOOTER placeholder
                    sp = copy.deepcopy(ph._element)
                    cNvPr = sp.find(qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                    if cNvPr is not None:
                        cNvPr.set('id', str(next_id))
                        next_id += 1
                    txBody = sp.find(qn('p:txBody'))
                    if txBody is not None:
                        for p in txBody.findall(qn('a:p')):
                            txBody.remove(p)
                        ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                        safe_text = xml_escape(self._footer_text)
                        p_xml = f'<a:p {ns}><a:r><a:t>{safe_text}</a:t></a:r></a:p>'
                        txBody.append(parse_xml(p_xml))
                    spTree.append(sp)

                elif idx == 12 and self._show_slide_numbers:  # SLIDE_NUMBER placeholder
                    sp = copy.deepcopy(ph._element)
                    cNvPr = sp.find(qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                    if cNvPr is not None:
                        cNvPr.set('id', str(next_id))
                        next_id += 1
                    spTree.append(sp)

        if missing_footer:
            # Previously this just produced a deck with no footer and no
            # explanation of why the argument appeared to do nothing.
            self.warnings.append(
                f"footer_text was dropped on {missing_footer} slide(s): their layout "
                "has no footer placeholder."
            )

        logger.debug("Applied footer/slide numbers to all slides")

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------

    def save(self) -> io.BytesIO:
        """Save presentation to a BytesIO object.

        Returns:
            BytesIO containing the presentation.

        Raises:
            RuntimeError: If saving fails.
        """
        logger.info("Saving PowerPoint to memory buffer")
        try:
            buffer = io.BytesIO()
            self.presentation.save(buffer)
            buffer.seek(0)
            return buffer
        except Exception as e:
            logger.error("Failed to save PowerPoint presentation: %s", e, exc_info=True)
            raise RuntimeError(f"Failed to save presentation: {e}") from e
