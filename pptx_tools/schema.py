"""Typed slide schema for the PowerPoint tool.

Why this exists
---------------
The tool used to take ``slides: List[dict]`` and describe the whole contract in
a paragraph of prose. Nothing was validated: a typo such as ``text`` instead of
``slide_text`` produced an empty slide and a success response, and an unknown
field was silently ignored. Modelling each slide type and joining them with a
discriminated union gives the client a real JSON schema (``oneOf`` plus a
``discriminator``), and rejects bad input with a path — ``slides.3.chart.series
.0.values.2`` — before any rendering happens.

Design notes
------------
* **Strict structure, loose values.** ``extra="forbid"`` turns a misspelled key
  into an error instead of a silently missing element, while table cells accept
  ``str | int | float | bool | None`` and colours accept ``#RRGGBB``,
  ``RRGGBB`` or a theme name, because those are the shapes a model actually
  emits.
* **Markdown where the model already thinks in markdown.** ``body`` takes
  either a markdown bullet string or an explicit list of :class:`Bullet`.
* **Backwards compatible.** :func:`migrate_legacy_slide` maps the previous key
  names (``slide_type``, ``slide_title``, ``slide_text``,
  ``indentation_level``, …) onto the new ones. It runs as a *before* validator
  on the list, ahead of union discrimination, so an old-style payload still
  validates. Deprecated keys are logged once per call.
"""

from __future__ import annotations

import logging
import re
from typing import Annotated, Any, Dict, List, Literal, Optional, Union

from pydantic import (
    BaseModel,
    BeforeValidator,
    ConfigDict,
    Field,
    TypeAdapter,
    ValidationError,
    field_validator,
)

from .constants import MAX_INDENT_LEVEL, TABLE_FONT_SIZE_RANGE

logger = logging.getLogger(__name__)


# =============================================================================
# Scalars
# =============================================================================

# 6-digit hex with or without '#', or one of the template's theme colours.
THEME_COLORS = (
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "dark1", "dark2", "light1", "light2",
)
_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")

CellValue = Union[str, int, float, bool, None]


def _validate_color(value: Optional[str]) -> Optional[str]:
    """Accept ``#RRGGBB``, ``RRGGBB`` or a theme colour name."""
    if value is None:
        return None
    text = str(value).strip()
    if _HEX_RE.match(text) or text.lower() in THEME_COLORS:
        return text
    raise ValueError(
        f"{value!r} is not a colour: use 6-digit hex ('#4172C4' or '4172C4') "
        f"or a theme name ({', '.join(THEME_COLORS)})"
    )


Color = Annotated[str, BeforeValidator(_validate_color)]


# =============================================================================
# Building blocks
# =============================================================================

def _clamp_level(value: Any) -> Any:
    """Clamp an indent level into range instead of rejecting it.

    Structure is strict here but values are not: a model that emits level 7,
    or the string "2", meant a deep bullet, not a failed presentation. The
    declared bounds stay on the field so the JSON schema still advertises the
    supported range.
    """
    if value is None:
        return 1
    try:
        level = int(value)
    except (TypeError, ValueError):
        logger.warning("[pptx] Invalid bullet level %r; using 1", value)
        return 1
    return max(1, min(level, MAX_INDENT_LEVEL))


class Bullet(BaseModel):
    """One bullet line. ``level`` is 1-based; 1 is the outermost."""

    model_config = ConfigDict(extra="forbid")

    text: str = Field(description="Bullet text. Inline markdown is supported.")
    level: Annotated[int, BeforeValidator(_clamp_level)] = Field(
        default=1, ge=1, le=MAX_INDENT_LEVEL,
        description=f"Indent depth, 1 (outermost) to {MAX_INDENT_LEVEL}.",
    )


# A body is either a markdown bullet string or explicit bullets.
Body = Union[str, List[Bullet]]


class Column(BaseModel):
    """One side of a two-column slide."""

    model_config = ConfigDict(extra="forbid")

    heading: Optional[str] = Field(default=None, description="Optional column subheading.")
    body: Body = Field(default_factory=list, description="Markdown bullets, or explicit bullet objects.")


class Kpi(BaseModel):
    """One headline figure on a KPI slide."""

    model_config = ConfigDict(extra="forbid")

    value: str = Field(description="The figure itself, e.g. '€4.2M' or '18%'. Kept as text so units and symbols survive.")
    label: str = Field(description="What the figure measures, e.g. 'ARR' or 'Churn'.")
    delta: Optional[str] = Field(
        default=None,
        description="Optional change, e.g. '+12% vs Q2'. Rendered in the accent colour.",
    )


class Step(BaseModel):
    """One step of a timeline or process slide."""

    model_config = ConfigDict(extra="forbid")

    label: str = Field(description="Short step name, e.g. 'Discovery' or 'Q1'.")
    detail: Optional[str] = Field(default=None, description="One line of detail under the label.")


class Series(BaseModel):
    """One data series of a category chart."""

    model_config = ConfigDict(extra="forbid")

    name: str = Field(description="Series name, shown in the legend.")
    values: List[Optional[float]] = Field(
        description="One number per category. Use null for a gap."
    )


class XySeries(BaseModel):
    """One data series of a scatter (XY) chart."""

    model_config = ConfigDict(extra="forbid")

    name: str = Field(description="Series name, shown in the legend.")
    points: List[List[float]] = Field(
        description="List of [x, y] pairs, e.g. [[1, 4.5], [2, 6.1]]."
    )

    @field_validator("points")
    @classmethod
    def _pairs(cls, points: List[List[float]]) -> List[List[float]]:
        for i, point in enumerate(points):
            if len(point) != 2:
                raise ValueError(f"point {i} must be exactly [x, y], got {len(point)} values")
        return points


# =============================================================================
# Slides
# =============================================================================

class SlideBase(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: Optional[str] = Field(default=None, description="Slide title.")
    notes: Optional[str] = Field(default=None, description="Speaker notes.")
    layout: Optional[str] = Field(
        default=None,
        description="Layout name from the active template, overriding the default for this slide type.",
    )


class TitleSlide(SlideBase):
    type: Literal["title"]
    subtitle: Optional[str] = Field(default=None, description="Subtitle: author, tagline, date…")


class SectionSlide(SlideBase):
    type: Literal["section"]


class ContentSlide(SlideBase):
    type: Literal["content"]
    body: Body = Field(default_factory=list, description="Markdown bullets, or explicit bullet objects.")


class TableSlide(SlideBase):
    type: Literal["table"]
    rows: List[List[CellValue]] = Field(
        description="Rows of cells; the first row is the header. Numbers and nulls are accepted."
    )
    align: Optional[List[Literal["left", "center", "right"]]] = Field(
        default=None, description="Per-column alignment."
    )
    header_color: Optional[Color] = Field(default=None, description="Header fill colour.")
    zebra: bool = Field(default=True, description="Shade alternating rows.")
    font_size: Optional[int] = Field(
        default=None, ge=TABLE_FONT_SIZE_RANGE[0], le=TABLE_FONT_SIZE_RANGE[1],
        description="Cell font size in points.",
    )


class ChartSlide(SlideBase):
    type: Literal["chart"]
    body: Optional[Body] = Field(
        default=None,
        description="Optional takeaways placed beside the chart, which is then drawn at half width.",
    )
    chart_type: Literal[
        "bar", "bar_stacked", "column", "column_stacked",
        "line", "line_markers", "pie", "doughnut",
        "area", "area_stacked", "radar",
    ] = Field(description="Category chart type. For an XY chart use the 'scatter' slide type.")
    categories: List[str] = Field(description="Category axis labels.")
    series: List[Series] = Field(description="One or more data series.")
    legend: Literal["right", "left", "top", "bottom", "none"] = Field(
        default="right", description="Legend position, or 'none' to hide it."
    )
    data_labels: bool = Field(default=False, description="Print the value on each data point.")
    number_format: Optional[str] = Field(
        default=None, description="Excel number format for data labels, e.g. '#,##0' or '0.0%'."
    )
    chart_title: Optional[str] = Field(default=None, description="Title drawn inside the chart.")
    x_title: Optional[str] = Field(default=None, description="Category axis title.")
    y_title: Optional[str] = Field(default=None, description="Value axis title.")


class ScatterSlide(SlideBase):
    type: Literal["scatter"]
    series: List[XySeries] = Field(description="One or more series of [x, y] points.")
    legend: Literal["right", "left", "top", "bottom", "none"] = Field(default="right")
    chart_title: Optional[str] = Field(default=None)
    x_title: Optional[str] = Field(default=None, description="X axis title.")
    y_title: Optional[str] = Field(default=None, description="Y axis title.")


class ImageSlide(SlideBase):
    type: Literal["image"]
    source: str = Field(
        description="Image https URL, or a data URI (data:image/png;base64,...)."
    )
    caption: Optional[str] = Field(default=None, description="Caption under the image.")
    body: Optional[Body] = Field(
        default=None,
        description="Optional text beside the picture, which is then drawn at half width.",
    )


class TwoColumnSlide(SlideBase):
    type: Literal["two_column"]
    left: Column = Field(default_factory=Column, description="Left column.")
    right: Column = Field(default_factory=Column, description="Right column.")


class KpiSlide(SlideBase):
    type: Literal["kpi"]
    items: List[Kpi] = Field(
        min_length=1, max_length=6,
        description="Two to four figures read best; more than six will not fit.",
    )


class AgendaSlide(SlideBase):
    type: Literal["agenda"]
    items: Optional[List[str]] = Field(
        default=None,
        description=(
            "Agenda entries. Omit to generate them from the deck's own section "
            "slides, in order — the usual case, and it cannot drift from the deck."
        ),
    )


class ClosingSlide(SlideBase):
    type: Literal["closing"]
    subtitle: Optional[str] = Field(default=None, description="Line under the closing title.")
    contact: Optional[List[str]] = Field(
        default=None, description="Contact lines, e.g. a name, an email, a phone number."
    )


# ---------------------------------------------------------------------------
# Blank slide with positioned elements — the escape hatch
# ---------------------------------------------------------------------------

_POSITION_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*(%|in)?\s*$")


def _validate_position(value: Any) -> Any:
    """A length as inches (number, or "1.5in") or a share of the slide ("40%").

    Kept as written and resolved against the real slide size at build time;
    validation here is only that it can be resolved at all.
    """
    if isinstance(value, bool):
        raise ValueError("position must be a number or a string, not a boolean")
    if isinstance(value, (int, float)):
        if value < 0:
            raise ValueError("position must not be negative")
        return float(value)
    if isinstance(value, str) and _POSITION_RE.match(value):
        return value.strip()
    raise ValueError(
        f"position {value!r} must be inches (2, 1.5, '1.5in') or a percentage ('40%')"
    )


Position = Annotated[Union[float, str], BeforeValidator(_validate_position)]


class ElementBase(BaseModel):
    """One thing placed on a blank slide, at an explicit position."""
    model_config = ConfigDict(extra="forbid")

    x: Position = Field(description="Left edge: inches, or a percentage of the slide width.")
    y: Position = Field(description="Top edge: inches, or a percentage of the slide height.")
    w: Position = Field(description="Width: inches, or a percentage of the slide width.")
    h: Optional[Position] = Field(
        default=None,
        description="Height: inches or a percentage. Images keep their aspect ratio "
                    "within it; text boxes default to one inch; shapes default to their width.",
    )


class TextElement(ElementBase):
    kind: Literal["text"]
    text: str = Field(description="Text, with the same inline markdown as a bullet.")
    font_size: Optional[int] = Field(default=None, ge=6, le=120, description="Points.")
    align: Literal["left", "center", "right"] = Field(default="left")
    bold: bool = Field(default=False)


class ImageElement(ElementBase):
    kind: Literal["image"]
    source: str = Field(description="Image https URL, or a data URI (data:image/png;base64,...).")


class ShapeElement(ElementBase):
    kind: Literal["shape"]
    shape: Literal["rectangle", "rounded_rectangle", "ellipse", "chevron", "arrow"] = Field(
        default="rectangle"
    )
    fill: Optional[Color] = Field(default=None, description="Fill: hex or a theme colour name.")
    text: Optional[str] = Field(default=None, description="Text centred inside the shape.")


Element = Annotated[
    Union[TextElement, ImageElement, ShapeElement],
    Field(discriminator="kind"),
]


class BlankSlide(SlideBase):
    """Positioned elements on an empty layout, for the slide no other type fits.

    Every other type decides where things go. This one takes coordinates and
    draws exactly what it is given, which is the right tool for a one-off
    layout and the wrong one for anything the typed slides can express.
    """
    type: Literal["blank"]
    elements: List[Element] = Field(
        min_length=1, max_length=20,
        description="Elements drawn in order; later ones sit on top.",
    )


class TimelineSlide(SlideBase):
    type: Literal["timeline"]
    steps: List[Step] = Field(
        min_length=2, max_length=6,
        description="Three to five steps read best; more than six will not fit across the slide.",
    )
    style: Literal["chevron", "box"] = Field(
        default="chevron", description="Chevrons imply sequence; boxes imply parallel items."
    )


class QuoteSlide(SlideBase):
    type: Literal["quote"]
    text: str = Field(description="The quotation, without surrounding quote marks.")
    attribution: Optional[str] = Field(default=None, description="Who said it.")


_SLIDE_MODELS = (
    TitleSlide, SectionSlide, ContentSlide, TableSlide,
    ChartSlide, ScatterSlide, ImageSlide, TwoColumnSlide, QuoteSlide,
    KpiSlide, AgendaSlide, ClosingSlide, TimelineSlide, BlankSlide,
)

SLIDE_TYPES = tuple(sorted(m.model_fields["type"].annotation.__args__[0] for m in _SLIDE_MODELS))

AnySlide = Annotated[
    Union[
        TitleSlide, SectionSlide, ContentSlide, TableSlide,
        ChartSlide, ScatterSlide, ImageSlide, TwoColumnSlide, QuoteSlide,
        KpiSlide, AgendaSlide, ClosingSlide, TimelineSlide, BlankSlide,
    ],
    Field(discriminator="type"),
]


# =============================================================================
# Legacy key migration
# =============================================================================
# The previous schema is still what existing client prompts emit. Map it onto
# the new keys BEFORE the union discriminates, since discrimination reads
# "type" and would otherwise reject every old-style slide outright.

_LEGACY_COMMON = {
    "slide_type": "type",
    "slide_title": "title",
    "speaker_notes": "notes",
}

_LEGACY_BY_TYPE = {
    "content": {"slide_text": "body"},
    "table": {"table_data": "rows", "alternate_rows": "zebra"},
    "image": {"image_url": "source", "image_caption": "caption"},
    "quote": {"quote_text": "text", "quote_author": "attribution"},
}


def _migrate_bullets(value: Any) -> Any:
    """Map ``indentation_level`` onto ``level`` and accept bare strings."""
    if not isinstance(value, list):
        return value
    out = []
    for item in value:
        if isinstance(item, str):
            out.append({"text": item})
        elif isinstance(item, dict):
            item = dict(item)
            if "indentation_level" in item and "level" not in item:
                item["level"] = item.pop("indentation_level")
            else:
                item.pop("indentation_level", None)
            out.append(item)
        else:
            out.append(item)
    return out


def migrate_legacy_slide(slide: Any, seen: Optional[set] = None) -> Any:
    """Return *slide* with any pre-Phase-1 keys renamed to their new names.

    Unrecognised keys are left untouched so ``extra="forbid"`` still reports
    them; this only translates the spellings the tool used to document.
    """
    if not isinstance(slide, dict):
        return slide

    out = dict(slide)
    used_legacy = []

    for old, new in _LEGACY_COMMON.items():
        if old in out:
            used_legacy.append(old)
            value = out.pop(old)
            out.setdefault(new, value)

    slide_type = out.get("type")

    for old, new in _LEGACY_BY_TYPE.get(slide_type, {}).items():
        if old in out:
            used_legacy.append(old)
            value = out.pop(old)
            out.setdefault(new, value)

    # Bullet lists carry their own legacy key.
    if slide_type == "content" and isinstance(out.get("body"), list):
        out["body"] = _migrate_bullets(out["body"])

    # two_column: left_heading/left_column -> left: {heading, body}
    if slide_type == "two_column":
        for side in ("left", "right"):
            heading_key, body_key = f"{side}_heading", f"{side}_column"
            if heading_key in out or body_key in out:
                used_legacy.extend(k for k in (heading_key, body_key) if k in out)
                column: Dict[str, Any] = {}
                heading = out.pop(heading_key, None)
                body = out.pop(body_key, None)
                if heading:
                    column["heading"] = heading
                if body is not None:
                    column["body"] = _migrate_bullets(body)
                out.setdefault(side, column)
        for side in ("left", "right"):
            value = out.get(side)
            if isinstance(value, dict) and isinstance(value.get("body"), list):
                value["body"] = _migrate_bullets(value["body"])

    # chart: chart_data{categories,series} -> categories/series; legend flags.
    if slide_type == "chart":
        chart_data = out.pop("chart_data", None)
        if isinstance(chart_data, dict):
            used_legacy.append("chart_data")
            if "categories" in chart_data:
                out.setdefault("categories", chart_data["categories"])
            if "series" in chart_data:
                out.setdefault("series", chart_data["series"])
        has_legend = out.pop("has_legend", None)
        legend_position = out.pop("legend_position", None)
        if has_legend is False:
            used_legacy.append("has_legend")
            # Contradictory input: "no legend" plus a position for it. The old
            # renderer tested has_legend first and never read the position, so
            # keep that precedence — but record the key that lost, or the
            # deprecation log claims to have accepted something it discarded.
            if legend_position:
                used_legacy.append("legend_position")
            out.setdefault("legend", "none")
        elif legend_position:
            used_legacy.append("legend_position")
            out.setdefault("legend", legend_position)
        elif has_legend is True:
            used_legacy.append("has_legend")

    if used_legacy and seen is not None:
        seen.update(used_legacy)

    return out


def _migrate_list(value: Any) -> Any:
    if not isinstance(value, list):
        return value
    seen: set = set()
    migrated = [migrate_legacy_slide(slide, seen) for slide in value]
    if seen:
        logger.info(
            "[pptx] Accepted deprecated slide keys (%s). The current names are "
            "documented in the tool description; support will be removed in a "
            "future release.",
            ", ".join(sorted(seen)),
        )
    return migrated


Slides = Annotated[List[AnySlide], BeforeValidator(_migrate_list)]

_SLIDES_ADAPTER: TypeAdapter = TypeAdapter(Slides)


# =============================================================================
# Entry point for non-MCP callers
# =============================================================================

def _describe_error(error: Dict[str, Any]) -> str:
    """Render one pydantic error as 'slide 2 -> rows.0: message'."""
    loc = [str(part) for part in error.get("loc", ())]
    # Drop the union-member tag pydantic injects so the path reads naturally.
    index = loc[0] if loc else "?"
    rest = [part for part in loc[1:] if part not in SLIDE_TYPES]
    where = f"slide {index}"
    if rest:
        where += " -> " + ".".join(rest)
    return f"{where}: {error.get('msg', 'invalid')}"


def coerce_slides(slides: Any) -> List[Any]:
    """Validate *slides* into typed models, raising a readable ``ValueError``.

    MCP callers never reach the error path — FastMCP validates against the tool
    signature first — but direct callers (tests, ``create_presentation``) do,
    and a raw pydantic dump is poor feedback for a model trying to correct
    itself.
    """
    if isinstance(slides, list) and slides and all(
        isinstance(slide, SlideBase) for slide in slides
    ):
        return list(slides)

    try:
        return _SLIDES_ADAPTER.validate_python(slides)
    except ValidationError as exc:
        problems = [_describe_error(error) for error in exc.errors()]
        # A wrong discriminator produces one error per slide; lead with the
        # valid types so the fix is obvious.
        hint = ""
        if any("does not match any of the expected tags" in problem for problem in problems):
            hint = f" Valid slide types: {', '.join(SLIDE_TYPES)}."
        raise ValueError(
            "Invalid slides: " + "; ".join(problems[:8]) + hint
        ) from exc
