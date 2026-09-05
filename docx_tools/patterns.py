"""Compiled regex patterns and block-level markdown detection.

All block-level patterns are centralised here so that every module in the
docx_tools package can import them without circular dependencies.
"""

import re

from inline_markdown import ESCAPE_RE, LINK_RE, build_inline_pattern

# ---------------------------------------------------------------------------
# Block-level patterns (compiled once, used by many modules)
# ---------------------------------------------------------------------------

ORDERED_LIST_PATTERN = re.compile(r'^\d+\.\s+')
UNORDERED_LIST_PATTERN = re.compile(r'^[-*+]\s+')
# Capture variants used by process_list_items() to extract the item text.
# The ordered pattern captures the explicit number (group 1) so the renderer can
# restart numbering when "1." reappears at a level; group 2 is the item text.
# Item text is (.*) — matching the detection patterns above — so a marker with no
# text (e.g. "1." or "-") still captures (as empty) rather than failing the match.
ORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^(\d+)\.\s+(.*)')
UNORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^[-*+]\s+(.*)')
# Comment directive: <!-- key --> or <!-- key: value --> placed on its own line
# directly above the block it modifies. One mechanism for all block directives
# (borderless, widths, style, …). Group 1 = key, group 2 = optional value.
COMMENT_DIRECTIVE_PATTERN = re.compile(r'^<!--\s*([\w-]+)(?:\s*:\s*(.*?))?\s*-->$',
                                       re.IGNORECASE)
HEADING_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$')
PAGE_BREAK_PATTERN = re.compile(r'^-{3,}\s*$')
HORIZONTAL_LINE_PATTERN = re.compile(r'^\*{3,}\s*$')
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
TABLE_LINE_PATTERN = re.compile(r'^\|.+\|$')
# Fenced code block opener: 3+ backticks or tildes, optional info/language string.
# Group 1 is the fence run (its char/length identify the matching close).
CODE_FENCE_PATTERN = re.compile(r'^(`{3,}|~{3,})(.*)$')

# All block-level patterns checked by contains_block_markdown
_BLOCK_PATTERNS = [
    ORDERED_LIST_PATTERN, UNORDERED_LIST_PATTERN, HEADING_PATTERN,
    PAGE_BREAK_PATTERN, HORIZONTAL_LINE_PATTERN, IMAGE_PATTERN,
    TABLE_LINE_PATTERN, CODE_FENCE_PATTERN,
]

# ---------------------------------------------------------------------------
# Inline formatting patterns
# ---------------------------------------------------------------------------

# The emphasis grammar is shared with the PowerPoint renderer — see
# inline_markdown.py for the rules and why there is one copy. Word can draw
# every span, so it asks for all of them. These names are kept so the five
# modules that import them need not change.
_INLINE_FORMAT_RE = build_inline_pattern(highlight=True, superscript=True, subscript=True)
_LINK_RE = LINK_RE
_ESCAPE_RE = ESCAPE_RE
# Literal newline escape sequences ("\n", "\r\n", "\r" written as text) that LLMs
# often emit instead of a real newline. Longest alternative first so "\r\n"
# collapses to a single break rather than two.
_ESCAPED_NEWLINE_RE = re.compile(r'\\r\\n|\\[nr]')

# ---------------------------------------------------------------------------
# Alignment patterns
# ---------------------------------------------------------------------------

# Inline (single-line):  <center>text</center>  or  <div align="x">text</div>
_ALIGN_INLINE_RE = re.compile(
    r'^(?:<center>(.*)</center>'
    r'|<div\s+align="(right|center|justify|left)">(.*)</div>)$',
    re.IGNORECASE,
)
# Block open:  <center>  or  <div align="x">  (content on following lines)
_ALIGN_OPEN_RE = re.compile(
    r'^(?:<center>'
    r'|<div\s+align="(right|center|justify|left)">)\s*$',
    re.IGNORECASE,
)
# Block close:  </center>  or  </div>
_ALIGN_CLOSE_RE = re.compile(r'^</(?:center|div)>\s*$', re.IGNORECASE)

# ---------------------------------------------------------------------------
# Word field / header-footer patterns
# ---------------------------------------------------------------------------

_PAGE_TOKEN_RE = re.compile(r'(\{page}|\{pages})')

# HTML <br> tag variants (used for line breaks inside table cells etc.)
_BR_RE = re.compile(r'<br\s*/?>', re.IGNORECASE)


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def normalize_escaped_newlines(text: str) -> str:
    """Turn literal newline escape sequences into real newlines.

    LLMs frequently emit a newline as the two literal characters ``\\n`` (or
    ``\\r\\n``) inside a tool argument instead of a real line break — often
    because tool/argument descriptions demonstrate ``\\n`` as if it were syntax.
    Left untouched, the backslash handler would strip the slash and leave a
    stray ``n`` while the paragraph break is lost. Converting these sequences to
    real newlines up front makes a literal ``\\n`` behave exactly like a genuine
    newline (line/paragraph break) throughout the renderer.
    """
    if not text:
        return text
    return _ESCAPED_NEWLINE_RE.sub('\n', text)


def ordered_list_is_genuine(lines, idx) -> bool:
    """Return True if the ordered-list marker at ``lines[idx]`` should start a list.

    A numbered line begins an ordered list only when its number is ``1`` (a list
    may legitimately have a single item) OR a continuation follows — another
    sibling ordered item at the same indent, or a more-indented nested list item.
    This stops a standalone numbered line such as a date ("23. června 2026") from
    being misread as an ordered list.

    Note: a day-1 date ("1. června 2026") still matches the ``number == 1`` case
    and must be escaped ("1\\. června 2026") to render as prose; dates on days
    2–31 are handled automatically here.
    """
    raw = lines[idx]
    match = ORDERED_LIST_CAPTURE_PATTERN.match(raw.strip())
    if not match:
        return False
    if int(match.group(1)) == 1:
        return True
    base_indent = len(raw) - len(raw.lstrip())
    for nxt in lines[idx + 1:]:
        stripped = nxt.strip()
        if not stripped:
            return False  # blank line ends the run before any continuation
        indent = len(nxt) - len(nxt.lstrip())
        if indent > base_indent:
            # A more-indented list item nested under this one is a continuation.
            return bool(ORDERED_LIST_PATTERN.match(stripped)
                        or UNORDERED_LIST_PATTERN.match(stripped))
        if indent == base_indent:
            return bool(ORDERED_LIST_PATTERN.match(stripped))  # sibling ordered item
        return False  # dedent ends the run
    return False


def _segment_is_block(segments, idx) -> bool:
    """Return True if ``segments[idx]`` begins a block element.

    A heading or unordered-list marker always counts; an ordered-list marker
    counts only when :func:`ordered_list_is_genuine` accepts it within the
    *segments* list (so a lone number that is really a date does not).
    """
    seg = segments[idx]
    if HEADING_PATTERN.match(seg) or UNORDERED_LIST_PATTERN.match(seg):
        return True
    return bool(ORDERED_LIST_PATTERN.match(seg)
                and ordered_list_is_genuine(segments, idx))


def expand_br_to_block_breaks(text: str) -> str:
    """Promote ``<br>`` to a real newline when it borders block content.

    ``<br>`` is normally an *inline* soft line break (see
    :func:`docx_tools.inline_formatting.parse_inline_formatting`). But the
    block-level parser splits only on real newlines, so a list or heading that a
    model separated from surrounding text with ``<br>`` is never recognised — a
    ``"1. ..."`` after a ``<br>`` would render as literal text instead of a
    numbered list. This pre-pass splits a physical line on ``<br>`` *only* when
    one of the resulting segments is itself a block element (a heading, or a
    genuine ordered/unordered list item); plain-prose ``<br>`` is left untouched
    so it still renders as a within-paragraph soft break.

    Table rows (``| ... |``) and fenced code blocks are never touched: ``<br>``
    inside a table cell is an in-cell break handled by ``add_table_to_doc`` and
    code is taken verbatim. Idempotent — running it on already-expanded text is a
    no-op, so it is safe to call on both the base and placeholder paths.
    """
    if not text or not _BR_RE.search(text):
        return text
    out = []
    in_code = False
    for line in text.split('\n'):
        stripped = line.strip()
        if CODE_FENCE_PATTERN.match(stripped):
            in_code = not in_code
            out.append(line)
            continue
        if in_code or TABLE_LINE_PATTERN.match(stripped) or not _BR_RE.search(line):
            out.append(line)
            continue
        segments = [seg.strip() for seg in _BR_RE.split(line)]
        if len(segments) > 1 and any(_segment_is_block(segments, idx)
                                     for idx in range(len(segments))):
            out.extend(segments)
        else:
            out.append(line)
    return '\n'.join(out)


def contains_block_markdown(value: str) -> bool:
    """Return True if *value* contains block-level markdown content."""
    from .block_elements import detect_alignment  # deferred to avoid circular

    lines = value.split('\n')
    for idx, line in enumerate(lines):
        stripped = line.strip()
        for pattern in _BLOCK_PATTERNS:
            if not pattern.match(stripped):
                continue
            # A lone numbered line (e.g. a date "23. června 2026") is not a list
            # unless it starts at 1 or has a continuation — keep it inline prose.
            if pattern is ORDERED_LIST_PATTERN and not ordered_list_is_genuine(lines, idx):
                continue
            return True
        if detect_alignment(stripped) is not None:
            return True
    return False

