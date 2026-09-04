"""The inline-markdown grammar shared by the Word and PowerPoint renderers.

Each renderer used to carry its own copy of the emphasis regex, and the two
had drifted: Word had no flanking rules, so ``5 * 3 * 2 = 30`` italicised the
3 and ``10 ~ 20 ~ 30`` subscripted the 20; PowerPoint had flanking but
mis-parsed a bold span ending in italics (``**a *b***``), and its escape
pattern swallowed the backslash before *any* character, turning ``C:\\new``
into ``C:new``. Fixing one copy left the other wrong.

The grammar lives here once. A renderer asks for the pattern with the spans
it can actually draw (Word can highlight; PowerPoint cannot) and dispatches
on the token shapes, which are identical for both. The token *shapes* are the
contract: ``**…**``, ``*…*``, ``~~…~~``, ``__…__``, `````…`````,
``^…^``, ``~…~``, ``==…==`` and ``[label](target)``.

Flanking follows CommonMark: an opening marker must be followed, and a
closing marker preceded, by something that is neither whitespace nor the
marker character itself. That is what keeps arithmetic and prose from being
read as emphasis. Nesting is stated structurally rather than with a closing
lookbehind, because a span whose last element is itself a nested span ends in
a marker character and a lookbehind would reject it.
"""

from __future__ import annotations

import re
import string

# A backslash escapes only the ASCII punctuation markdown uses as markers. It
# must not swallow the backslash before other characters: a literal "\n" is
# the two characters backslash and n, and r'\\(.)' collapsed it to a stray
# "n" — and corrupted "\t" and Windows paths like C:\new the same way.
ESCAPE_RE = re.compile(r"\\([" + re.escape(string.punctuation) + r"])")

# [label](target): the label is parsed for emphasis like any other segment;
# the target takes no nesting and no whitespace. Anchored, because a renderer
# applies it to a whole token the inline pattern has already isolated.
LINK_RE = re.compile(r"^\[([^\]\n]+)\]\(([^)\s]+)\)$")

# One nested italic unit, usable inside a bold span: flanked on both sides.
_NESTED_ITALIC = r"\*[^\s*][^*]*?(?<=[^\s*])\*"
# One nested bold unit, usable inside an italic span.
_NESTED_BOLD = r"\*\*[^*]+\*\*"

_BOLD_ITALIC = r"\*{3}(?=[^\s*])(?:[^*]|\*(?!\*{2}))+?(?<=[^\s*])\*{3}"
# Bold, allowing a nested *italic* — including as the very last thing before
# the closer, which a plain lookbehind would reject ("**a *b***").
_BOLD = (
    r"\*\*(?=[^\s*])(?:[^*]|" + _NESTED_ITALIC + r")*?"
    r"(?:[^\s*]|" + _NESTED_ITALIC + r")\*\*"
)
_STRIKE = r"~~(?=[^\s~]).+?(?<=[^\s~])~~"
_HIGHLIGHT = r"==(?=[^\s=]).+?(?<=[^\s=])=="
_UNDERLINE = r"__(?=[^\s_]).+?(?<=[^\s_])__"
# Italic, allowing a nested **bold** — same structural closer as bold.
_ITALIC = (
    r"\*(?=[^\s*])(?:[^*]|" + _NESTED_BOLD + r")*?"
    r"(?:[^\s*]|" + _NESTED_BOLD + r")\*"
)
_CODE = r"`[^`]+`"
_SUPERSCRIPT = r"\^(?=[^\s^])[^^]+?(?<=[^\s^])\^"
# Single tilde only: the strikethrough branch has already claimed "~~".
_SUBSCRIPT = r"~(?!~)(?=[^\s~])[^~]+?(?<=[^\s~])~"
_LINK = r"\[[^\]\n]+\]\([^)\s]+\)"


def build_inline_pattern(
    *, highlight: bool = False, superscript: bool = False, subscript: bool = False
) -> re.Pattern:
    """Compile the inline pattern with the spans a renderer can draw.

    The result has one capturing group around the whole alternation, so
    ``pattern.split(text)`` yields the tokens interleaved with plain text —
    the calling convention both renderers already use. Branch order matters:
    longer markers first, so ``***`` is not read as ``**`` + ``*``, and
    strikethrough before subscript so ``~~`` is never read as two ``~``.
    """
    branches = [_BOLD_ITALIC, _BOLD, _STRIKE]
    if highlight:
        branches.append(_HIGHLIGHT)
    branches += [_UNDERLINE, _ITALIC, _CODE]
    if superscript:
        branches.append(_SUPERSCRIPT)
    if subscript:
        branches.append(_SUBSCRIPT)
    branches.append(_LINK)
    return re.compile("(" + "|".join(branches) + ")")
