"""Post-process LLM-generated draft bodies into clean Outlook-ready HTML.

Two responsibilities:

1. Convert markdown emphasis (**bold**, *italic*) and dash-prefixed bullet
   lines into HTML so Outlook renders them instead of showing literal
   asterisks / bullet characters.
2. Replace em-dashes and en-dashes with safer punctuation, and smart quotes
   with ASCII quotes. Em/en dashes are an obvious AI-authorship tell and
   smart quotes occasionally cause encoding glitches in older mail clients.

This module is the single source of truth for "what do we send to Graph as
HTML." Every outgoing draft / send path in graph_client.py routes through
``sanitize_draft_body`` before the wrapping div is applied, so the same
cleanup applies to smart_responder, bas_reminder, debtor_followup,
client_outreach, annual_review, engagement_letter — anything that ends up
calling ``GraphClient.create_draft`` / ``send_email``.

Idempotent — safe to run twice on the same body.
"""

from __future__ import annotations

import re

# --- bold/italic --------------------------------------------------------
# Order matters: bold (**) must be processed before italic (*) so the
# italic regex doesn't eat half of a bold marker.
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.DOTALL)
_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", re.DOTALL)

# Markdown bullet lines: leading "-" or "*" followed by a space.
_BULLET_LINE_RE = re.compile(r"^[\-\*]\s+(.+)$", re.MULTILINE)

# Heuristic: if the body already contains any of these block/inline tags
# we treat it as already-HTML and skip markdown→HTML conversion.
_LOOKS_HTML_RE = re.compile(
    r"<(p|div|ul|ol|li|strong|em|b|i|h[1-6]|br|table|tr|td)\b",
    re.IGNORECASE,
)


def strip_typographic_punctuation(text: str) -> str:
    """Replace em/en dashes and smart quotes with plain ASCII equivalents.

    Dash policy (changed from the previous all-to-hyphen behaviour):
      * en-dash between digits (e.g. ``2024–2025``)  → hyphen ``-``
      * em-dash anywhere                              → ``, `` (comma + space)
      * en-dash anywhere else                         → ``, `` (comma + space)
      * HTML entities ``&mdash;`` / ``&ndash;`` follow the em-dash rule.

    Smart quotes always normalise to straight ASCII quotes.

    Idempotent: calling twice produces the same result, since neither commas
    nor hyphens match the source patterns.
    """
    if not text:
        return text

    # Numeric ranges first so 2024–2025 → 2024-2025 (preserved as a range).
    text = re.sub(r"(\d)\s*–\s*(\d)", r"\1-\2", text)

    # Em-dash variants → comma. Collapse surrounding whitespace so
    # "word — word" and "word—word" both become "word, word".
    text = re.sub(r"\s*—\s*", ", ", text)
    text = re.sub(r"\s*&mdash;\s*", ", ", text)

    # Any remaining en-dashes → comma. (Numeric-range case handled above.)
    text = re.sub(r"\s*–\s*", ", ", text)
    text = re.sub(r"\s*&ndash;\s*", ", ", text)

    # Tidy doubled punctuation produced by the substitutions:
    #   "foo, ." → "foo."   "foo, ;" → "foo;"
    text = re.sub(r",\s*([.;:!?])", r"\1", text)
    # Collapse runs of internal spaces introduced by the comma insertion.
    text = re.sub(r"  +", " ", text)

    # Smart quotes → ASCII.
    smart_quotes = {
        "“": '"',   # left double
        "”": '"',   # right double
        "‘": "'",   # left single
        "’": "'",   # right single
        "&ldquo;": '"',
        "&rdquo;": '"',
        "&lsquo;": "'",
        "&rsquo;": "'",
    }
    for src, dst in smart_quotes.items():
        text = text.replace(src, dst)

    return text


def looks_like_already_html(text: str) -> bool:
    """Heuristic: if the body already has block/inline tags, skip the
    markdown→HTML pipeline (we'd just double-wrap)."""
    if not text:
        return False
    return bool(_LOOKS_HTML_RE.search(text))


def _markdown_to_html(text: str) -> str:
    """Convert the small subset of markdown the model actually produces."""
    text = _BOLD_RE.sub(r"<strong>\1</strong>", text)
    text = _ITALIC_RE.sub(r"<em>\1</em>", text)
    return text


def _bullets_to_html_list(text: str) -> str:
    """Convert contiguous '- item' / '* item' lines to a single <ul>.

    Inserts blank lines around the <ul> block so the later paragraph
    splitter treats the list as its own block instead of folding it into
    a surrounding <p>.
    """
    lines = text.split("\n")
    out: list[str] = []
    in_list = False
    for line in lines:
        m = _BULLET_LINE_RE.match(line)
        if m:
            if not in_list:
                if out and out[-1].strip() != "":
                    out.append("")
                out.append("<ul>")
                in_list = True
            out.append(f"  <li>{m.group(1)}</li>")
        else:
            if in_list:
                out.append("</ul>")
                out.append("")
                in_list = False
            out.append(line)
    if in_list:
        out.append("</ul>")
    return "\n".join(out)


def _paragraphs_to_html(text: str) -> str:
    """Wrap blank-line-separated blocks in <p>; leave block-level tags alone."""
    blocks = re.split(r"\n\s*\n", text.strip())
    out_blocks: list[str] = []
    for block in blocks:
        stripped = block.strip()
        if not stripped:
            continue
        if re.match(r"^\s*<(ul|ol|p|div|table|h[1-6])\b", stripped, re.IGNORECASE):
            out_blocks.append(stripped)
        else:
            with_breaks = stripped.replace("\n", "<br>\n")
            out_blocks.append(f"<p>{with_breaks}</p>")
    return "\n\n".join(out_blocks)


def sanitize_draft_body(raw: str) -> str:
    """Convert an LLM draft into Outlook-ready HTML.

    Pipeline:
      1. Strip em/en dashes and smart quotes (always — applies to HTML and
         markdown bodies alike, since they may have leaked in either form).
      2. If the body already looks like HTML, stop here. We don't want to
         re-wrap pre-formatted content from plugins that build HTML directly.
      3. Otherwise treat as loose markdown:
           a. Bullet lines → <ul><li>
           b. **bold** / *italic* → <strong>/<em>
           c. Blank-line-separated blocks → <p>
    """
    if not raw:
        return ""

    text = strip_typographic_punctuation(raw)

    if looks_like_already_html(text):
        return text

    text = _bullets_to_html_list(text)
    text = _markdown_to_html(text)
    text = _paragraphs_to_html(text)
    return text
