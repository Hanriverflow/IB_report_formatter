"""
Converted Markdown cleanup for term-sheet input (opt-in MD -> MD).

External HWP/Word converters write a term sheet's chapter titles as bold
numbered sentences, turn an "appendix" band into a one-row table and leave the
cover (confidentiality label, title, date, logo, disclaimer) at the top of the
body. `clean_converted_term_sheet` rewrites such Markdown into `term-sheet`
input and reports every line it moves, rewrites, removes or keeps undecided.

The parser never promotes short numbered text to headings by default
(AGENTS.md), so this cleanup runs only when a user asks for it:
`md-format input.md output.md --converted-term-sheet`.

Changelog (converted input):
    - NEW: chapter headings from bold numbered lines or numbered lines followed
      by a table; one-row band tables become headings; cover lines before the
      first chapter move to frontmatter; unclear lines stay and are reported.
    - Only top-level blocks change: indented code, list and quote blocks
      (nested ones included), fences and whole HTML tables (balanced, comments
      and attributes skipped) are left as written, except that a one-row table
      with plain text (HTML: one `<p>` and formatting tags at most) can become a
      band heading; any ATX heading ends a paragraph and the cover.
    - Frontmatter follows the parser: `---` closers (`...` is rewritten and
      reported), case-insensitive keys, and a block-YAML rewrite (reported)
      for flow style or when new keys cannot be appended; escaped `\\<br>` is
      not a hard break; multi-line changes report their line range.
"""

import html
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import yaml

# ═══════════════════════════════════════════════════════════════════════════════
# PATTERNS
# ═══════════════════════════════════════════════════════════════════════════════

_FENCE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})")
_ATX_RE = re.compile(r"^ {0,3}#{1,6}(?:\s|$)")
_DELIMITER_CELL_RE = re.compile(r"^\s*:?-+:?\s*$")
_CELL_SPLIT_RE = re.compile(r"(?<!\\)\|")
_CHAPTER_RE = re.compile(r"^(\d{1,2})\\?\.\s+(.{1,40})$")
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_NUMBERED_H1_RE = re.compile(r"^#\s+\d{1,2}\\?\.\s")
_BR_TAG_RE = re.compile(r"<br\s*/?>$", re.IGNORECASE)
_IMAGE_RE = re.compile(
    r"!\[[^\]]*\]\((?:<[^>]*>|[^)\s]*)(?:\s+\"[^\"]*\")?\)(?:\{[^}]*\})?|<img\b[^>]*>",
    re.IGNORECASE,
)
_IMAGE_SOURCE_RE = re.compile(r"\]\(<?([^)\s>]*)|\bsrc\s*=\s*[\"']([^\"']*)", re.IGNORECASE)
_EMPHASIS_EDGE_RE = re.compile(r"^[*_\s]+|[*_\s]+$")
_CONFIDENTIAL_RE = re.compile(r"confidential|대외비", re.IGNORECASE)
_DATE_RE = re.compile(
    r"^\d{4}\s*(?:[.\-/]|년)\s*\d{1,2}\s*(?:[.\-/]|월)\s*(?:\d{1,2}\s*(?:일|\.)?)?\s*"
    r"(?:\([월화수목금토일]\))?$"
)
_CONTAINER_RE = re.compile(r"^ {0,3}(?:>|[-+*](?:\s|$))")  # block quote or bullet item
_ORDERED_ITEM_RE = re.compile(r"^ {0,3}\d{1,9}[.)](?:\s|$)")
_TABLE_OPEN_RE = re.compile(r"<table\b", re.IGNORECASE)
# Comments first, then whole tags (quoted attribute values may hold `>` or `</table>`).
_HTML_TOKEN_RE = re.compile(r"<!--.*?-->|<(/?)([A-Za-z][\w:-]*)(?:[^>\"']|\"[^\"]*\"|'[^']*')*>", re.DOTALL)
# Characters the inline parser may read as syntax (emphasis, LaTeX, super/subscript, code,
# links, escapes, HTML); a band heading holding any of them could change meaning.
_INLINE_SYNTAX_RE = re.compile(r"[<>*_`\[\]\\|$~^]|&[#\w]+;")
_BLOCK_KEY_RE = re.compile(r"^[^\s{\[#][^:]*:(?:\s|$)")  # `key:` at column 0 (block-style YAML)
_HTML_ROW_RE = re.compile(r"<tr\b", re.IGNORECASE)
_HTML_CELL_RE = re.compile(
    r"<t([dh])\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*>(.*?)</t\1\s*>", re.IGNORECASE | re.DOTALL
)
_HTML_PARAGRAPH_RE = re.compile(r"<p\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*>(.*)</p\s*>", re.IGNORECASE | re.DOTALL)
# Formatting-only tags whose text reads the same inside a bold heading; any other tag
# (breaks, super/subscript, links, blocks) keeps the table as written.
_BAND_INLINE_TAGS = frozenset({"span", "font", "b", "strong", "i", "em", "u"})
_TABLE_STRUCTURE_TAGS = frozenset({"table", "thead", "tbody", "tfoot", "tr", "colgroup", "col"})

_CONFIDENTIAL_MAX_CHARS = 40
_DISCLAIMER_MIN_CHARS = 80
_QUOTE_CHARS = 60
_FRONTMATTER_ORDER = ("profile", "title", "subtitle", "date", "confidential_label", "disclaimer")


# ═══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class CleanupNote:
    """One reported change or decision.

    Attributes:
        line: One-based first input line; 0 when not tied to a line.
        message: What was moved, rewritten, removed, kept or is missing.
        last: One-based last input line of a multi-line change; 0 for one line.
    """

    line: int
    message: str
    last: int = 0


@dataclass
class ConvertedCleanupReport:
    """Every change the cleanup made, plus the lines it left undecided."""

    notes: List[CleanupNote] = field(default_factory=list)

    def add(self, line: int, message: str, last: int = 0) -> None:
        """Record one note; lines are one-based, 0 for document-level notes."""
        self.notes.append(CleanupNote(line, message, last if last > line else 0))

    def lines(self) -> List[str]:
        """Printable report lines: line-bound notes in input order, then the rest."""
        ordered = sorted(self.notes, key=lambda note: (note.line == 0, note.line))
        return [
            (f"lines {note.line}-{note.last}: " if note.last else f"line {note.line}: ") + note.message
            if note.line else note.message
            for note in ordered
        ]


@dataclass
class _Block:
    """Consecutive source lines of one kind: text, pipe, html, fence or indented."""

    kind: str
    start: int  # zero-based index into the body lines
    lines: List[str]
    gap: List[str]  # blank lines before the block
    removed: List[bool] = field(default_factory=list)
    band: bool = False  # a one-row table kept as a table; it still ends the cover

    def __post_init__(self) -> None:
        self.removed = [False] * len(self.lines)

    @property
    def top_level(self) -> bool:
        """True when the block starts at column 0 (not a list continuation)."""
        return not self.lines[0][:1].isspace()

    @property
    def container(self) -> bool:
        """A block quote or bullet item, an ordered item with indented continuation
        lines, or any block holding an indented list item or quote.

        A lone ordered-looking line stays eligible: converters write dates
        (`2026. 9. 29.`) and chapter titles (`1. 개요`) that way.
        """
        first, rest = self.lines[0], self.lines[1:]
        if _CONTAINER_RE.match(first):
            return True
        if any(
            line[:1].isspace() and (_CONTAINER_RE.match(line.lstrip()) or _ORDERED_ITEM_RE.match(line.lstrip()))
            for line in rest
        ):
            return True
        return bool(_ORDERED_ITEM_RE.match(first)) and any(line[:1].isspace() for line in rest)


@dataclass
class _CoverMove:
    """Cover lines that become one frontmatter value, applied after the scan."""

    block: _Block
    indexes: Tuple[int, ...]
    key: str
    value: str


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════


def clean_converted_term_sheet(text: str) -> Tuple[str, ConvertedCleanupReport]:
    """Rewrite converter-shaped Markdown into term-sheet input.

    Only top-level blocks change. Chapter headings: a one-line paragraph
    `1. Title` (at most 40 characters after the number; `1\\.` too) becomes
    `## 1. Title` when it is bold or the next block is a table. A table with a
    header row and no body rows (an HTML one only when its cells hold plain
    text and nothing is outside them) becomes a `##` heading of its non-empty
    cells joined by ` | `. Blocks before the first heading of level 2 or
    more (or a numbered level-1 heading) are the cover; list and quote blocks
    there stay. The first
    confidentiality line becomes `confidential_label` (identical repeats are
    removed), the leading run of bold lines (or a level-1 heading) becomes
    `title` then `subtitle`, a date line `date`, paragraphs of 80 characters
    or more `disclaimer`; images are removed and reported for the house
    `style.logo`. Other lines stay and are reported. `prepared_by` is never
    inferred. Existing frontmatter lines and keys (case-insensitive) are kept,
    and a cover line whose key is already set stays in the body. Fenced and
    indented code is never changed.

    Args:
        text: Converted Markdown, with or without frontmatter.

    Returns:
        The cleaned Markdown and the report of every change.

    Raises:
        ValueError: Existing frontmatter is not valid YAML or not a mapping.
    """
    report = ConvertedCleanupReport()
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    front_lines, existing, offset = _read_frontmatter(lines, report)
    blocks, trailing = _split_blocks(lines[offset:])

    _promote_chapter_headings(blocks, offset, report)
    _convert_band_tables(blocks, offset, report)
    fields = _extract_cover(blocks, existing, offset, report)

    if "profile" not in existing:
        fields = {"profile": "term-sheet", **fields}
        report.add(0, 'frontmatter profile <- "term-sheet"')
    elif existing["profile"] != "term-sheet":
        report.add(0, f"frontmatter profile is {_quote(str(existing['profile']))}; term-sheet input needs \"term-sheet\"")
    _report_missing(existing, fields, report)

    front = _write_frontmatter(front_lines, fields, report)
    output = front + ([""] if front else []) + _join_blocks(blocks, trailing)
    return "\n".join(output).rstrip("\n") + "\n", report


# ═══════════════════════════════════════════════════════════════════════════════
# BLOCKS
# ═══════════════════════════════════════════════════════════════════════════════


def _read_frontmatter(
    lines: List[str], report: ConvertedCleanupReport
) -> Tuple[List[str], Dict[str, Any], int]:
    """Return the frontmatter lines (with fences), its lower-cased mapping and the body start.

    The parser only closes frontmatter with `---`, so a YAML `...` closer is
    rewritten to `---` and reported.
    """
    if not lines or lines[0].strip() != "---":
        return [], {}, 0
    for index in range(1, len(lines)):
        closer = lines[index].strip()
        if closer in ("---", "..."):
            try:
                data = yaml.safe_load("\n".join(lines[1:index])) or {}
            except yaml.YAMLError as exc:
                raise ValueError(f"existing frontmatter is not valid YAML: {exc}") from exc
            if not isinstance(data, dict):
                raise ValueError("existing frontmatter must be a YAML mapping")
            front = lines[: index + 1]
            if closer == "...":
                front[-1] = "---"
                report.add(index + 1, 'frontmatter closer "..." -> "---" (the parser closes frontmatter with ---)')
            existing = {str(key).strip().lower(): value for key, value in data.items() if key is not None}
            return front, existing, index + 1
    return [], {}, 0


def _indent(line: str) -> int:
    return len(line.expandtabs(4)) - len(line.expandtabs(4).lstrip())


def _is_pipe_row(line: str) -> bool:
    return _indent(line) <= 3 and line.lstrip().startswith("|")


def _row_cells(line: str) -> List[str]:
    cells = _CELL_SPLIT_RE.split(line.strip())
    if cells and not cells[0].strip():
        cells = cells[1:]
    if cells and not cells[-1].strip():
        cells = cells[:-1]
    return [cell.strip() for cell in cells]


def _is_delimiter_row(line: str) -> bool:
    cells = _row_cells(line)
    return "|" in line and bool(cells) and all(_DELIMITER_CELL_RE.match(cell) for cell in cells)


def _starts_pipe_table(lines: List[str], index: int) -> bool:
    return _is_pipe_row(lines[index]) and index + 1 < len(lines) and _is_delimiter_row(lines[index + 1])


def _starts_html_table(line: str) -> bool:
    return _indent(line) <= 3 and line.lstrip().lower().startswith("<table")


def _closes_fence(line: str, marker: str) -> bool:
    stripped = line.strip()
    return len(stripped) >= len(marker) and set(stripped) == {marker[0]} and _indent(line) <= 3


def _html_table_end(lines: List[str], index: int) -> int:
    """Exclusive end line of the balanced HTML table opening at `index`.

    Comments and quoted attribute values are skipped; an unclosed table runs
    to the end of the document.
    """
    text = "\n".join(lines[index:])
    depth = 0
    for token in _HTML_TOKEN_RE.finditer(text):
        if (token.group(2) or "").lower() != "table":
            continue
        depth += -1 if token.group(1) else 1
        if depth == 0:
            return index + text.count("\n", 0, token.end()) + 1
    return len(lines)


def _starts_block(lines: List[str], index: int) -> bool:
    """A line that interrupts a paragraph: fence, ATX heading or table start."""
    line = lines[index]
    return bool(
        _FENCE_RE.match(line) or _ATX_RE.match(line) or _starts_html_table(line)
        or _starts_pipe_table(lines, index)
    )


def _split_blocks(lines: List[str]) -> Tuple[List[_Block], List[str]]:
    """Split body lines into blocks; returns the blocks and the trailing blank lines."""
    blocks: List[_Block] = []
    gap: List[str] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        if not line.strip():
            gap.append(line)
            index += 1
            continue
        fence = _FENCE_RE.match(line)
        end = index + 1
        if fence:
            while end < len(lines) and not _closes_fence(lines[end], fence.group(1)):
                end += 1
            kind, end = "fence", min(end + 1, len(lines))
        elif _indent(line) >= 4:  # indented code or a list continuation: never changed
            while end < len(lines) and lines[end].strip() and _indent(lines[end]) >= 4:
                end += 1
            kind = "indented"
        elif _starts_html_table(line):
            kind, end = "html", _html_table_end(lines, index)
        elif _starts_pipe_table(lines, index):
            end = index + 2
            while end < len(lines) and _is_pipe_row(lines[end]):
                end += 1
            kind = "pipe"
        else:
            if not _ATX_RE.match(line):  # a heading is a block of its own
                # A quote or bullet line at column 0 starts its own (container)
                # block, which keeps following quote/bullet lines and lazy
                # continuations; indented ones stay with the block they continue.
                in_container = bool(_CONTAINER_RE.match(line))
                while (
                    end < len(lines)
                    and lines[end].strip()
                    and not _starts_block(lines, end)
                    and (in_container or lines[end][:1].isspace() or not _CONTAINER_RE.match(lines[end]))
                ):
                    end += 1
            kind = "text"
        blocks.append(_Block(kind, index, lines[index:end], gap))
        gap = []
        index = end
    return blocks, gap


def _join_blocks(blocks: List[_Block], trailing: List[str]) -> List[str]:
    """Reassemble kept lines; a removed block takes its preceding blank lines with it."""
    output: List[str] = []
    for block in blocks:
        kept = [line for line, removed in zip(block.lines, block.removed) if not removed]
        if kept:
            output.extend((block.gap if output else []) + kept)
    return output + trailing


def _span(block: _Block, offset: int) -> Tuple[int, int]:
    """One-based first and last input lines of a block."""
    first = offset + block.start + 1
    return first, first + len(block.lines) - 1


# ═══════════════════════════════════════════════════════════════════════════════
# HEADINGS
# ═══════════════════════════════════════════════════════════════════════════════


def _unbold(line: str) -> Optional[str]:
    """Inner text of a line that is one bold span, otherwise None."""
    stripped = line.strip()
    if len(stripped) > 4 and stripped.startswith("**") and stripped.endswith("**"):
        inner = stripped[2:-2].strip()
        if inner and "**" not in inner:
            return inner
    return None


def _promote_chapter_headings(blocks: List[_Block], offset: int, report: ConvertedCleanupReport) -> None:
    """`**1. Title**`, or `1. Title` right before a table, becomes `## 1. Title`."""
    for position, block in enumerate(blocks):
        if block.kind != "text" or len(block.lines) != 1 or not block.top_level:
            continue
        source = block.lines[0]
        inner = _unbold(source)
        match = _CHAPTER_RE.match(inner if inner is not None else source.strip())
        if match is None:
            continue
        following = blocks[position + 1] if position + 1 < len(blocks) else None
        before_table = following is not None and following.kind in ("pipe", "html") and following.top_level
        if inner is None and not before_table:
            # Left as list text; a chapter title the rules missed is worth a look.
            report.add(
                offset + block.start + 1,
                f"kept numbered line {_quote(source)}: not bold and no table follows; "
                "mark it as a heading if it is a chapter",
            )
            continue
        heading = f"## {match.group(1)}. {match.group(2).strip()}"
        block.lines[0] = heading
        reason = "bold numbered line" if inner is not None else "numbered line before a table"
        report.add(offset + block.start + 1, f"heading {_quote(source)} -> {_quote(heading)} ({reason})")


def _convert_band_tables(blocks: List[_Block], offset: int, report: ConvertedCleanupReport) -> None:
    """A top-level table with a header row and no body rows becomes a `##` heading."""
    for block in blocks:
        if not block.top_level:
            continue
        first, last = _span(block, offset)
        cells: Optional[List[str]] = None
        if block.kind == "pipe" and len(block.lines) == 2:
            cells = [_unbold(cell) or cell for cell in _row_cells(block.lines[0])]
        elif block.kind == "html":
            cells, why = _html_band_cells("\n".join(block.lines))
            if why:
                block.band = True
                report.add(first, f"kept a one-row HTML table: {why}", last)
                continue
        if cells is None:
            continue
        texts = [cell for cell in cells if cell]
        if not texts:
            block.band = True
            report.add(first, "kept a one-row table without text", last)
            continue
        heading = "## " + " | ".join(texts)
        if _INLINE_SYNTAX_RE.search(heading[3:].replace(" | ", " ")):
            # Joined cells (or decoded HTML) could open emphasis, LaTeX or markup in a heading.
            block.band = True
            report.add(first, "kept a one-row table: its text has Markdown or HTML characters", last)
            continue
        block.kind, block.lines, block.removed = "text", [heading], [False]
        report.add(first, f"band table -> {_quote(heading)}", last)


def _html_cell_text(content: str) -> Optional[str]:
    """Text of a cell holding at most one `<p>` and formatting-only tags, else None.

    Formatting tags join their text (`1<span>00</span>` is `100`); a break,
    superscript, link, image, comment or block tag would change the text as a
    heading, so such a cell is not converted.
    """
    inner = content.strip()
    paragraph = _HTML_PARAGRAPH_RE.fullmatch(inner)
    if paragraph and not re.search(r"<p\b", paragraph.group(1), re.IGNORECASE):
        inner = paragraph.group(1)
    for token in _HTML_TOKEN_RE.finditer(inner):
        if token.group(2) is None or token.group(2).lower() not in _BAND_INLINE_TAGS:
            return None
    return " ".join(html.unescape(_HTML_TOKEN_RE.sub("", inner)).split())


def _html_band_cells(source: str) -> Tuple[Optional[List[str]], str]:
    """Cell texts of a lone single-row HTML table.

    Returns:
        `(cells, "")` when it can become a heading; `(None, reason)` when it is
        a one-row table that must stay; `(None, "")` when it is not one.
    """
    if len(_TABLE_OPEN_RE.findall(source)) != 1 or len(_HTML_ROW_RE.findall(source)) != 1:
        return None, ""
    if not re.search(r"</table\s*>\s*$", source, re.IGNORECASE):
        return None, "text follows the table"
    outside = _HTML_CELL_RE.sub("", source)
    for token in _HTML_TOKEN_RE.finditer(outside):
        if token.group(2) is None or token.group(2).lower() not in _TABLE_STRUCTURE_TAGS:
            return None, "it has content outside its cells"
    if html.unescape(_HTML_TOKEN_RE.sub("", outside)).strip():
        return None, "it has text outside its cells"
    cells = []
    for _, content in _HTML_CELL_RE.findall(source):
        text = _html_cell_text(content)
        if text is None:
            return None, "a cell has markup other than simple formatting"
        cells.append(text)
    return cells or None, ""


# ═══════════════════════════════════════════════════════════════════════════════
# COVER
# ═══════════════════════════════════════════════════════════════════════════════


def _break_start(line: str) -> Optional[int]:
    """Index of an unescaped trailing hard-break marker (`\\` or `<br>`), else None."""
    text = line.rstrip()
    tag = _BR_TAG_RE.search(text)
    if tag:
        start = tag.start()
    elif text.endswith("\\"):
        start = len(text) - 1
    else:
        return None
    escapes = len(text[:start]) - len(text[:start].rstrip("\\"))
    return start if escapes % 2 == 0 else None


def _core(line: str) -> str:
    """Line without a trailing hard-break marker."""
    start = _break_start(line)
    return (line[:start] if start is not None else line).strip()


def _plain(line: str) -> str:
    """Line text without surrounding emphasis markers."""
    return _EMPHASIS_EDGE_RE.sub("", line)


def _image_sources(line: str) -> Optional[List[str]]:
    """Image sources when the line holds only images (and emphasis), else None."""
    images = _IMAGE_RE.findall(line)
    if not images or _plain(_IMAGE_RE.sub("", line)):
        return None
    sources = []
    for image in images:
        match = _IMAGE_SOURCE_RE.search(image)
        source = next((group for group in match.groups() if group), "") if match else ""
        sources.append("a data: URI" if source.lower().startswith("data:") else source or "an image")
    return sources


def _title_text(line: str) -> Optional[str]:
    """Title candidate: a line that is one bold span, or a level-1 heading."""
    inner = _unbold(line)
    if inner is not None:
        return inner
    heading = _HEADING_RE.match(line)
    if heading and len(heading.group(1)) == 1 and heading.group(2):
        return heading.group(2)
    return None


def _cover_kind(core: str) -> Optional[str]:
    """Classify one cover line; None when the cleanup cannot tell."""
    plain = _plain(core)
    if _image_sources(core):
        return "image"
    if _CONFIDENTIAL_RE.search(plain) and len(plain) <= _CONFIDENTIAL_MAX_CHARS:
        return "confidential"
    if _DATE_RE.match(plain):
        return "date"
    if len(plain) >= _DISCLAIMER_MIN_CHARS:
        return "long"
    if _title_text(core) is not None:
        return "title"
    return None


def _paragraph_text(lines: List[str]) -> str:
    """Soft-wrapped lines joined by spaces; hard breaks become line breaks."""
    parts: List[str] = []
    for index, line in enumerate(lines):
        if index:
            parts.append("\n" if _break_start(lines[index - 1]) is not None else " ")
        parts.append(_core(line))
    return "".join(parts)


def _is_chapter(block: _Block) -> bool:
    """An ATX heading of level 2 or more, or a numbered level-1 heading (up to 3 spaces in).

    Any such heading ends the cover, indented or not, and so does a one-row
    band table kept as a table; ending it early only moves fewer lines.
    """
    if block.band:
        return True
    if block.kind != "text" or len(block.lines) != 1 or not _ATX_RE.match(block.lines[0]):
        return False
    heading = block.lines[0].lstrip()
    return heading.startswith("##") or bool(_NUMBERED_H1_RE.match(heading))


def _extract_cover(
    blocks: List[_Block], existing: Dict[str, Any], offset: int, report: ConvertedCleanupReport
) -> Dict[str, str]:
    """Move top-level cover lines before the first chapter heading into frontmatter values."""
    first = next((position for position, block in enumerate(blocks) if _is_chapter(block)), None)
    if first is None:
        report.add(0, "no chapter heading found: cover lines were not moved")
        return {}
    moves: List[_CoverMove] = []
    titles: List[_CoverMove] = []
    title_open = True
    label: Optional[str] = None

    def line_no(block: _Block, index: int) -> int:
        return offset + block.start + index + 1

    def keep(block: _Block, index: int, why: str) -> None:
        report.add(line_no(block, index), f"kept {why}: {_quote(block.lines[index])}")

    for block in blocks[:first]:
        if block.kind != "text" or not block.top_level or block.container:
            title_open = title_open and not titles
            if block.kind == "text":
                what = "list or quote block" if block.top_level else "indented block"
            else:
                what = {"fence": "code block", "indented": "indented block"}.get(block.kind, "table")
            first_line, last_line = _span(block, offset)
            report.add(first_line, f"kept a {what} before the first chapter", last_line)
            continue
        cores = [_core(line) for line in block.lines]
        kinds = [_cover_kind(core) for core in cores]
        if set(kinds) <= {None, "long"} and len(_plain(" ".join(cores))) >= _DISCLAIMER_MIN_CHARS:
            indexes = tuple(range(len(block.lines)))
            moves.append(_CoverMove(block, indexes, "disclaimer", _paragraph_text(block.lines)))
            title_open = title_open and not titles
            continue
        for index, (core, kind) in enumerate(zip(cores, kinds)):
            plain = _plain(core)
            if block.lines[index][:1].isspace():  # a continuation line is never moved on its own
                title_open = title_open and not titles
                keep(block, index, "an indented line")
            elif kind == "image":
                block.removed[index] = True
                for source in _image_sources(core) or []:
                    report.add(
                        line_no(block, index),
                        f"removed cover image {_quote(source)}: set it as the house style.logo if it is the logo",
                    )
            elif kind == "confidential":
                if label is None:
                    label = plain
                    moves.append(_CoverMove(block, (index,), "confidential_label", plain))
                elif plain == label and "confidential_label" not in existing:
                    block.removed[index] = True
                    report.add(line_no(block, index), f"removed repeated confidentiality line {_quote(plain)}")
                else:
                    keep(block, index, "another confidentiality line")
            elif kind == "date":
                if any(move.key == "date" for move in moves):
                    keep(block, index, "a second date line")
                else:
                    moves.append(_CoverMove(block, (index,), "date", plain))
            elif kind == "title" and title_open:
                key = "subtitle" if titles else "title"
                titles.append(_CoverMove(block, (index,), key, _title_text(core) or plain))
            elif kind == "long":
                moves.append(_CoverMove(block, (index,), "disclaimer", core))
                title_open = title_open and not titles
            else:
                title_open = title_open and not titles
                keep(block, index, "a bold line after the title" if kind == "title" else "unclassified cover text")
    return _apply_moves(moves + titles, existing, offset, report)


def _apply_moves(
    moves: List[_CoverMove], existing: Dict[str, Any], offset: int, report: ConvertedCleanupReport
) -> Dict[str, str]:
    """Remove moved lines and collect values; a key already set keeps its lines."""
    values: Dict[str, List[str]] = {}
    for move in moves:
        first = offset + move.block.start + move.indexes[0] + 1
        last = offset + move.block.start + move.indexes[-1] + 1
        if move.key in existing:
            source = move.block.lines[move.indexes[0]]
            report.add(first, f"kept (frontmatter already sets {move.key}): {_quote(source)}", last)
            continue
        for index in move.indexes:
            move.block.removed[index] = True
        values.setdefault(move.key, []).append(move.value)
        report.add(first, f"frontmatter {move.key} <- {_quote(move.value)}", last)
    joiners = {"subtitle": " ", "disclaimer": "\n"}
    return {key: joiners.get(key, "").join(values[key]) for key in _FRONTMATTER_ORDER if key in values}


def _report_missing(existing: Dict[str, Any], fields: Dict[str, str], report: ConvertedCleanupReport) -> None:
    if "title" not in existing and "title" not in fields:
        report.add(0, "title not found: add title to the frontmatter")
    if "prepared_by" not in existing:
        report.add(0, "prepared_by is not inferred: add it to the house file or the frontmatter")
    if "disclaimer" not in existing and "disclaimer" not in fields:
        report.add(0, "disclaimer not found: the house file or the frontmatter must supply it")


# ═══════════════════════════════════════════════════════════════════════════════
# OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════


class _FrontmatterDumper(yaml.SafeDumper):
    """Safe dumper writing multi-line strings as literal blocks (local, not global)."""


def _represent_str(dumper: yaml.SafeDumper, value: str) -> yaml.Node:
    return dumper.represent_scalar("tag:yaml.org,2002:str", value, style="|" if "\n" in value else None)


_FrontmatterDumper.add_representer(str, _represent_str)


def _dump(data: Dict[Any, Any]) -> List[str]:
    text = yaml.dump(data, Dumper=_FrontmatterDumper, allow_unicode=True, sort_keys=False, width=4096)
    return text.rstrip("\n").split("\n")


def _write_frontmatter(
    front_lines: List[str], fields: Dict[str, str], report: ConvertedCleanupReport
) -> List[str]:
    """Existing frontmatter lines unchanged, new keys before the closing fence.

    Flow-style frontmatter (`{title: ...}`), which the parser does not read,
    and frontmatter that appending would not extend to the expected mapping
    are rewritten as block YAML, and the rewrite is reported.
    """
    if not front_lines:
        return ["---", *_dump(fields), "---"] if fields else []
    body = front_lines[1:-1]
    first = next((line for line in body if line.strip() and not line.lstrip().startswith("#")), "")
    block_style = not first or bool(_BLOCK_KEY_RE.match(first))
    if block_style and not fields:
        return front_lines
    original = yaml.safe_load("\n".join(body)) or {}
    if block_style:
        appended = body + _dump(fields)
        try:
            merged = yaml.safe_load("\n".join(appended))
        except yaml.YAMLError:
            merged = None
        if merged == {**original, **fields}:
            return front_lines[:1] + appended + front_lines[-1:]
    report.add(0, "rewrote the frontmatter as block YAML the parser reads; its comments and layout were not kept")
    return ["---", *_dump({**original, **fields}), "---"]


def _quote(text: str) -> str:
    text = " ".join(text.split())
    return '"' + (text if len(text) <= _QUOTE_CHARS else text[: _QUOTE_CHARS - 1] + "…") + '"'
