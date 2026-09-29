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
"""

import html
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import yaml

# ═══════════════════════════════════════════════════════════════════════════════
# PATTERNS
# ═══════════════════════════════════════════════════════════════════════════════

_FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})")
_DELIMITER_CELL_RE = re.compile(r"^\s*:?-+:?\s*$")
_CELL_SPLIT_RE = re.compile(r"(?<!\\)\|")
_CHAPTER_RE = re.compile(r"^(\d{1,2})\\?\.\s+(.{1,40})$")
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_NUMBERED_H1_RE = re.compile(r"^#\s+\d{1,2}\\?\.\s")
_HARD_BREAK_RE = re.compile(r"(?:\\|<br\s*/?>)\s*$", re.IGNORECASE)
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
_HTML_ROW_RE = re.compile(r"<tr\b", re.IGNORECASE)
_HTML_CELL_RE = re.compile(r"<t([dh])\b[^>]*>(.*?)</t\1\s*>", re.IGNORECASE | re.DOTALL)
_HTML_TAG_RE = re.compile(r"<[^>]+>")

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
        line: One-based line in the input; 0 when not tied to a line.
        message: What was moved, rewritten, removed, kept or is missing.
    """

    line: int
    message: str


@dataclass
class ConvertedCleanupReport:
    """Every change the cleanup made, plus the cover lines it left undecided."""

    notes: List[CleanupNote] = field(default_factory=list)

    def add(self, line: int, message: str) -> None:
        """Record one note; `line` is one-based, 0 for document-level notes."""
        self.notes.append(CleanupNote(line, message))

    def lines(self) -> List[str]:
        """Printable report lines: line-bound notes in input order, then the rest."""
        ordered = sorted(self.notes, key=lambda note: (note.line == 0, note.line))
        return [f"line {note.line}: {note.message}" if note.line else note.message for note in ordered]


@dataclass
class _Block:
    """Consecutive non-blank source lines of one kind: text, pipe, html or fence."""

    kind: str
    start: int  # zero-based index into the body lines
    lines: List[str]
    gap: List[str]  # blank lines before the block
    removed: List[bool] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.removed = [False] * len(self.lines)


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

    Chapter headings: a one-line paragraph `1. Title` (at most 40 characters
    after the number; `1\\.` too) becomes `## 1. Title` when it is bold or the
    next block is a table. A table with a header row and no body rows becomes
    a `##` heading of its non-empty cells joined by ` | `. Blocks before the
    first chapter heading are the cover: the first confidentiality line becomes
    `confidential_label` (identical repeats are removed), the leading run of
    bold lines (or a level-1 heading) becomes `title` then `subtitle`, a date
    line `date`, paragraphs of 80 characters or more `disclaimer`; images are
    removed and reported for the house `style.logo`. Other cover lines stay
    and are reported.
    `prepared_by` is never inferred. Existing frontmatter lines and keys are
    kept, and a cover line whose key is already set stays in the body. Fenced
    code is never changed.

    Args:
        text: Converted Markdown, with or without frontmatter.

    Returns:
        The cleaned Markdown and the report of every change.

    Raises:
        ValueError: Existing frontmatter is not valid YAML or not a mapping.
    """
    report = ConvertedCleanupReport()
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    front_lines, existing, offset = _read_frontmatter(lines)
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

    front = _write_frontmatter(front_lines, fields)
    output = front + ([""] if front else []) + _join_blocks(blocks, trailing)
    return "\n".join(output).rstrip("\n") + "\n", report


# ═══════════════════════════════════════════════════════════════════════════════
# BLOCKS
# ═══════════════════════════════════════════════════════════════════════════════


def _read_frontmatter(lines: List[str]) -> Tuple[List[str], Dict[str, Any], int]:
    """Return the frontmatter lines (with fences), its mapping and the body start."""
    if not lines or lines[0].strip() != "---":
        return [], {}, 0
    for index in range(1, len(lines)):
        if lines[index].strip() in ("---", "..."):
            try:
                data = yaml.safe_load("\n".join(lines[1:index])) or {}
            except yaml.YAMLError as exc:
                raise ValueError(f"existing frontmatter is not valid YAML: {exc}") from exc
            if not isinstance(data, dict):
                raise ValueError("existing frontmatter must be a YAML mapping")
            return lines[: index + 1], data, index + 1
    return [], {}, 0


def _is_pipe_row(line: str) -> bool:
    return line.lstrip().startswith("|")


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
    return line.lstrip().lower().startswith("<table")


def _closes_fence(line: str, marker: str) -> bool:
    stripped = line.strip()
    return (
        len(stripped) >= len(marker)
        and set(stripped) == {marker[0]}
        and len(line) - len(line.lstrip()) <= 3
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
        if fence:
            end = index + 1
            while end < len(lines) and not _closes_fence(lines[end], fence.group(1)):
                end += 1
            kind, end = "fence", min(end + 1, len(lines))
        elif _starts_html_table(line):
            end = index
            while end < len(lines) and "</table>" not in lines[end].lower():
                end += 1
            kind, end = "html", min(end + 1, len(lines))
        elif _starts_pipe_table(lines, index):
            end = index + 2
            while end < len(lines) and _is_pipe_row(lines[end]):
                end += 1
            kind = "pipe"
        else:
            end = index + 1
            while (
                end < len(lines)
                and lines[end].strip()
                and not _FENCE_RE.match(lines[end])
                and not _starts_html_table(lines[end])
                and not _starts_pipe_table(lines, end)
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
        if block.kind != "text" or len(block.lines) != 1:
            continue
        source = block.lines[0]
        if source.startswith(("    ", "\t")):  # indented code or list continuation
            continue
        inner = _unbold(source)
        match = _CHAPTER_RE.match(inner if inner is not None else source.strip())
        if match is None:
            continue
        next_kind = blocks[position + 1].kind if position + 1 < len(blocks) else ""
        if inner is None and next_kind not in ("pipe", "html"):
            continue
        heading = f"## {match.group(1)}. {match.group(2).strip()}"
        block.lines[0] = heading
        reason = "bold numbered line" if inner is not None else "numbered line before a table"
        report.add(offset + block.start + 1, f"heading {_quote(source)} -> {_quote(heading)} ({reason})")


def _convert_band_tables(blocks: List[_Block], offset: int, report: ConvertedCleanupReport) -> None:
    """A table with a header row and no body rows becomes a `##` heading."""
    for block in blocks:
        cells: Optional[List[str]] = None
        if block.kind == "pipe" and len(block.lines) == 2:
            cells = [_unbold(cell) or cell for cell in _row_cells(block.lines[0])]
        elif block.kind == "html":
            cells = _html_band_cells("\n".join(block.lines))
        if cells is None:
            continue
        line = offset + block.start + 1
        texts = [cell for cell in cells if cell]
        if not texts:
            report.add(line, "kept a one-row table without text")
            continue
        heading = "## " + " | ".join(texts)
        block.kind, block.lines, block.removed = "text", [heading], [False]
        report.add(line, f"band table -> {_quote(heading)}")


def _html_band_cells(source: str) -> Optional[List[str]]:
    """Cell texts of a lone single-row HTML table without images or nested tables."""
    lowered = source.strip().lower()
    if (
        not lowered.endswith("</table>")
        or lowered.count("<table") != 1
        or len(_HTML_ROW_RE.findall(source)) != 1
        or "<img" in lowered
    ):
        return None
    cells = [
        " ".join(html.unescape(_HTML_TAG_RE.sub(" ", content)).split())
        for _, content in _HTML_CELL_RE.findall(source)
    ]
    return cells or None


# ═══════════════════════════════════════════════════════════════════════════════
# COVER
# ═══════════════════════════════════════════════════════════════════════════════


def _core(line: str) -> str:
    """Line without a trailing hard-break marker (`\\` or `<br>`)."""
    return _HARD_BREAK_RE.sub("", line).strip()


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
            parts.append("\n" if _HARD_BREAK_RE.search(lines[index - 1]) else " ")
        parts.append(_core(line))
    return "".join(parts)


def _is_chapter(block: _Block) -> bool:
    if block.kind != "text" or len(block.lines) != 1:
        return False
    line = block.lines[0]
    return line.startswith("##") or bool(_NUMBERED_H1_RE.match(line))


def _extract_cover(
    blocks: List[_Block], existing: Dict[str, Any], offset: int, report: ConvertedCleanupReport
) -> Dict[str, str]:
    """Move cover lines before the first chapter heading into frontmatter values."""
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
        if block.kind != "text":
            title_open = title_open and not titles
            what = "code block" if block.kind == "fence" else "table"
            report.add(line_no(block, 0), f"kept a {what} before the first chapter")
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
            if kind == "image":
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
        line = offset + move.block.start + move.indexes[0] + 1
        if move.key in existing:
            source = move.block.lines[move.indexes[0]]
            report.add(line, f"kept (frontmatter already sets {move.key}): {_quote(source)}")
            continue
        for index in move.indexes:
            move.block.removed[index] = True
        values.setdefault(move.key, []).append(move.value)
        report.add(line, f"frontmatter {move.key} <- {_quote(move.value)}")
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


def _write_frontmatter(front_lines: List[str], fields: Dict[str, str]) -> List[str]:
    """Existing frontmatter lines unchanged; new keys go before the closing fence."""
    if not fields:
        return front_lines
    added = yaml.dump(
        fields, Dumper=_FrontmatterDumper, allow_unicode=True, sort_keys=False, width=4096
    ).rstrip("\n").split("\n")
    if front_lines:
        return front_lines[:-1] + added + front_lines[-1:]
    return ["---", *added, "---"]


def _quote(text: str) -> str:
    text = " ".join(text.split())
    return '"' + (text if len(text) <= _QUOTE_CHARS else text[: _QUOTE_CHARS - 1] + "…") + '"'
