"""Shared document model for Markdown-to-Word rendering.

Changelog (quality hardening):
    - Preserve inferred IB subtitle headings for cover-free rendering.
    - Add chart specifications with retained source for opt-in rendering.
    - Retain original confirmation fences for lossless term-sheet fallback.
    - Carry an optional table note rendered below the table in every profile.

Changelog (captions):
    - Tables and figures with a reference ID (`NumberedTable`, `Figure`) and the
      render-time caption number run; subclasses keep models without IDs unchanged.
    - `CrossReference` runs for `@tbl-…` / `@fig-…` in body text.
    - Review fixes: `MarkedHeading` and `MarkedBlockquote` keep escaped `@`
      apart in text that was unescaped at parse time; `table_is_drawn` and
      `figure_id_problem` are the one rule for what gets a number.
    - Second review: a `CrossReference` is made only for a reference that
      resolves, so nothing is joined back; `numbered_items` is the one list of
      numbered tables and figures for the parser and the renderer.
"""

import re
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field, fields
from enum import Enum, auto
from pathlib import Path
from typing import Any, Optional, Union

from input_safety import DEFAULT_LIMITS, InputLimits


class ElementType(Enum):
    """Types of markdown elements"""

    HEADING_1 = auto()
    HEADING_2 = auto()
    HEADING_3 = auto()
    HEADING_4 = auto()
    NUMBERED_HEADING = auto()  # **1. 제목** format
    PARAGRAPH = auto()
    BULLET_LIST = auto()
    NUMBERED_LIST = auto()
    TABLE = auto()
    BLOCKQUOTE = auto()
    IMAGE = auto()
    SEPARATOR = auto()
    EMPTY = auto()
    CODE_BLOCK = auto()
    # ── NEW (v3) ────────────────────────────────────────────────────────────
    LATEX_BLOCK = auto()  # $$ ... $$ (display math)
    LATEX_INLINE = auto()  # standalone inline math rendered as paragraph
    # ── NEW (v5) ────────────────────────────────────────────────────────────
    DIAGRAM = auto()  # ```diagram:type ... ``` code block
    CHART = auto()  # ```chart YAML with lossless code-panel fallback
    CONFIRMATION = auto()  # term-sheet ```confirmation``` block (customer sign-off box)


class TableType(Enum):
    """Types of tables for specialized rendering"""

    GENERIC = auto()
    FINANCIAL = auto()
    BEP_SENSITIVITY = auto()
    RISK_MATRIX = auto()
    UPSIDE_DOWNSIDE = auto()


# ═══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class TextRun:
    """A run of text with formatting"""

    text: str
    bold: bool = False
    italic: bool = False
    superscript: bool = False
    subscript: bool = False
    color_hex: str | None = None
    is_latex: bool = False  # NEW (v3): marks this run as inline LaTeX
    hyperlink: str | None = None
    footnote_id: int | None = None
    term_key: str | None = None  # set when the text is a substituted {{term}} value
    code: bool = False  # inline code span content (rendered monospace, without backticks)
    image: Optional["Image"] = None  # inline image (e.g. in a table cell); the text is empty


@dataclass
class CaptionNumber(TextRun):
    """The `표 1` at the start of a numbered caption, built at render time.

    `text` holds the label and number as displayed; the renderer writes the
    label, a `SEQ` field whose cached result is `number`, and a hidden bookmark
    around both.
    """

    label: str = ""
    number: int = 0
    bookmark: str = ""


@dataclass
class CrossReference(TextRun):
    """A `@tbl-…` / `@fig-…` that resolves to a numbered table or figure.

    The parser makes one only for a reference whose target is numbered and
    that stands where a Word field can go (`TextParser._read_references`);
    `text` is the `@…` as written and `target` the ID. The renderer writes it
    as a `REF` field. Text that is not such a reference keeps its run.
    """

    target: str = ""


TEXT_RUN_FIELDS = tuple(item.name for item in fields(TextRun))


@dataclass
class LaTeXEquation:
    """A LaTeX equation element (NEW v3)"""

    expression: str
    is_block: bool = True  # True = display ($$), False = inline ($)


@dataclass
class CodeBlock:
    """A fenced code block element."""

    code: str
    language: str = ""
    is_ascii_art: bool = False

    _BOX_CHARS = set("┌┐└┘│─├┤┬┴┼╔╗╚╝║═╠╣╦╩╬→←↑↓▶◀▲▼►◄─━")

    @staticmethod
    def detect_ascii_art(text: str, threshold: int = 20) -> bool:
        """Return True if text contains enough box-drawing characters."""
        count = sum(1 for ch in text if ch in CodeBlock._BOX_CHARS)
        return count >= threshold


@dataclass
class ConfirmationBlock:
    """An empty term-sheet confirmation fence with its complete source."""

    source: str


@dataclass
class DiagramBox:
    """A box in a flow diagram."""

    id: str
    label: str
    pos: list[float] = field(default_factory=lambda: [0, 0])
    style: str = "default"


@dataclass
class ChartSeries:
    """One named sequence of chart values, in the supplied units."""

    name: str
    values: list[float]


@dataclass
class ChartSpec:
    """Validated chart data; waterfall values are deltas from zero."""

    chart_type: str
    title: str | None
    labels: list[str]
    series: list[ChartSeries]
    y_label: str | None = None
    source: str | None = None
    total_label: str | None = None
    unit: str | None = None
    number_format: str | None = None


@dataclass
class Chart:
    """Chart fence with its source and any deferred validation diagnostic.

    Parsing recognizes charts independently of rendering policy. Disabled charts
    always use the original code panel, including invalid specifications.
    """

    code: str
    spec: ChartSpec | None = None
    error: str | None = None
    label: str = "Chart"


@dataclass
class DiagramArrow:
    """An arrow connecting two boxes in a flow diagram."""

    from_id: str
    to_id: str
    label: str = ""
    style: str = "solid"


@dataclass
class Diagram:
    """A flow diagram element parsed from ```diagram:flow code blocks."""

    diagram_type: str = "flow"
    title: str = ""
    boxes: list[DiagramBox] = field(default_factory=list)
    arrows: list[DiagramArrow] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class TableCell:
    """A cell in a table"""

    content: str
    runs: list[TextRun] = field(default_factory=list)
    alignment: str = "left"  # left, center, right
    is_header: bool = False
    is_numeric: bool = False
    is_negative: bool = False
    is_base_case: bool = False
    risk_level: str | None = None  # high, medium, low
    merge: str | None = None  # resolved span marker: "up" (^^) or "left" (<<)
    literal: bool = False  # runs are final (built from HTML): the text is never a span marker


@dataclass
class TableRow:
    """A row in a table"""

    cells: list[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    """A parsed table"""

    rows: list[TableRow] = field(default_factory=list)
    table_type: TableType = TableType.GENERIC
    col_count: int = 0
    alignments: list[str] = field(default_factory=list)
    column_types: list[str] = field(default_factory=list)
    caption: str = ""
    unit: str = ""
    source: str = ""
    as_of: str = ""
    landscape: bool = False
    warnings: list[str] = field(default_factory=list)
    spans: bool | None = None  # table spec `spans`; None inherits the profile default
    label_columns: int | None = None  # shaded leading label columns (term-sheet layout)
    note: str = ""  # table spec `note`: short right-aligned text below the table
    header_rows: int = 1  # leading rows drawn and repeated as the header (spec `header_rows`)
    schedule: dict[str, Any] | None = None  # validated table spec `schedule` (repayment checks)


@dataclass
class NumberedTable(Table):
    """A table whose specification gives a reference ID (`tables[].id`).

    A subclass rather than a field, so tables without an ID keep their model
    exactly as before.
    """

    ref_id: str = ""


@dataclass
class Heading:
    """A heading element"""

    level: int
    text: str
    is_numbered: bool = False
    runs: list[TextRun] = field(default_factory=list)  # set only when terms are substituted


@dataclass
class MarkedHeading(Heading):
    """A heading whose `text` was unescaped at parse time and held an escaped `@`.

    `marked_text` is `text` with each escaped `@` written as
    `TextParser.ESCAPED_AT`, so a `\\@tbl-…` is never read as a reference.
    """

    marked_text: str = ""


@dataclass
class Paragraph:
    """A paragraph element"""

    text: str
    runs: list[TextRun] = field(default_factory=list)
    has_inline_latex: bool = False  # NEW (v3)


@dataclass
class ListItem:
    """A list item"""

    text: str
    runs: list[TextRun] = field(default_factory=list)
    indent_level: int = 0


@dataclass
class BulletList:
    """A bullet list"""

    items: list[ListItem] = field(default_factory=list)


@dataclass
class NumberedList:
    """A numbered list"""

    items: list[ListItem] = field(default_factory=list)


@dataclass
class Blockquote:
    """A blockquote (callout)"""

    text: str
    title: str = "KEY INSIGHT"
    runs: list[TextRun] = field(default_factory=list)  # set only when terms are substituted


@dataclass
class MarkedBlockquote(Blockquote):
    """A callout whose source held an escaped `@` (see `MarkedHeading`)."""

    marked_text: str = ""


@dataclass
class Image:
    """An image reference — supports file paths and Base64 (v3)"""

    alt_text: str
    path: str
    base64_data: str | None = None  # NEW (v3): Base64-encoded image data
    mime_type: str = "image/png"  # NEW (v3): MIME type for Base64


@dataclass
class Figure(Image):
    """An image line with a reference ID (`![설명](파일){#fig-…}`), numbered in its caption."""

    ref_id: str = ""


# ═══════════════════════════════════════════════════════════════════════════════
# CAPTION NUMBERS AND REFERENCE IDS
# ═══════════════════════════════════════════════════════════════════════════════

# Word's caption labels, the same in every profile (owner decision 2026-11-14).
CAPTION_LABELS = {"tbl": "표", "fig": "그림"}
# Lowercase ASCII words joined by hyphens: the ID maps one-to-one onto a Word
# bookmark name (`-` becomes `_`), and a Korean particle may follow it directly.
REFERENCE_ID_RE = re.compile(r"(tbl|fig)-[a-z0-9]+(?:-[a-z0-9]+)*")
# Word allows 40 characters in a bookmark name; `_Ref_` takes five.
REFERENCE_ID_MAX_LENGTH = 35


def reference_id_problem(ref_id: str, kind: str) -> str | None:
    """Explain why `ref_id` cannot name a table or figure, or return None.

    Args:
        ref_id: The ID as written, without `#` or `@`.
        kind: `tbl` for a table, `fig` for a figure.

    Returns:
        A diagnostic sentence, or None for a usable ID.
    """
    match = REFERENCE_ID_RE.fullmatch(ref_id)
    if match is None or match.group(1) != kind:
        return (
            f"{ref_id!r} is not a {'table' if kind == 'tbl' else 'figure'} id: write {kind}- "
            "followed by lowercase letters and digits, words joined by hyphens"
        )
    if len(ref_id) > REFERENCE_ID_MAX_LENGTH:
        return f"{ref_id!r} is longer than {REFERENCE_ID_MAX_LENGTH} characters"
    return None


def caption_bookmark_name(ref_id: str) -> str:
    """Return the hidden Word bookmark around a caption's label and number."""
    return "_Ref_" + ref_id.replace("-", "_")


def figure_id_problem(ref_id: str, path: str, description: str) -> str | None:
    """Explain why an image with `{#ref_id}` cannot be a numbered figure, or None.

    Args:
        ref_id: The ID as written after `#`.
        path: The image destination.
        description: The image description (alt text).

    Returns:
        A diagnostic starting with the ID, or None when the image can be numbered.
    """
    problem = reference_id_problem(ref_id, "fig")
    if problem is None and path.startswith("data:"):
        problem = f"{ref_id!r} is on an embedded image; figure ids need an image file"
    if problem is None and not description.strip():
        problem = f"{ref_id!r} needs an image description: its number is written before it"
    return problem


def table_is_drawn(table: "Table") -> bool:
    """Whether the renderer draws `table` (and so whether its caption gets a number)."""
    return bool(table.rows) and table.col_count > 0


def numbered_items(
    elements: Sequence["Element"],
) -> Iterator[tuple["NumberedTable | Figure", str, str | None]]:
    """Each table and figure with a reference ID, in document order.

    The parser takes the IDs that resolve from here and the renderer numbers
    the same items, so both agree on what a reference can point to.

    Args:
        elements: Document elements.

    Yields:
        The table or figure, its kind (`tbl` or `fig`) and why it gets no
        number (None when it is numbered).
    """
    for element in elements:
        content = element.content
        if element.element_type == ElementType.TABLE and isinstance(content, NumberedTable):
            kind, caption = "tbl", content.caption
        elif element.element_type == ElementType.IMAGE and isinstance(content, Figure):
            kind, caption = "fig", content.alt_text
        else:
            continue
        problem = reference_id_problem(content.ref_id, kind)
        if problem is None and not caption.strip():
            problem = f"{content.ref_id!r} has no caption to number"
        if problem is None and isinstance(content, NumberedTable) and not table_is_drawn(content):
            # The renderer skips a table without rows or columns, so it gets
            # no number and references to it have no target.
            problem = f"{content.ref_id!r} is on a table with no columns, which is not drawn"
        yield content, kind, problem


def numbered_targets(elements: Sequence["Element"]) -> frozenset[str]:
    """IDs that a reference can resolve to: those of numbered tables and figures."""
    return frozenset(
        item.ref_id for item, _, problem in numbered_items(elements) if problem is None
    )


# ─────────────────────────────────────────────────────────────────────────────
# Union type for Element.content
# ─────────────────────────────────────────────────────────────────────────────
ElementContent = Union[
    Heading,
    Paragraph,
    Table,
    ListItem,
    tuple[str, ListItem],  # numbered list: (number, ListItem)
    Blockquote,
    Image,
    CodeBlock,
    ConfirmationBlock,
    Chart,
    LaTeXEquation,  # NEW (v3)
    "Diagram",  # NEW (v5)
    None,
]


@dataclass
class Element:
    """A generic document element"""

    element_type: ElementType
    content: ElementContent
    raw_text: str = ""
    inferred_subtitle: bool = False  # Render in the body unless used on an IB cover.


@dataclass
class Section:
    """A document section"""

    heading: Heading | None = None
    elements: list[Element] = field(default_factory=list)


@dataclass
class Footnote:
    """A footnote reference"""

    number: int
    text: str


@dataclass
class DocumentMetadata:
    """Document metadata from frontmatter"""

    title: str = "IB Report"
    subtitle: str = ""
    company: str = "Korea Development Bank"
    ticker: str = ""
    sector: str = "SECTOR"
    analyst: str = "DCM Team 1"
    extra: dict[str, Any] = field(default_factory=dict)
    profile: str = "ib-report"
    # Parsed runs for title/subtitle/date when terms are substituted; empty otherwise.
    display_runs: dict[str, list[TextRun]] = field(default_factory=dict)


@dataclass
class DocumentModel:
    """The complete parsed document"""

    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    sections: list[Section] = field(default_factory=list)
    elements: list[Element] = field(default_factory=list)
    footnotes: dict[int, str] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    # None means a hand-built model; parsed models retain their parsing policy.
    parsed_profile: str | None = None
    # Folder of the source file; None for text or stream input. Relative link
    # targets refer to it until a save rebases them on the output folder.
    source_dir: Path | None = None
    # Trusted caller context, deliberately separate from frontmatter metadata.
    asset_roots: tuple[Path, ...] = ()
    limits: InputLimits = DEFAULT_LIMITS


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════
