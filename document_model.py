"""Shared document model for Markdown-to-Word rendering.

Changelog (quality hardening):
    - Preserve inferred IB subtitle headings for cover-free rendering.
    - Add chart specifications with retained source for opt-in rendering.
    - Retain original confirmation fences for lossless term-sheet fallback.
    - Carry an optional table note rendered below the table in every profile.
"""

from dataclasses import dataclass, field
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
class Heading:
    """A heading element"""

    level: int
    text: str
    is_numbered: bool = False
    runs: list[TextRun] = field(default_factory=list)  # set only when terms are substituted


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
class Image:
    """An image reference — supports file paths and Base64 (v3)"""

    alt_text: str
    path: str
    base64_data: str | None = None  # NEW (v3): Base64-encoded image data
    mime_type: str = "image/png"  # NEW (v3): MIME type for Base64


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


class DefaultTitle(str):
    """A profile's stand-in title, not one the document gave.

    Any title read from frontmatter or the body is a plain `str`, even when its
    text equals the stand-in, so a renderer can tell the two apart.
    """


@dataclass
class DocumentMetadata:
    """Document metadata from frontmatter"""

    title: str = DefaultTitle("IB Report")
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
