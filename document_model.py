"""Shared document model for Markdown-to-Word rendering."""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any, Dict, List, Optional, Tuple, Union


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
    color_hex: Optional[str] = None
    is_latex: bool = False  # NEW (v3): marks this run as inline LaTeX
    hyperlink: Optional[str] = None
    footnote_id: Optional[int] = None


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
class DiagramBox:
    """A box in a flow diagram."""

    id: str
    label: str
    pos: List[float] = field(default_factory=lambda: [0, 0])
    style: str = "default"


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
    boxes: List[DiagramBox] = field(default_factory=list)
    arrows: List[DiagramArrow] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)


@dataclass
class TableCell:
    """A cell in a table"""

    content: str
    runs: List[TextRun] = field(default_factory=list)
    alignment: str = "left"  # left, center, right
    is_header: bool = False
    is_numeric: bool = False
    is_negative: bool = False
    is_base_case: bool = False
    risk_level: Optional[str] = None  # high, medium, low


@dataclass
class TableRow:
    """A row in a table"""

    cells: List[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    """A parsed table"""

    rows: List[TableRow] = field(default_factory=list)
    table_type: TableType = TableType.GENERIC
    col_count: int = 0
    alignments: List[str] = field(default_factory=list)
    column_types: List[str] = field(default_factory=list)
    caption: str = ""
    unit: str = ""
    source: str = ""
    as_of: str = ""
    landscape: bool = False
    warnings: List[str] = field(default_factory=list)


@dataclass
class Heading:
    """A heading element"""

    level: int
    text: str
    is_numbered: bool = False


@dataclass
class Paragraph:
    """A paragraph element"""

    text: str
    runs: List[TextRun] = field(default_factory=list)
    has_inline_latex: bool = False  # NEW (v3)


@dataclass
class ListItem:
    """A list item"""

    text: str
    runs: List[TextRun] = field(default_factory=list)
    indent_level: int = 0


@dataclass
class BulletList:
    """A bullet list"""

    items: List[ListItem] = field(default_factory=list)


@dataclass
class NumberedList:
    """A numbered list"""

    items: List[ListItem] = field(default_factory=list)


@dataclass
class Blockquote:
    """A blockquote (callout)"""

    text: str
    title: str = "KEY INSIGHT"


@dataclass
class Image:
    """An image reference — supports file paths and Base64 (v3)"""

    alt_text: str
    path: str
    base64_data: Optional[str] = None  # NEW (v3): Base64-encoded image data
    mime_type: str = "image/png"  # NEW (v3): MIME type for Base64


# ─────────────────────────────────────────────────────────────────────────────
# Union type for Element.content
# ─────────────────────────────────────────────────────────────────────────────
ElementContent = Union[
    Heading,
    Paragraph,
    Table,
    ListItem,
    Tuple[str, ListItem],  # numbered list: (number, ListItem)
    Blockquote,
    Image,
    CodeBlock,
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


@dataclass
class Section:
    """A document section"""

    heading: Optional[Heading] = None
    elements: List[Element] = field(default_factory=list)


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
    extra: Dict[str, Any] = field(default_factory=dict)
    profile: str = "ib-report"


@dataclass
class DocumentModel:
    """The complete parsed document"""

    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    sections: List[Section] = field(default_factory=list)
    elements: List[Element] = field(default_factory=list)
    footnotes: Dict[int, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════
