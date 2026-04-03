"""
MD Parser Module for IB Style Word Report Converter
Handles parsing of Markdown files including frontmatter, elements, tables,
LaTeX equations, Base64 images, and footnotes.

Changelog (v3):
    - NEW: LaTeX block equation parsing ($$ ... $$, multi-line)
    - NEW: LaTeX inline equation detection within paragraphs ($ ... $)
    - NEW: ElementType.LATEX_BLOCK / LATEX_INLINE
    - NEW: LaTeXEquation dataclass
    - NEW: Base64 embedded image parsing (data:image/... URI)
    - NEW: Image.base64_data / Image.mime_type fields
    - NEW: Paragraph.has_inline_latex flag
    - ENHANCED: TextRun with is_latex flag for inline math
    - ENHANCED: Encoding detection with charset_normalizer fallback
    - ENHANCED: Table column truncation warning (no silent data loss)
    - FIXED: heading level mapping (## → level=2, ### → level=3, #### → level=4)
    - FIXED: numbered list consuming heading-like lines
    - FIXED: multi-line blockquote merging
    - FIXED: skip_references flag logic
    - OPTIMIZED: cleanup_text with compiled regex
    - OPTIMIZED: table column count normalization (pad short rows)

Dependencies:
    Required: pyyaml
    Optional: charset-normalizer (better encoding detection)
"""

import logging
import re
from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Dict, List, Optional, Tuple, Union

import yaml

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# ENUMS
# ═══════════════════════════════════════════════════════════════════════════════


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
    extra: Dict[str, str] = field(default_factory=dict)


@dataclass
class DocumentModel:
    """The complete parsed document"""

    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    sections: List[Section] = field(default_factory=list)
    elements: List[Element] = field(default_factory=list)
    footnotes: Dict[int, str] = field(default_factory=dict)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ═══════════════════════════════════════════════════════════════════════════════


class FrontmatterParser:
    """Parses YAML frontmatter from markdown"""

    _MARKDOWN_HEADING_RE = re.compile(r"^#{1,6}\s+")
    _BOLD_LABEL_RE = re.compile(r"\*\*[^*]+:\*\*")
    _SIMPLE_KEY_VALUE_RE = re.compile(r"^[A-Za-z0-9_.-]+\s*:\s*.*$")

    @staticmethod
    def parse(lines: List[str]) -> Tuple[DocumentMetadata, List[str]]:
        """
        Parse YAML frontmatter and return metadata + remaining lines.

        Args:
            lines: All lines from the markdown file

        Returns:
            Tuple of (DocumentMetadata, remaining_lines)
        """
        metadata = DocumentMetadata()

        if not lines or lines[0].strip() != "---":
            return metadata, lines

        # Find end of frontmatter
        content_start_idx = 0
        frontmatter_lines: List[str] = []

        for i, line in enumerate(lines[1:], 1):
            if line.strip() == "---":
                content_start_idx = i + 1
                break
            frontmatter_lines.append(line)

        if content_start_idx == 0:
            return metadata, lines

        if not FrontmatterParser._is_valid_frontmatter(frontmatter_lines):
            logger.debug("Frontmatter markers found, but content is not YAML frontmatter")
            return metadata, lines

        # Parse YAML content
        yaml_content = "\n".join(frontmatter_lines)
        try:
            parsed_data = yaml.safe_load(yaml_content) or {}
        except yaml.YAMLError:
            # Fallback to simple key: value parsing only for simple YAML-like blocks
            if not FrontmatterParser._is_simple_key_value_block(frontmatter_lines):
                logger.debug(
                    "Frontmatter YAML parse failed and fallback is not safe; "
                    "treating block as document content"
                )
                return metadata, lines
            parsed_data = FrontmatterParser._parse_simple_key_values(frontmatter_lines)

        if not isinstance(parsed_data, dict):
            logger.debug(
                "Frontmatter parsed to %s (not mapping); treating as document content",
                type(parsed_data).__name__,
            )
            return metadata, lines

        data = {
            str(key).strip().lower(): value for key, value in parsed_data.items() if key is not None
        }

        # Map to metadata fields
        metadata.title = str(data.get("title", metadata.title))
        metadata.subtitle = str(data.get("subtitle", metadata.subtitle))
        metadata.company = str(data.get("company", metadata.company))
        metadata.ticker = str(data.get("ticker", metadata.ticker))
        metadata.sector = str(data.get("sector", metadata.sector))
        metadata.analyst = str(data.get("analyst", metadata.analyst))

        # Store extra fields
        known_keys = {"title", "subtitle", "company", "ticker", "sector", "analyst"}
        metadata.extra = {k: str(v) for k, v in data.items() if k not in known_keys}

        return metadata, lines[content_start_idx:]

    @staticmethod
    def _is_valid_frontmatter(frontmatter_lines: List[str]) -> bool:
        """
        Validate that a frontmatter block looks like YAML metadata, not markdown content.

        This prevents accidental content loss when documents begin with horizontal rules.
        """
        if not frontmatter_lines:
            return False

        non_empty_count = 0
        key_value_count = 0
        empty_streak = 0

        for line in frontmatter_lines:
            stripped = line.strip()

            if not stripped:
                empty_streak += 1
                if empty_streak >= 2:
                    return False
                continue

            empty_streak = 0
            non_empty_count += 1

            # Markdown content indicators
            if FrontmatterParser._MARKDOWN_HEADING_RE.match(stripped):
                return False
            if FrontmatterParser._BOLD_LABEL_RE.search(stripped):
                return False
            if "**" in stripped:
                return False

            if FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                _, value = stripped.split(":", 1)
                if len(value.strip()) > 120:
                    return False
                key_value_count += 1
                continue

            # Long prose line without key-value shape is likely document content
            if len(stripped) > 120 and ":" not in stripped:
                return False

        return non_empty_count > 0 and key_value_count > 0

    @staticmethod
    def _is_simple_key_value_block(frontmatter_lines: List[str]) -> bool:
        """Check if all non-empty lines look like simple key: value pairs."""
        has_key_value = False

        for line in frontmatter_lines:
            stripped = line.strip()
            if not stripped:
                continue
            if not FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                return False
            has_key_value = True

        return has_key_value

    @staticmethod
    def _parse_simple_key_values(frontmatter_lines: List[str]) -> Dict[str, str]:
        """Parse simple key: value lines when YAML parsing fails."""
        data: Dict[str, str] = {}

        for line in frontmatter_lines:
            stripped = line.strip()
            if not stripped or not FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                continue
            key, value = stripped.split(":", 1)
            data[key.strip().lower()] = value.strip().strip('"').strip("'")

        return data


class TextParser:
    """Parses inline text formatting (bold, italic, inline LaTeX, etc.)"""

    # Compiled once — used by cleanup_text
    _ESCAPE_RE = re.compile(r'\\([~.*"\'()\[\]{}|_-])')

    # Inline formatting patterns
    _SUBSCRIPT_PATTERN = r"(?<!~)~[A-Za-z0-9]{1,8}~(?!~)"
    _INLINE_FORMAT_SPLIT_RE = re.compile(
        r"(\*\*[^*\n]+?\*\*|\^[^^\n]+?\^|"
        + _SUBSCRIPT_PATTERN
        + r"|(?<!\*)\*[^*\n]+?\*(?!\*)|(?<!\w)_[^_\n]+?_(?!\w))"
    )
    _COLOR_SPAN_RE = re.compile(
        r"<span\s+style=(['\"])(.*?)\1\s*>(.*?)</span>",
        re.IGNORECASE | re.DOTALL,
    )
    _COLOR_STYLE_RE = re.compile(r"color\s*:\s*(#[0-9A-Fa-f]{6})", re.IGNORECASE)

    # Inline LaTeX: $...$ but not $$...$$
    _INLINE_LATEX_RE = re.compile(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)")

    @classmethod
    def parse_runs(cls, text: str) -> List[TextRun]:
        """
        Parse text into runs with formatting.
        Handles **bold**, inline $LaTeX$, and combinations.
        """
        runs: List[TextRun] = []

        # Normalize escaped asterisks to regular bold markers
        normalized_text = text.replace(r"\*\*", "**")

        # ── Phase 1: Split on color spans, then inline LaTeX boundaries ─────
        color_segments = cls._split_on_color_spans(normalized_text)

        for segment_text, color_hex in color_segments:
            inline_segments = cls._split_on_inline_latex(segment_text)

            for inline_text, is_latex in inline_segments:
                if is_latex:
                    runs.append(
                        TextRun(
                            text=inline_text,
                            bold=False,
                            italic=False,
                            color_hex=color_hex,
                            is_latex=True,
                        )
                    )
                else:
                    runs.extend(cls._apply_color(cls._parse_inline_formatting(inline_text), color_hex))

        return runs

    @classmethod
    def _split_on_inline_latex(cls, text: str) -> List[Tuple[str, bool]]:
        """
        Split text into (content, is_latex) segments.

        Returns:
            List of (text, is_latex) tuples preserving order
        """
        segments: List[Tuple[str, bool]] = []
        last_end = 0

        for m in cls._INLINE_LATEX_RE.finditer(text):
            # Text before this LaTeX
            before = text[last_end : m.start()]
            if before:
                segments.append((before, False))

            # LaTeX expression (without $ delimiters)
            segments.append((m.group(1), True))
            last_end = m.end()

        # Remaining text after last LaTeX
        after = text[last_end:]
        if after:
            segments.append((after, False))

        # If no LaTeX found, return entire text as non-LaTeX
        if not segments:
            segments.append((text, False))

        return segments

    @classmethod
    def parse_runs_plain(cls, text: str) -> List[TextRun]:
        """
        Parse runs WITHOUT LaTeX detection.
        Used for contexts where $ should not be interpreted as LaTeX
        (e.g., table cells with currency values).
        """
        normalized = text.replace(r"\*\*", "**")
        runs: List[TextRun] = []
        for segment_text, color_hex in cls._split_on_color_spans(normalized):
            runs.extend(cls._apply_color(cls._parse_inline_formatting(segment_text), color_hex))
        return runs

    @classmethod
    def _split_on_color_spans(cls, text: str) -> List[Tuple[str, Optional[str]]]:
        """Split text into color-scoped segments preserving original order."""
        segments: List[Tuple[str, Optional[str]]] = []
        last_end = 0

        for match in cls._COLOR_SPAN_RE.finditer(text):
            color_hex = cls._extract_color_from_style(match.group(2))
            if color_hex is None:
                continue

            before = text[last_end : match.start()]
            if before:
                segments.append((before, None))

            segments.append((match.group(3), color_hex))
            last_end = match.end()

        after = text[last_end:]
        if after:
            segments.append((after, None))

        if not segments:
            segments.append((text, None))

        return segments

    @classmethod
    def _extract_color_from_style(cls, style_attr: str) -> Optional[str]:
        """Extract a normalized #RRGGBB color from an HTML style attribute."""
        match = cls._COLOR_STYLE_RE.search(style_attr)
        if not match:
            return None
        return match.group(1).upper()

    @staticmethod
    def _apply_color(runs: List[TextRun], color_hex: Optional[str]) -> List[TextRun]:
        """Apply a color value to each run in a parsed segment."""
        if not color_hex:
            return runs
        for run in runs:
            run.color_hex = color_hex
        return runs

    @classmethod
    def _parse_inline_formatting(cls, text: str) -> List[TextRun]:
        """Parse bold, italic, superscript, and subscript runs from plain text."""
        runs: List[TextRun] = []
        parts = cls._INLINE_FORMAT_SPLIT_RE.split(text)

        for part in parts:
            if not part:
                continue

            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                content = cls.cleanup_text(part[2:-2])
                if content:
                    runs.append(TextRun(text=content, bold=True))
                continue

            if part.startswith("^") and part.endswith("^") and len(part) > 2:
                content = cls.cleanup_text(part[1:-1])
                if content:
                    runs.append(TextRun(text=content, superscript=True))
                continue

            if part.startswith("~") and part.endswith("~") and len(part) > 2:
                content = cls.cleanup_text(part[1:-1])
                if content:
                    runs.append(TextRun(text=content, subscript=True))
                continue

            if part.startswith("*") and part.endswith("*") and len(part) > 2:
                content = cls.cleanup_text(part[1:-1])
                if content:
                    runs.append(TextRun(text=content, italic=True))
                continue

            if part.startswith("_") and part.endswith("_") and len(part) > 2:
                content = cls.cleanup_text(part[1:-1])
                if content:
                    runs.append(TextRun(text=content, italic=True))
                continue

            cleaned = cls.cleanup_text_preserve_spacing(part)
            if cleaned:
                runs.append(TextRun(text=cleaned))

        return runs

    @classmethod
    def has_inline_latex(cls, text: str) -> bool:
        """Check if text contains inline LaTeX expressions"""
        return bool(cls._INLINE_LATEX_RE.search(text))

    @staticmethod
    def cleanup_text(text: str) -> str:
        """Remove markdown artifacts and escape characters (single-pass regex)"""
        return TextParser._ESCAPE_RE.sub(r"\1", text).strip()

    @staticmethod
    def cleanup_text_preserve_spacing(text: str) -> str:
        """Remove markdown escape characters while preserving surrounding spacing."""
        return TextParser._ESCAPE_RE.sub(r"\1", text)


class FootnoteParser:
    """Extracts footnotes and references from markdown"""

    # Pattern for inline footnote markers like .1 or superscript ^1^
    INLINE_PATTERN = re.compile(r"(?:\.(\d+)(?=\s|$|[,;:\-])|\^(\d+)\^)")

    # Pattern for reference definitions like "1. Citation text"
    REFERENCE_PATTERN = re.compile(r"^(\d+)[\\.]\s+(.+)$")

    # Keywords that signal start of references section
    REFERENCE_KEYWORDS = frozenset(
        [
            "works cited",
            "references",
            "sources",
            "citations",
            "참고문헌",
            "출처",
        ]
    )

    @staticmethod
    def extract_references(lines: List[str]) -> Dict[int, str]:
        """
        Extract references from the end of the document.
        Looks for patterns like "1. Citation text" after references section.

        Only considers the *last* references section to avoid false positives.
        """
        references: Dict[int, str] = {}
        in_references = False

        for line in lines:
            stripped = line.strip()
            line_lower = stripped.lower()

            # Detect references section start
            if any(kw in line_lower for kw in FootnoteParser.REFERENCE_KEYWORDS):
                in_references = True
                continue

            if in_references:
                match = FootnoteParser.REFERENCE_PATTERN.match(stripped)
                if match:
                    ref_num = int(match.group(1))
                    ref_text = match.group(2).strip()
                    references[ref_num] = ref_text
                elif not stripped:
                    # Allow blank lines inside references section
                    continue
                else:
                    # Non-reference, non-blank line — end of section
                    # Keep going in case there is another references section
                    in_references = False

        return references

    @staticmethod
    def find_inline_references(text: str) -> List[int]:
        """Find all inline reference numbers in text"""
        refs: List[int] = []
        for match in FootnoteParser.INLINE_PATTERN.finditer(text):
            group = match.group(1) or match.group(2)
            if group:
                refs.append(int(group))
        return refs


class TableParser:
    """Parses markdown tables"""

    # Keywords for table type detection
    BEP_KEYWORDS = [
        "bep",
        "sensitivity",
        "cmr",
        "contribution margin",
        "fixed cost",
        "손익분기",
        "민감도",
        "고정비",
        "변동비",
    ]

    RISK_KEYWORDS = [
        "risk",
        "impact",
        "probability",
        "likelihood",
        "리스크",
        "위험",
        "영향",
        "확률",
    ]

    FINANCIAL_KEYWORDS = [
        "revenue",
        "income",
        "ebitda",
        "profit",
        "margin",
        "expense",
        "매출",
        "수익",
        "이익",
        "손익",
        "순이익",
        "영업",
        "비용",
    ]

    YEAR_INDICATORS = [
        "2024",
        "2025",
        "2026",
        "yoy",
        "cagr",
        "a)",
        "b)",
        "e)",
        "년도",
        "연도",
        "실적",
    ]

    UPSIDE_DOWNSIDE_KEYWORDS = [
        "upside",
        "downside",
        "상승",
        "하락",
        "요인",
    ]

    @staticmethod
    def parse(lines: List[str]) -> Table:
        """
        Parse markdown table lines into a Table object.

        Args:
            lines: Lines that make up the table (starting with |)
        """
        table = Table()

        # Filter out separator lines (|---|---|)
        data_lines = [
            line for line in lines if not set(line.strip()).issubset({"|", "-", " ", ":"})
        ]

        if not data_lines:
            return table

        # Parse alignments from separator line
        alignments = TableParser._parse_alignments(lines)

        # Get column count from first row
        first_row_cells = TableParser._split_row(data_lines[0])
        table.col_count = len(first_row_cells)
        table.alignments = (
            alignments if len(alignments) == table.col_count else ["left"] * table.col_count
        )

        # Detect table type
        header_text = " ".join(first_row_cells).lower()
        table.table_type = TableParser._detect_type(header_text)

        # Parse all rows — normalise column count per row
        for i, line in enumerate(data_lines):
            cells = TableParser._split_row(line)
            is_header = i == 0

            # Pad short rows with empty cells
            while len(cells) < table.col_count:
                cells.append("")

            # Truncate extra cells with warning (v3: no silent data loss)
            if len(cells) > table.col_count:
                logger.warning(
                    "Table row %d has %d columns (expected %d) - extra columns dropped: %s",
                    i,
                    len(cells),
                    table.col_count,
                    cells[table.col_count :],
                )
                cells = cells[: table.col_count]

            row = TableRow(is_header=is_header)
            for j, cell_text in enumerate(cells):
                cell = TableParser._parse_cell(
                    cell_text,
                    is_header=is_header,
                    col_idx=j,
                    row_idx=i,
                    total_rows=len(data_lines),
                    table_type=table.table_type,
                    header_cells=first_row_cells,
                )
                cell.alignment = table.alignments[j] if j < len(table.alignments) else "left"
                row.cells.append(cell)

            table.rows.append(row)

        return table

    @staticmethod
    def _split_row(line: str) -> List[str]:
        """Split a table row into cell contents, preserving intentional blanks."""
        cells = line.split("|")

        # Remove only the markdown boundary pipes, not genuine empty cells.
        if line.startswith("|") and cells:
            cells = cells[1:]
        if line.endswith("|") and cells:
            cells = cells[:-1]

        return [cell.strip() for cell in cells]

    @staticmethod
    def _parse_alignments(lines: List[str]) -> List[str]:
        """Parse column alignments from separator line"""
        alignments: List[str] = []
        for line in lines:
            if "-" in line and set(line.strip()).issubset({"|", "-", " ", ":"}):
                cells = line.split("|")
                for cell in cells:
                    cell = cell.strip()
                    if not cell:
                        continue
                    if cell.startswith(":") and cell.endswith(":"):
                        alignments.append("center")
                    elif cell.endswith(":"):
                        alignments.append("right")
                    else:
                        alignments.append("left")
                break
        return alignments

    @staticmethod
    def _detect_type(header_text: str) -> TableType:
        """Detect table type from header row"""
        if any(kw in header_text for kw in TableParser.UPSIDE_DOWNSIDE_KEYWORDS):
            return TableType.UPSIDE_DOWNSIDE

        if any(kw in header_text for kw in TableParser.BEP_KEYWORDS):
            return TableType.BEP_SENSITIVITY

        if any(kw in header_text for kw in TableParser.RISK_KEYWORDS):
            return TableType.RISK_MATRIX

        has_financial = any(kw in header_text for kw in TableParser.FINANCIAL_KEYWORDS)
        has_year = any(yi in header_text for yi in TableParser.YEAR_INDICATORS)
        if has_financial or has_year:
            return TableType.FINANCIAL

        return TableType.GENERIC

    @staticmethod
    def _parse_cell(
        text: str,
        is_header: bool,
        col_idx: int,
        row_idx: int,
        total_rows: int,
        table_type: TableType,
        header_cells: List[str],
    ) -> TableCell:
        """Parse a single table cell"""
        cell = TableCell(content=text, is_header=is_header)

        # Prefer LaTeX-aware parsing only when a balanced inline expression exists.
        if TextParser.has_inline_latex(text):
            cell.runs = TextParser.parse_runs(text)
        else:
            cell.runs = TextParser.parse_runs_plain(text)

        # Detect numeric content
        cell.is_numeric = any(char.isdigit() for char in text) and col_idx > 0

        # Detect negative numbers
        cell.is_negative = (
            text.startswith("(") and text.endswith(")") and any(c.isdigit() for c in text)
        ) or (text.startswith("-") and any(c.isdigit() for c in text))

        # Table-type specific detection
        if table_type == TableType.BEP_SENSITIVITY:
            cell.is_base_case = (
                "base" in text.lower()
                or "기준" in text
                or (
                    row_idx == total_rows // 2 and col_idx == len(header_cells) // 2 and row_idx > 0
                )
            )

        elif table_type == TableType.RISK_MATRIX:
            header_lower = header_cells[col_idx].lower() if col_idx < len(header_cells) else ""
            risk_header_kw = ("impact", "probability", "영향", "확률")
            if any(kw in header_lower for kw in risk_header_kw):
                text_lower = text.lower()
                if "high" in text_lower or "높" in text:
                    cell.risk_level = "high"
                elif any(kw in text_lower for kw in ("medium", "moderate")) or "중" in text:
                    cell.risk_level = "medium"
                elif "low" in text_lower or "낮" in text:
                    cell.risk_level = "low"

        return cell


# ═══════════════════════════════════════════════════════════════════════════════
# LaTeX PARSER (NEW v3)
# ═══════════════════════════════════════════════════════════════════════════════


class LaTeXParser:
    """
    Parses LaTeX equations from markdown.

    Handles:
        - Block equations: $$ ... $$ (single-line and multi-line)
        - Inline equations: $ ... $ (detected within paragraphs)
        - Escaped dollar signs: \\$ (not treated as LaTeX)
    """

    # Single-line block equation: $$ E = mc^2 $$
    BLOCK_SINGLE_LINE_RE = re.compile(r"^\$\$(.+?)\$\$\s*$")

    # Block equation delimiter (start or end of multi-line)
    BLOCK_DELIMITER_RE = re.compile(r"^\$\$\s*$")

    # Inline LaTeX: $...$ but not $$...$$, not escaped \$
    INLINE_RE = re.compile(r"(?<!\$)(?<!\\)\$(?!\$)(.+?)(?<!\$)(?<!\\)\$(?!\$)")

    @classmethod
    def is_block_start(cls, line: str) -> bool:
        """Check if line starts a block equation"""
        stripped = line.strip()
        return bool(cls.BLOCK_DELIMITER_RE.match(stripped))

    @classmethod
    def is_block_single_line(cls, line: str) -> Optional[str]:
        """
        Check if line is a single-line block equation.

        Returns:
            The LaTeX expression if matched, None otherwise
        """
        m = cls.BLOCK_SINGLE_LINE_RE.match(line.strip())
        return m.group(1).strip() if m else None

    @classmethod
    def has_inline(cls, text: str) -> bool:
        """Check if text contains inline LaTeX"""
        return bool(cls.INLINE_RE.search(text))

    @classmethod
    def extract_inline_segments(cls, text: str) -> List[Tuple[str, bool]]:
        """
        Split text into (content, is_latex) segments.

        Returns:
            List of (text, is_latex) tuples preserving source order
        """
        segments: List[Tuple[str, bool]] = []
        last_end = 0

        for m in cls.INLINE_RE.finditer(text):
            before = text[last_end : m.start()]
            if before:
                segments.append((before, False))
            segments.append((m.group(1), True))
            last_end = m.end()

        after = text[last_end:]
        if after:
            segments.append((after, False))

        if not segments:
            segments.append((text, False))

        return segments


# ═══════════════════════════════════════════════════════════════════════════════
# BASE64 IMAGE PARSER (NEW v3)
# ═══════════════════════════════════════════════════════════════════════════════


class Base64ImageParser:
    """
    Parses Base64-encoded images embedded in markdown.

    Handles:
        ![alt text](data:image/png;base64,iVBORw0KGgo...)
        ![alt text](data:image/jpeg;base64,/9j/4AAQ...)
        ![alt text](data:image/svg+xml;base64,PHN2Zy...)
    """

    # Full pattern for Base64 image markdown
    PATTERN = re.compile(
        r"^!\[([^\]]*)\]"  # ![alt text]
        r"\("  # (
        r"data:(image/[a-zA-Z0-9+.-]+)"  #   data:image/type
        r";base64,"  #   ;base64,
        r"([A-Za-z0-9+/=\s]+)"  #   base64 data
        r"\)\s*$"  # )
    )

    # Looser pattern for detection (may span part of a longer line)
    DETECT_RE = re.compile(r"!\[[^\]]*\]\(data:image/[a-zA-Z0-9+.-]+;base64,")

    @classmethod
    def parse(cls, line: str) -> Optional[Image]:
        """
        Parse a Base64 image line.

        Args:
            line: A markdown line potentially containing a Base64 image

        Returns:
            Image object with base64_data populated, or None
        """
        m = cls.PATTERN.match(line.strip())
        if not m:
            return None

        alt_text = m.group(1)
        mime_type = m.group(2)
        b64_data = m.group(3).replace("\n", "").replace("\r", "").replace(" ", "")

        # Basic validation: Base64 length should be reasonable
        if len(b64_data) < 4:
            logger.warning(
                "Base64 image data too short (%d chars) — skipping",
                len(b64_data),
            )
            return None

        return Image(
            alt_text=alt_text,
            path="",
            base64_data=b64_data,
            mime_type=mime_type,
        )

    @classmethod
    def is_base64_image(cls, line: str) -> bool:
        """Quick check if line contains a Base64 image"""
        return bool(cls.DETECT_RE.search(line))

    @classmethod
    def parse_multiline(cls, lines: List[str], start_idx: int) -> Tuple[Optional[Image], int]:
        """
        Parse a Base64 image that may span multiple lines.

        Some editors wrap long Base64 data across lines. This method
        concatenates lines until the closing ) is found.

        Args:
            lines: All document lines
            start_idx: Index of the line containing ![

        Returns:
            Tuple of (Image or None, next_line_index)
        """
        if start_idx >= len(lines):
            return None, start_idx + 1

        # Try single-line first
        single = cls.parse(lines[start_idx])
        if single:
            return single, start_idx + 1

        # Multi-line: concatenate until closing parenthesis
        if not cls.DETECT_RE.search(lines[start_idx]):
            return None, start_idx + 1

        combined = lines[start_idx].rstrip()
        idx = start_idx + 1

        # Limit lookahead to prevent runaway concatenation
        max_lookahead = 50
        while idx < len(lines) and idx - start_idx < max_lookahead:
            line = lines[idx].strip()
            combined += line
            idx += 1
            if line.endswith(")"):
                break

        result = cls.parse(combined)
        if result:
            return result, idx
        else:
            logger.warning(
                "Failed to parse multi-line Base64 image starting at line %d",
                start_idx,
            )
            return None, start_idx + 1


# ═══════════════════════════════════════════════════════════════════════════════
# MARKDOWN PARSER (MAIN)
# ═══════════════════════════════════════════════════════════════════════════════


class MarkdownParser:
    """Main parser for markdown documents"""

    # ── Heading patterns ────────────────────────────────────────────────────
    H1_PATTERN = re.compile(r"^#\s+(.+)$")
    H2_PATTERN = re.compile(r"^##\s+(.+)$")
    H3_PATTERN = re.compile(r"^###\s+(.+)$")
    H4_PATTERN = re.compile(r"^####\s+(.+)$")

    # Numbered heading pattern: **1. Title** or **1\. Title**
    NUMBERED_HEADING_PATTERN = re.compile(r"^\*\*\d+(\.|\\.)")

    # ── List patterns ───────────────────────────────────────────────────────
    BULLET_PATTERN = re.compile(r"^([ \t]*)[-*]\s+(.+)$")
    NUMBERED_LIST_PATTERN = re.compile(r"^([ \t]*)(\d+)\.\s+(.+)$")

    # Heuristic: numbered line that looks like a section heading
    _NUMBERED_HEADING_HEURISTIC = re.compile(r"^(\d{1,2})\.\s+([가-힣A-Za-z][\w\s가-힣]{0,40})$")

    # ── Other patterns ──────────────────────────────────────────────────────
    BLOCKQUOTE_PATTERN = re.compile(r"^>\s+(.+)$")
    IMAGE_PATTERN = re.compile(r"^!\[(.*?)\]\((.*?)\)$")
    TABLE_START_PATTERN = re.compile(r"^\|")
    SEPARATOR_PATTERN = re.compile(r"^(---|## ---)$")
    CODE_FENCE_PATTERN = re.compile(r"^```([A-Za-z0-9_+-]*)\s*$")

    # Reference section keywords
    _REFERENCE_KEYWORDS = frozenset(
        [
            "works cited",
            "references",
            "sources",
            "citations",
            "참고문헌",
            "출처",
        ]
    )

    _HTML_BREAK_RE = re.compile(r"<br\s*/?>\s*$", re.IGNORECASE)
    _MULTISPACE_RE = re.compile(r"[ \t]{2,}")
    _OPEN_PAREN_SPACE_RE = re.compile(r"\(\s+")
    _CLOSE_PAREN_SPACE_RE = re.compile(r"\s+\)")

    def __init__(self, preserve_trailing_double_space_break: bool = False):
        """
        Initialize parser behavior flags.

        Args:
            preserve_trailing_double_space_break:
                When True, treat markdown trailing 2+ spaces as hard line breaks.
                Default False to avoid accidental Shift+Enter artifacts from noisy input.
        """
        self.preserve_trailing_double_space_break = preserve_trailing_double_space_break

    def parse(self, content: str) -> DocumentModel:
        """
        Parse markdown content into a DocumentModel.

        Args:
            content: The full markdown content

        Returns:
            A DocumentModel with all parsed elements
        """
        lines = content.split("\n")

        # Parse frontmatter
        metadata, remaining_lines = FrontmatterParser.parse(lines)

        # Extract footnotes/references
        footnotes = FootnoteParser.extract_references(remaining_lines)

        # Parse elements
        elements = self._parse_elements(remaining_lines)

        return DocumentModel(
            metadata=metadata,
            elements=elements,
            footnotes=footnotes,
        )

    # ── Element-level parsing ───────────────────────────────────────────────

    def _parse_elements(self, lines: List[str]) -> List[Element]:
        """Parse lines into elements"""
        elements: List[Element] = []
        i = 0
        in_references = False

        while i < len(lines):
            raw_line = lines[i]
            line = raw_line.strip()

            # ── References section gating ───────────────────────────────────
            if self._is_reference_header(line):
                in_references = True
                i += 1
                continue

            if in_references:
                if FootnoteParser.REFERENCE_PATTERN.match(line) or not line:
                    i += 1
                    continue
                if self._looks_like_heading(line):
                    in_references = False
                else:
                    i += 1
                    continue

            # ── Empty line ──────────────────────────────────────────────────
            if not line:
                i += 1
                continue

            # ── Separator ───────────────────────────────────────────────────
            if self.SEPARATOR_PATTERN.match(line):
                i += 1
                continue

            # ── Fenced code block ───────────────────────────────────────────
            code_match = self.CODE_FENCE_PATTERN.match(line)
            if code_match:
                language = code_match.group(1).strip()
                code_lines: List[str] = []
                i += 1
                while i < len(lines):
                    if self.CODE_FENCE_PATTERN.match(lines[i].strip()):
                        i += 1
                        break
                    code_lines.append(lines[i].rstrip("\n"))
                    i += 1

                elements.append(
                    Element(
                        element_type=ElementType.CODE_BLOCK,
                        content=CodeBlock(code="\n".join(code_lines).rstrip(), language=language),
                        raw_text=raw_line,
                    )
                )
                continue

            # ════════════════════════════════════════════════════════════════
            # NEW (v3): LaTeX block equations — checked early (high priority)
            # ════════════════════════════════════════════════════════════════

            # Case A: Single-line block: $$ E = mc^2 $$
            latex_expr = LaTeXParser.is_block_single_line(line)
            if latex_expr is not None:
                elements.append(
                    Element(
                        element_type=ElementType.LATEX_BLOCK,
                        content=LaTeXEquation(expression=latex_expr, is_block=True),
                        raw_text=line,
                    )
                )
                i += 1
                continue

            # Case B: Multi-line block: $$ (start delimiter)
            if LaTeXParser.is_block_start(line):
                latex_lines: List[str] = []
                i += 1
                while i < len(lines):
                    if LaTeXParser.is_block_start(lines[i]):
                        i += 1
                        break
                    latex_lines.append(lines[i])
                    i += 1

                expression = "\n".join(latex_lines).strip()
                if expression:
                    elements.append(
                        Element(
                            element_type=ElementType.LATEX_BLOCK,
                            content=LaTeXEquation(expression=expression, is_block=True),
                            raw_text=f"$$\n{expression}\n$$",
                        )
                    )
                else:
                    logger.warning("Empty LaTeX block equation at line %d — skipped", i)
                continue

            # ════════════════════════════════════════════════════════════════
            # NEW (v3): Base64 embedded images — checked before regular images
            # ════════════════════════════════════════════════════════════════

            if Base64ImageParser.is_base64_image(line):
                image, next_idx = Base64ImageParser.parse_multiline(lines, i)
                if image:
                    elements.append(
                        Element(
                            element_type=ElementType.IMAGE,
                            content=image,
                            raw_text=line[:100] + "..." if len(line) > 100 else line,
                        )
                    )
                    i = next_idx
                    continue
                # Fall through to regular parsing if Base64 parse failed

            # ── Table (collect all contiguous table lines) ──────────────────
            if self.TABLE_START_PATTERN.match(line):
                table_lines: List[str] = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    table_lines.append(lines[i].strip())
                    i += 1
                table = TableParser.parse(table_lines)
                elements.append(
                    Element(
                        element_type=ElementType.TABLE,
                        content=table,
                        raw_text="\n".join(table_lines),
                    )
                )
                continue

            # ── Headings (must be checked before numbered list) ─────────────
            element = self._try_parse_heading(line)
            if element:
                elements.append(element)
                i += 1
                continue

            # ── Blockquote (merge consecutive > lines) ──────────────────────
            match = self.BLOCKQUOTE_PATTERN.match(line)
            if match:
                bq_lines: List[str] = []
                while i < len(lines):
                    bq_match = self.BLOCKQUOTE_PATTERN.match(lines[i].strip())
                    if bq_match:
                        bq_lines.append(TextParser.cleanup_text(bq_match.group(1)))
                        i += 1
                    else:
                        break

                title, body = self._extract_blockquote_title(bq_lines)
                elements.append(
                    Element(
                        element_type=ElementType.BLOCKQUOTE,
                        content=Blockquote(text=body, title=title),
                        raw_text="\n".join(bq_lines),
                    )
                )
                continue

            # ── Bullet list ─────────────────────────────────────────────────
            match = self.BULLET_PATTERN.match(raw_line)
            if match:
                indent_level = self._get_indent_level(match.group(1))
                text = match.group(2)
                item = ListItem(
                    text=text,
                    runs=TextParser.parse_runs(text),
                    indent_level=indent_level,
                )
                elements.append(
                    Element(
                        element_type=ElementType.BULLET_LIST,
                        content=item,
                        raw_text=raw_line,
                    )
                )
                i += 1
                continue

            # ── Numbered list ───────────────────────────────────────────────
            match = self.NUMBERED_LIST_PATTERN.match(raw_line)
            if match and not self._is_numbered_heading(line):
                indent_level = self._get_indent_level(match.group(1))
                number = match.group(2)
                text = match.group(3)
                item = ListItem(
                    text=text,
                    runs=TextParser.parse_runs(text),
                    indent_level=indent_level,
                )
                elements.append(
                    Element(
                        element_type=ElementType.NUMBERED_LIST,
                        content=(number, item),
                        raw_text=raw_line,
                    )
                )
                i += 1
                continue

            # ── Numbered heading fallback (e.g. "1. 서론") ─────────────────
            if match and self._is_numbered_heading(line):
                full_text = TextParser.cleanup_text(line)
                elements.append(
                    Element(
                        element_type=ElementType.NUMBERED_HEADING,
                        content=Heading(level=2, text=full_text, is_numbered=True),
                        raw_text=line,
                    )
                )
                i += 1
                continue

            # ── Regular image (non-Base64) ──────────────────────────────────
            match = self.IMAGE_PATTERN.match(line)
            if match:
                elements.append(
                    Element(
                        element_type=ElementType.IMAGE,
                        content=Image(alt_text=match.group(1), path=match.group(2)),
                        raw_text=line,
                    )
                )
                i += 1
                continue

            # ── Paragraph (with inline LaTeX detection) ─────────────────────
            paragraph_text, next_idx = self._collect_paragraph(lines, i)
            para_element = self._parse_paragraph(paragraph_text)
            elements.append(para_element)
            i = next_idx

        return elements

    def _collect_paragraph(self, lines: List[str], start_idx: int) -> Tuple[str, int]:
        """
        Collect contiguous paragraph lines and normalize Markdown line breaks.

        CommonMark semantics:
            - soft line break (single newline in paragraph) -> space
            - hard line break (2+ trailing spaces, trailing backslash, <br>) -> newline
        """
        parts: List[str] = []
        i = start_idx
        use_hard_break = False

        while i < len(lines):
            raw_line = lines[i]
            line = raw_line.strip()

            if not line:
                break

            if i > start_idx and self._starts_new_block(line):
                break

            normalized, hard_break = self._normalize_paragraph_line(raw_line)
            if normalized:
                if parts:
                    parts.append("\n" if use_hard_break else " ")
                parts.append(normalized)

            use_hard_break = hard_break
            i += 1

        paragraph_text = "".join(parts).strip()
        return paragraph_text, i

    def _starts_new_block(self, line: str) -> bool:
        """Return True when line should start a new non-paragraph block."""
        return bool(
            self.SEPARATOR_PATTERN.match(line)
            or LaTeXParser.is_block_single_line(line) is not None
            or LaTeXParser.is_block_start(line)
            or Base64ImageParser.is_base64_image(line)
            or self.TABLE_START_PATTERN.match(line)
            or self._try_parse_heading(line)
            or self.BLOCKQUOTE_PATTERN.match(line)
            or self.BULLET_PATTERN.match(line)
            or self.NUMBERED_LIST_PATTERN.match(line)
            or self.IMAGE_PATTERN.match(line)
            or self._is_reference_header(line)
        )

    def _normalize_paragraph_line(self, raw_line: str) -> Tuple[str, bool]:
        """Normalize one paragraph source line and detect hard line breaks."""
        hard_break = False
        working = raw_line

        if self._HTML_BREAK_RE.search(working):
            hard_break = True
            working = self._HTML_BREAK_RE.sub("", working).rstrip()
        elif re.search(r"(?<!\\)\\\s*$", working):
            hard_break = True
            working = re.sub(r"\\\s*$", "", working).rstrip()
        elif self.preserve_trailing_double_space_break and re.search(r"[ \t]{2,}$", working):
            hard_break = True
            working = working.rstrip()

        return self._normalize_inline_spacing(working.strip()), hard_break

    @classmethod
    def _normalize_inline_spacing(cls, text: str) -> str:
        """Normalize excessive inline spacing for paragraph text."""
        if not text:
            return text

        normalized = cls._MULTISPACE_RE.sub(" ", text)
        normalized = cls._OPEN_PAREN_SPACE_RE.sub("(", normalized)
        normalized = cls._CLOSE_PAREN_SPACE_RE.sub(")", normalized)
        return normalized.strip()

    @staticmethod
    def _get_indent_level(prefix: str) -> int:
        """Convert leading markdown indentation to a list nesting level."""
        expanded = prefix.replace("\t", "    ")
        if not expanded:
            return 0
        return max(0, len(expanded) // 2)

    # ── Paragraph parsing (ENHANCED v3) ─────────────────────────────────────

    def _parse_paragraph(self, line: str) -> Element:
        """
        Parse a line as a paragraph, detecting inline LaTeX if present.

        If the line contains inline $...$ expressions, the resulting
        TextRuns will have is_latex=True for those segments, allowing
        the renderer to handle them appropriately.
        """
        has_latex = LaTeXParser.has_inline(line)

        para = Paragraph(
            text=line,
            runs=TextParser.parse_runs(line),
            has_inline_latex=has_latex,
        )

        return Element(
            element_type=ElementType.PARAGRAPH,
            content=para,
            raw_text=line,
        )

    # ── Heading parsing helpers ─────────────────────────────────────────────

    def _try_parse_heading(self, line: str) -> Optional[Element]:
        """Try to parse line as a heading"""

        # Numbered heading (**1. Title**)
        if self.NUMBERED_HEADING_PATTERN.match(line):
            text = TextParser.cleanup_text(line)
            return Element(
                element_type=ElementType.NUMBERED_HEADING,
                content=Heading(level=1, text=text, is_numbered=True),
                raw_text=line,
            )

        # H4 (check longer prefixes first to avoid partial match)
        match = self.H4_PATTERN.match(line)
        if match:
            text = TextParser.cleanup_text(match.group(1))
            return Element(
                element_type=ElementType.HEADING_4,
                content=Heading(level=4, text=text),
                raw_text=line,
            )

        # H3
        match = self.H3_PATTERN.match(line)
        if match:
            text = TextParser.cleanup_text(match.group(1))
            return Element(
                element_type=ElementType.HEADING_3,
                content=Heading(level=3, text=text),
                raw_text=line,
            )

        # H2
        match = self.H2_PATTERN.match(line)
        if match:
            text = TextParser.cleanup_text(match.group(1))
            return Element(
                element_type=ElementType.HEADING_2,
                content=Heading(level=2, text=text),
                raw_text=line,
            )

        # H1
        match = self.H1_PATTERN.match(line)
        if match:
            text = TextParser.cleanup_text(match.group(1))
            return Element(
                element_type=ElementType.HEADING_1,
                content=Heading(level=1, text=text),
                raw_text=line,
            )

        return None

    def _is_numbered_heading(self, line: str) -> bool:
        """
        Determine if a numbered line (e.g. "1. 서론") is a section heading
        rather than a list item.

        Heuristics:
            - Short title (≤ ~40 chars after the number)
            - Starts with Korean or uppercase English
            - Does NOT contain sentence-ending punctuation mid-line
            - Number ≤ 20 (unlikely section numbers above this)
        """
        match = self._NUMBERED_HEADING_HEURISTIC.match(line.strip())
        if not match:
            return False

        number = int(match.group(1))
        title_part = match.group(2).strip()

        # Reject unreasonably high section numbers
        if number > 20:
            return False

        if len(title_part) <= 40:
            # Reject if it ends with sentence-ending patterns
            if re.search(r"[다요음함임됨것수점]\.$", title_part):
                return False
            return True

        return False

    def _looks_like_heading(self, line: str) -> bool:
        """Quick check if a line looks like any kind of heading"""
        return bool(
            self.H1_PATTERN.match(line)
            or self.H2_PATTERN.match(line)
            or self.H3_PATTERN.match(line)
            or self.H4_PATTERN.match(line)
            or self.NUMBERED_HEADING_PATTERN.match(line)
        )

    def _is_reference_header(self, line: str) -> bool:
        """Check if line is a references section header"""
        stripped = line.strip().lower()
        cleaned = re.sub(r"^#{1,4}\s+", "", stripped)
        return any(kw in cleaned for kw in self._REFERENCE_KEYWORDS)

    # ── Blockquote helpers ──────────────────────────────────────────────────

    @staticmethod
    def _extract_blockquote_title(
        bq_lines: List[str],
    ) -> Tuple[str, str]:
        """
        Extract title from blockquote lines.
        If first line matches [시사점], [참고], etc., use it as title.

        Returns:
            (title, body_text)
        """
        title = "KEY INSIGHT"
        body_lines = bq_lines

        if bq_lines:
            first = bq_lines[0]
            label_match = re.match(
                r"^\[(시사점|참고|주의|결론|요약|핵심|"
                r"KEY INSIGHT|NOTE|WARNING)\]\s*(.*)",
                first,
                re.IGNORECASE,
            )
            if label_match:
                title = label_match.group(1).upper()
                remainder = label_match.group(2).strip()
                body_lines = ([remainder] if remainder else []) + bq_lines[1:]

        body = " ".join(body_lines).strip()
        return title, body


# ═══════════════════════════════════════════════════════════════════════════════
# ENCODING UTILITIES (ENHANCED v3)
# ═══════════════════════════════════════════════════════════════════════════════


def _read_with_encoding(file_path: str) -> str:
    """
    Read file content with intelligent encoding detection.

    Strategy:
        1. charset_normalizer (if available) — statistical detection
        2. BOM detection — UTF-8 BOM
        3. Sequential fallback — UTF-8 → EUC-KR → CP949

    Args:
        file_path: Path to file

    Returns:
        File content as string

    Raises:
        UnicodeDecodeError: If all detection methods fail
    """
    from pathlib import Path

    raw_bytes = Path(file_path).read_bytes()

    # Strategy 1: charset_normalizer (optional dependency)
    try:
        from charset_normalizer import from_bytes  # pyright: ignore[reportMissingImports]

        result = from_bytes(raw_bytes).best()
        if result and result.encoding:
            logger.debug(
                "Encoding detected by charset_normalizer: %s (confidence: %.1f%%)",
                result.encoding,
                result.encoding if hasattr(result, "encoding") else 0,
            )
            return str(result)
    except ImportError:
        pass
    except Exception as e:
        logger.debug("charset_normalizer failed: %s — falling back", e)

    # Strategy 2: BOM detection
    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        return raw_bytes.decode("utf-8-sig")

    # Strategy 3: Sequential fallback
    encodings = ["utf-8", "euc-kr", "cp949"]
    for enc in encodings:
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue

    raise UnicodeDecodeError(
        "multiple",
        b"",
        0,
        1,
        f"Failed to decode {file_path} with encodings: {encodings}",
    )


# ═══════════════════════════════════════════════════════════════════════════════
# CONVENIENCE FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════


def _decode_bytes(raw_bytes: bytes, label: str = "<stream>") -> str:
    """Decode raw bytes with intelligent encoding detection.

    Same strategy as _read_with_encoding but works on bytes directly.
    """
    # charset_normalizer (optional)
    try:
        from charset_normalizer import from_bytes  # pyright: ignore[reportMissingImports]

        result = from_bytes(raw_bytes).best()
        if result and result.encoding:
            return str(result)
    except ImportError:
        pass
    except Exception:
        pass

    # BOM detection
    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        return raw_bytes.decode("utf-8-sig")

    # Sequential fallback
    encodings = ["utf-8", "euc-kr", "cp949"]
    for enc in encodings:
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue

    raise UnicodeDecodeError(
        "multiple", b"", 0, 1,
        f"Failed to decode {label} with encodings: {encodings}",
    )


def parse_markdown_file(source: "Union[str, BinaryIO]") -> DocumentModel:
    """
    Parse a markdown file or stream into a DocumentModel.

    Accepts a file path (str) or a binary stream (BinaryIO).
    Tries intelligent encoding detection, falls back through
    UTF-8 → EUC-KR → CP949.

    Args:
        source: Path to the markdown file, or a readable binary stream.

    Returns:
        A DocumentModel containing all parsed content

    Raises:
        FileNotFoundError: If source is a path and the file does not exist
        UnicodeDecodeError: If none of the attempted encodings work
    """
    from stream_utils import is_stream

    if is_stream(source):
        raw_bytes = source.read()
        content = _decode_bytes(raw_bytes, label="<stream>")
        label = "<stream>"
    else:
        content = _read_with_encoding(source)
        label = source

    lines = content.splitlines()
    _, remaining_lines = FrontmatterParser.parse(lines)
    frontmatter_present = remaining_lines is not lines

    parser = MarkdownParser()
    model = parser.parse(content)
    _infer_metadata_from_elements(model, allow_company_inference=not frontmatter_present)

    logger.info(
        "Parsed %s: %d elements, %d footnotes, latex_blocks=%d, images=%d",
        label,
        len(model.elements),
        len(model.footnotes),
        sum(1 for e in model.elements if e.element_type == ElementType.LATEX_BLOCK),
        sum(1 for e in model.elements if e.element_type == ElementType.IMAGE),
    )

    return model


def _infer_metadata_from_elements(
    model: DocumentModel,
    allow_company_inference: bool = True,
) -> None:
    """Backfill metadata from document content when frontmatter is absent or partial."""
    metadata = model.metadata

    if metadata.title == "IB Report":
        first_heading = next(
            (
                element.content.text.strip()
                for element in model.elements
                if element.element_type == ElementType.HEADING_1
                and isinstance(element.content, Heading)
                and element.content.text.strip()
            ),
            "",
        )
        if first_heading:
            metadata.title = first_heading

    if allow_company_inference and metadata.company == "Korea Development Bank":
        inferred_company = _infer_company_from_title(metadata.title)
        if inferred_company:
            metadata.extra.setdefault("subject_company", inferred_company)

    label_map = {
        "기준일": ("extra", "date"),
        "as of": ("extra", "date"),
        "작성일": ("extra", "date"),
        "report date": ("extra", "date"),
        "date": ("extra", "date"),
        "분석 대상 기간": ("extra", "analysis_period"),
        "analysis period": ("extra", "analysis_period"),
        "분석 기준": ("extra", "analysis_basis"),
        "analysis basis": ("extra", "analysis_basis"),
        "prepared by": ("analyst", None),
        "작성자": ("analyst", None),
        "analyst": ("analyst", None),
        "institution": ("company", None),
        "company": ("company", None),
        "기관": ("company", None),
        "sector": ("sector", None),
        "업종": ("sector", None),
        "prepared for": ("extra", "recipient"),
        "recipient": ("extra", "recipient"),
        "수신": ("extra", "recipient"),
    }

    for element in model.elements:
        if element.element_type not in {ElementType.PARAGRAPH, ElementType.HEADING_2}:
            continue
        if element.element_type == ElementType.HEADING_2:
            break
        if not isinstance(element.content, Paragraph):
            continue

        extracted = _extract_leading_metadata_pair(element.content)
        if extracted is None:
            continue

        label, value = extracted
        mapped = label_map.get(label.lower())
        if not mapped or not value:
            continue

        field_name, extra_key = mapped
        if field_name == "extra" and extra_key:
            metadata.extra.setdefault(extra_key, value)
        elif field_name and getattr(metadata, field_name, "") in {"", "IB Report", "SECTOR", "DCM Team 1", "Korea Development Bank"}:
            setattr(metadata, field_name, value)


def _extract_leading_metadata_pair(paragraph: Paragraph) -> Optional[Tuple[str, str]]:
    """Extract a bold-label metadata pair like '**작성일:** 2026-03-20' from a paragraph."""
    if not paragraph.runs:
        return None

    first_run = paragraph.runs[0]
    if not first_run.bold:
        return None

    label = first_run.text.strip().rstrip(":").rstrip("：").strip()
    if not label:
        return None

    value = "".join(run.text for run in paragraph.runs[1:]).strip()
    if not value and ":" in first_run.text:
        raw_label, raw_value = first_run.text.split(":", 1)
        label = raw_label.strip().rstrip("：").strip()
        value = raw_value.strip()
    if not value:
        return None

    return label, value


def _infer_company_from_title(title: str) -> str:
    """Infer the company/subject name from a report-style title."""
    if not title or title == "IB Report":
        return ""

    for delimiter in (" — ", " - ", " – ", ": "):
        if delimiter in title:
            candidate = title.split(delimiter, 1)[0].strip()
            if candidate:
                return candidate

    suffix_patterns = [
        r"\s+수익성\s+변화\s+분석\s+보고서$",
        r"\s+수익성\s+분석\s+보고서$",
        r"\s+분석\s+보고서$",
        r"\s+보고서$",
        r"\s+리포트$",
    ]

    for pattern in suffix_patterns:
        candidate = re.sub(pattern, "", title).strip()
        if candidate and candidate != title:
            return candidate

    return ""
