"""
MD Parser Module for IB Style Word Report Converter
Handles parsing of Markdown files including frontmatter, elements, tables,
LaTeX equations, Base64 images, and footnotes.

Changelog (2026-09-29):
    - FIXED: Preserve escaped inline syntax and non-reference body content.
    - FIXED: Share fence boundaries and normalize Markdown block input.
    - FIXED: Retain inferred IB subtitle headings and separate header metadata lines.

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
from pathlib import Path
from typing import BinaryIO, Dict, List, Match, Optional, Set, Tuple, Union, cast

import yaml

from document_model import (
    Blockquote as Blockquote,
)
from document_model import (
    BulletList as BulletList,
)
from document_model import (
    CodeBlock as CodeBlock,
)
from document_model import (
    Diagram as Diagram,
)
from document_model import (
    DiagramArrow as DiagramArrow,
)
from document_model import (
    DiagramBox as DiagramBox,
)
from document_model import (
    DocumentMetadata as DocumentMetadata,
)
from document_model import (
    DocumentModel as DocumentModel,
)
from document_model import (
    Element as Element,
)
from document_model import (
    ElementContent as ElementContent,
)

# ═══════════════════════════════════════════════════════════════════════════════
# ENUMS
# ═══════════════════════════════════════════════════════════════════════════════
from document_model import (
    ElementType as ElementType,
)
from document_model import (
    Footnote as Footnote,
)
from document_model import (
    Heading as Heading,
)
from document_model import (
    Image as Image,
)
from document_model import (
    LaTeXEquation as LaTeXEquation,
)
from document_model import (
    ListItem as ListItem,
)
from document_model import (
    NumberedList as NumberedList,
)
from document_model import (
    Paragraph as Paragraph,
)
from document_model import (
    Section as Section,
)
from document_model import (
    Table as Table,
)
from document_model import (
    TableCell as TableCell,
)
from document_model import (
    TableRow as TableRow,
)
from document_model import (
    TableType as TableType,
)
from document_model import (
    TextRun as TextRun,
)
from document_profiles import apply_table_specs, default_metadata, get_profile

logger = logging.getLogger(__name__)


class FrontmatterParser:
    """Parses YAML frontmatter from markdown"""

    _MARKDOWN_HEADING_RE = re.compile(r"^#{1,6}\s+")
    _BOLD_LABEL_RE = re.compile(r"\*\*[^*]+:\*\*")
    _SIMPLE_KEY_VALUE_RE = re.compile(r"^[A-Za-z0-9_.-]+\s*:\s*.*$")

    @staticmethod
    def parse(
        lines: List[str], profile: Optional[str] = None
    ) -> Tuple[DocumentMetadata, List[str]]:
        """
        Parse YAML frontmatter and return metadata + remaining lines.

        Args:
            lines: All lines from the markdown file

        Returns:
            Tuple of (DocumentMetadata, remaining_lines)
        """
        metadata = default_metadata(profile or "ib-report")

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

        declared = any(re.match(r"^(profile|layout|tables|sender):", line, re.IGNORECASE) for line in frontmatter_lines)
        if not declared and not FrontmatterParser._is_valid_frontmatter(frontmatter_lines):
            logger.debug("Frontmatter markers found, but content is not YAML frontmatter")
            return metadata, lines

        # Parse YAML content
        yaml_content = "\n".join(frontmatter_lines)
        try:
            parsed_data = yaml.safe_load(yaml_content) or {}
        except yaml.YAMLError as exc:
            if declared:
                raise ValueError(f"Invalid YAML frontmatter: {exc}") from exc
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

        metadata = default_metadata(profile or str(data.get("profile", "ib-report")))

        # Map to metadata fields
        metadata.title = str(data.get("title", metadata.title))
        metadata.subtitle = str(data.get("subtitle", metadata.subtitle))
        metadata.company = str(data.get("company", metadata.company))
        metadata.ticker = str(data.get("ticker", metadata.ticker))
        metadata.sector = str(data.get("sector", metadata.sector))
        metadata.analyst = str(data.get("analyst", metadata.analyst))

        # Store extra fields
        known_keys = {"title", "subtitle", "company", "ticker", "sector", "analyst", "profile"}
        structured = {"layout", "tables", "sender", "recipients", "cc", "attachments", "attendees", "letter"}
        metadata.extra = {
            k: v if k in structured else str(v) for k, v in data.items() if k not in known_keys
        }

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
    _ESCAPE_RE = re.compile(r'\\([\\$`^~.*"\'()\[\]{}|_-])')
    _HTML_BREAK_RE = re.compile(r"(?<!\\)<br\s*/?>", re.IGNORECASE)
    _ESCAPED_HTML_BREAK_RE = re.compile(r"\\(<br\s*/?>)", re.IGNORECASE)
    _CODE_SPAN_RE = re.compile(r"(?<![\\`])(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.DOTALL)

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
    _INLINE_LATEX_RE = re.compile(r"(?<![\\$])\$(?!\$)((?:\\.|[^$\\\n])+?)\$(?!\$)")

    @classmethod
    def parse_runs(cls, text: str) -> List[TextRun]:
        """
        Parse text into runs with formatting.
        Handles **bold**, inline $LaTeX$, and combinations.
        """
        runs: List[TextRun] = []
        text, code_spans = cls._protect_code_spans(text)

        # ── Phase 1: Split on color spans, then inline LaTeX boundaries ─────
        color_segments = cls._split_on_color_spans(text)

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
                    runs.extend(
                        cls._apply_color(cls._parse_inline_formatting(inline_text), color_hex)
                    )

        return cls._restore_code_spans(runs, code_spans)

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
        text, code_spans = cls._protect_code_spans(text)
        runs: List[TextRun] = []
        for segment_text, color_hex in cls._split_on_color_spans(text):
            runs.extend(cls._apply_color(cls._parse_inline_formatting(segment_text), color_hex))
        return cls._restore_code_spans(runs, code_spans)

    @classmethod
    def _protect_code_spans(cls, text: str) -> Tuple[str, Dict[str, str]]:
        """Shield literals before emphasis, HTML, links and math tokenize them."""
        literals: Dict[str, str] = {}
        prefix = "\ue000CODE"
        while prefix in text:
            prefix += "X"

        def replace(match):
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(0)
            return token

        return cls._CODE_SPAN_RE.sub(replace, text), literals

    @classmethod
    def _protect_escapes(cls, text: str) -> Tuple[str, Dict[str, str]]:
        """Keep escaped punctuation literal until inline tokenization finishes."""
        literals: Dict[str, str] = {}
        prefix = "\ue000ESC"
        while prefix in text:
            prefix += "X"

        def replace(match: Match[str]) -> str:
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(1)
            return token

        return cls._ESCAPE_RE.sub(replace, text), literals

    @staticmethod
    def _restore_code_spans(runs: List[TextRun], literals: Dict[str, str]) -> List[TextRun]:
        """Restore original text without manufacturing code-font semantics."""
        for run in runs:
            for token, literal in literals.items():
                run.text = run.text.replace(token, literal)
                if run.hyperlink:
                    run.hyperlink = run.hyperlink.replace(token, literal)
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

    _INLINE_REFERENCE_RE = re.compile(
        r"(?<!!)\[([^\]\n]+)\]\((https?://[^\s)]+|mailto:[^\s)]+)\)|\[\^(\d+)\]"
    )

    @classmethod
    def _parse_inline_formatting(cls, text: str) -> List[TextRun]:
        """Parse links and numeric footnote references alongside inline emphasis."""
        text, literals = cls._protect_escapes(text)
        runs: List[TextRun] = []
        offset = 0
        for match in cls._INLINE_REFERENCE_RE.finditer(text):
            runs.extend(cls._parse_plain_formatting(text[offset : match.start()]))
            if match.group(3):
                runs.append(TextRun(text=match.group(3), superscript=True, footnote_id=int(match.group(3))))
            else:
                labels = cls._parse_plain_formatting(match.group(1))
                for run in labels:
                    run.hyperlink = match.group(2)
                runs.extend(labels)
            offset = match.end()
        runs.extend(cls._parse_plain_formatting(text[offset:]))
        # Destination strings are opaque; Markdown escapes apply to visible text.
        for run in runs:
            if run.hyperlink:
                for token, literal in literals.items():
                    run.hyperlink = run.hyperlink.replace(token, "\\" + literal)
        return cls._restore_code_spans(runs, literals)

    @classmethod
    def _parse_plain_formatting(cls, text: str) -> List[TextRun]:
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
        text, _ = cls._protect_code_spans(text)
        return bool(cls._INLINE_LATEX_RE.search(text))

    @staticmethod
    def cleanup_text(text: str) -> str:
        """Remove markdown artifacts and escape characters (single-pass regex)"""
        return TextParser.normalize_html_breaks(TextParser._ESCAPE_RE.sub(r"\1", text).strip())

    @staticmethod
    def cleanup_text_preserve_spacing(text: str) -> str:
        """Remove markdown escape characters while preserving surrounding spacing."""
        return TextParser.normalize_html_breaks(TextParser._ESCAPE_RE.sub(r"\1", text))

    @classmethod
    def normalize_html_breaks(cls, text: str) -> str:
        """Convert explicit HTML breaks, preserving literal code and escaped tags.

        Run after emphasis/link tokenization so breaks inside bold text retain
        their formatting and link destinations are never rewritten.
        """
        pieces: List[str] = []
        offset = 0
        for match in cls._CODE_SPAN_RE.finditer(text):
            part = cls._HTML_BREAK_RE.sub("\n", text[offset:match.start()])
            pieces.append(cls._ESCAPED_HTML_BREAK_RE.sub(r"\1", part))
            pieces.append(match.group(0))
            offset = match.end()
        part = cls._HTML_BREAK_RE.sub("\n", text[offset:])
        pieces.append(cls._ESCAPED_HTML_BREAK_RE.sub(r"\1", part))
        return "".join(pieces)


class FenceScanner:
    """Share fenced-code boundaries across block parsing and metadata scanning."""

    OPEN_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})([^\r\n]*)$")
    CLOSE_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})[ \t]*$")

    @classmethod
    def opening(cls, line: str) -> Optional[Tuple[str, str]]:
        """Read an opening delimiter and info string.

        Args:
            line: A source line with its indentation preserved.

        Returns:
            Delimiter and info string, or None for ordinary text.
        """
        match = cls.OPEN_RE.match(line)
        if not match or (match.group(1)[0] == "`" and "`" in match.group(2)):
            return None
        return match.group(1), match.group(2).strip()

    @classmethod
    def scan(cls, lines: List[str], start: int) -> Optional[Tuple[str, List[str], int]]:
        """Read a fence through a matching close or the end of the document.

        Args:
            lines: Source lines.
            start: Index of the possible opening fence.

        Returns:
            Info string, literal body lines and next index, or None.
        """
        opening = cls.opening(lines[start])
        if opening is None:
            return None
        delimiter, language = opening
        end = start + 1
        while end < len(lines):
            close = cls.CLOSE_RE.match(lines[end])
            if close and close.group(1)[0] == delimiter[0] and len(close.group(1)) >= len(delimiter):
                return language, lines[start + 1:end], end + 1
            end += 1
        return language, lines[start + 1:end], end

    @classmethod
    def protected_indices(cls, lines: List[str]) -> Set[int]:
        """Find all fence lines, including delimiters, for non-code processing.

        Args:
            lines: Source lines.

        Returns:
            Indices that must remain literal.
        """
        indices: Set[int] = set()
        index = 0
        while index < len(lines):
            block = cls.scan(lines, index)
            if block is None:
                index += 1
            else:
                end = block[2]
                indices.update(range(index, end))
                index = end
        return indices


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
            reference_label = re.sub(r"^#{1,6}\s+", "", line_lower).strip("* :：")
            if reference_label in FootnoteParser.REFERENCE_KEYWORDS:
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
    def parse(lines: List[str], financial_rules: bool = True) -> Table:
        """
        Parse markdown table lines into a Table object.

        Args:
            lines: Lines that make up the table (starting with |)
        """
        table = Table()

        # Only the second line can be the structural delimiter row.
        data_lines = [
            line
            for index, line in enumerate(lines)
            if not (index == 1 and TableParser._is_delimiter_row(line))
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
        table.table_type = (
            TableParser._detect_type(header_text) if financial_rules else TableType.GENERIC
        )

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
                    "Table row %d has %d columns (expected %d) — extra columns dropped: %s",
                    i,
                    len(cells),
                    table.col_count,
                    cells[table.col_count :],
                )
                table.warnings.append(f"Table row {i} has extra cells; content was dropped")
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
        """Split a table row into cell contents"""
        text = line.strip()
        cells: List[str] = []
        current: List[str] = []
        index = 0
        while index < len(text):
            char = text[index]
            if char == "\\" and index + 1 < len(text) and text[index + 1] in {"|", "\\"}:
                current.append(text[index + 1])
                index += 2
                continue
            if char == "|":
                cells.append("".join(current).strip())
                current = []
            else:
                current.append(char)
            index += 1
        cells.append("".join(current).strip())
        if text.startswith("|"):
            cells = cells[1:]
        if text.endswith("|") and cells and cells[-1] == "":
            backslashes = len(text[:-1]) - len(text[:-1].rstrip("\\"))
            if backslashes % 2 == 0:
                cells = cells[:-1]
        return cells

    @staticmethod
    def _is_delimiter_row(line: str) -> bool:
        """Recognize a structural table delimiter without classifying body rows."""
        return "-" in line and set(line.strip()).issubset({"|", "-", " ", "\t", ":"})

    @staticmethod
    def _parse_alignments(lines: List[str]) -> List[str]:
        """Parse column alignments from separator line"""
        alignments: List[str] = []
        for line in lines[1:2]:
            if TableParser._is_delimiter_row(line):
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
        visible_text = "".join(run.text for run in cell.runs)
        cell.is_negative = (
            visible_text.startswith("(")
            and visible_text.endswith(")")
            and any(c.isdigit() for c in visible_text)
        ) or (visible_text.startswith("-") and any(c.isdigit() for c in visible_text))

        # Table-type specific detection
        if table_type == TableType.RISK_MATRIX:
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
    INLINE_RE = TextParser._INLINE_LATEX_RE

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
        return TextParser.has_inline_latex(text)

    @classmethod
    def extract_inline_segments(cls, text: str) -> List[Tuple[str, bool]]:
        """
        Split text into (content, is_latex) segments.

        Returns:
            List of (text, is_latex) tuples preserving source order
        """
        protected, literals = TextParser._protect_code_spans(text)
        segments = TextParser._split_on_inline_latex(protected)
        return [
            (TextParser._restore_code_spans([TextRun(text=value)], literals)[0].text, is_math)
            for value, is_math in segments
        ]


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
    IMAGE_PATTERN = re.compile(
        r"^!\[(.*?)\]\(\s*(?:<([^>\n]+)>|([^\s]+?))"
        r"(?:\s+(?:\"[^\"\n]*\"|'[^'\n]*'))?\s*\)$"
    )
    TABLE_START_PATTERN = re.compile(r"^\|")
    SEPARATOR_PATTERN = re.compile(r"^(---|## ---)$")
    CODE_FENCE_PATTERN = FenceScanner.OPEN_RE

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

    _HTML_BREAK_RE = re.compile(r"(?<!\\)<br\s*/?>\s*$", re.IGNORECASE)
    _HTML_ANCHOR_RE = re.compile(
        r"<a\b[^>]*\bid\s*=\s*(['\"]).*?\1[^>]*>\s*</a>",
        re.IGNORECASE,
    )
    _MULTISPACE_RE = re.compile(r"[ \t]{2,}")
    _OPEN_PAREN_SPACE_RE = re.compile(r"\(\s+")
    _CLOSE_PAREN_SPACE_RE = re.compile(r"\s+\)")

    _EXT_FOOTNOTE_DEF_RE = re.compile(r"^\[\^(\d+)\]:\s*(.+)$")
    _SETEXT_H1_RE = re.compile(r"^=+\s*$")
    _LINK_DEFINITION_RE = re.compile(
        r"^[ \t]*\[(?!\^)([^\]\n]+)\]:[ \t]*(?:<([^>\n]+)>|(\S+))"
        r"(?:[ \t]+(?:\"[^\"\n]*\"|'[^'\n]*'))?[ \t]*$", re.MULTILINE
    )
    _REFERENCE_LINK_RE = re.compile(r"(?<!!)(?<!\\)\[(?!\^)([^\]\n]+)\](?:\[([^\]\n]*)\])?(?!\()")
    _REFERENCE_LABEL_SPACE_RE = re.compile(r"\s+")

    def __init__(
        self, preserve_trailing_double_space_break: bool = False, profile: Optional[str] = None
    ):
        """
        Initialize parser behavior flags.

        Args:
            preserve_trailing_double_space_break:
                When True, treat markdown trailing 2+ spaces as hard line breaks.
                Default False to avoid accidental Shift+Enter artifacts from noisy input.
        """
        self.preserve_trailing_double_space_break = preserve_trailing_double_space_break
        self.profile = profile
        self._financial_rules = True

    def parse(self, content: str) -> DocumentModel:
        """
        Parse markdown content into a DocumentModel.

        Args:
            content: The full markdown content

        Returns:
            A DocumentModel with all parsed elements
        """
        lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")

        # Parse frontmatter
        metadata, remaining_lines = FrontmatterParser.parse(lines, self.profile)
        self._financial_rules = get_profile(metadata.profile).is_ib

        # Detect whether YAML frontmatter was present
        has_frontmatter = len(remaining_lines) < len(lines)
        remaining_lines = self._strip_comments(remaining_lines)
        remaining_lines = self._resolve_reference_links(remaining_lines)

        # Extract footnotes/references
        fenced_indices = FenceScanner.protected_indices(remaining_lines)
        prose_lines = [
            line if index not in fenced_indices else ""
            for index, line in enumerate(remaining_lines)
        ]
        # Preserve line indices while masking code spans that cross source lines.
        scan_lines = TextParser._CODE_SPAN_RE.sub(
            lambda match: "CODE" + "\n" * match.group(0).count("\n"),
            "\n".join(prose_lines),
        ).split("\n")
        footnotes = (
            FootnoteParser.extract_references(prose_lines) if self._financial_rules else {}
        )
        input_warnings: List[str] = []
        explicit_references: Set[int] = set()
        content_lines: List[str] = []
        for index, raw_line in enumerate(remaining_lines):
            if index in fenced_indices:
                content_lines.append(raw_line)
                continue
            match = self._EXT_FOOTNOTE_DEF_RE.match(raw_line.strip())
            if not self._EXT_FOOTNOTE_DEF_RE.match(scan_lines[index].strip()):
                match = None
            if match:
                number = int(match.group(1))
                if number <= 0:
                    raise ValueError("Footnote IDs must be positive integers")
                if number in footnotes:
                    input_warnings.append(f"Duplicate footnote definition: {number}")
                footnotes[number] = match.group(2)
            else:
                explicit_references.update(
                    run.footnote_id for run in TextParser.parse_runs(scan_lines[index])
                    if run.footnote_id is not None
                )
                content_lines.append(raw_line)
        remaining_lines = content_lines

        # Parse elements
        elements = self._parse_elements(remaining_lines)

        # If no YAML frontmatter, extract metadata from document header
        if not has_frontmatter and self._financial_rules:
            elements = self._extract_header_metadata(metadata, elements)

        input_warnings.extend(
            f"Undefined footnote: {number}"
            for number in sorted(explicit_references - set(footnotes))
        )
        model = DocumentModel(
            metadata=metadata,
            elements=elements,
            footnotes=footnotes,
            warnings=input_warnings,
            parsed_profile=metadata.profile,
        )
        if not self._financial_rules:
            first_heading = next(
                (
                    e.content.text
                    for e in elements
                    if e.element_type == ElementType.HEADING_1 and isinstance(e.content, Heading)
                ),
                "",
            )
            if metadata.title in {"", "Document", "IB Report"} and first_heading:
                metadata.title = first_heading
        apply_table_specs(model)
        model.warnings.extend(
            warning
            for element in elements
            if element.element_type == ElementType.TABLE and isinstance(element.content, Table)
            for warning in element.content.warnings
        )
        return model

    @staticmethod
    def _strip_comments(lines: List[str]) -> List[str]:
        """Omit HTML comments while retaining fenced and inline code verbatim."""
        text = "\n".join(lines)
        offsets: List[int] = []
        offset = 0
        for line in lines:
            offsets.append(offset)
            offset += len(line) + 1
        line_starts = {start: index for index, start in enumerate(offsets)}
        pieces: List[str] = []
        position = 0
        while position < len(text):
            if position in line_starts:
                block = FenceScanner.scan(lines, line_starts[position])
                if block:
                    end = offsets[block[2]] if block[2] < len(offsets) else len(text)
                    pieces.append(text[position:end])
                    position = end
                    continue
            code = TextParser._CODE_SPAN_RE.match(text, position) if text[position] == "`" else None
            if code:
                pieces.append(code.group(0))
                position = code.end()
                continue
            if text.startswith("<!--", position) and (position == 0 or text[position - 1] != "\\"):
                close = text.find("-->", position + 4)
                end = close + 3 if close >= 0 else len(text)
                pieces.append("\n" * text[position:end].count("\n"))
                position = end
                continue
            pieces.append(text[position])
            position += 1
        return "".join(pieces).split("\n")

    @classmethod
    def _resolve_reference_links(cls, lines: List[str]) -> List[str]:
        """Collect link definitions and expand resolved references outside code."""
        literals: Dict[str, str] = {}
        prefix = "\ue000REF"
        source = "\n".join(lines)
        while prefix in source:
            prefix += "X"

        def protect(value: str) -> str:
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = value
            return token

        pieces: List[str] = []
        index = 0
        while index < len(lines):
            block = FenceScanner.scan(lines, index)
            if block:
                pieces.append(protect("\n".join(lines[index:block[2]])))
                index = block[2]
            else:
                pieces.append(lines[index])
                index += 1
        source = TextParser._CODE_SPAN_RE.sub(lambda match: protect(match.group(0)), "\n".join(pieces))
        source = TextParser._ESCAPE_RE.sub(lambda match: protect(match.group(0)), source)
        source = TextParser._INLINE_REFERENCE_RE.sub(lambda match: protect(match.group(0)), source)
        definitions: Dict[str, str] = {}

        def label(value: str) -> str:
            return cls._REFERENCE_LABEL_SPACE_RE.sub(" ", value.strip()).casefold()

        def define(match: Match[str]) -> str:
            definitions.setdefault(label(match.group(1)), match.group(2) or match.group(3))
            return ""

        source = cls._LINK_DEFINITION_RE.sub(define, source)

        def resolve(match: Match[str]) -> str:
            destination = definitions.get(label(match.group(2) or match.group(1)))
            if destination is None:
                return match.group(0)
            return f"[{match.group(1)}]({destination})"

        source = cls._REFERENCE_LINK_RE.sub(resolve, source)
        # Later protection layers may contain tokens from earlier layers.
        for token, value in reversed(list(literals.items())):
            source = source.replace(token, value)
        return source.split("\n")

    # ── Element-level parsing ───────────────────────────────────────────────

    def _parse_elements(self, lines: List[str]) -> List[Element]:
        """Parse lines into elements"""
        elements: List[Element] = []
        i = 0
        in_references = False

        while i < len(lines):
            raw_line = lines[i]
            line = self._strip_html_anchors(raw_line.strip()).strip()

            # ── References section gating ───────────────────────────────────
            if self._financial_rules and self._is_reference_header(line):
                in_references = True
                i += 1
                continue

            if in_references:
                if FootnoteParser.REFERENCE_PATTERN.match(line) or not line:
                    i += 1
                    continue
                in_references = False

            # ── Empty line ──────────────────────────────────────────────────
            if not line:
                i += 1
                continue

            # ── Separator ───────────────────────────────────────────────────
            if self.SEPARATOR_PATTERN.match(line):
                elements.append(
                    Element(
                        element_type=ElementType.SEPARATOR,
                        content=None,
                        raw_text=raw_line,
                    )
                )
                i += 1
                continue

            # ── Fenced code block ───────────────────────────────────────────
            code_block = FenceScanner.scan(lines, i)
            if code_block:
                language, code_lines, i = code_block
                code_text = "\n".join(code_lines).rstrip()

                # ── diagram:flow → Diagram object ──────────────────────
                if language.startswith("diagram:"):
                    diagram = self._parse_diagram(language, code_text)
                    if diagram:
                        elements.append(
                            Element(
                                element_type=ElementType.DIAGRAM,
                                content=diagram,
                                raw_text=raw_line,
                            )
                        )
                        continue

                # ── ASCII art detection ────────────────────────────────
                is_ascii = CodeBlock.detect_ascii_art(code_text)

                elements.append(
                    Element(
                        element_type=ElementType.CODE_BLOCK,
                        content=CodeBlock(code=code_text, language=language, is_ascii_art=is_ascii),
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
                table = TableParser.parse(table_lines, financial_rules=self._financial_rules)
                elements.append(
                    Element(
                        element_type=ElementType.TABLE,
                        content=table,
                        raw_text="\n".join(table_lines),
                    )
                )
                continue

            # ── Headings (must be checked before numbered list) ─────────────
            if i + 1 < len(lines) and self._SETEXT_H1_RE.fullmatch(lines[i + 1].strip()):
                if not self._starts_new_block(line):
                    elements.append(Element(
                        element_type=ElementType.HEADING_1,
                        content=Heading(level=1, text=line),
                        raw_text=raw_line + "\n" + lines[i + 1],
                    ))
                    i += 2
                    continue
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
                text, next_idx = self._collect_paragraph(lines, i, first_line=match.group(2))
                item = ListItem(
                    text=text,
                    runs=TextParser.parse_runs(text),
                    indent_level=indent_level,
                )
                elements.append(
                    Element(
                        element_type=ElementType.BULLET_LIST,
                        content=item,
                        raw_text="\n".join(lines[i:next_idx]),
                    )
                )
                i = next_idx
                continue

            # ── Numbered list ───────────────────────────────────────────────
            match = self.NUMBERED_LIST_PATTERN.match(raw_line)
            if match and not self._is_numbered_heading(line):
                indent_level = self._get_indent_level(match.group(1))
                number = match.group(2)
                text, next_idx = self._collect_paragraph(lines, i, first_line=match.group(3))
                item = ListItem(
                    text=text,
                    runs=TextParser.parse_runs(text),
                    indent_level=indent_level,
                )
                elements.append(
                    Element(
                        element_type=ElementType.NUMBERED_LIST,
                        content=(number, item),
                        raw_text="\n".join(lines[i:next_idx]),
                    )
                )
                i = next_idx
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
                        content=Image(alt_text=match.group(1), path=match.group(2) or match.group(3)),
                        raw_text=line,
                    )
                )
                i += 1
                continue

            # ── Paragraph (with inline LaTeX detection) ─────────────────────
            paragraph_text, next_idx = self._collect_paragraph(lines, i)
            para_element = self._parse_paragraph(paragraph_text)
            # Header metadata needs source line boundaries before soft-wrap merging.
            para_element.raw_text = "\n".join(lines[i:next_idx])
            elements.append(para_element)
            i = next_idx

        return elements

    def _collect_paragraph(
        self, lines: List[str], start_idx: int, first_line: Optional[str] = None
    ) -> Tuple[str, int]:
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
            raw_line = first_line if i == start_idx and first_line is not None else lines[i]
            line = self._strip_html_anchors(raw_line).strip()

            if not line:
                break

            if i > start_idx and self._starts_new_block(line):
                break
            if i > start_idx and i + 1 < len(lines) and self._SETEXT_H1_RE.fullmatch(lines[i + 1].strip()):
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
            or FenceScanner.opening(line)
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
        working = self._strip_html_anchors(raw_line)

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

        text, literals = TextParser._protect_code_spans(text)
        normalized = cls._MULTISPACE_RE.sub(" ", text)
        normalized = cls._OPEN_PAREN_SPACE_RE.sub("(", normalized)
        normalized = cls._CLOSE_PAREN_SPACE_RE.sub(")", normalized)
        return TextParser._restore_code_spans([TextRun(text=normalized.strip())], literals)[0].text

    @classmethod
    def _strip_html_anchors(cls, text: str) -> str:
        """Remove HTML anchor tags that should not surface in rendered output."""
        if "<a" not in text.lower():
            return text
        return cls._HTML_ANCHOR_RE.sub("", text)

    def _get_indent_level(self, prefix: str) -> int:
        """Convert leading markdown indentation to a list nesting level."""
        expanded = prefix.replace("\t", "    ")
        if not expanded:
            return 0
        return max(0, len(expanded) // (2 if self._financial_rules else 4))

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
        if self._financial_rules and self.NUMBERED_HEADING_PATTERN.match(line):
            text = TextParser.cleanup_text(line)
            return Element(
                element_type=ElementType.NUMBERED_HEADING,
                content=Heading(level=1, text=text, is_numbered=True),
                raw_text=line,
            )

        # H4 (check longer prefixes first to avoid partial match)
        match = self.H4_PATTERN.match(line)
        if match:
            text = match.group(1).strip()
            return Element(
                element_type=ElementType.HEADING_4,
                content=Heading(level=4, text=text),
                raw_text=line,
            )

        # H3
        match = self.H3_PATTERN.match(line)
        if match:
            text = match.group(1).strip()
            return Element(
                element_type=ElementType.HEADING_3,
                content=Heading(level=3, text=text),
                raw_text=line,
            )

        # H2
        match = self.H2_PATTERN.match(line)
        if match:
            text = match.group(1).strip()
            return Element(
                element_type=ElementType.HEADING_2,
                content=Heading(level=2, text=text),
                raw_text=line,
            )

        # H1
        match = self.H1_PATTERN.match(line)
        if match:
            text = match.group(1).strip()
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
        if not self._financial_rules:
            return False
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

    # ── Diagram parsing ─────────────────────────────────────────────────

    # Matches **key**: value  OR  **key:** value (colon inside or outside bold)
    _BOLD_KV_PATTERN = re.compile(r"^\*\*([^*:：]+)[:：]?\*\*\s*[:：]?\s*(.+)$")
    # Only fields actually consumed by CoverRenderer / office metadata rendering.
    _HEADER_METADATA_FIELDS = {
        "기준일": ("extra", "date"),
        "as of": ("extra", "date"),
        "작성일": ("extra", "date"),
        "report date": ("extra", "date"),
        "date": ("extra", "date"),
        "분석 대상 기간": ("extra", "analysis_period"),
        "analysis period": ("extra", "analysis_period"),
        "analysis_period": ("extra", "analysis_period"),
        "분석 기준": ("extra", "analysis_basis"),
        "analysis basis": ("extra", "analysis_basis"),
        "analysis_basis": ("extra", "analysis_basis"),
        "prepared by": ("analyst", ""),
        "작성자": ("analyst", ""),
        "analyst": ("analyst", ""),
        "institution": ("company", ""),
        "company": ("company", ""),
        "기관": ("company", ""),
        "sector": ("sector", ""),
        "업종": ("sector", ""),
        "ticker": ("ticker", ""),
        "prepared for": ("extra", "recipient"),
        "recipient": ("extra", "recipient"),
        "수신": ("extra", "recipient"),
        "subject_company": ("extra", "subject_company"),
        "report_type": ("extra", "report_type"),
    }

    def _parse_diagram(self, language: str, code_text: str) -> Optional[Diagram]:
        """Parse a diagram:flow YAML block into a Diagram object."""
        try:
            data = yaml.safe_load(code_text)
            if not isinstance(data, dict):
                return None

            diagram = Diagram(diagram_type=language.split(":", 1)[1] if ":" in language else "flow")
            diagram.title = data.get("title", "")
            diagram.notes = data.get("notes", [])

            for b in data.get("boxes", []):
                box = DiagramBox(
                    id=b["id"],
                    label=b.get("label", b["id"]),
                    pos=b.get("pos", [0, 0]),
                    style=b.get("style", "default"),
                )
                diagram.boxes.append(box)

            for a in data.get("arrows", []):
                arrow = DiagramArrow(
                    from_id=a["from"],
                    to_id=a["to"],
                    label=a.get("label", ""),
                    style=a.get("style", "solid"),
                )
                diagram.arrows.append(arrow)

            return diagram
        except Exception as e:
            logger.warning("Failed to parse diagram: %s", e)
            return None

    def _extract_header_metadata(
        self, metadata: DocumentMetadata, elements: List[Element]
    ) -> List[Element]:
        """Extract IB header metadata while retaining body content.

        Args:
            metadata: Metadata to populate when YAML frontmatter is absent.
            elements: Parsed elements retaining paragraph source line boundaries.

        Returns:
            Body elements, including the marked inferred subtitle and unknown labels.
        """
        filtered: List[Element] = []
        in_header = True
        title_found = False

        # Clear defaults that don't make sense without frontmatter
        metadata.company = ""

        for elem in elements:
            if in_header and elem.element_type in (
                ElementType.HEADING_1,
                ElementType.HEADING_2,
                ElementType.PARAGRAPH,
            ):
                # First H1 → title
                if (
                    not title_found
                    and elem.element_type == ElementType.HEADING_1
                    and isinstance(elem.content, Heading)
                ):
                    metadata.title = elem.content.text
                    title_found = True
                    continue

                # Retain the inferred subtitle heading for cover-free rendering.
                if (
                    title_found
                    and not metadata.subtitle
                    and elem.element_type == ElementType.HEADING_2
                    and isinstance(elem.content, Heading)
                ):
                    metadata.subtitle = elem.content.text
                    elem.inferred_subtitle = True
                    filtered.append(elem)
                    continue

                if elem.element_type == ElementType.HEADING_2:
                    in_header = False

                # Each leading bold key-value source line is a separate field.
                if elem.element_type == ElementType.PARAGRAPH:
                    source_lines = elem.raw_text.split("\n")
                    line_index = 0
                    while line_index < len(source_lines):
                        normalized, _ = self._normalize_paragraph_line(source_lines[line_index])
                        kv_match = self._BOLD_KV_PATTERN.match(normalized)
                        if not kv_match:
                            break
                        key = kv_match.group(1).strip()
                        value = kv_match.group(2).strip()
                        mapped = self._HEADER_METADATA_FIELDS.get(key.lower())
                        if mapped:
                            field_name, extra_key = mapped
                            if field_name == "extra":
                                metadata.extra[extra_key] = value
                            else:
                                setattr(metadata, field_name, value)
                        else:
                            if line_index + 1 < len(source_lines):
                                following, _ = self._normalize_paragraph_line(source_lines[line_index + 1])
                                if not self._BOLD_KV_PATTERN.match(following):
                                    # Unknown labels remain ordinary prose, including soft wraps.
                                    break
                            filtered.append(self._parse_paragraph(normalized))
                        line_index += 1
                    if line_index == len(source_lines):
                        continue
                    if line_index:
                        # Ordinary prose ends the header and retains normal soft wrapping.
                        text, _ = self._collect_paragraph(source_lines, line_index)
                        elem = self._parse_paragraph(text)
                        elem.raw_text = "\n".join(source_lines[line_index:])
                    in_header = False
            else:
                if in_header:
                    in_header = False

            filtered.append(elem)

        return filtered

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
        return cleaned.strip("* :：") in self._REFERENCE_KEYWORDS

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
    """Read a file without statistically misdetecting valid UTF-8 Korean text."""
    return _decode_bytes(Path(file_path).read_bytes(), label=file_path)


# ═══════════════════════════════════════════════════════════════════════════════
# CONVENIENCE FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════


def _decode_bytes(raw_bytes: bytes, label: str = "<stream>") -> str:
    """Prefer deterministic UTF-8/Korean decoding, then optional detection."""
    for encoding in ("utf-8-sig", "euc-kr", "cp949"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    try:
        from charset_normalizer import from_bytes

        result = from_bytes(raw_bytes).best()
        if result and result.encoding:
            logger.debug("Encoding detected by charset_normalizer: %s", result.encoding)
            return str(result)
    except ImportError:
        pass
    except Exception as exc:
        logger.debug("Encoding detector failed: %s", exc)
    raise UnicodeDecodeError("multiple", raw_bytes, 0, min(1, len(raw_bytes)), f"Failed to decode {label}")


def parse_markdown_file(
    source: "Union[str, BinaryIO]", profile: Optional[str] = None
) -> DocumentModel:
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
        raw_bytes = cast(BinaryIO, source).read()
        content = _decode_bytes(raw_bytes, label="<stream>")
        label = "<stream>"
    else:
        content = _read_with_encoding(str(source))
        label = str(source)

    lines = content.splitlines()
    _, remaining_lines = FrontmatterParser.parse(lines, profile=profile)
    frontmatter_present = remaining_lines is not lines

    parser = MarkdownParser(profile=profile)
    model = parser.parse(content)
    if get_profile(model.metadata.profile).is_ib:
        _infer_metadata_from_elements(model, allow_company_inference=not frontmatter_present)
    if not is_stream(source) and model.metadata.extra.get("theme") not in {None, "default", "mono"}:
        theme_path = Path(model.metadata.extra["theme"])
        if not theme_path.is_absolute():
            model.metadata.extra["theme"] = str(Path(str(source)).resolve().parent / theme_path)

    if not is_stream(source):
        source_dir = Path(str(source)).resolve().parent
        for element in model.elements:
            if isinstance(element.content, Image):
                image = element.content
                if image.path and not image.base64_data and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", image.path):
                    image_path = Path(image.path)
                    if not image_path.is_absolute():
                        image.path = str(source_dir / image_path)

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
        elif field_name and getattr(metadata, field_name, "") in {
            "",
            "IB Report",
            "SECTOR",
            "DCM Team 1",
            "Korea Development Bank",
        }:
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
