"""Block-level scanners: code fences, footnotes, LaTeX equations and embedded Base64 images."""

import logging
import re

from document_model import Image, TextRun
from md_parse_text import TextParser, _document_int

logger = logging.getLogger("md_parser")


class FenceScanner:
    """Share fenced-code boundaries across block parsing and metadata scanning."""

    OPEN_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})([^\r\n]*)$")
    CLOSE_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})[ \t]*$")

    @classmethod
    def opening(cls, line: str) -> tuple[str, str] | None:
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
    def scan(cls, lines: list[str], start: int) -> tuple[str, list[str], int] | None:
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
            if (
                close
                and close.group(1)[0] == delimiter[0]
                and len(close.group(1)) >= len(delimiter)
            ):
                return language, lines[start + 1 : end], end + 1
            end += 1
        return language, lines[start + 1 : end], end

    @classmethod
    def protected_indices(cls, lines: list[str]) -> set[int]:
        """Find all fence lines, including delimiters, for non-code processing.

        Args:
            lines: Source lines.

        Returns:
            Indices that must remain literal.
        """
        indices: set[int] = set()
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
    def extract_references(lines: list[str]) -> dict[int, str]:
        """
        Extract references from the end of the document.
        Looks for patterns like "1. Citation text" after references section.

        Only considers the *last* references section to avoid false positives.
        """
        references: dict[int, str] = {}
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
                    ref_num = _document_int(match.group(1))
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
    def find_inline_references(text: str) -> list[int]:
        """Find all inline reference numbers in text"""
        refs: list[int] = []
        for match in FootnoteParser.INLINE_PATTERN.finditer(text):
            group = match.group(1) or match.group(2)
            if group:
                refs.append(_document_int(group))
        return refs


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
    def is_block_single_line(cls, line: str) -> str | None:
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
    def extract_inline_segments(cls, text: str) -> list[tuple[str, bool]]:
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
    def parse(cls, line: str) -> Image | None:
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
    def parse_multiline(cls, lines: list[str], start_idx: int) -> tuple[Image | None, int]:
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
