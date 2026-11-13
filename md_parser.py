"""
MD Parser Module for IB Style Word Report Converter
Handles parsing of Markdown files including frontmatter, elements, tables,
LaTeX equations, Base64 images, and footnotes.

Changelog (property-test regressions):
    - Use documented rejections for malformed declared YAML and invalid footnote IDs.
    - Diagnose XML-incompatible input before strict rendering reaches metadata.

Changelog (term variables):
    - NEW: `terms:` frontmatter values substituted for `{{key}}` as value runs
      (paragraphs, lists, cells, headings, quotes, title/subtitle/date); code,
      inline math, `\\{{` escapes and link destinations stay literal.
    - NEW: Undefined keys and malformed references become model warnings.

Changelog (table structure):
    - NEW: `Table.header_rows` (table spec `header_rows`); HTML tables keep the
      rows their header spans cover as header rows instead of joining them.

Changelog (input loss found on a converted term sheet):
    - FIXED: `*`/`_` emphasis needs flanking delimiters, so `2 * 3 * 4` and
      spaced note markers keep their asterisks.
    - NEW: Inline images in text and table cells (`TextRun.image`), protected
      as one unit before math, links, emphasis and term substitution.
    - NEW: HTML `<table>` blocks built directly as runs (spans, block breaks,
      nested formatting, links, images, terms, nested tables) and standalone
      `<img>` lines; unclosed tables and unplaceable text warn.

Changelog (memo rendering):
    - NEW: Inline code spans become literal `code` runs without their backticks.
    - NEW: Local file links (angle brackets, drive or ./ paths, document
      extensions) become hyperlinks; paths inside the file's folder are relative.

Changelog (2026-09-29):
    - FIXED: Preserve escaped inline syntax and non-reference body content.
    - FIXED: Share fence boundaries and normalize Markdown block input.
    - FIXED: Retain inferred IB subtitle headings and separate header metadata lines.
    - NEW: Validate table span groups and preserve term-sheet confirmation fences.
    - NEW: Resolve file-relative house paths without opening house files while parsing.

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
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser
from pathlib import Path
from re import Match
from typing import (
    BinaryIO,
    cast,
)

from chart_renderer import ChartSpecError, parse_chart_spec
from document_model import (
    Blockquote as Blockquote,
)
from document_model import (
    BulletList as BulletList,
)
from document_model import Chart as Chart
from document_model import (
    CodeBlock as CodeBlock,
)
from document_model import ConfirmationBlock as ConfirmationBlock
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
from document_profiles import (
    apply_table_specs,
    check_schedules,
    get_profile,
    validate_term_sheet_metadata,
)
from input_safety import (
    DEFAULT_LIMITS,
    InputLimits,
    RejectionCode,
    StrictValidationError,
    document_validation,
    load_yaml,
    read_input_file,
    read_input_stream,
    resolve_asset_roots,
    validate_xml_strings,
)
from md_parse_frontmatter import FrontmatterParser as FrontmatterParser
from md_parse_frontmatter import _extract_leading_metadata_pair as _extract_leading_metadata_pair
from md_parse_frontmatter import _infer_company_from_title as _infer_company_from_title
from md_parse_frontmatter import _infer_metadata_from_elements as _infer_metadata_from_elements
from md_parse_tables import TableParser as TableParser
from md_parse_tables import TableSpanResolver as TableSpanResolver
from md_parse_text import _DRIVE_PATH_RE as _DRIVE_PATH_RE
from md_parse_text import _URI_DRIVE_PATH_RE as _URI_DRIVE_PATH_RE
from md_parse_text import _URI_SCHEME_RE as _URI_SCHEME_RE
from md_parse_text import TextParser as TextParser
from md_parse_text import _document_int as _document_int
from md_parse_text import _parse_term_runs as _parse_term_runs
from md_parse_text import _rebase_link_target as _rebase_link_target
from md_parse_text import _substitute_text as _substitute_text
from md_parse_text import file_uri as file_uri
from md_parse_text import rebase_link_target as rebase_link_target
from numeric_checks import evaluate_checks, parse_checks
from term_variables import (
    TermResolver,
    TokenMap,
    restore_terms,
    split_term_runs,
    validate_terms,
)

logger = logging.getLogger(__name__)


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
# HTML TABLES (converter output)
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class _HtmlCell:
    """One `<td>`/`<th>` of the outermost table while it is being read."""

    colspan: int = 1
    rowspan: int = 1
    runs: list[TextRun] = field(default_factory=list)


def _html_image(source: str, alt: str) -> Image | None:
    """Return the image an HTML `src` names; a `data:` URI keeps its bytes.

    Args:
        source: The `src` attribute.
        alt: The `alt` attribute.

    Returns:
        The image, or None without a usable source.
    """
    source = source.strip()
    if source.lower().startswith("data:"):
        header, _, payload = source.partition(",")
        if not header.lower().endswith(";base64") or not payload.strip():
            return None
        return Image(
            alt_text=alt,
            path="",
            base64_data=payload.strip(),
            mime_type=header[5:-7] or "image/png",
        )
    if not source or any(character in source for character in "<>\n"):
        return None
    return Image(alt_text=alt, path=source)


class _HtmlTableReader(HTMLParser):
    """Collect the outermost table's cells as runs; HTML text is always literal.

    Formatting tags nest (`<b><i>x</i></b>` is bold italic) and end with their
    cell. Block tags (`p`, `div`, `li`, headings, nested rows and tables) start
    and end cell lines; every `<br>` breaks a line.
    """

    _STYLE_TAGS = {
        "b": "bold",
        "strong": "bold",
        "i": "italic",
        "em": "italic",
        "code": "code",
        "tt": "code",
        "kbd": "code",
        "sup": "superscript",
        "sub": "subscript",
    }
    _BLOCK_TAGS = frozenset(
        {
            "p",
            "div",
            "li",
            "ul",
            "ol",
            "tr",
            "blockquote",
            "pre",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
        }
    )
    _MAX_SPAN = 63  # Word's column limit bounds any sensible span

    def __init__(self, terms: TermResolver | None = None) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[_HtmlCell]] = []
        self.stray: list[str] = []
        self.extra_tables = 0  # further top-level tables in the same block (not read)
        self._terms = terms
        self._depth = 0
        self._cell: _HtmlCell | None = None
        self._styles: dict[str, int] = {}
        self._links: list[str] = []
        self._colors: list[str | None] = []
        self._line_break = False  # a block boundary: the next content starts a new line
        self._separator = False  # a nested cell boundary: the next content is spaced

    @classmethod
    def _span(cls, value: str | None) -> int:
        try:
            return max(1, int(str(value).strip()))
        except ValueError:
            return 1

    def _line_has_content(self) -> bool:
        for run in reversed(self._cell.runs if self._cell is not None else []):
            if run.text == "\n":
                return False
            if run.image is not None or run.text.strip():
                return True
        return False

    def _add(self, run: TextRun) -> None:
        """Append content, first realizing a pending line break or nested-cell space."""
        if self._cell is None:
            return
        if self._line_break:
            if self._line_has_content():
                self._cell.runs.append(TextRun(text="\n"))
        elif self._separator and self._line_has_content():
            self._cell.runs.append(TextRun(text=" "))
        self._line_break = self._separator = False
        self._cell.runs.append(run)

    def _formatted(self, text: str) -> TextRun:
        """A run of literal text under the open formatting tags."""
        link = next((href for href in reversed(self._links) if href), "")
        return TextRun(
            text=text,
            bold=self._styles.get("bold", 0) > 0,
            italic=self._styles.get("italic", 0) > 0,
            superscript=self._styles.get("superscript", 0) > 0,
            subscript=self._styles.get("subscript", 0) > 0,
            code=self._styles.get("code", 0) > 0,
            color_hex=next((color for color in reversed(self._colors) if color), None),
            hyperlink=TextParser.link_target(link) if link else None,
        )

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = {name: value or "" for name, value in attrs}
        if tag == "table":
            if self._depth == 0 and self.rows:
                self.extra_tables += 1
            elif self._cell is not None:
                self._line_break = True  # a nested table starts on its own line
            self._depth += 1
            return
        if self.extra_tables or self._depth == 0:
            return
        if self._depth == 1 and tag == "tr":
            self.rows.append([])
            self._cell = None
            return
        if self._depth == 1 and tag in ("td", "th"):
            if not self.rows:
                self.rows.append([])
            self._cell = _HtmlCell(
                self._span(attributes.get("colspan")), self._span(attributes.get("rowspan"))
            )
            self._styles, self._links, self._colors = {}, [], []
            self._line_break = self._separator = False
            self.rows[-1].append(self._cell)
            return
        if self._cell is None:
            return
        if tag in self._BLOCK_TAGS:
            self._line_break = True
            if tag == "li":
                self._add(TextRun(text="• "))
        elif tag in ("td", "th"):
            self._separator = True
        elif tag == "br":
            self._line_break = self._separator = False
            self._cell.runs.append(TextRun(text="\n"))
        elif tag in self._STYLE_TAGS:
            name = self._STYLE_TAGS[tag]
            self._styles[name] = self._styles.get(name, 0) + 1
        elif tag == "a":
            self._links.append(attributes.get("href", "").strip())
        elif tag in ("span", "font"):
            match = TextParser._COLOR_STYLE_RE.search(attributes.get("style", ""))
            color = match.group(1) if match else attributes.get("color", "")
            self._colors.append(color.upper() if re.fullmatch(r"#[0-9A-Fa-f]{6}", color) else None)
        elif tag == "img":
            image = _html_image(attributes.get("src", ""), attributes.get("alt", ""))
            if image is not None:
                self._add(TextRun(text="", image=image))

    def handle_endtag(self, tag: str) -> None:
        if tag == "table":
            self._depth = max(self._depth - 1, 0)
            if self._cell is not None:
                self._line_break = True  # text after a nested table starts a new line
            return
        if self.extra_tables or self._depth == 0:
            return
        if self._depth == 1 and tag in ("td", "th"):
            self._cell = None
            return
        if self._cell is None:
            return
        if tag in self._BLOCK_TAGS:
            self._line_break = True
        elif tag in self._STYLE_TAGS:
            name = self._STYLE_TAGS[tag]
            self._styles[name] = max(self._styles.get(name, 0) - 1, 0)
        elif tag == "a" and self._links:
            self._links.pop()
        elif tag in ("span", "font") and self._colors:
            self._colors.pop()

    def handle_data(self, data: str) -> None:
        if self.extra_tables:
            return
        if self._cell is None:
            if data.strip():
                self.stray.append(data.strip())
            return
        text = re.sub(r"[ \t\r\n\f]+", " ", data)
        if not text.strip() and (
            self._line_break or self._separator or not self._line_has_content()
        ):
            return  # layout whitespace between blocks
        self._add(self._formatted(text))

    def finish(self, runs: list[TextRun]) -> list[TextRun]:
        """Make one cell's final runs, as HTML shows them.

        Adjacent runs with the same formatting are merged first, so an amount
        or a `{{key}}` reference split by a neutral tag such as `<span>` is
        whole; then references are substituted, spaces HTML would not show are
        removed and line breaks at the cell edges dropped. Code and term values
        keep their text, including an empty value.

        Args:
            runs: Runs collected for one cell.

        Returns:
            The cell's runs.
        """
        merged: list[TextRun] = []
        for run in runs:
            previous = merged[-1] if merged else None
            if (
                previous is not None
                and previous.image is None
                and run.image is None
                and "\n" not in (previous.text, run.text)
                and replace(previous, text="") == replace(run, text="")
            ):
                tail = run.text.lstrip(" ") if previous.text.endswith(" ") else run.text
                merged[-1] = replace(previous, text=previous.text + tail)
            else:
                merged.append(run)
        substituted: list[TextRun] = []
        for run in merged:
            if self._terms is None or run.code or "{{" not in run.text:
                substituted.append(run)
                continue
            self._terms.reserve(run.text)  # literal token-shaped text never becomes a value
            tokens: TokenMap = {}
            tokenized = TextParser._tokenize_references(run.text, {}, {}, self._terms, tokens)
            substituted.extend(split_term_runs([replace(run, text=tokenized)], tokens))

        def fixed(run: TextRun) -> bool:
            return run.image is not None or run.term_key is not None or run.code or run.text == "\n"

        cleaned: list[TextRun] = []
        for run in substituted:
            text = run.text
            if not fixed(run):
                previous = cleaned[-1] if cleaned else None
                if (
                    previous is None
                    or previous.text == "\n"
                    or (previous.image is None and previous.text.endswith(" "))
                ):
                    text = text.lstrip(" ")
            if text or run.image is not None or run.term_key is not None:
                cleaned.append(replace(run, text=text))
        for index, run in enumerate(cleaned):
            following = cleaned[index + 1] if index + 1 < len(cleaned) else None
            if (following is None or following.text == "\n") and not fixed(run):
                cleaned[index] = replace(run, text=run.text.rstrip(" "))
        cleaned = [
            run for run in cleaned if run.text or run.image is not None or run.term_key is not None
        ]
        while cleaned and cleaned[0].text == "\n":
            cleaned.pop(0)
        while cleaned and cleaned[-1].text == "\n":
            cleaned.pop()
        return cleaned


class HtmlTableParser:
    """Read an HTML `<table>` block, as document converters write them.

    `colspan`/`rowspan` become the engine's `<<`/`^^` span markers, so the
    table renders merged whatever the profile's span default. Cells are built
    as runs, so HTML text is literal (never re-read as Markdown): `<br>`, `<p>`,
    `<div>` and `<li>` start cell lines; `<b>`/`<strong>`, `<i>`/`<em>`,
    `<code>`, `<sup>`/`<sub>`, colour spans and `<a href>` format text; `<img>`
    is an inline image; `{{key}}` term references are substituted. A nested
    table is flattened into its cell, one line per row. Empty rows are kept
    while spans are placed, then dropped when nothing covers them.

    The header is the first row plus any rows its row spans cover; every header
    row is drawn and repeated as a header (`Table.header_rows`).
    """

    START_RE = re.compile(r"^\s*<table\b", re.IGNORECASE)
    _OPEN_RE = re.compile(r"<table\b", re.IGNORECASE)
    _CLOSE_RE = re.compile(r"</table\s*>", re.IGNORECASE)

    @classmethod
    def block_end(cls, lines: list[str], start: int) -> int | None:
        """Return the index after the line closing the table opened at `start`.

        Args:
            lines: Document lines.
            start: Index of the line that opens the table.

        Returns:
            The end index, or None when the table is never closed.
        """
        depth = 0
        for index in range(start, len(lines)):
            depth += len(cls._OPEN_RE.findall(lines[index])) - len(
                cls._CLOSE_RE.findall(lines[index])
            )
            if depth <= 0:
                return index + 1
        return None

    @staticmethod
    def _table_cell(runs: list[TextRun]) -> TableCell:
        """A prepared literal cell; `content` is its text with breaks as `<br>`."""
        content = "".join(run.text for run in runs).replace("\n", "<br>")
        return TableCell(content=content, runs=runs, literal=True)

    @classmethod
    def parse(
        cls,
        html: str,
        terms: TermResolver | None = None,
        limits: InputLimits = DEFAULT_LIMITS,
    ) -> tuple[list[list[str | TableCell]], int, list[str]]:
        """Convert a complete table block into rows of cells.

        Args:
            html: The `<table>...</table>` source.
            terms: Active term resolver, or None when the document has no `terms:`.
            limits: Operator-only expanded-grid limit, checked before allocating it.

        Returns:
            The rows (covered positions hold span markers), how many leading
            rows form the header, and warnings about content that could not be
            placed.
        """
        reader = _HtmlTableReader(terms)
        reader.feed(html)
        reader.close()
        warnings = [
            f"Text outside the HTML table cells was dropped: {text}" for text in reader.stray
        ]
        if reader.extra_tables:
            warnings.append(
                f"{reader.extra_tables} further HTML table(s) on the closing line were dropped; "
                "start each table on its own line"
            )

        runs: dict[tuple[int, int], list[TextRun]] = {}
        anchors: dict[tuple[int, int], tuple[int, int]] = {}
        spans: dict[tuple[int, int], int] = {}
        rows = reader.rows
        if any(
            cell.colspan > reader._MAX_SPAN or cell.rowspan > reader._MAX_SPAN
            for row in rows
            for cell in row
        ):
            cls._check_unclamped_grid(rows, limits)
        for row_index, html_row in enumerate(rows):
            position = 0
            for cell in html_row:
                while (row_index, position) in anchors:
                    position += 1
                # Retain the existing Word span cap, after checking the written grid.
                rowspan = min(cell.rowspan, reader._MAX_SPAN, len(rows) - row_index)
                colspan = min(cell.colspan, reader._MAX_SPAN)
                limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(rows) * (position + colspan))
                spans[(row_index, position)] = rowspan
                runs[(row_index, position)] = reader.finish(cell.runs)
                for down in range(rowspan):
                    for right in range(colspan):
                        anchors[(row_index + down, position + right)] = (row_index, position)
                position += colspan
        if not anchors:
            return [], 0, warnings
        col_count = max(covered for _, covered in anchors) + 1
        kept = [
            row
            for row in range(len(rows))
            if any((row, column) in anchors for column in range(col_count))
        ]

        header_end = kept[0] + 1
        grown = True
        while grown:
            grown = False
            for (top, _), rowspan in spans.items():
                if top < header_end < top + rowspan:
                    header_end, grown = top + rowspan, True

        def source(row: int, column: int) -> str | TableCell:
            anchor = anchors.get((row, column))
            if anchor is None:
                return ""
            if anchor == (row, column):
                return cls._table_cell(runs[anchor])
            return "^^" if column == anchor[1] else "<<"

        cells = [[source(row, column) for column in range(col_count)] for row in kept]
        return cells, sum(1 for row in kept if row < header_end), warnings

    @staticmethod
    def _check_unclamped_grid(rows: list[list[_HtmlCell]], limits: InputLimits) -> None:
        """Count written spans before Word's existing span cap can reduce them."""
        occupied: set[tuple[int, int]] = set()
        for row_index, row in enumerate(rows):
            column = 0
            for cell in row:
                while (row_index, column) in occupied:
                    column += 1
                limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(rows) * (column + cell.colspan))
                occupied.update(
                    (row_index + down, column + right)
                    for down in range(min(cell.rowspan, len(rows) - row_index))
                    for right in range(cell.colspan)
                )
                column += cell.colspan

    @staticmethod
    def image(line: str) -> Image | None:
        """Return the image of a line that is a single HTML `<img>` tag.

        Args:
            line: Stripped document line.

        Returns:
            The image, or None when the line is not one `<img>` with a usable `src`.
        """
        found: list[dict[str, str]] = []

        class _Reader(HTMLParser):
            def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
                if tag == "img":
                    found.append({name: value or "" for name, value in attrs})

        reader = _Reader(convert_charrefs=True)
        reader.feed(line)
        reader.close()
        if len(found) != 1:
            return None
        return _html_image(found[0].get("src", ""), found[0].get("alt", ""))


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
    # Levels 5 and 6 have no style of their own; they render as the deepest styled level.
    DEEP_HEADING_PATTERN = re.compile(r"^#{5,6}\s+(.+)$")

    # Numbered heading pattern: **1. Title** or **1\. Title**
    NUMBERED_HEADING_PATTERN = re.compile(r"^\*\*\d+(\.|\\.)")

    # ── List patterns ───────────────────────────────────────────────────────
    BULLET_PATTERN = re.compile(r"^([ \t]*)[-*+]\s+(.+)$")
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
    _HTML_IMAGE_LINE_RE = re.compile(r"^<img\b[^<>]*>$", re.IGNORECASE)
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
    _INVALID_XML_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
    _OPEN_PAREN_SPACE_RE = re.compile(r"\(\s+")
    _CLOSE_PAREN_SPACE_RE = re.compile(r"\s+\)")

    _EXT_FOOTNOTE_DEF_RE = re.compile(r"^\[\^(\d+)\]:\s*(.+)$")
    _SETEXT_H1_RE = re.compile(r"^=+\s*$")
    _LINK_DEFINITION_RE = re.compile(
        r"^[ \t]*\[(?!\^)([^\]\n]+)\]:[ \t]*(?:<([^>\n]+)>|(\S+))"
        r"(?:[ \t]+(?:\"[^\"\n]*\"|'[^'\n]*'))?[ \t]*$",
        re.MULTILINE,
    )
    _REFERENCE_LINK_RE = re.compile(r"(?<!!)(?<!\\)\[(?!\^)([^\]\n]+)\](?:\[([^\]\n]*)\])?(?!\()")
    _REFERENCE_LABEL_SPACE_RE = re.compile(r"\s+")

    def __init__(
        self,
        preserve_trailing_double_space_break: bool = False,
        profile: str | None = None,
        base_dir: Path | None = None,
        asset_roots: Sequence[str | Path] = (),
        limits: InputLimits = DEFAULT_LIMITS,
    ):
        """
        Initialize parser behavior flags.

        Args:
            preserve_trailing_double_space_break:
                When True, treat markdown trailing 2+ spaces as hard line breaks.
                Default False to avoid accidental Shift+Enter artifacts from noisy input.
            profile: Explicit profile override.
            base_dir: Base for relative destinations; does not itself grant asset access.
            asset_roots: Operator-supplied folders carried to rendering. Text has no default.
            limits: Immutable caller policy carried to rendering; text size uses UTF-8.
        """
        self.preserve_trailing_double_space_break = preserve_trailing_double_space_break
        self.profile = profile
        # Folder of the source file: relative link paths are resolved against it
        # when the document is saved (see `rebase_local_links`).
        self.base_dir = base_dir
        self.asset_roots = resolve_asset_roots(asset_roots)
        self.limits = limits
        self._financial_rules = True
        self._term_sheet = False
        self._parse_warnings: list[str] = []
        # Per-parse `terms:` state; None keeps `{{...}}` ordinary text.
        self._terms: TermResolver | None = None

    def parse(self, content: str) -> DocumentModel:
        """
        Parse markdown content into a DocumentModel.

        Args:
            content: The full markdown content

        Returns:
            A DocumentModel with all parsed elements
        """
        self.limits.check(RejectionCode.LIMIT_INPUT_BYTES, len(content.encode("utf-8")))
        return self._parse_content(content)

    def _parse_content(self, content: str) -> DocumentModel:
        """Parse text after the public text/file/stream boundary checked original bytes."""
        lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")

        # Parse frontmatter
        metadata, remaining_lines = FrontmatterParser.parse(lines, self.profile)
        # Required-text checks precede heading inference; length waits for terms.
        with document_validation():
            validate_term_sheet_metadata(metadata, check_length=False)
        self._financial_rules = get_profile(metadata.profile).is_ib
        self._term_sheet = metadata.profile == "term-sheet"
        self._parse_warnings = []
        # An empty mapping enables checking: every reference is then undefined.
        with document_validation():
            self._terms = (
                TermResolver(validate_terms(metadata.extra["terms"]), corpus=content)
                if "terms" in metadata.extra
                else None
            )
            validate_xml_strings(metadata)

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
        footnotes = FootnoteParser.extract_references(prose_lines) if self._financial_rules else {}
        input_warnings: list[str] = []
        invalid_character = self._INVALID_XML_CHAR_RE.search(content)
        if invalid_character is not None:
            input_warnings.append(
                f"XML-incompatible character U+{ord(invalid_character.group()):04X} in Markdown input"
            )
        explicit_references: set[int] = set()
        content_lines: list[str] = []
        for index, raw_line in enumerate(remaining_lines):
            if index in fenced_indices:
                content_lines.append(raw_line)
                continue
            match = self._EXT_FOOTNOTE_DEF_RE.match(raw_line.strip())
            if not self._EXT_FOOTNOTE_DEF_RE.match(scan_lines[index].strip()):
                match = None
            if match:
                number = _document_int(match.group(1))
                if number <= 0:
                    raise StrictValidationError("", ["Footnote IDs must be positive integers"])
                if number in footnotes:
                    input_warnings.append(f"Duplicate footnote definition: {number}")
                footnotes[number] = match.group(2)
            else:
                explicit_references.update(
                    run.footnote_id
                    for run in TextParser.parse_runs(scan_lines[index])
                    if run.footnote_id is not None
                )
                content_lines.append(raw_line)
        remaining_lines = content_lines

        # Parse elements
        elements = self._parse_elements(remaining_lines)
        input_warnings.extend(self._parse_warnings)

        # If no YAML frontmatter, extract metadata from document header
        if not has_frontmatter and self._financial_rules:
            elements = self._extract_header_metadata(metadata, elements)

        input_warnings.extend(
            f"Undefined footnote: {number}"
            for number in sorted(explicit_references - set(footnotes))
        )
        # Title inference follows what the author wrote, not a substituted value.
        written_title = metadata.title
        self._substitute_metadata(metadata)
        with document_validation():
            validate_term_sheet_metadata(metadata)
        model = DocumentModel(
            metadata=metadata,
            elements=elements,
            footnotes=footnotes,
            warnings=input_warnings,
            parsed_profile=metadata.profile,
            source_dir=self.base_dir,
            asset_roots=self.asset_roots,
            limits=self.limits,
        )
        if not self._financial_rules:
            first_heading = next(
                (
                    e.content
                    for e in elements
                    if e.element_type == ElementType.HEADING_1 and isinstance(e.content, Heading)
                ),
                None,
            )
            if (
                written_title in {"", "Document", "IB Report"}
                and first_heading
                and first_heading.text
            ):
                metadata.title = first_heading.text
                if first_heading.runs:
                    metadata.display_runs["title"] = [replace(run) for run in first_heading.runs]
        with document_validation(errors=(ValueError, TypeError)):
            apply_table_specs(model)
        # Span markers are classified on raw cell source, before term text changes.
        for element in elements:
            if element.element_type == ElementType.TABLE and isinstance(element.content, Table):
                TableSpanResolver.resolve(element.content, metadata.profile)
        self._substitute_table_text(elements)
        model.warnings.extend(
            warning
            for element in elements
            if element.element_type == ElementType.TABLE and isinstance(element.content, Table)
            for warning in element.content.warnings
        )
        check_schedules(model)
        if "checks" in metadata.extra:
            with document_validation():
                checks = parse_checks(metadata.extra["checks"])
            values = dict(self._terms.values) if self._terms is not None else {}
            model.warnings.extend(evaluate_checks(checks, values))
        if self._terms is not None:
            model.warnings.extend(self._terms.warnings())
            unused = self._terms.unused()
            if unused:
                logger.info("Unused terms (defined but never referenced): %s", ", ".join(unused))
        return model

    # ── Term substitution ───────────────────────────────────────────────────

    def _term_runs(
        self,
        text: str,
        parse: Callable[[str], list[TextRun]],
        latex: bool = True,
    ) -> tuple[str, list[TextRun], bool]:
        """Parse one field's runs with this document's terms (see `_parse_term_runs`)."""
        return _parse_term_runs(text, parse, self._terms, latex=latex)

    def _substitute_heading(self, element: Element) -> Element:
        """Carry value runs for a heading that references terms.

        Resolve escapes once so escaped braces cannot become term references.

        Args:
            element: A parsed heading element.

        Returns:
            The same element; text and runs change only when a value is inserted.
        """
        heading = element.content
        if self._terms is None or not isinstance(heading, Heading):
            return element
        tokens: TokenMap = {}
        if element.element_type == ElementType.NUMBERED_HEADING:
            # Its text is already unescaped, so substitute from the source line.
            tokenized = TextParser.cleanup_text(
                TextParser.tokenize_terms(element.raw_text, self._terms, tokens)
            )
        else:
            tokenized = TextParser.tokenize_terms(heading.text, self._terms, tokens)
        if tokens:
            heading.text = restore_terms(tokenized, tokens)
            # HeadingRenderer drops bold markers because it bolds every run.
            heading.runs = split_term_runs(
                TextParser.parse_runs(tokenized.replace("**", "").strip()),
                tokens,
            )
        return element

    def _substitute_metadata(self, metadata: DocumentMetadata) -> None:
        """Substitute title, subtitle and date, keeping their runs for display."""
        if self._terms is None:
            return
        for name in ("title", "subtitle", "date"):
            source = metadata.extra.get(name) if name == "date" else getattr(metadata, name)
            if not isinstance(source, str):
                continue
            text, runs, substituted = self._term_runs(source, TextParser.parse_runs)
            if not substituted:
                continue
            metadata.display_runs[name] = runs
            if name == "date":
                metadata.extra[name] = text
            else:
                setattr(metadata, name, text)

    def _substitute_table_text(self, elements: list[Element]) -> None:
        """Substitute table captions, units, sources, dates and notes as untagged text.

        The renderer reads these fields as inline Markdown, so values are
        escaped to stay literal there.
        """
        if self._terms is None:
            return
        for element in elements:
            if element.element_type != ElementType.TABLE or not isinstance(element.content, Table):
                continue
            for name in ("caption", "unit", "source", "as_of", "note"):
                tokens: TokenMap = {}
                tokenized = TextParser.tokenize_terms(
                    getattr(element.content, name), self._terms, tokens
                )
                if tokens:
                    setattr(
                        element.content,
                        name,
                        restore_terms(tokenized, tokens, TextParser.escape_literal),
                    )

    @staticmethod
    def _strip_comments(lines: list[str]) -> list[str]:
        """Omit HTML comments while retaining fenced and inline code verbatim."""
        text = "\n".join(lines)
        offsets: list[int] = []
        offset = 0
        for line in lines:
            offsets.append(offset)
            offset += len(line) + 1
        line_starts = {start: index for index, start in enumerate(offsets)}
        pieces: list[str] = []
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

    @staticmethod
    def _wrap_destination(destination: str) -> str:
        """Angle-bracket a definition's destination so any path form stays a link."""
        if any(character in destination for character in "<>\n"):
            return destination
        return "<" + destination + ">"

    @classmethod
    def _resolve_reference_links(cls, lines: list[str]) -> list[str]:
        """Collect link definitions and expand resolved references outside code."""
        literals: dict[str, str] = {}
        prefix = "\ue000REF"
        source = "\n".join(lines)
        while prefix in source:
            prefix += "X"

        def protect(value: str) -> str:
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = value
            return token

        pieces: list[str] = []
        index = 0
        while index < len(lines):
            block = FenceScanner.scan(lines, index)
            if block:
                pieces.append(protect("\n".join(lines[index : block[2]])))
                index = block[2]
            else:
                pieces.append(lines[index])
                index += 1
        source = TextParser._CODE_SPAN_RE.sub(
            lambda match: protect(match.group(0)), "\n".join(pieces)
        )
        source = TextParser._ESCAPE_RE.sub(lambda match: protect(match.group(0)), source)
        source = TextParser._INLINE_REFERENCE_RE.sub(lambda match: protect(match.group(0)), source)
        definitions: dict[str, str] = {}

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
            return f"[{match.group(1)}]({cls._wrap_destination(destination)})"

        source = cls._REFERENCE_LINK_RE.sub(resolve, source)
        # Later protection layers may contain tokens from earlier layers.
        for token, value in reversed(list(literals.items())):
            source = source.replace(token, value)
        return source.split("\n")

    # ── Element-level parsing ───────────────────────────────────────────────

    def _parse_elements(self, lines: list[str]) -> list[Element]:
        """Parse lines into elements"""
        elements: list[Element] = []
        i = 0
        in_references = False
        chart_number = 0

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
                fence_start = i
                language, code_lines, i = code_block
                code_text = "\n".join(code_lines).rstrip()

                if self._term_sheet and language == "confirmation":
                    source = "\n".join(lines[fence_start:i])
                    closed = i - fence_start == len(code_lines) + 2
                    if closed and not code_text.strip():
                        elements.append(
                            Element(ElementType.CONFIRMATION, ConfirmationBlock(source), source)
                        )
                    else:
                        reason = "non-empty" if closed else "unclosed"
                        self._parse_warnings.append(
                            f"Invalid confirmation fence ({reason}); original source kept as code"
                        )
                        elements.append(
                            Element(
                                ElementType.CODE_BLOCK,
                                CodeBlock(code=source, language=language),
                                source,
                            )
                        )
                    continue

                if language == "chart":
                    chart_number += 1
                    chart = Chart(code=code_text, label=f"Chart {chart_number}")
                    try:
                        chart.spec = parse_chart_spec(code_text)
                        if chart.spec.title:
                            chart.label += f" ({chart.spec.title})"
                    except ChartSpecError as exc:
                        chart.error = str(exc)
                    elements.append(Element(ElementType.CHART, chart, raw_line))
                    continue

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
                # A closed fence is literal: its trailing blank lines and spaces stay.
                # An unclosed one runs to the end of the input, so its tail is trimmed.
                if i - fence_start == len(code_lines) + 2:
                    code_text = "\n".join(code_lines)

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
                latex_lines: list[str] = []
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

            # ── HTML table (converter output; spans become span markers) ────
            if HtmlTableParser.START_RE.match(line):
                end = HtmlTableParser.block_end(lines, i)
                if end is not None:
                    source = "\n".join(lines[i:end])
                    cells, header_rows, html_warnings = HtmlTableParser.parse(
                        source, self._terms, limits=self.limits
                    )
                    self._parse_warnings.extend(html_warnings)
                    if cells:
                        table = TableParser.from_cells(
                            cells,
                            [],
                            financial_rules=self._financial_rules,
                            terms=self._terms,
                            header_rows=header_rows,
                            limits=self.limits,
                        )
                        table.spans = True
                        elements.append(
                            Element(element_type=ElementType.TABLE, content=table, raw_text=source)
                        )
                    else:
                        self._parse_warnings.append(
                            f"HTML table at line {i + 1} has no cells; it was dropped"
                        )
                    i = end
                    continue
                self._parse_warnings.append(f"Unclosed HTML table at line {i + 1} was kept as text")

            # ── Table (collect all contiguous table lines) ──────────────────
            if self.TABLE_START_PATTERN.match(line):
                table_lines: list[str] = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    table_lines.append(lines[i].strip())
                    i += 1
                table = TableParser.parse(
                    table_lines,
                    financial_rules=self._financial_rules,
                    terms=self._terms,
                    limits=self.limits,
                )
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
                    elements.append(
                        self._substitute_heading(
                            Element(
                                element_type=ElementType.HEADING_1,
                                content=Heading(level=1, text=line),
                                raw_text=raw_line + "\n" + lines[i + 1],
                            )
                        )
                    )
                    i += 2
                    continue
            element = self._try_parse_heading(line)
            if element:
                elements.append(self._substitute_heading(element))
                i += 1
                continue

            # ── Blockquote (merge consecutive > lines) ──────────────────────
            match = self.BLOCKQUOTE_PATTERN.match(line)
            if match:
                bq_lines: list[str] = []
                quote_tokens: TokenMap = {}
                while i < len(lines):
                    bq_match = self.BLOCKQUOTE_PATTERN.match(lines[i].strip())
                    if bq_match:
                        source = bq_match.group(1)
                        if self._terms is not None:
                            # Callouts render their body without inline math.
                            source = TextParser.tokenize_terms(
                                source,
                                self._terms,
                                quote_tokens,
                                latex=False,
                            )
                        bq_lines.append(TextParser.cleanup_text(source))
                        i += 1
                    else:
                        break

                title, body = self._extract_blockquote_title(bq_lines)
                quote = Blockquote(text=body, title=title)
                if quote_tokens:
                    quote.text = restore_terms(body, quote_tokens)
                    quote.runs = split_term_runs(TextParser.parse_runs_plain(body), quote_tokens)
                elements.append(
                    Element(
                        element_type=ElementType.BLOCKQUOTE,
                        content=quote,
                        raw_text=restore_terms("\n".join(bq_lines), quote_tokens),
                    )
                )
                continue

            # ── Bullet list ─────────────────────────────────────────────────
            match = self.BULLET_PATTERN.match(raw_line)
            if match:
                indent_level = self._get_indent_level(match.group(1))
                text, next_idx = self._collect_paragraph(lines, i, first_line=match.group(2))
                text, runs, _ = self._term_runs(text, TextParser.parse_runs)
                item = ListItem(
                    text=text,
                    runs=runs,
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
                text, runs, _ = self._term_runs(text, TextParser.parse_runs)
                item = ListItem(
                    text=text,
                    runs=runs,
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
            file_image: Image | None = None
            if match:
                file_image = Image(alt_text=match.group(1), path=match.group(2) or match.group(3))
            elif self._HTML_IMAGE_LINE_RE.match(line):
                file_image = HtmlTableParser.image(line)
            if file_image is not None:
                elements.append(
                    Element(element_type=ElementType.IMAGE, content=file_image, raw_text=line)
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
        self, lines: list[str], start_idx: int, first_line: str | None = None
    ) -> tuple[str, int]:
        """
        Collect contiguous paragraph lines and normalize Markdown line breaks.

        CommonMark semantics:
            - soft line break (single newline in paragraph) -> space
            - hard line break (2+ trailing spaces, trailing backslash, <br>) -> newline
        """
        parts: list[str] = []
        i = start_idx
        use_hard_break = False

        while i < len(lines):
            raw_line = first_line if i == start_idx and first_line is not None else lines[i]
            line = self._strip_html_anchors(raw_line).strip()

            if not line:
                break

            if i > start_idx and self._starts_new_block(line):
                break
            if (
                i > start_idx
                and i + 1 < len(lines)
                and self._SETEXT_H1_RE.fullmatch(lines[i + 1].strip())
            ):
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
            or HtmlTableParser.START_RE.match(line)
            or self._HTML_IMAGE_LINE_RE.match(line)
            or self._try_parse_heading(line)
            or self.BLOCKQUOTE_PATTERN.match(line)
            or self.BULLET_PATTERN.match(line)
            or self.NUMBERED_LIST_PATTERN.match(line)
            or self.IMAGE_PATTERN.match(line)
            or self._is_reference_header(line)
        )

    def _normalize_paragraph_line(self, raw_line: str) -> tuple[str, bool]:
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
        level = max(0, len(expanded) // (2 if self._financial_rules else 4))
        self.limits.check(RejectionCode.LIMIT_LIST_DEPTH, level + 1)
        return level

    # ── Paragraph parsing (ENHANCED v3) ─────────────────────────────────────

    def _parse_paragraph(self, line: str) -> Element:
        """
        Parse a line as a paragraph, detecting inline LaTeX if present.

        If the line contains inline $...$ expressions, the resulting
        TextRuns will have is_latex=True for those segments, allowing
        the renderer to handle them appropriately.
        """
        has_latex = LaTeXParser.has_inline(line)
        text, runs, _ = self._term_runs(line, TextParser.parse_runs)

        para = Paragraph(
            text=text,
            runs=runs,
            has_inline_latex=has_latex,
        )

        return Element(
            element_type=ElementType.PARAGRAPH,
            content=para,
            raw_text=line,
        )

    # ── Heading parsing helpers ─────────────────────────────────────────────

    def _try_parse_heading(self, line: str) -> Element | None:
        """Try to parse line as a heading"""

        # Numbered heading (**1. Title**)
        if self._financial_rules and self.NUMBERED_HEADING_PATTERN.match(line):
            text = TextParser.cleanup_text(line)
            return Element(
                element_type=ElementType.NUMBERED_HEADING,
                content=Heading(level=1, text=text, is_numbered=True),
                raw_text=line,
            )

        # H5 and H6 share the deepest styled level (check longer prefixes first)
        match = self.DEEP_HEADING_PATTERN.match(line) or self.H4_PATTERN.match(line)
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

        try:
            number = int(match.group(1))
        except ValueError:
            # Too long for Python to convert; no section number is that long.
            return False
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

    def _parse_diagram(self, language: str, code_text: str) -> Diagram | None:
        """Parse a diagram:flow YAML block into a Diagram object."""
        try:
            data = load_yaml(code_text)
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
        self, metadata: DocumentMetadata, elements: list[Element]
    ) -> list[Element]:
        """Extract IB header metadata while retaining body content.

        Args:
            metadata: Metadata to populate when YAML frontmatter is absent.
            elements: Parsed elements retaining paragraph source line boundaries.

        Returns:
            Body elements, including the marked inferred subtitle and unknown labels.
        """
        filtered: list[Element] = []
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
                                following, _ = self._normalize_paragraph_line(
                                    source_lines[line_index + 1]
                                )
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
            or self.DEEP_HEADING_PATTERN.match(line)
            or self.NUMBERED_HEADING_PATTERN.match(line)
        )

    def _is_reference_header(self, line: str) -> bool:
        """Check if line is a references section header"""
        stripped = line.strip().lower()
        cleaned = re.sub(r"^#{1,6}\s+", "", stripped)
        return cleaned.strip("* :：") in self._REFERENCE_KEYWORDS

    # ── Blockquote helpers ──────────────────────────────────────────────────

    @staticmethod
    def _extract_blockquote_title(
        bq_lines: list[str],
    ) -> tuple[str, str]:
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


def _read_with_encoding(file_path: str, limits: InputLimits = DEFAULT_LIMITS) -> str:
    """Read a file without statistically misdetecting valid UTF-8 Korean text."""
    return _decode_bytes(read_input_file(Path(file_path), limits), label=file_path)


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
    raise UnicodeDecodeError(
        "multiple", raw_bytes, 0, min(1, len(raw_bytes)), f"Failed to decode {label}"
    )


def parse_markdown_file(
    source: "str | BinaryIO",
    profile: str | None = None,
    *,
    asset_roots: Sequence[str | Path] = (),
    limits: InputLimits = DEFAULT_LIMITS,
) -> DocumentModel:
    """
    Parse a markdown file or stream into a DocumentModel.

    Accepts a file path (str) or a binary stream (BinaryIO).
    Tries intelligent encoding detection, falls back through
    UTF-8 → EUC-KR → CP949.

    Args:
        source: Path to the markdown file, or a readable binary stream.
        asset_roots: Operator-supplied folders added to the source-file folder.
        limits: Immutable caller policy; the original bytes are checked before decoding.

    Returns:
        A DocumentModel containing all parsed content

    Raises:
        FileNotFoundError: If source is a path and the file does not exist
        UnicodeDecodeError: If none of the attempted encodings work
    """
    from stream_utils import is_stream

    if is_stream(source):
        raw_bytes = read_input_stream(cast(BinaryIO, source), limits)
        content = _decode_bytes(raw_bytes, label="<stream>")
        label = "<stream>"
    else:
        content = _read_with_encoding(str(source), limits)
        label = str(source)

    lines = content.splitlines()
    _, remaining_lines = FrontmatterParser.parse(lines, profile=profile)
    frontmatter_present = remaining_lines is not lines

    base_dir = None if is_stream(source) else Path(str(source)).resolve().parent
    roots = (*((base_dir,) if base_dir is not None else ()), *asset_roots)
    parser = MarkdownParser(profile=profile, base_dir=base_dir, asset_roots=roots, limits=limits)
    model = parser._parse_content(content)
    if get_profile(model.metadata.profile).is_ib:
        _infer_metadata_from_elements(model, allow_company_inference=not frontmatter_present)
    if not is_stream(source) and model.metadata.extra.get("theme") not in {None, "default", "mono"}:
        theme_path = Path(model.metadata.extra["theme"])
        if not theme_path.is_absolute():
            model.metadata.extra["theme"] = str(Path(str(source)).resolve().parent / theme_path)

    house = model.metadata.extra.get("house")
    if not is_stream(source) and isinstance(house, str) and house.strip():
        house_path = Path(house)
        if not house_path.is_absolute():
            model.metadata.extra["house"] = str(
                (Path(str(source)).resolve().parent / house_path).resolve()
            )

    style = model.metadata.extra.get("style")
    if base_dir is not None and isinstance(style, dict) and isinstance(style.get("logo"), str):
        logo = style["logo"].strip()
        if logo and not logo.lower().startswith("data:") and not Path(logo).is_absolute():
            # Anchor the written destination without following links yet: the
            # asset helper must resolve written and decoded candidates separately.
            with document_validation(errors=(ValueError, OSError)):
                model.metadata.extra["style"] = {
                    **style,
                    "logo": str(Path(str(base_dir / logo)).absolute()),
                }

    if not is_stream(source):
        source_dir = Path(str(source)).resolve().parent
        for element in model.elements:
            if isinstance(element.content, Image):
                image = element.content
                if (
                    image.path
                    and not image.base64_data
                    and not re.match(
                        r"^(?:[a-zA-Z][a-zA-Z0-9+.-]*://|data:)", image.path, re.IGNORECASE
                    )
                ):
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
