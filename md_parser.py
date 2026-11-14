"""
MD Parser Module for IB Style Word Report Converter
Handles parsing of Markdown files including frontmatter, elements, tables,
LaTeX equations, Base64 images, and footnotes.

Changelog (captions):
    - NEW: `![설명](파일){#fig-…}` is a numbered figure; an unusable `{#…}` ending
      keeps the line as text and is reported by the renderer wherever it stands.
    - Callouts and numbered headings, unescaped here, keep escaped `@` apart.
    - Second review: a document that may hold references is parsed a second
      time once its numbered tables and figures are known, so only resolvable
      references split runs; title, subtitle, date, image descriptions and
      footnotes are checked as text that cannot hold a field.

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
from dataclasses import replace
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
    CrossReference,
    Figure,
    MarkedBlockquote,
    MarkedHeading,
    figure_id_problem,
    numbered_items,
    numbered_targets,
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
from md_parse_blocks import Base64ImageParser as Base64ImageParser
from md_parse_blocks import FenceScanner as FenceScanner
from md_parse_blocks import FootnoteParser as FootnoteParser
from md_parse_blocks import LaTeXParser as LaTeXParser
from md_parse_frontmatter import FrontmatterParser as FrontmatterParser
from md_parse_frontmatter import _extract_leading_metadata_pair as _extract_leading_metadata_pair
from md_parse_frontmatter import _infer_company_from_title as _infer_company_from_title
from md_parse_frontmatter import _infer_metadata_from_elements as _infer_metadata_from_elements
from md_parse_html_tables import HtmlTableParser as HtmlTableParser
from md_parse_html_tables import _html_image as _html_image
from md_parse_html_tables import _HtmlCell as _HtmlCell
from md_parse_html_tables import _HtmlTableReader as _HtmlTableReader
from md_parse_tables import TableParser as TableParser
from md_parse_tables import TableSpanResolver as TableSpanResolver
from md_parse_text import _DRIVE_PATH_RE as _DRIVE_PATH_RE
from md_parse_text import _URI_DRIVE_PATH_RE as _URI_DRIVE_PATH_RE
from md_parse_text import _URI_SCHEME_RE as _URI_SCHEME_RE
from md_parse_text import ReferenceScope as ReferenceScope
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
    # The `{#…}` ending of a figure line; IMAGE_PATTERN then reads what precedes it.
    _FIGURE_ID_RE = re.compile(r"\{#([^{}\s]*)\}$")
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
        # Term runs of a heading whose references resolve, read without them,
        # for the title when that heading becomes it.
        self._title_runs: dict[int, list[TextRun]] = {}

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
        """Parse text after the public text/file/stream boundary checked original bytes.

        Text without `@tbl-` or `@fig-` is parsed once, outside any reference
        scope, as before references existed. Otherwise it is parsed in a scope
        in which no reference resolves yet; when that parse finds numbered
        tables or figures, the document is parsed again in a scope holding their
        IDs. Only resolvable references split runs; every other one is reported.
        """
        if "@tbl-" not in content and "@fig-" not in content:
            with TextParser.reference_scope(None):
                return self._parse_document(content)
        found: list[str] = []

        def report(message: str) -> None:
            if message not in found:
                found.append(message)

        with TextParser.reference_scope(ReferenceScope(report=report)):
            model = self._parse_document(content)
        targets = numbered_targets(model.elements)
        if targets:
            found.clear()
            with TextParser.reference_scope(ReferenceScope(targets=targets, report=report)):
                model = self._parse_document(content)
        # An ID that numbers nothing leaves its references without a target;
        # say why first (the renderer reports the same when nothing refers to it).
        problems = [
            "Reference id " + problem
            for _, _, problem in numbered_items(model.elements)
            if problem is not None
        ]
        for message in [*problems, *found]:
            if message not in model.warnings:
                model.warnings.append(message)
                logger.warning("%s", message)
        return model

    def _parse_document(self, content: str) -> DocumentModel:
        """Parse the document once under the active reference scope."""
        lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")

        # Parse frontmatter
        metadata, remaining_lines = FrontmatterParser.parse(lines, self.profile)
        # Required-text checks precede heading inference; length waits for terms.
        with document_validation():
            validate_term_sheet_metadata(metadata, check_length=False)
        self._financial_rules = get_profile(metadata.profile).is_ib
        self._term_sheet = metadata.profile == "term-sheet"
        self._parse_warnings = []
        self._title_runs = {}
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
                # A raw line is only scanned for footnote marks here.
                with TextParser.reference_scope(None):
                    runs = TextParser.parse_runs(scan_lines[index])
                explicit_references.update(
                    run.footnote_id for run in runs if run.footnote_id is not None
                )
                content_lines.append(raw_line)
        remaining_lines = content_lines
        for text in footnotes.values():
            TextParser.check_text(text, "footnote")

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
        self._check_metadata_references(metadata)
        with TextParser.reference_scope(None):  # metadata never holds a field
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
                if metadata.profile != "plain":  # where the heading becomes the title
                    TextParser.check_text(
                        first_heading.text,
                        "document title",
                        markdown=metadata.profile == "term-sheet",
                        escaped_ats=first_heading.escaped_ats
                        if isinstance(first_heading, MarkedHeading)
                        else (),
                    )
                if first_heading.runs:
                    metadata.display_runs["title"] = [
                        replace(run)
                        for run in self._title_runs.get(id(first_heading), first_heading.runs)
                    ]
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
        escaped_ats: tuple[int, ...] = ()
        if element.element_type == ElementType.NUMBERED_HEADING:
            # Its text is already unescaped, so substitute from the source line.
            source = TextParser.tokenize_terms(element.raw_text, self._terms, tokens)
            tokenized = TextParser.cleanup_text(source)
            escaped_ats = TextParser.escaped_ats(source)
        else:
            tokenized = TextParser.tokenize_terms(heading.text, self._terms, tokens)
        if tokens:
            if isinstance(heading, MarkedHeading):
                # Its offsets are those of the text without values; the runs
                # made here carry what they protected.
                heading = Heading(
                    level=heading.level, text=heading.text, is_numbered=heading.is_numbered
                )
                element.content = heading
            heading.text = restore_terms(tokenized, tokens)
            # HeadingRenderer drops bold markers because it bolds every run; the
            # references are judged on the text with them, as in body text.
            text = tokenized.replace("**", "").strip()
            runs = TextParser.parse_judged_runs(text, tokenized, escaped_ats)
            heading.runs = split_term_runs(runs, tokens)
            if any(isinstance(run, CrossReference) for run in runs):
                with TextParser.reference_scope(None):
                    plain = TextParser.parse_runs(text)
                self._title_runs[id(heading)] = split_term_runs(plain, tokens)
        return element

    def _check_metadata_references(self, metadata: DocumentMetadata) -> None:
        """Report references in the title, subtitle and date, as written.

        These values cannot hold a field and may take term values, which are
        never read as references, so they are checked here before substitution.
        The title and subtitle are always checked; the date wherever a profile
        shows it (every profile but `plain`). A term sheet reads them as inline
        Markdown, the other profiles show them as plain text.

        Args:
            metadata: Frontmatter metadata before term substitution.
        """
        markdown = metadata.profile == "term-sheet"
        places: list[tuple[str, object]] = [
            ("document title", metadata.title),
            ("subtitle", metadata.subtitle),
        ]
        if metadata.profile != "plain":
            places.append(("frontmatter date", metadata.extra.get("date")))
        for place, value in places:
            if isinstance(value, str):
                TextParser.check_text(value, place, markdown=markdown)

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
                # The callout body is unescaped here and parsed again when
                # rendered, so where it held an escaped `@` is kept apart.
                sources: list[str] = []
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
                        sources.append(source)
                        i += 1
                    else:
                        break

                title, body = self._extract_blockquote_title(bq_lines)
                escaped_ats = self._quote_escaped_ats(sources)
                quote: Blockquote
                if quote_tokens:
                    # The runs made here carry what the offsets protected.
                    quote = Blockquote(text=restore_terms(body, quote_tokens), title=title)
                    quote.runs = split_term_runs(
                        TextParser.parse_judged_runs(body, body, escaped_ats, plain=True),
                        quote_tokens,
                    )
                elif escaped_ats:
                    quote = MarkedBlockquote(text=body, title=title, escaped_ats=escaped_ats)
                else:
                    quote = Blockquote(text=body, title=title)
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
                elements.append(
                    Element(
                        element_type=ElementType.NUMBERED_HEADING,
                        content=self._numbered_heading(line, 2),
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
            else:
                # A `{#…}` that cannot number this image keeps the line as text;
                # the renderer reports it wherever such an image is written.
                file_image = self._parse_figure_line(line)
            if file_image is not None:
                if match or isinstance(file_image, Figure):
                    # The description is a caption line that cannot hold a field.
                    TextParser.check_text(file_image.alt_text, "image description")
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
            or self._parse_figure_line(line) is not None
            or self._is_reference_header(line)
        )

    @classmethod
    def _parse_figure_line(cls, line: str) -> Figure | None:
        """Read an image line ending in a Quarto-style ID, `![설명](파일){#fig-…}`.

        Only an ID that `figure_id_problem` accepts makes a figure. Any other
        `{#…}` ending keeps today's reading of the line as text; the renderer
        reports an image followed by `{#…}` wherever it stands, so the result
        does not depend on the lines around it.

        Args:
            line: One stripped source line.

        Returns:
            The figure, or None.
        """
        attribute = cls._FIGURE_ID_RE.search(line)
        if attribute is None:
            return None
        image = cls.IMAGE_PATTERN.match(line[: attribute.start()])
        if image is None:
            return None
        ref_id = attribute.group(1)
        path = image.group(2) or image.group(3)
        if figure_id_problem(ref_id, path, image.group(1)) is not None:
            return None
        return Figure(alt_text=image.group(1), path=path, ref_id=ref_id)

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

    @staticmethod
    def _numbered_heading(line: str, level: int) -> Heading:
        """A numbered heading, whose text is unescaped here (see `MarkedHeading`)."""
        text = TextParser.cleanup_text(line)
        escaped_ats = TextParser.escaped_ats(line)
        if escaped_ats:
            return MarkedHeading(level=level, text=text, is_numbered=True, escaped_ats=escaped_ats)
        return Heading(level=level, text=text, is_numbered=True)

    def _try_parse_heading(self, line: str) -> Element | None:
        """Try to parse line as a heading"""

        # Numbered heading (**1. Title**)
        if self._financial_rules and self.NUMBERED_HEADING_PATTERN.match(line):
            return Element(
                element_type=ElementType.NUMBERED_HEADING,
                content=self._numbered_heading(line, 1),
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

    @classmethod
    def _quote_escaped_ats(cls, sources: list[str]) -> tuple[int, ...]:
        """Offsets in a callout body of each `@` its lines escape.

        As in `TextParser.escaped_ats`, the lines are unescaped and joined twice
        with an escaped `@` written as two different characters; the offsets
        where the two bodies differ are those `@`.

        Args:
            sources: The callout's lines as written (after `> `).

        Returns:
            The offsets in the body `_extract_blockquote_title` makes, in order.
        """
        if not any("\\@" in source for source in sources):
            return ()
        _, first = cls._extract_blockquote_title(
            [TextParser.cleanup_text(source, at_as="0") for source in sources]
        )
        _, second = cls._extract_blockquote_title(
            [TextParser.cleanup_text(source, at_as="1") for source in sources]
        )
        return tuple(
            index
            for index, (one, other) in enumerate(zip(first, second, strict=True))
            if one != other
        )


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
