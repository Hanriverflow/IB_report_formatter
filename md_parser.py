"""
MD Parser Module for IB Style Word Report Converter
Handles parsing of Markdown files including frontmatter, elements, tables,
LaTeX equations, Base64 images, and footnotes.

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
import os
import re
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser
from pathlib import Path
from typing import (
    BinaryIO,
    Callable,
    Dict,
    List,
    Match,
    Optional,
    Sequence,
    Set,
    Tuple,
    Union,
    cast,
)
from urllib.parse import quote, unquote, urlsplit

import yaml

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
    default_metadata,
    get_profile,
    header_labels,
    validate_term_sheet_metadata,
)
from numeric_checks import evaluate_checks, parse_checks
from term_variables import (
    TERM_REFERENCE_RE,
    TERM_TOKEN_PATTERN,
    TermResolver,
    TokenMap,
    restore_terms,
    split_term_runs,
    validate_terms,
)

logger = logging.getLogger(__name__)


class FrontmatterParser:
    """Parses YAML frontmatter from markdown"""

    _MARKDOWN_HEADING_RE = re.compile(r"^#{1,6}\s+")
    _BOLD_LABEL_RE = re.compile(r"\*\*[^*]+:\*\*")
    _SIMPLE_KEY_VALUE_RE = re.compile(r"^[A-Za-z0-9_.-]+\s*:\s*.*$")
    _DECLARED_RE = re.compile(
        r"^(profile|layout|tables|sender|charts|preset|house|terms|confirmation):", re.IGNORECASE
    )

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

        declared = any(FrontmatterParser._DECLARED_RE.match(line) for line in frontmatter_lines)
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
        if metadata.profile == "term-sheet":
            # Preserve types so the profile contract can reject malformed YAML.
            metadata.title = data.get("title", metadata.title)
            metadata.subtitle = data.get("subtitle", metadata.subtitle)
        else:
            metadata.title = str(data.get("title", metadata.title))
            metadata.subtitle = str(data.get("subtitle", metadata.subtitle))
        metadata.company = str(data.get("company", metadata.company))
        metadata.ticker = str(data.get("ticker", metadata.ticker))
        metadata.sector = str(data.get("sector", metadata.sector))
        metadata.analyst = str(data.get("analyst", metadata.analyst))

        # Store extra fields
        known_keys = {"title", "subtitle", "company", "ticker", "sector", "analyst", "profile"}
        structured = {
            "layout", "tables", "sender", "recipients", "cc", "attachments", "attendees", "letter",
            "charts", "preset", "house", "terms", "confirmation", "checks",
        }
        if metadata.profile == "term-sheet":
            structured.update({"prepared_by", "disclaimer", "confidential_label"})
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


# ═══════════════════════════════════════════════════════════════════════════════
# LOCAL LINK TARGETS
# ═══════════════════════════════════════════════════════════════════════════════

_DRIVE_PATH_RE = re.compile(r"^[A-Za-z]:[\\/]")
_URI_DRIVE_PATH_RE = re.compile(r"^/[A-Za-z]:/")
_URI_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")


def file_uri(path: str) -> str:
    """Return the `file:` URI of an absolute local path; `%` and `#` are literal.

    Args:
        path: Absolute Windows (`C:\\...`, `C:/...`), UNC (`\\\\server\\share`)
            or POSIX (`/...`) path.

    Returns:
        The percent-encoded URI, as `Path.as_uri()` writes it; a UNC server
        becomes the URI host.
    """
    normalized = path if path[:1] == "/" else path.replace("\\", "/")
    if _DRIVE_PATH_RE.match(normalized):
        return "file:///" + normalized[:2] + quote(normalized[2:], safe="/")
    if normalized[:2] == "//":
        return "file:" + quote(normalized, safe="/")
    return "file://" + quote(normalized, safe="/")


def rebase_link_target(target: str, source_dir: Path, output_dir: Path) -> str:
    """Re-express a parsed link target for a DOCX saved in `output_dir`.

    A relative target refers to the source folder, so it is rewritten to reach
    the same file from the output folder (a `file:` URI when no relative path
    exists, e.g. on another drive); its `#` fragment is kept. An absolute local
    `file:` target becomes relative, keeping its query and fragment, only when
    the file lies inside the output folder, so the author's directory is not
    exposed; otherwise it is kept and a warning says recipients may not be able
    to open it. Network shares, other URLs, in-document anchors and targets
    that are not valid local paths are returned unchanged.

    Args:
        target: Hyperlink target produced by `TextParser.link_target`.
        source_dir: Folder of the source Markdown file.
        output_dir: Folder of the saved DOCX.

    Returns:
        The target to store in the saved document.
    """
    try:
        return _rebase_link_target(target, source_dir, output_dir.resolve())
    except (OSError, ValueError) as error:  # NUL bytes, invalid UTF-8 escapes, bad hosts
        logger.warning("Link target kept as written; it is not a usable local path (%s): %s", error, target)
        return target


def _rebase_link_target(target: str, source_dir: Path, output_dir: Path) -> str:
    """`rebase_link_target` without its guard; raises on malformed targets."""
    if target[:5].lower() == "file:":
        parts = urlsplit(target)
        location = unquote(parts.path, errors="strict")
        if _URI_DRIVE_PATH_RE.match(location):
            location = location[1:]
        path = Path(location)
        if parts.netloc.lower() not in ("", "localhost") or not path.is_absolute():
            return target  # a network share, or another platform's path
        resolved = path.resolve()
        try:
            relative = resolved.relative_to(output_dir)
        except ValueError:
            logger.warning(
                "Link points to a local path outside the output folder; recipients "
                "may not be able to open it: %s", location,
            )
            return target
        query = "?" + parts.query if parts.query else ""
        fragment = "#" + parts.fragment if parts.fragment else ""
        return quote(relative.as_posix(), safe="/") + query + fragment
    if not target or target[:1] == "#" or _URI_SCHEME_RE.match(target):
        return target
    location, hash_mark, fragment = target.partition("#")
    local = (source_dir / unquote(location, errors="strict")).resolve()
    try:
        reached = Path(os.path.relpath(local, output_dir)).as_posix()
    except ValueError:  # another drive: no relative path exists
        return file_uri(str(local)) + hash_mark + fragment
    return quote(reached, safe="/") + hash_mark + fragment


class TextParser:
    """Parses inline text formatting (bold, italic, inline LaTeX, etc.)"""

    # Compiled once — used by cleanup_text
    _ESCAPE_RE = re.compile(r'\\([\\$`^~.*"\'()\[\]{}|_-])')
    _HTML_BREAK_RE = re.compile(r"(?<!\\)<br\s*/?>", re.IGNORECASE)
    _ESCAPED_HTML_BREAK_RE = re.compile(r"\\(<br\s*/?>)", re.IGNORECASE)
    _CODE_SPAN_RE = re.compile(r"(?<![\\`])(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.DOTALL)
    # Everything `_ESCAPE_RE` unescapes, plus HTML breaks: escaping these makes
    # parse_runs reproduce inserted text verbatim.
    _ESCAPABLE_RE = re.compile(r'([\\$`^~.*"\'()\[\]{}|_-])')
    _LITERAL_BREAK_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)

    # Inline formatting patterns. A term token counts as one subscript unit, so a
    # value between tildes is a subscript whatever it contains (never re-parsed).
    _SUBSCRIPT_PATTERN = r"(?<!~)~(?:[A-Za-z0-9]|" + TERM_TOKEN_PATTERN + r"){1,8}~(?!~)"
    # Emphasis delimiters must flank their text (CommonMark): an opening `*`/`_`
    # is not followed by whitespace and a closing one is not preceded by it, so
    # `2 * 3 * 4` and spaced note markers (`매출처* ...`) stay literal.
    _INLINE_FORMAT_SPLIT_RE = re.compile(
        r"(\*\*(?![\s*])[^*\n]*?[^\s*]\*\*|\^[^^\n]+?\^|"
        + _SUBSCRIPT_PATTERN
        + r"|(?<!\*)\*(?![\s*])[^*\n]*?[^\s*]\*(?!\*)|(?<!\w)_(?![\s_])[^_\n]*?[^\s_]_(?!\w))"
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
        text, images = cls._protect_images(text)

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

        runs = cls._split_image_runs(runs, images, code_spans)
        return cls._encode_link_targets(cls._split_code_runs(runs, code_spans))

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
        text, images = cls._protect_images(text)
        runs: List[TextRun] = []
        for segment_text, color_hex in cls._split_on_color_spans(text):
            runs.extend(cls._apply_color(cls._parse_inline_formatting(segment_text), color_hex))
        runs = cls._split_image_runs(runs, images, code_spans)
        return cls._encode_link_targets(cls._split_code_runs(runs, code_spans))

    @classmethod
    def _protect_images(cls, text: str) -> Tuple[str, Dict[str, Tuple[str, str, str]]]:
        """Shield inline images before colour spans, math, links and emphasis see them.

        An image is one unit, so `$` in its path is not math and emphasis around
        it still pairs. An image-like text inside a link's destination belongs
        to that link; an image in a link's label stays an image. A `!` after an
        odd number of backslashes is escaped.

        Args:
            text: Text whose code spans are already protected.

        Returns:
            Text with images replaced by tokens, and token -> (source, alt, destination).
        """
        images: Dict[str, Tuple[str, str, str]] = {}
        if "![" not in text:
            return text, images
        prefix = "IMG"
        while prefix in text:
            prefix += "X"
        destinations = [
            match.span(2) for match in cls._INLINE_REFERENCE_RE.finditer(text)
            if match.group(2) and not cls._escaped(text, match.start())
        ]
        pieces: List[str] = []
        last = 0
        for match in cls._INLINE_IMAGE_RE.finditer(text):
            if cls._escaped(text, match.start()) or any(
                start <= match.start() < end for start, end in destinations
            ):
                continue
            token = f"{prefix}{len(images)}"
            images[token] = (match.group(0), match.group(1), match.group(2) or match.group(3))
            pieces.extend((text[last:match.start()], token))
            last = match.end()
        pieces.append(text[last:])
        return "".join(pieces), images

    @staticmethod
    def _escaped(text: str, index: int) -> bool:
        """Whether the character at `index` follows an odd number of backslashes."""
        count = 0
        while index - count > 0 and text[index - count - 1] == "\\":
            count += 1
        return count % 2 == 1

    @classmethod
    def _split_image_runs(
        cls, runs: List[TextRun], images: Dict[str, Tuple[str, str, str]], code_spans: Dict[str, str],
    ) -> List[TextRun]:
        """Turn image tokens into image runs that keep the surrounding formatting.

        Math stays literal, so a token inside an equation gets its source back.

        Args:
            runs: Runs parsed from text whose images were replaced by tokens.
            images: Token to (source, alt text, destination) from `_protect_images`.
            code_spans: Code-span tokens, restored inside image fields.

        Returns:
            Runs with each image as its own run whose text is empty.
        """
        if not images:
            return runs

        def restore(value: str) -> str:
            for token, literal in code_spans.items():
                value = value.replace(token, literal)
            return value

        tokens = re.compile("|".join(re.escape(token) for token in images))
        result: List[TextRun] = []
        for run in runs:
            if run.is_latex:
                run.text = tokens.sub(lambda match: images[match.group(0)][0], run.text)
            if run.is_latex or not tokens.search(run.text):
                result.append(run)
                continue
            offset = 0
            for match in tokens.finditer(run.text):
                if match.start() > offset:
                    result.append(replace(run, text=run.text[offset:match.start()]))
                _, alt, destination = images[match.group(0)]
                image = Image(alt_text=cls.cleanup_text(restore(alt)), path=restore(destination))
                result.append(replace(run, text="", image=image))
                offset = match.end()
            if offset < len(run.text):
                result.append(replace(run, text=run.text[offset:]))
        return result

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

    @classmethod
    def _split_code_runs(cls, runs: List[TextRun], literals: Dict[str, str]) -> List[TextRun]:
        """Turn protected code spans into literal code runs without their backticks.

        The code text is never parsed; the run keeps the surrounding emphasis,
        colour and link. Link destinations and equations keep the source text.

        Args:
            runs: Runs parsed from text whose code spans were replaced by tokens.
            literals: Token to original code span (backticks included).

        Returns:
            Runs with each code span as its own run flagged `code`.
        """
        if not literals:
            return runs
        tokens = re.compile("|".join(re.escape(token) for token in literals))
        result: List[TextRun] = []
        for run in runs:
            for token, literal in literals.items():
                if run.hyperlink:
                    run.hyperlink = run.hyperlink.replace(token, literal)
                if run.image is not None:
                    run.image.path = run.image.path.replace(token, literal)
                    run.image.alt_text = run.image.alt_text.replace(token, literal)
            if run.is_latex or not tokens.search(run.text):
                if run.is_latex:
                    for token, literal in literals.items():
                        run.text = run.text.replace(token, literal)
                result.append(run)
                continue
            offset = 0
            for match in tokens.finditer(run.text):
                if match.start() > offset:
                    result.append(replace(run, text=run.text[offset:match.start()]))
                content = cls._code_span_content(literals[match.group(0)])
                result.append(replace(run, text=content, code=True))
                offset = match.end()
            if offset < len(run.text):
                result.append(replace(run, text=run.text[offset:]))
        return result

    @staticmethod
    def _code_span_content(literal: str) -> str:
        """Return a code span's text without its backtick fence (CommonMark rules)."""
        fence = len(literal) - len(literal.lstrip("`"))
        content = literal[fence:len(literal) - fence].replace("\n", " ")
        if len(content) >= 2 and content[0] == " " and content[-1] == " " and content.strip():
            content = content[1:-1]
        return content

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

    _LINK_DESTINATION = (
        r"<[^<>\n]+>"
        r"|https?://[^\s)]+|mailto:[^\s)]+"
        r"|(?:[A-Za-z]:[\\/]|\.{1,2}[\\/])[^\s()<>]*"
        r"|[^\s()<>]+\.(?i:md|markdown|docx?|xlsx?|xlsm|pptx?|pdf|hwpx?|txt|csv|png|jpe?g|gif|svg)"
        r"(?:#[^\s()<>]*)?"
    )
    _INLINE_REFERENCE_RE = re.compile(
        r"(?<!!)\[([^\]\n]+)\]\((" + _LINK_DESTINATION + r")\)|\[\^(\d+)\]"
    )
    # An image inside text or a table cell; a bare destination may hold one level
    # of balanced parentheses (`images/(2026)/a.png`), as converters write them.
    _INLINE_IMAGE_RE = re.compile(
        r"!\[((?:\\.|[^\]\\\n])*)\]\(\s*(?:<([^<>\n]+)>|((?:[^\s()<>]|\([^\s()<>]*\))+))"
        r"(?:\s+(?:\"[^\"\n]*\"|'[^'\n]*'))?\s*\)"
    )
    @classmethod
    def link_target(cls, destination: str) -> str:
        """Return a Word relationship target for a Markdown link destination.

        URLs are kept as written. An absolute local path (`C:\\...`, UNC or `/...`)
        is a file name, so it becomes a `file:///` URI in which `%` and `#` are
        literal characters. A relative path is a URL reference to the source
        folder: existing `%` escapes and a `#` fragment are kept. Spaces and
        non-ASCII characters are percent-encoded.

        Args:
            destination: Destination as written, optionally in angle brackets.

        Returns:
            The hyperlink target.
        """
        target = destination[1:-1] if destination[:1] == "<" and destination[-1:] == ">" else destination
        if _DRIVE_PATH_RE.match(target) or target[:1] in ("/", "\\"):
            return file_uri(target)
        if _URI_SCHEME_RE.match(target):
            return target
        return quote(target.replace("\\", "/"), safe="/#%")

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
    def _encode_link_targets(cls, runs: List[TextRun]) -> List[TextRun]:
        """Encode destinations once every protected literal is back in place."""
        for run in runs:
            if run.hyperlink:
                run.hyperlink = cls.link_target(run.hyperlink)
        return runs

    @classmethod
    def _parse_plain_formatting(cls, text: str) -> List[TextRun]:
        """Parse bold, italic, superscript, and subscript runs from plain text."""
        runs: List[TextRun] = []
        parts = cls._INLINE_FORMAT_SPLIT_RE.split(text)

        for part in parts:
            if not part:
                continue

            if cls.flanked(part, "**"):
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

            if cls.flanked(part, "*") or cls.flanked(part, "_"):
                content = cls.cleanup_text(part[1:-1])
                if content:
                    runs.append(TextRun(text=content, italic=True))
                continue

            cleaned = cls.cleanup_text_preserve_spacing(part)
            if cleaned:
                runs.append(TextRun(text=cleaned))

        return runs

    @staticmethod
    def flanked(part: str, delimiter: str) -> bool:
        """Whether `part` is text wrapped in `delimiter` that flanks it on both sides.

        Args:
            part: Candidate emphasis span, delimiters included.
            delimiter: `**`, `*` or `_`.

        Returns:
            True when the span opens and closes with the delimiter, has content,
            and neither delimiter touches whitespace on its inner side.
        """
        size = len(delimiter)
        return (
            len(part) > 2 * size
            and part.startswith(delimiter)
            and part.endswith(delimiter)
            and not part[size].isspace()
            and not part[-size - 1].isspace()
            and (size == 2 or part[1] != delimiter)
        )

    @classmethod
    def has_inline_latex(cls, text: str) -> bool:
        """Check if text contains inline LaTeX expressions"""
        text, _ = cls._protect_code_spans(text)
        text, _ = cls._protect_images(text)
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

    # ── Term references ─────────────────────────────────────────────────────

    @classmethod
    def escape_literal(cls, text: str) -> str:
        """Escape text so that parse_runs reproduces it verbatim.

        Args:
            text: Literal text inserted into a field parsed as inline Markdown.

        Returns:
            Text whose inline syntax, including HTML breaks, is escaped.
        """
        escaped = cls._ESCAPABLE_RE.sub(r"\\\1", text)
        return cls._LITERAL_BREAK_RE.sub(lambda match: "\\" + match.group(0), escaped)

    @classmethod
    def tokenize_terms(
        cls, text: str, resolver: TermResolver, tokens: TokenMap, latex: bool = True,
    ) -> str:
        """Replace substitutable `{{key}}` references with private-use tokens.

        The stages mirror parse_runs: references in code spans, inline math,
        escaped braces (`\\{{`) and link destinations stay literal. A token then
        passes through inline parsing like ordinary text, so its value inherits
        the surrounding emphasis, colour and link; `split_term_runs` inserts the
        value afterwards, so the value itself is never parsed.

        Args:
            text: Markdown source of one field.
            resolver: Document term values and diagnostics.
            tokens: Receives each new token of this field.
            latex: Whether the field is parsed with inline math (parse_runs).

        Returns:
            The source with each defined reference replaced by its token; all
            other text, including undefined references, is unchanged.
        """
        if "{{" not in text:
            return text
        resolver.reserve(text)
        protected, code_spans = cls._protect_code_spans(text)
        # An image's alt text and destination are never substituted, like a link destination.
        protected, images = cls._protect_images(protected)
        # Inline math is split before links are parsed, so a `$` in a URL could
        # otherwise expose part of the destination as text.
        protected, destinations = cls._protect_link_destinations(protected)
        pieces: List[str] = []
        last = 0
        for match in cls._COLOR_SPAN_RE.finditer(protected):
            if cls._extract_color_from_style(match.group(2)) is None:
                continue
            pieces.append(cls._tokenize_segment(
                protected[last:match.start()], code_spans, resolver, tokens, latex,
            ))
            pieces.append(protected[match.start():match.start(3)])
            pieces.append(cls._tokenize_segment(match.group(3), code_spans, resolver, tokens, latex))
            pieces.append(protected[match.end(3):match.end()])
            last = match.end()
        pieces.append(cls._tokenize_segment(protected[last:], code_spans, resolver, tokens, latex))
        result = "".join(pieces)
        for token, literal in destinations.items():
            result = result.replace(token, literal)
        for token, (source, _, _) in images.items():
            result = result.replace(token, source)
        for token, literal in code_spans.items():
            result = result.replace(token, literal)
        return result

    @classmethod
    def _protect_link_destinations(cls, text: str) -> Tuple[str, Dict[str, str]]:
        """Shield each complete inline-link destination; labels stay visible.

        Args:
            text: Text whose code spans are already protected.

        Returns:
            Text with destinations replaced by tokens, and token -> destination.
        """
        literals: Dict[str, str] = {}
        prefix = "\ue000URL"
        while prefix in text:
            prefix += "X"

        def replace(match: Match[str]) -> str:
            index = match.start()
            while index and text[index - 1] == "\\":
                index -= 1
            if not match.group(2) or (match.start() - index) % 2:
                return match.group(0)  # a footnote reference or an escaped bracket
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(2)
            return f"[{match.group(1)}]({token})"

        return cls._INLINE_REFERENCE_RE.sub(replace, text), literals

    @classmethod
    def _tokenize_segment(
        cls, text: str, code_spans: Dict[str, str], resolver: TermResolver,
        tokens: TokenMap, latex: bool,
    ) -> str:
        """Tokenize one colour segment, keeping inline math verbatim."""
        pieces: List[str] = []
        last = 0
        if latex:
            for match in cls._INLINE_LATEX_RE.finditer(text):
                pieces.append(cls._tokenize_inline(text[last:match.start()], code_spans, resolver, tokens))
                pieces.append(match.group(0))
                last = match.end()
        pieces.append(cls._tokenize_inline(text[last:], code_spans, resolver, tokens))
        return "".join(pieces)

    @classmethod
    def _tokenize_inline(
        cls, text: str, code_spans: Dict[str, str], resolver: TermResolver, tokens: TokenMap,
    ) -> str:
        """Tokenize text and link labels, keeping escapes and destinations verbatim."""
        protected, escapes = cls._protect_escapes(text)
        pieces: List[str] = []
        offset = 0
        for match in cls._INLINE_REFERENCE_RE.finditer(protected):
            pieces.append(cls._tokenize_references(
                protected[offset:match.start()], code_spans, escapes, resolver, tokens,
            ))
            if match.group(3):
                pieces.append(match.group(0))
            else:
                label = cls._tokenize_references(match.group(1), code_spans, escapes, resolver, tokens)
                pieces.append(f"[{label}]({match.group(2)})")
            offset = match.end()
        pieces.append(cls._tokenize_references(protected[offset:], code_spans, escapes, resolver, tokens))
        result = "".join(pieces)
        for token, literal in escapes.items():
            result = result.replace(token, "\\" + literal)
        return result

    @staticmethod
    def _tokenize_references(
        text: str, code_spans: Dict[str, str], escapes: Dict[str, str],
        resolver: TermResolver, tokens: TokenMap,
    ) -> str:
        """Replace references in literal-free text; escapes inside a key are resolved."""

        def replace_reference(match: Match[str]) -> str:
            written = key = match.group(1)
            for token, literal in escapes.items():
                written = written.replace(token, "\\" + literal)
                key = key.replace(token, literal)
            for token, literal in code_spans.items():
                written = written.replace(token, literal)
                key = key.replace(token, literal)
            replacement = resolver.token("{{" + written + "}}", key.strip(), tokens)
            return replacement if replacement is not None else match.group(0)

        return TERM_REFERENCE_RE.sub(replace_reference, text)


def _parse_term_runs(
    text: str,
    parse: Callable[[str], List[TextRun]],
    terms: Optional[TermResolver],
    latex: bool = True,
) -> Tuple[str, List[TextRun], bool]:
    """Parse runs, substituting defined `{{key}}` references when terms are active.

    Args:
        text: Markdown source of one field.
        parse: The TextParser function the renderer-facing runs come from.
        terms: Active term resolver, or None when the document has no `terms:`.
        latex: Whether `parse` recognizes inline math.

    Returns:
        Substituted source text, runs with value runs split out, and whether a
        value was inserted. Text and runs are unchanged without a substitution.
    """
    if terms is None:
        return text, parse(text), False
    tokens: TokenMap = {}
    tokenized = TextParser.tokenize_terms(text, terms, tokens, latex=latex)
    if not tokens:
        return text, parse(text), False
    return restore_terms(tokenized, tokens), split_term_runs(parse(tokenized), tokens), True


def _substitute_text(text: str, terms: TermResolver) -> str:
    """Return a table cell's source with defined references replaced, as if typed.

    Args:
        text: Raw cell source.
        terms: Active term resolver.

    Returns:
        The source with values inserted where the cell's runs show them.
    """
    tokens: TokenMap = {}
    latex = TextParser.has_inline_latex(text)
    return restore_terms(TextParser.tokenize_terms(text, terms, tokens, latex=latex), tokens)


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
    def parse(
        lines: List[str], financial_rules: bool = True, terms: Optional[TermResolver] = None,
    ) -> Table:
        """
        Parse markdown table lines into a Table object.

        Args:
            lines: Lines that make up the table (starting with |)
            financial_rules: Whether IB table semantics are inferred.
            terms: Active term resolver; cell runs receive values, content stays raw.
        """
        # Only the second line can be the structural delimiter row.
        data_lines = [
            line
            for index, line in enumerate(lines)
            if not (index == 1 and TableParser._is_delimiter_row(line))
        ]
        return TableParser.from_cells(
            [TableParser._split_row(line) for line in data_lines],
            TableParser._parse_alignments(lines),
            financial_rules=financial_rules,
            terms=terms,
        )

    @staticmethod
    def from_cells(
        rows: Sequence[Sequence[Union[str, TableCell]]],
        alignments: List[str],
        financial_rules: bool = True,
        terms: Optional[TermResolver] = None,
        header_rows: int = 1,
    ) -> Table:
        """
        Build a Table from cell sources; the leading `header_rows` rows are the header.

        Args:
            rows: Each cell as inline Markdown (span markers allowed) or as a
                prepared cell whose runs are final (HTML tables).
            alignments: Column alignments; ignored unless one per column.
            financial_rules: Whether IB table semantics are inferred.
            terms: Active term resolver; cell runs receive values, content stays raw.
            header_rows: Number of header rows (at least one).
        """
        table = Table()
        if not rows:
            return table
        table.header_rows = max(1, min(header_rows, len(rows)))

        # Get column count from first row
        first_row_cells = rows[0]
        table.col_count = len(first_row_cells)
        table.alignments = (
            alignments if len(alignments) == table.col_count else ["left"] * table.col_count
        )

        # Header semantics follow the values the reader sees across every header
        # row, merged labels included; cells keep raw content.
        header_cells = header_labels([
            [
                "".join(run.text for run in source.runs) if isinstance(source, TableCell)
                else source if terms is None else _substitute_text(source, terms)
                for source in header_row
            ]
            for header_row in rows[:table.header_rows]
        ])

        # Detect table type
        header_text = " ".join(header_cells).lower()
        table.table_type = (
            TableParser._detect_type(header_text) if financial_rules else TableType.GENERIC
        )

        # Parse all rows — normalise column count per row
        for i, row_cells in enumerate(rows):
            cells = list(row_cells)
            is_header = i < table.header_rows

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
            for j, cell_source in enumerate(cells):
                if isinstance(cell_source, TableCell):
                    cell = replace(cell_source, is_header=is_header)
                    TableParser._detect_flags(
                        cell, "".join(run.text for run in cell.runs), j, table.table_type, header_cells,
                    )
                else:
                    cell = TableParser._parse_cell(
                        cell_source,
                        is_header=is_header,
                        col_idx=j,
                        row_idx=i,
                        total_rows=len(rows),
                        table_type=table.table_type,
                        header_cells=header_cells,
                        terms=terms,
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
        terms: Optional[TermResolver] = None,
    ) -> TableCell:
        """Parse a single table cell.

        The raw `content` is kept for span markers and width estimation; runs
        and content-derived flags see substituted term values as typed text.
        """
        cell = TableCell(content=text, is_header=is_header)

        # Prefer LaTeX-aware parsing only when a balanced inline expression exists.
        latex = TextParser.has_inline_latex(text)
        parse = TextParser.parse_runs if latex else TextParser.parse_runs_plain
        detected_text, cell.runs, _ = _parse_term_runs(text, parse, terms, latex=latex)
        TableParser._detect_flags(cell, detected_text, col_idx, table_type, header_cells)
        return cell

    @staticmethod
    def _detect_flags(
        cell: TableCell, detected_text: str, col_idx: int, table_type: TableType, header_cells: List[str],
    ) -> None:
        """Set a cell's numeric, negative and risk flags from its text.

        Args:
            cell: Cell with final runs.
            detected_text: Text the numeric and risk detection reads.
            col_idx: Zero-based column.
            table_type: Detected or configured table type.
            header_cells: Header texts, for risk columns.
        """
        # Detect numeric content
        cell.is_numeric = any(char.isdigit() for char in detected_text) and col_idx > 0

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
                text_lower = detected_text.lower()
                if "high" in text_lower or "높" in detected_text:
                    cell.risk_level = "high"
                elif any(kw in text_lower for kw in ("medium", "moderate")) or "중" in detected_text:
                    cell.risk_level = "medium"
                elif "low" in text_lower or "낮" in detected_text:
                    cell.risk_level = "low"


class TableSpanResolver:
    """Resolve opt-in span markers while retaining invalid groups literally."""

    @staticmethod
    def resolve(table: Table, profile: str = "ib-report") -> None:
        """Validate each connected marker group and infer leading label columns.

        References point only above or left, so row-major traversal resolves a
        cell's group from an already visited target without recursive walks.
        Invalid groups never receive merge directives; neighboring groups remain
        independent even when their bounding rectangles overlap.

        Args:
            table: A parsed table after ordered specifications have been applied.
            profile: Selected document profile, including its default span policy.
        """
        active = table.spans if table.spans is not None else profile == "term-sheet"
        first_header_width = 1
        if active:
            markers: Dict[Tuple[int, int], str] = {}
            roots: Dict[Tuple[int, int], Tuple[int, int]] = {}
            groups: Dict[Tuple[int, int], List[Tuple[int, int]]] = {}
            outside: Set[Tuple[int, int]] = set()

            for row_index, row in enumerate(table.rows):
                for column_index, cell in enumerate(row.cells):
                    coordinate = (row_index, column_index)
                    raw = "" if cell.literal else cell.content.strip()
                    root = coordinate
                    if raw in {r"\^^", r"\<<"}:
                        # Classify raw source before unescaping either representation.
                        cell.content = raw[1:]
                        cell.runs = [TextRun(text=raw[1:])]
                    elif raw in {"^^", "<<"}:
                        markers[coordinate] = "up" if raw == "^^" else "left"
                        target = (
                            (row_index - 1, column_index)
                            if raw == "^^"
                            else (row_index, column_index - 1)
                        )
                        if target in roots:
                            root = roots[target]
                        else:
                            outside.add(coordinate)
                    roots[coordinate] = root
                    groups.setdefault(root, []).append(coordinate)

            for coordinates in groups.values():
                marker_cells = [position for position in coordinates if position in markers]
                if not marker_cells:
                    continue
                top = min(row for row, _ in coordinates)
                bottom = max(row for row, _ in coordinates)
                left = min(column for _, column in coordinates)
                right = max(column for _, column in coordinates)
                anchors = [position for position in coordinates if position not in markers]
                error = ""
                if any(position in outside for position in coordinates):
                    error = "a marker points outside the table"
                elif len({table.rows[row].is_header for row, _ in coordinates}) > 1:
                    error = "a merge crosses the header/body boundary"
                elif anchors != [(top, left)]:
                    error = "a merge requires one top-left anchor"
                elif len(coordinates) != (bottom - top + 1) * (right - left + 1):
                    error = "a merge must fill one complete rectangle"
                if error:
                    warning = (
                        f"Invalid table span group at row {top + 1}, column {left + 1}: "
                        f"{error}; markers kept literally"
                    )
                    table.warnings.append(warning)
                    logger.warning("%s", warning)
                    continue
                for marker_row, marker_column in marker_cells:
                    table.rows[marker_row].cells[marker_column].merge = markers[
                        (marker_row, marker_column)
                    ]
                if (top, left) == (0, 0) and table.rows[0].is_header:
                    first_header_width = right + 1

        if table.label_columns is None and (active or profile == "term-sheet"):
            table.label_columns = min(first_header_width, max(table.col_count - 1, 0))


# ═══════════════════════════════════════════════════════════════════════════════
# HTML TABLES (converter output)
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class _HtmlCell:
    """One `<td>`/`<th>` of the outermost table while it is being read."""

    colspan: int = 1
    rowspan: int = 1
    runs: List[TextRun] = field(default_factory=list)


def _html_image(source: str, alt: str) -> Optional[Image]:
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
        return Image(alt_text=alt, path="", base64_data=payload.strip(), mime_type=header[5:-7] or "image/png")
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
        "b": "bold", "strong": "bold", "i": "italic", "em": "italic",
        "code": "code", "tt": "code", "kbd": "code", "sup": "superscript", "sub": "subscript",
    }
    _BLOCK_TAGS = frozenset({
        "p", "div", "li", "ul", "ol", "tr", "blockquote", "pre", "h1", "h2", "h3", "h4", "h5", "h6",
    })
    _MAX_SPAN = 63  # Word's column limit bounds any sensible span

    def __init__(self, terms: Optional[TermResolver] = None) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: List[List[_HtmlCell]] = []
        self.stray: List[str] = []
        self.extra_tables = 0  # further top-level tables in the same block (not read)
        self._terms = terms
        self._depth = 0
        self._cell: Optional[_HtmlCell] = None
        self._styles: Dict[str, int] = {}
        self._links: List[str] = []
        self._colors: List[Optional[str]] = []
        self._line_break = False  # a block boundary: the next content starts a new line
        self._separator = False  # a nested cell boundary: the next content is spaced

    @classmethod
    def _span(cls, value: Optional[str]) -> int:
        try:
            return max(1, min(int(str(value).strip()), cls._MAX_SPAN))
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

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
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
            self._cell = _HtmlCell(self._span(attributes.get("colspan")), self._span(attributes.get("rowspan")))
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
        if not text.strip() and (self._line_break or self._separator or not self._line_has_content()):
            return  # layout whitespace between blocks
        self._add(self._formatted(text))

    def finish(self, runs: List[TextRun]) -> List[TextRun]:
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
        merged: List[TextRun] = []
        for run in runs:
            previous = merged[-1] if merged else None
            if (
                previous is not None and previous.image is None and run.image is None
                and "\n" not in (previous.text, run.text)
                and replace(previous, text="") == replace(run, text="")
            ):
                tail = run.text.lstrip(" ") if previous.text.endswith(" ") else run.text
                merged[-1] = replace(previous, text=previous.text + tail)
            else:
                merged.append(run)
        substituted: List[TextRun] = []
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

        cleaned: List[TextRun] = []
        for run in substituted:
            text = run.text
            if not fixed(run):
                previous = cleaned[-1] if cleaned else None
                if previous is None or previous.text == "\n" or (
                    previous.image is None and previous.text.endswith(" ")
                ):
                    text = text.lstrip(" ")
            if text or run.image is not None or run.term_key is not None:
                cleaned.append(replace(run, text=text))
        for index, run in enumerate(cleaned):
            following = cleaned[index + 1] if index + 1 < len(cleaned) else None
            if (following is None or following.text == "\n") and not fixed(run):
                cleaned[index] = replace(run, text=run.text.rstrip(" "))
        cleaned = [run for run in cleaned if run.text or run.image is not None or run.term_key is not None]
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
    def block_end(cls, lines: List[str], start: int) -> Optional[int]:
        """Return the index after the line closing the table opened at `start`.

        Args:
            lines: Document lines.
            start: Index of the line that opens the table.

        Returns:
            The end index, or None when the table is never closed.
        """
        depth = 0
        for index in range(start, len(lines)):
            depth += len(cls._OPEN_RE.findall(lines[index])) - len(cls._CLOSE_RE.findall(lines[index]))
            if depth <= 0:
                return index + 1
        return None

    @staticmethod
    def _table_cell(runs: List[TextRun]) -> TableCell:
        """A prepared literal cell; `content` is its text with breaks as `<br>`."""
        content = "".join(run.text for run in runs).replace("\n", "<br>")
        return TableCell(content=content, runs=runs, literal=True)

    @classmethod
    def parse(
        cls, html: str, terms: Optional[TermResolver] = None,
    ) -> Tuple[List[List[Union[str, TableCell]]], int, List[str]]:
        """Convert a complete table block into rows of cells.

        Args:
            html: The `<table>...</table>` source.
            terms: Active term resolver, or None when the document has no `terms:`.

        Returns:
            The rows (covered positions hold span markers), how many leading
            rows form the header, and warnings about content that could not be
            placed.
        """
        reader = _HtmlTableReader(terms)
        reader.feed(html)
        reader.close()
        warnings = [f"Text outside the HTML table cells was dropped: {text}" for text in reader.stray]
        if reader.extra_tables:
            warnings.append(
                f"{reader.extra_tables} further HTML table(s) on the closing line were dropped; "
                "start each table on its own line"
            )

        runs: Dict[Tuple[int, int], List[TextRun]] = {}
        anchors: Dict[Tuple[int, int], Tuple[int, int]] = {}
        spans: Dict[Tuple[int, int], int] = {}
        rows = reader.rows
        for row_index, html_row in enumerate(rows):
            position = 0
            for cell in html_row:
                while (row_index, position) in anchors:
                    position += 1
                rowspan = min(cell.rowspan, len(rows) - row_index)
                spans[(row_index, position)] = rowspan
                runs[(row_index, position)] = reader.finish(cell.runs)
                for down in range(rowspan):
                    for right in range(cell.colspan):
                        anchors[(row_index + down, position + right)] = (row_index, position)
                position += cell.colspan
        if not anchors:
            return [], 0, warnings
        col_count = max(covered for _, covered in anchors) + 1
        kept = [row for row in range(len(rows)) if any((row, column) in anchors for column in range(col_count))]

        header_end = kept[0] + 1
        grown = True
        while grown:
            grown = False
            for (top, _), rowspan in spans.items():
                if top < header_end < top + rowspan:
                    header_end, grown = top + rowspan, True

        def source(row: int, column: int) -> Union[str, TableCell]:
            anchor = anchors.get((row, column))
            if anchor is None:
                return ""
            if anchor == (row, column):
                return cls._table_cell(runs[anchor])
            return "^^" if column == anchor[1] else "<<"

        cells = [[source(row, column) for column in range(col_count)] for row in kept]
        return cells, sum(1 for row in kept if row < header_end), warnings

    @staticmethod
    def image(line: str) -> Optional[Image]:
        """Return the image of a line that is a single HTML `<img>` tag.

        Args:
            line: Stripped document line.

        Returns:
            The image, or None when the line is not one `<img>` with a usable `src`.
        """
        found: List[Dict[str, str]] = []

        class _Reader(HTMLParser):
            def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
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
        self,
        preserve_trailing_double_space_break: bool = False,
        profile: Optional[str] = None,
        base_dir: Optional[Path] = None,
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
        # Folder of the source file: relative link paths are resolved against it
        # when the document is saved (see `rebase_local_links`).
        self.base_dir = base_dir
        self._financial_rules = True
        self._term_sheet = False
        self._parse_warnings: List[str] = []
        # Per-parse `terms:` state; None keeps `{{...}}` ordinary text.
        self._terms: Optional[TermResolver] = None

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
        # Required-text checks precede heading inference; length waits for terms.
        validate_term_sheet_metadata(metadata, check_length=False)
        self._financial_rules = get_profile(metadata.profile).is_ib
        self._term_sheet = metadata.profile == "term-sheet"
        self._parse_warnings = []
        # An empty mapping enables checking: every reference is then undefined.
        self._terms = (
            TermResolver(validate_terms(metadata.extra["terms"]), corpus=content)
            if "terms" in metadata.extra else None
        )

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
        validate_term_sheet_metadata(metadata)
        model = DocumentModel(
            metadata=metadata,
            elements=elements,
            footnotes=footnotes,
            warnings=input_warnings,
            parsed_profile=metadata.profile,
            source_dir=self.base_dir,
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
            if written_title in {"", "Document", "IB Report"} and first_heading and first_heading.text:
                metadata.title = first_heading.text
                if first_heading.runs:
                    metadata.display_runs["title"] = [replace(run) for run in first_heading.runs]
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
        self, text: str, parse: Callable[[str], List[TextRun]], latex: bool = True,
    ) -> Tuple[str, List[TextRun], bool]:
        """Parse one field's runs with this document's terms (see `_parse_term_runs`)."""
        return _parse_term_runs(text, parse, self._terms, latex=latex)

    def _substitute_heading(self, element: Element) -> Element:
        """Carry value runs for a heading that references terms.

        The runs follow HeadingRenderer's parse of the heading text but resolve
        escapes once: its second unescape pass would turn `\\{{` into a reference.

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
                TextParser.parse_runs(tokenized.replace("**", "").strip()), tokens,
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

    def _substitute_table_text(self, elements: List[Element]) -> None:
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
                tokenized = TextParser.tokenize_terms(getattr(element.content, name), self._terms, tokens)
                if tokens:
                    setattr(
                        element.content, name,
                        restore_terms(tokenized, tokens, TextParser.escape_literal),
                    )

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

    @staticmethod
    def _wrap_destination(destination: str) -> str:
        """Angle-bracket a definition's destination so any path form stays a link."""
        if any(character in destination for character in "<>\n"):
            return destination
        return "<" + destination + ">"

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
            return f"[{match.group(1)}]({cls._wrap_destination(destination)})"

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

            # ── HTML table (converter output; spans become span markers) ────
            if HtmlTableParser.START_RE.match(line):
                end = HtmlTableParser.block_end(lines, i)
                if end is not None:
                    source = "\n".join(lines[i:end])
                    cells, header_rows, html_warnings = HtmlTableParser.parse(source, self._terms)
                    self._parse_warnings.extend(html_warnings)
                    if cells:
                        table = TableParser.from_cells(
                            cells, [], financial_rules=self._financial_rules, terms=self._terms,
                            header_rows=header_rows,
                        )
                        table.spans = True
                        elements.append(Element(element_type=ElementType.TABLE, content=table, raw_text=source))
                    else:
                        self._parse_warnings.append(f"HTML table at line {i + 1} has no cells; it was dropped")
                    i = end
                    continue
                self._parse_warnings.append(f"Unclosed HTML table at line {i + 1} was kept as text")

            # ── Table (collect all contiguous table lines) ──────────────────
            if self.TABLE_START_PATTERN.match(line):
                table_lines: List[str] = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    table_lines.append(lines[i].strip())
                    i += 1
                table = TableParser.parse(
                    table_lines, financial_rules=self._financial_rules, terms=self._terms,
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
                    elements.append(self._substitute_heading(Element(
                        element_type=ElementType.HEADING_1,
                        content=Heading(level=1, text=line),
                        raw_text=raw_line + "\n" + lines[i + 1],
                    )))
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
                bq_lines: List[str] = []
                quote_tokens: TokenMap = {}
                while i < len(lines):
                    bq_match = self.BLOCKQUOTE_PATTERN.match(lines[i].strip())
                    if bq_match:
                        source = bq_match.group(1)
                        if self._terms is not None:
                            # Callouts render their body without inline math.
                            source = TextParser.tokenize_terms(
                                source, self._terms, quote_tokens, latex=False,
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
            file_image: Optional[Image] = None
            if match:
                file_image = Image(alt_text=match.group(1), path=match.group(2) or match.group(3))
            elif self._HTML_IMAGE_LINE_RE.match(line):
                file_image = HtmlTableParser.image(line)
            if file_image is not None:
                elements.append(Element(element_type=ElementType.IMAGE, content=file_image, raw_text=line))
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
            or HtmlTableParser.START_RE.match(line)
            or self._HTML_IMAGE_LINE_RE.match(line)
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

    base_dir = None if is_stream(source) else Path(str(source)).resolve().parent
    parser = MarkdownParser(profile=profile, base_dir=base_dir)
    model = parser.parse(content)
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

    if not is_stream(source):
        source_dir = Path(str(source)).resolve().parent
        for element in model.elements:
            if isinstance(element.content, Image):
                image = element.content
                if (
                    image.path and not image.base64_data
                    and not re.match(r"^(?:[a-zA-Z][a-zA-Z0-9+.-]*://|data:)", image.path, re.IGNORECASE)
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


def _infer_metadata_from_elements(
    model: DocumentModel,
    allow_company_inference: bool = True,
) -> None:
    """Backfill metadata from document content when frontmatter is absent or partial."""
    metadata = model.metadata

    # Display runs exist only for a title written with a term reference; such a
    # title is explicit even when its value equals the default.
    if metadata.title == "IB Report" and "title" not in metadata.display_runs:
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
