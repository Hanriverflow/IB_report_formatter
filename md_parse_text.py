"""Inline Markdown parsing, local link targets and term-variable substitution in text.

Changelog (captions):
    - `@tbl-…` / `@fig-…` in text become `CrossReference` runs; code, math, link
      destinations, escaped `\\@` and e-mail-like `a@tbl-x` stay literal.
    - Review fixes: one recognizer (`_reference_candidates`) for Markdown and
      HTML runs and for text that cannot hold a field (`check_text`); the `@`
      must follow whitespace, an opening bracket or quote, or the start of the
      text, the ID is read whole, a domain-like `@tbl-x.example` is not a
      reference.
    - Second review: the recognizer reads the visible text of all runs at once
      and splits a run only for a reference that resolves in the active
      `ReferenceScope`; every other reference is reported without touching the
      runs, so text without a resolvable reference parses exactly as before.
"""

import logging
import os
import re
from bisect import bisect_right
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, replace
from pathlib import Path
from re import Match
from urllib.parse import quote, unquote, urlsplit

from document_model import (
    TEXT_RUN_FIELDS,
    CrossReference,
    Image,
    TextRun,
    reference_id_problem,
)
from input_safety import StrictValidationError
from term_variables import (
    TERM_REFERENCE_RE,
    TERM_TOKEN_PATTERN,
    TermResolver,
    TokenMap,
    restore_terms,
    split_term_runs,
)

logger = logging.getLogger("md_parser")


# ═══════════════════════════════════════════════════════════════════════════════
# CROSS-REFERENCE SCOPE
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class ReferenceScope:
    """What `@tbl-…` / `@fig-…` text may become while text is parsed in a context.

    Outside any scope no text is read as a reference: the parser behaves as it
    did before cross-references existed.

    Attributes:
        targets: IDs of the numbered tables and figures; only these resolve.
        report: Receives each diagnostic (strict mode rejects them); None reads
            silently.
        place: Set for text that cannot hold a Word field; it names the place
            in the diagnostic and no reference resolves there.
        allowed: When set, only the `@` with these ordinals (counted over the
            visible `@` of the text) may resolve: a heading is judged on its
            source with emphasis markers and rendered without them.
        kept: Receives the ordinal of each `@` that resolves.
    """

    targets: frozenset[str] = frozenset()
    report: Callable[[str], None] | None = None
    place: str | None = None
    allowed: frozenset[int] | None = None
    kept: Callable[[int], None] | None = None


_REFERENCE_SCOPE: ContextVar[ReferenceScope | None] = ContextVar("ib_reference_scope", default=None)


@dataclass(frozen=True)
class _Candidate:
    """A `@tbl-…` / `@fig-…` that the grammar reads as a reference.

    Attributes:
        ordinal: Position of its `@` among the visible `@` of the text.
        run: Index of the run holding the `@`.
        start: Offset of the `@` in that run.
        end: Offset after the ID in that run (meaningful when `whole`).
        target: The whole visible ID after `@`.
        whole: The `@` and the whole ID lie in one run (no formatting splits it).
    """

    ordinal: int
    run: int
    start: int
    end: int
    target: str
    whole: bool


def _document_int(digits: str) -> int:
    """Convert digits written in a document; a number too long to convert is a rejection.

    Args:
        digits: A run of decimal digits matched in the Markdown.

    Returns:
        The number.
    """
    try:
        return int(digits)
    except ValueError as error:
        raise StrictValidationError(
            "", [f"Number is too long to read ({len(digits)} digits): {digits[:12]}..."]
        ) from error


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
        logger.warning(
            "Link target kept as written; it is not a usable local path (%s): %s", error, target
        )
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
                "may not be able to open it: %s",
                location,
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

    # Compiled once — used by cleanup_text. A backslash escapes any ASCII
    # punctuation character, as in CommonMark; an escaped HTML break is left
    # for normalize_html_breaks, which keeps the tag as text.
    _ESCAPE_RE = re.compile(r"\\(?!<br\s*/?>)([!-/:-@\[-`{-~])", re.IGNORECASE)
    # The set resolved before the escape set was widened. cleanup_text still
    # applies only this set inside code spans, inline math and link destinations,
    # so LaTeX commands such as \% and \, keep their backslash there.
    _LEGACY_ESCAPE_RE = re.compile(r'\\([\\$`^~.*"\'()\[\]{}|_-])')
    _HTML_BREAK_RE = re.compile(r"(?<!\\)<br\s*/?>", re.IGNORECASE)
    _ESCAPED_HTML_BREAK_RE = re.compile(r"\\(<br\s*/?>)", re.IGNORECASE)
    _CODE_SPAN_RE = re.compile(r"(?<![\\`])(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.DOTALL)
    # Everything `_ESCAPE_RE` unescapes, plus HTML breaks: escaping these makes
    # parse_runs reproduce inserted text verbatim.
    _ESCAPABLE_RE = re.compile(r'([\\$`^~.*"\'()\[\]{}|_-])')
    _LITERAL_BREAK_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
    # Every visible `@`; a reference starts where `tbl-` or `fig-` follows one.
    _AT_RE = re.compile("@")
    # Characters an ID is read through (the whole token is checked afterwards,
    # so `@tbl-a-` and `@tbl-` are reported instead of shortened).
    _REFERENCE_ID_CHARS = frozenset(
        "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
    )
    # Characters that may stand right before a reference's `@` besides
    # whitespace and the start of the text: opening brackets and quotes.
    _REFERENCE_OPENERS = frozenset("([{<\"'“‘「『《〈【〔（［｛")
    # Stands for a Markdown `\@`: in text that was unescaped before it is parsed
    # (callouts, numbered headings), and in runs until references are read, so
    # that `@` is never read as a reference. `ESCAPED_DOLLAR` does the same for
    # `\$`, so an escaped dollar never closes math around a reference.
    ESCAPED_AT = "\ue010"
    ESCAPED_DOLLAR = "\ue011"
    # One visible unit that is not prose: code, math, an image, a footnote mark,
    # a term value or any other token. It is neither a reference edge nor an
    # ID character.
    _OPAQUE = "\ue012"
    # Characters of a token left in run text (term values and the like); the two
    # escape marks above are the only private-use characters that stay readable.
    _TOKEN_CHAR_RE = re.compile("[\ue000-\ue00f\ue012-\uf8ff]")

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
    # Spans whose characters are not prose: code, inline math, a link destination.
    _LITERAL_SPAN_RE = re.compile(
        _CODE_SPAN_RE.pattern + "|" + _INLINE_LATEX_RE.pattern + r"|(?<=\])\([^()\s]*\)",
        re.DOTALL,
    )

    @classmethod
    def parse_runs(cls, text: str, escaped_at: bool = False) -> list[TextRun]:
        """
        Parse text into runs with formatting.
        Handles **bold**, inline $LaTeX$, and combinations.

        Args:
            text: Inline Markdown.
            escaped_at: The text was unescaped already and writes each escaped
                `@` as `ESCAPED_AT`; it is shown as `@` and never a reference.
        """
        runs: list[TextRun] = []
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
                        cls._apply_color(
                            cls._parse_inline_formatting(inline_text, escaped_at), color_hex
                        )
                    )

        runs = cls._split_image_runs(runs, images, code_spans, escaped_at)
        runs = cls._read_references(cls._split_code_runs(runs, code_spans), dollars=False)
        return cls._encode_link_targets(runs)

    @classmethod
    def _split_on_inline_latex(cls, text: str) -> list[tuple[str, bool]]:
        """
        Split text into (content, is_latex) segments.

        Returns:
            List of (text, is_latex) tuples preserving order
        """
        segments: list[tuple[str, bool]] = []
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
    def parse_runs_plain(cls, text: str, escaped_at: bool = False) -> list[TextRun]:
        """
        Parse runs WITHOUT LaTeX detection.
        Used for contexts where $ should not be interpreted as LaTeX
        (e.g., table cells with currency values). Text between dollar signs is
        still never read as a cross-reference. `escaped_at` as in `parse_runs`.
        """
        text, code_spans = cls._protect_code_spans(text)
        text, images = cls._protect_images(text)
        runs: list[TextRun] = []
        for segment_text, color_hex in cls._split_on_color_spans(text):
            runs.extend(
                cls._apply_color(cls._parse_inline_formatting(segment_text, escaped_at), color_hex)
            )
        runs = cls._split_image_runs(runs, images, code_spans, escaped_at)
        runs = cls._read_references(cls._split_code_runs(runs, code_spans), dollars=True)
        return cls._encode_link_targets(runs)

    @classmethod
    def _protect_images(cls, text: str) -> tuple[str, dict[str, tuple[str, str, str]]]:
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
        images: dict[str, tuple[str, str, str]] = {}
        if "![" not in text:
            return text, images
        prefix = "IMG"
        while prefix in text:
            prefix += "X"
        destinations = [
            match.span(2)
            for match in cls._INLINE_REFERENCE_RE.finditer(text)
            if match.group(2) and not cls._escaped(text, match.start())
        ]
        pieces: list[str] = []
        last = 0
        for match in cls._INLINE_IMAGE_RE.finditer(text):
            if cls._escaped(text, match.start()) or any(
                start <= match.start() < end for start, end in destinations
            ):
                continue
            token = f"{prefix}{len(images)}"
            images[token] = (match.group(0), match.group(1), match.group(2) or match.group(3))
            pieces.extend((text[last : match.start()], token))
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
        cls,
        runs: list[TextRun],
        images: dict[str, tuple[str, str, str]],
        code_spans: dict[str, str],
        escaped_at: bool = False,
    ) -> list[TextRun]:
        """Turn image tokens into image runs that keep the surrounding formatting.

        Math stays literal, so a token inside an equation gets its source back.
        A description cannot hold a Word field: its cross-references, read with
        the rules of body text, are reported (`check_text`).

        Args:
            runs: Runs parsed from text whose images were replaced by tokens.
            images: Token to (source, alt text, destination) from `_protect_images`.
            code_spans: Code-span tokens, restored inside image fields.
            escaped_at: Whether `ESCAPED_AT` stands for an escaped `@` (see `parse_runs`).

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
        result: list[TextRun] = []
        for run in runs:
            if run.is_latex:
                run.text = tokens.sub(lambda match: images[match.group(0)][0], run.text)
            if run.is_latex or not tokens.search(run.text):
                result.append(run)
                continue
            offset = 0
            for match in tokens.finditer(run.text):
                if match.start() > offset:
                    result.append(replace(run, text=run.text[offset : match.start()]))
                _, alt, destination = images[match.group(0)]
                source = restore(alt)
                description = cls.cleanup_text(source)
                path = restore(destination)
                if escaped_at:
                    description = description.replace(cls.ESCAPED_AT, "@")
                    path = path.replace(cls.ESCAPED_AT, "@")
                cls.check_text(source, "image description", escaped_at=escaped_at)
                image = Image(alt_text=description, path=path)
                result.append(replace(run, text="", image=image))
                offset = match.end()
            if offset < len(run.text):
                result.append(replace(run, text=run.text[offset:]))
        return result

    @classmethod
    def _protect_code_spans(cls, text: str) -> tuple[str, dict[str, str]]:
        """Shield literals before emphasis, HTML, links and math tokenize them."""
        literals: dict[str, str] = {}
        prefix = "\ue000CODE"
        while prefix in text:
            prefix += "X"

        def replace(match):
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(0)
            return token

        return cls._CODE_SPAN_RE.sub(replace, text), literals

    @classmethod
    def _protect_escapes(cls, text: str, escaped_at: bool = False) -> tuple[str, dict[str, str]]:
        """Keep escaped punctuation literal until inline tokenization finishes.

        With `escaped_at`, each `ESCAPED_AT` is an escaped `@` too.
        """
        literals: dict[str, str] = {}
        prefix = "\ue000ESC"
        while prefix in text:
            prefix += "X"

        def replace(match: Match[str]) -> str:
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(1)
            return token

        if escaped_at and cls.ESCAPED_AT in text:
            parts = text.split(cls.ESCAPED_AT)
            text = parts[0]
            for part in parts[1:]:
                token = prefix + str(len(literals)) + "\ue001"
                literals[token] = "@"
                text += token + part

        # Text between dollar signs keeps the legacy set whether or not this
        # context renders math, so `$12\%$` reads the same in a callout or a cell.
        pieces: list[str] = []
        offset = 0
        for span in cls._INLINE_LATEX_RE.finditer(text):
            pieces.append(cls._ESCAPE_RE.sub(replace, text[offset : span.start()]))
            pieces.append(cls._LEGACY_ESCAPE_RE.sub(replace, span.group(0)))
            offset = span.end()
        pieces.append(cls._ESCAPE_RE.sub(replace, text[offset:]))
        return "".join(pieces), literals

    @classmethod
    def _split_code_runs(cls, runs: list[TextRun], literals: dict[str, str]) -> list[TextRun]:
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
        result: list[TextRun] = []
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
                    result.append(replace(run, text=run.text[offset : match.start()]))
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
        content = literal[fence : len(literal) - fence].replace("\n", " ")
        if len(content) >= 2 and content[0] == " " and content[-1] == " " and content.strip():
            content = content[1:-1]
        return content

    @classmethod
    def _restore_code_spans(
        cls, runs: list[TextRun], literals: dict[str, str], mark: bool = False
    ) -> list[TextRun]:
        """Restore original text without manufacturing code-font semantics.

        Args:
            runs: Runs whose text and destinations hold tokens.
            literals: Token to the text it stands for.
            mark: Write an escaped `@` or `$` in run text as `ESCAPED_AT` or
                `ESCAPED_DOLLAR`; `_read_references` writes them back.

        Returns:
            The same runs.
        """
        marks = {"@": cls.ESCAPED_AT, "$": cls.ESCAPED_DOLLAR} if mark else {}
        for run in runs:
            for token, literal in literals.items():
                run.text = run.text.replace(token, marks.get(literal, literal))
                if run.hyperlink:
                    run.hyperlink = run.hyperlink.replace(token, literal)
        return runs

    @classmethod
    def _split_on_color_spans(cls, text: str) -> list[tuple[str, str | None]]:
        """Split text into color-scoped segments preserving original order."""
        segments: list[tuple[str, str | None]] = []
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
    def _extract_color_from_style(cls, style_attr: str) -> str | None:
        """Extract a normalized #RRGGBB color from an HTML style attribute."""
        match = cls._COLOR_STYLE_RE.search(style_attr)
        if not match:
            return None
        return match.group(1).upper()

    @staticmethod
    def _apply_color(runs: list[TextRun], color_hex: str | None) -> list[TextRun]:
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
        target = (
            destination[1:-1] if destination[:1] == "<" and destination[-1:] == ">" else destination
        )
        if _DRIVE_PATH_RE.match(target) or target[:1] in ("/", "\\"):
            return file_uri(target)
        if _URI_SCHEME_RE.match(target):
            return target
        return quote(target.replace("\\", "/"), safe="/#%")

    @classmethod
    def _parse_inline_formatting(cls, text: str, escaped_at: bool = False) -> list[TextRun]:
        """Parse links and numeric footnote references alongside inline emphasis.

        Args:
            text: One segment of inline Markdown (no math, one colour).
            escaped_at: Whether `ESCAPED_AT` stands for an escaped `@`.

        Returns:
            The segment's runs. An escaped `@` or `$` is written as `ESCAPED_AT`
            or `ESCAPED_DOLLAR` until `_read_references` has read the text.
        """
        text, literals = cls._protect_escapes(text, escaped_at)
        runs: list[TextRun] = []
        offset = 0
        for match in cls._INLINE_REFERENCE_RE.finditer(text):
            runs.extend(cls._parse_plain_formatting(text[offset : match.start()]))
            if match.group(3):
                runs.append(
                    TextRun(
                        text=match.group(3),
                        superscript=True,
                        footnote_id=_document_int(match.group(3)),
                    )
                )
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
        return cls._restore_code_spans(runs, literals, mark=True)

    # ── Cross-references ─────────────────────────────────────────────────────

    @staticmethod
    @contextmanager
    def reference_scope(scope: ReferenceScope | None) -> Iterator[ReferenceScope | None]:
        """Parse text in this context under `scope` (None: no text is a reference).

        Args:
            scope: Resolvable IDs, diagnostic sink and place, or None.

        Yields:
            The scope.
        """
        token = _REFERENCE_SCOPE.set(scope)
        try:
            yield scope
        finally:
            _REFERENCE_SCOPE.reset(token)

    @classmethod
    def check_text(
        cls, text: str, place: str, markdown: bool = True, escaped_at: bool = False
    ) -> None:
        """Report the references in text that cannot hold a Word field.

        Nothing is reported outside a scope. Inside one, every reference the
        grammar reads in `text` is reported as staying text in `place`.

        Args:
            text: The text as written (before any term value is inserted).
            place: Where the text is shown, for the diagnostic.
            markdown: Read `text` as inline Markdown; otherwise as the plain
                text it is shown as.
            escaped_at: Whether `ESCAPED_AT` stands for an escaped `@`.
        """
        scope = _REFERENCE_SCOPE.get()
        if scope is None or scope.report is None or "@" not in text:
            return
        with cls.reference_scope(ReferenceScope(report=scope.report, place=place)):
            if markdown:
                cls.parse_runs(text, escaped_at=escaped_at)
            else:
                cls._read_references([TextRun(text=text)], dollars=True)

    @classmethod
    def read_references(cls, runs: list[TextRun]) -> list[TextRun]:
        """Read references in runs built without Markdown (HTML table cells).

        The grammar is that of Markdown text; such runs hold no escapes, and
        text between dollar signs is not read.

        Args:
            runs: Finished runs of one cell.

        Returns:
            The runs, a run holding a resolvable reference split around it.
        """
        return cls._read_references(runs, dollars=True)

    @classmethod
    def parse_judged_runs(cls, text: str, source: str, escaped_at: bool = False) -> list[TextRun]:
        """Parse `text` with the references that resolve in `source`.

        A heading is written without its `**` markers (every heading run is
        bold) but judged like body text, on its source with the markers, so an
        ID that emphasis splits is never a reference. Diagnostics come from the
        source; the `@` in both texts are the same, counted in order.

        Args:
            text: The text to parse into runs.
            source: The same text with its emphasis markers.
            escaped_at: Whether `ESCAPED_AT` stands for an escaped `@`.

        Returns:
            The runs of `text`.
        """
        scope = _REFERENCE_SCOPE.get()
        if scope is None or "@" not in source:
            return cls.parse_runs(text, escaped_at=escaped_at)
        kept: list[int] = []
        with cls.reference_scope(replace(scope, kept=kept.append)):
            cls.parse_runs(source, escaped_at=escaped_at)
        judged = replace(scope, report=None, allowed=frozenset(kept), kept=None)
        with cls.reference_scope(judged):
            return cls.parse_runs(text, escaped_at=escaped_at)

    @classmethod
    def _read_references(cls, runs: list[TextRun], dollars: bool) -> list[TextRun]:
        """Split runs around resolvable references, report the others, unmark escapes.

        Outside a scope the runs are only unmarked, exactly as text without
        references. Inside one, each reference the grammar reads either
        resolves (its run is split around a `CrossReference`) or is reported
        without touching the runs.

        Args:
            runs: Finished runs; escaped `@` and `$` still marked.
            dollars: Text between dollar signs is not read (math is not rendered
                here; where it is, math is already its own run).

        Returns:
            The runs with escapes written back.
        """
        scope = _REFERENCE_SCOPE.get()
        if scope is not None and any("@" in run.text for run in runs):
            resolved: list[_Candidate] = []
            for candidate in cls._reference_candidates(runs, dollars):
                problem = cls._reference_problem(candidate, runs[candidate.run], scope)
                if problem is None:
                    resolved.append(candidate)
                    if scope.kept is not None:
                        scope.kept(candidate.ordinal)
                elif scope.report is not None:
                    scope.report(problem)
            if resolved:
                runs = cls._split_references(runs, resolved)
        for run in runs:
            if cls.ESCAPED_AT in run.text or cls.ESCAPED_DOLLAR in run.text:
                run.text = run.text.replace(cls.ESCAPED_AT, "@").replace(cls.ESCAPED_DOLLAR, "$")
        return runs

    @classmethod
    def _reference_candidates(cls, runs: list[TextRun], dollars: bool) -> list[_Candidate]:
        """Find what the grammar reads as references in the visible text of `runs`.

        The text is read as a reader sees it, across every run: emphasis and
        colour do not interrupt it, an empty run adds nothing, and code, math,
        images, footnote marks and term values are opaque units. A reference
        is `@` followed by `tbl-` or `fig-`, where the visible character before
        `@` is the start of the text, whitespace or an opening bracket or quote;
        its ID is every ID character after `@`, read whole even when formatting
        changes inside it; and an ID followed by `.` and an ASCII letter or
        digit is a domain name, not a reference. An escaped `@` is never one,
        and with `dollars` neither is an `@` between dollar signs.

        Args:
            runs: Finished runs; escaped `@` and `$` still marked.
            dollars: Skip text between dollar signs.

        Returns:
            The references in reading order.
        """
        pieces: list[str] = []
        starts: list[int] = []  # where each run's units begin in the visible text
        length = 0
        for run in runs:
            starts.append(length)
            if (
                run.image is not None
                or run.code
                or run.is_latex
                or run.footnote_id is not None
                or run.term_key is not None
            ):
                piece = cls._OPAQUE
            else:
                # A token (a term value, another literal) is opaque, one unit per character.
                piece = cls._TOKEN_CHAR_RE.sub(cls._OPAQUE, run.text)
            pieces.append(piece)
            length += len(piece)
        visible = "".join(pieces)
        math = [match.span() for match in cls._INLINE_LATEX_RE.finditer(visible)] if dollars else []
        math_starts = [low for low, _ in math]
        found: list[_Candidate] = []
        for ordinal, match in enumerate(cls._AT_RE.finditer(visible)):
            at = match.start()
            if visible[at + 1 : at + 5] not in ("tbl-", "fig-"):
                continue
            inside = bisect_right(math_starts, at) - 1
            if inside >= 0 and at < math[inside][1]:
                continue
            before = visible[at - 1] if at else ""
            if not (before == "" or before.isspace() or before in cls._REFERENCE_OPENERS):
                continue
            end = at + 1
            while end < len(visible) and visible[end] in cls._REFERENCE_ID_CHARS:
                end += 1  # always moves on, so reading ends with the text
            if (
                visible[end : end + 1] == "."
                and visible[end + 1 : end + 2].isascii()
                and visible[end + 1 : end + 2].isalnum()
            ):
                continue  # `@tbl-x.example`: a domain, not a reference
            run_index = bisect_right(starts, at) - 1
            found.append(
                _Candidate(
                    ordinal=ordinal,
                    run=run_index,
                    start=at - starts[run_index],
                    end=end - starts[run_index],
                    target=visible[at + 1 : end],
                    # Runs are consecutive blocks: the ID is whole if it ends in the same run.
                    whole=bisect_right(starts, end - 1) - 1 == run_index,
                )
            )
        return found

    @staticmethod
    def _reference_problem(
        candidate: _Candidate, run: TextRun, scope: ReferenceScope
    ) -> str | None:
        """Why a reference stays text in `scope`, or None when it resolves.

        Args:
            candidate: A reference the grammar read.
            run: The run holding its `@`.
            scope: The active scope.

        Returns:
            The diagnostic, or None.
        """
        written = "@" + candidate.target
        if scope.place is not None:
            return f"Cross-reference {written} in the {scope.place} stays text"
        if not candidate.whole:
            return f"Cross-reference {written} is split by formatting and stays text"
        problem = reference_id_problem(candidate.target, candidate.target[:3])
        if problem is not None:
            return f"Cross-reference {written}: {problem}"
        if run.hyperlink:
            return f"Cross-reference {written} is inside link text and stays text"
        if scope.allowed is not None and candidate.ordinal not in scope.allowed:
            return f"Cross-reference {written} does not resolve where it is written"
        if candidate.target not in scope.targets:
            return f"Cross-reference {written} has no table or figure with that id"
        return None

    @staticmethod
    def _split_references(runs: list[TextRun], resolved: list[_Candidate]) -> list[TextRun]:
        """Split the runs holding resolvable references around `CrossReference` runs.

        Args:
            runs: Finished runs.
            resolved: References that resolve, in reading order (each whole).

        Returns:
            The runs; the text around a reference keeps its run's formatting.
        """
        by_run: dict[int, list[_Candidate]] = {}
        for candidate in resolved:
            by_run.setdefault(candidate.run, []).append(candidate)
        result: list[TextRun] = []
        for index, run in enumerate(runs):
            if index not in by_run:
                result.append(run)
                continue
            values = {name: getattr(run, name) for name in TEXT_RUN_FIELDS if name != "text"}
            cursor = 0
            for candidate in by_run[index]:
                if candidate.start > cursor:
                    result.append(replace(run, text=run.text[cursor : candidate.start]))
                text = run.text[candidate.start : candidate.end]
                result.append(CrossReference(text=text, **values, target=candidate.target))
                cursor = candidate.end
            if cursor < len(run.text):
                result.append(replace(run, text=run.text[cursor:]))
        return result

    @classmethod
    def _encode_link_targets(cls, runs: list[TextRun]) -> list[TextRun]:
        """Encode destinations once every protected literal is back in place."""
        for run in runs:
            if run.hyperlink:
                run.hyperlink = cls.link_target(run.hyperlink)
        return runs

    @classmethod
    def _parse_plain_formatting(cls, text: str) -> list[TextRun]:
        """Parse bold, italic, superscript, and subscript runs from plain text."""
        runs: list[TextRun] = []
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

    @classmethod
    def _unescape(cls, text: str, mark_at: bool = False) -> str:
        """Resolve escapes in text that is not tokenized; literal spans keep the legacy set.

        Args:
            text: Raw inline Markdown of a heading, quote line or emphasis part.
            mark_at: Write an escaped `@` as `ESCAPED_AT` instead of `@`.

        Returns:
            Text with escapes resolved once.
        """

        def unescape(match: Match[str]) -> str:
            char = match.group(1)
            return cls.ESCAPED_AT if mark_at and char == "@" else char

        pieces: list[str] = []
        offset = 0
        for match in cls._LITERAL_SPAN_RE.finditer(text):
            pieces.append(cls._ESCAPE_RE.sub(unescape, text[offset : match.start()]))
            pieces.append(cls._LEGACY_ESCAPE_RE.sub(r"\1", match.group(0)))
            offset = match.end()
        pieces.append(cls._ESCAPE_RE.sub(unescape, text[offset:]))
        return "".join(pieces)

    @staticmethod
    def cleanup_text(text: str, mark_at: bool = False) -> str:
        """Remove markdown artifacts and escape characters (single-pass regex).

        With `mark_at`, an escaped `@` becomes `ESCAPED_AT` (see `parse_runs`).
        """
        return TextParser.normalize_html_breaks(TextParser._unescape(text, mark_at).strip())

    @staticmethod
    def cleanup_text_preserve_spacing(text: str) -> str:
        """Remove markdown escape characters while preserving surrounding spacing."""
        return TextParser.normalize_html_breaks(TextParser._unescape(text))

    @classmethod
    def normalize_html_breaks(cls, text: str) -> str:
        """Convert explicit HTML breaks, preserving literal code and escaped tags.

        Run after emphasis/link tokenization so breaks inside bold text retain
        their formatting and link destinations are never rewritten.
        """
        pieces: list[str] = []
        offset = 0
        for match in cls._CODE_SPAN_RE.finditer(text):
            part = cls._HTML_BREAK_RE.sub("\n", text[offset : match.start()])
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
        cls,
        text: str,
        resolver: TermResolver,
        tokens: TokenMap,
        latex: bool = True,
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
        pieces: list[str] = []
        last = 0
        for match in cls._COLOR_SPAN_RE.finditer(protected):
            if cls._extract_color_from_style(match.group(2)) is None:
                continue
            pieces.append(
                cls._tokenize_segment(
                    protected[last : match.start()],
                    code_spans,
                    resolver,
                    tokens,
                    latex,
                )
            )
            pieces.append(protected[match.start() : match.start(3)])
            pieces.append(
                cls._tokenize_segment(match.group(3), code_spans, resolver, tokens, latex)
            )
            pieces.append(protected[match.end(3) : match.end()])
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
    def _protect_link_destinations(cls, text: str) -> tuple[str, dict[str, str]]:
        """Shield each complete inline-link destination; labels stay visible.

        Args:
            text: Text whose code spans are already protected.

        Returns:
            Text with destinations replaced by tokens, and token -> destination.
        """
        literals: dict[str, str] = {}
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
        cls,
        text: str,
        code_spans: dict[str, str],
        resolver: TermResolver,
        tokens: TokenMap,
        latex: bool,
    ) -> str:
        """Tokenize one colour segment, keeping inline math verbatim."""
        pieces: list[str] = []
        last = 0
        if latex:
            for match in cls._INLINE_LATEX_RE.finditer(text):
                pieces.append(
                    cls._tokenize_inline(text[last : match.start()], code_spans, resolver, tokens)
                )
                pieces.append(match.group(0))
                last = match.end()
        pieces.append(cls._tokenize_inline(text[last:], code_spans, resolver, tokens))
        return "".join(pieces)

    @classmethod
    def _tokenize_inline(
        cls,
        text: str,
        code_spans: dict[str, str],
        resolver: TermResolver,
        tokens: TokenMap,
    ) -> str:
        """Tokenize text and link labels, keeping escapes and destinations verbatim."""
        protected, escapes = cls._protect_escapes(text)
        pieces: list[str] = []
        offset = 0
        for match in cls._INLINE_REFERENCE_RE.finditer(protected):
            pieces.append(
                cls._tokenize_references(
                    protected[offset : match.start()],
                    code_spans,
                    escapes,
                    resolver,
                    tokens,
                )
            )
            if match.group(3):
                pieces.append(match.group(0))
            else:
                label = cls._tokenize_references(
                    match.group(1), code_spans, escapes, resolver, tokens
                )
                pieces.append(f"[{label}]({match.group(2)})")
            offset = match.end()
        pieces.append(
            cls._tokenize_references(protected[offset:], code_spans, escapes, resolver, tokens)
        )
        result = "".join(pieces)
        for token, literal in escapes.items():
            result = result.replace(token, "\\" + literal)
        return result

    @staticmethod
    def _tokenize_references(
        text: str,
        code_spans: dict[str, str],
        escapes: dict[str, str],
        resolver: TermResolver,
        tokens: TokenMap,
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
    parse: Callable[[str], list[TextRun]],
    terms: TermResolver | None,
    latex: bool = True,
) -> tuple[str, list[TextRun], bool]:
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
