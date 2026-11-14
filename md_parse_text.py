"""Inline Markdown parsing, local link targets and term-variable substitution in text.

Changelog (captions):
    - `@tbl-…` / `@fig-…` in text become `CrossReference` runs; code, math, link
      destinations, escaped `\\@` and e-mail-like `a@tbl-x` stay literal.
"""

import logging
import os
import re
from collections.abc import Callable
from dataclasses import fields, replace
from pathlib import Path
from re import Match
from urllib.parse import quote, unquote, urlsplit

from document_model import CrossReference, Image, TextRun
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
    # `@tbl-…` / `@fig-…` that starts a word: not after a letter or digit (an
    # e-mail address), a path or URL character, another `@` or a backslash. The
    # ID is read generously and validated later, so `@tbl-Perf` is reported
    # rather than silently left as text; it ends on a letter or digit.
    CROSS_REFERENCE_RE = re.compile(r"(?<![\w@./\\-])@((?:tbl|fig)-[A-Za-z0-9_-]*[A-Za-z0-9])")

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
    def parse_runs(cls, text: str) -> list[TextRun]:
        """
        Parse text into runs with formatting.
        Handles **bold**, inline $LaTeX$, and combinations.
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
                        cls._apply_color(cls._parse_inline_formatting(inline_text), color_hex)
                    )

        runs = cls._split_image_runs(runs, images, code_spans)
        return cls._encode_link_targets(cls._split_code_runs(runs, code_spans))

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
    def parse_runs_plain(cls, text: str) -> list[TextRun]:
        """
        Parse runs WITHOUT LaTeX detection.
        Used for contexts where $ should not be interpreted as LaTeX
        (e.g., table cells with currency values).
        """
        text, code_spans = cls._protect_code_spans(text)
        text, images = cls._protect_images(text)
        runs: list[TextRun] = []
        for segment_text, color_hex in cls._split_on_color_spans(text):
            runs.extend(cls._apply_color(cls._parse_inline_formatting(segment_text), color_hex))
        runs = cls._split_image_runs(runs, images, code_spans)
        return cls._encode_link_targets(cls._split_code_runs(runs, code_spans))

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
    ) -> list[TextRun]:
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
                image = Image(alt_text=cls.cleanup_text(restore(alt)), path=restore(destination))
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
    def _protect_escapes(cls, text: str) -> tuple[str, dict[str, str]]:
        """Keep escaped punctuation literal until inline tokenization finishes."""
        literals: dict[str, str] = {}
        prefix = "\ue000ESC"
        while prefix in text:
            prefix += "X"

        def replace(match: Match[str]) -> str:
            token = prefix + str(len(literals)) + "\ue001"
            literals[token] = match.group(1)
            return token

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

    @staticmethod
    def _restore_code_spans(runs: list[TextRun], literals: dict[str, str]) -> list[TextRun]:
        """Restore original text without manufacturing code-font semantics."""
        for run in runs:
            for token, literal in literals.items():
                run.text = run.text.replace(token, literal)
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
    def _parse_inline_formatting(cls, text: str) -> list[TextRun]:
        """Parse links and numeric footnote references alongside inline emphasis."""
        text, literals = cls._protect_escapes(text)
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
        # Escaped characters are still tokens here, so `\@tbl-x` is not a reference.
        return cls._restore_code_spans(cls._split_cross_references(runs), literals)

    @classmethod
    def _split_cross_references(cls, runs: list[TextRun]) -> list[TextRun]:
        """Give each `@tbl-…` / `@fig-…` its own `CrossReference` run.

        The reference keeps the surrounding formatting (and a link, which the
        renderer then reports). Footnote, term, code, math and image runs are
        never split.

        Args:
            runs: Runs whose escapes and code spans are still protected tokens.

        Returns:
            The runs with every reference as its own run.
        """
        result: list[TextRun] = []
        for run in runs:
            if (
                "@" not in run.text
                or run.footnote_id is not None
                or run.term_key is not None
                or run.code
                or run.is_latex
                or run.image is not None
            ):
                result.append(run)
                continue
            offset = 0
            for match in cls.CROSS_REFERENCE_RE.finditer(run.text):
                if match.start() > offset:
                    result.append(replace(run, text=run.text[offset : match.start()]))
                values = {item.name: getattr(run, item.name) for item in fields(TextRun)}
                values["text"] = match.group(0)
                result.append(CrossReference(**values, target=match.group(1)))
                offset = match.end()
            if offset == 0:
                result.append(run)
            elif offset < len(run.text):
                result.append(replace(run, text=run.text[offset:]))
        return result

    @classmethod
    def cross_reference_targets(cls, text: str) -> list[str]:
        """IDs of the `@tbl-…` / `@fig-…` references written in plain text."""
        return [match.group(1) for match in cls.CROSS_REFERENCE_RE.finditer(text)]

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
    def _unescape(cls, text: str) -> str:
        """Resolve escapes in text that is not tokenized; literal spans keep the legacy set.

        Args:
            text: Raw inline Markdown of a heading, quote line or emphasis part.

        Returns:
            Text with escapes resolved once.
        """
        pieces: list[str] = []
        offset = 0
        for match in cls._LITERAL_SPAN_RE.finditer(text):
            pieces.append(cls._ESCAPE_RE.sub(r"\1", text[offset : match.start()]))
            pieces.append(cls._LEGACY_ESCAPE_RE.sub(r"\1", match.group(0)))
            offset = match.end()
        pieces.append(cls._ESCAPE_RE.sub(r"\1", text[offset:]))
        return "".join(pieces)

    @staticmethod
    def cleanup_text(text: str) -> str:
        """Remove markdown artifacts and escape characters (single-pass regex)"""
        return TextParser.normalize_html_breaks(TextParser._unescape(text).strip())

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
