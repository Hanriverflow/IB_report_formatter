"""Inline text rendering: runs, links, footnote references and inline images and equations.

Changelog (captions):
    - Caption numbers: the label, a `SEQ` field with its cached number and a hidden
      bookmark around both, from a render-scoped table of numbers.
    - Cross-references: `REF <bookmark> \\h \\* CHARFORMAT` with the target's
      `표 1` cached; an unresolvable reference stays text and is reported.
    - Review fixes: a `{#…}` after a non-figure image is reported; every
      written caption number is recorded for the order check.
    - Second review: the parser splits runs only for references that resolve
      (`reference_scope` gives it the render's numbers and diagnostic sink), so
      nothing is joined back here.
"""

import logging
import os
import re
from collections.abc import Callable, Iterator, Sequence
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field, replace
from typing import cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from document_model import (
    CaptionNumber,
    CrossReference,
    figure_id_problem,
)
from ib_render_fonts import FontStyler
from ib_render_footnotes import FootnoteRenderer
from ib_render_media import ImageRenderer, LaTeXRenderer
from input_safety import DEFAULT_LIMITS, AssetOutsideRootError, InputLimitError
from md_parser import Image, ReferenceScope, TextParser, TextRun
from render_styles import STYLE
from term_variables import register_term_run

logger = logging.getLogger("ib_renderer")


# ═══════════════════════════════════════════════════════════════════════════════
# CAPTION NUMBERS
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class CaptionNumbers:
    """The caption numbers of one render, worked out before composition.

    Attributes:
        by_object: Caption number of each numbered table or figure, keyed by the
            `id()` of the model object in the render's own copy of the model.
        by_id: The number each reference ID resolves to (its first definition).
        next_bookmark_id: The document's bookmark ID allocator.
        errors: The render's diagnostic sink.
    """

    by_object: dict[int, CaptionNumber] = field(default_factory=dict)
    by_id: dict[str, CaptionNumber] = field(default_factory=dict)
    next_bookmark_id: Callable[[], int] = lambda: 0
    errors: list[str] = field(default_factory=list)
    # Every caption number actually written, in document order.
    written: list[CaptionNumber] = field(default_factory=list)
    # The messages of the first `reported_count` entries of `errors`, so a
    # repeat is found without searching the list.
    reported: set[str] = field(default_factory=set)
    reported_count: int = 0


_CAPTION_NUMBERS: ContextVar[CaptionNumbers | None] = ContextVar("ib_caption_numbers", default=None)


@contextmanager
def use_caption_numbers(numbers: CaptionNumbers) -> Iterator[CaptionNumbers]:
    """Make `numbers` the caption numbers of the render running in this context."""
    token = _CAPTION_NUMBERS.set(numbers)
    try:
        yield numbers
    finally:
        _CAPTION_NUMBERS.reset(token)


def caption_prefix(item: object) -> list[TextRun]:
    """Return the `표 1` and `. ` runs that open the caption of `item`, if numbered.

    Args:
        item: A table or image of the model being rendered.

    Returns:
        The number run and its separator, or an empty list outside a render or
        for an item without a number.
    """
    numbers = _CAPTION_NUMBERS.get()
    number = numbers.by_object.get(id(item)) if numbers is not None else None
    if number is None:
        return []
    return [replace(number), TextRun(text=". ")]


def _report(numbers: CaptionNumbers, message: str) -> None:
    """Record a diagnostic once per render."""
    errors = numbers.errors
    if numbers.reported_count < len(errors):  # entries added without `_report`
        numbers.reported.update(errors[numbers.reported_count :])
        numbers.reported_count = len(errors)
    if message not in numbers.reported:
        errors.append(message)
        numbers.reported.add(message)
        numbers.reported_count += 1
        logger.warning("%s", message)


def reference_scope(numbers: CaptionNumbers) -> ReferenceScope:
    """The reference scope of a render: its numbered IDs and diagnostic sink."""
    return ReferenceScope(
        targets=frozenset(numbers.by_id), report=lambda message: _report(numbers, message)
    )


def resolve_cross_reference(reference: CrossReference, report: bool = True) -> CaptionNumber | None:
    """Return the caption a reference points to, or None to write its text.

    The parser makes a `CrossReference` only for an ID that resolves, so None
    means a model built by hand (or changed after parsing); inside a render
    that is reported.

    Args:
        reference: A resolvable `@tbl-…` / `@fig-…` run.
        report: Record the diagnostic when it does not resolve.

    Returns:
        The caption number of the target, or None.
    """
    numbers = _CAPTION_NUMBERS.get()
    if numbers is None:
        return None
    number = numbers.by_id.get(reference.target)
    if number is None and report:
        _report(numbers, f"Cross-reference @{reference.target} has no table or figure with that id")
    return number


def report_stray_figure_id(image: Image, following: str) -> None:
    """Report `{#…}` written right after an image that is not a numbered figure."""
    numbers = _CAPTION_NUMBERS.get()
    match = _FIGURE_ID_AFTER_IMAGE_RE.match(following)
    if numbers is None or match is None:
        return
    ref_id = match.group(1)
    problem = figure_id_problem(ref_id, image.path or "data:", image.alt_text)
    detail = problem or f"{ref_id!r} works only on an image line of its own"
    _report(numbers, f"Figure id {detail}; the image is not numbered")


_FIGURE_ID_AFTER_IMAGE_RE = re.compile(r"\{#([^{}\s]*)\}")


def reference_preview_text(runs: list[TextRun]) -> str | None:
    """Plain text of `runs` with resolvable references as their `표 1`, or None.

    Returns None when no reference resolves, so callers keep the text they
    showed before.
    """
    if not any(isinstance(run, CrossReference) for run in runs):
        return None
    pieces = []
    for run in runs:
        number = resolve_cross_reference(run, False) if isinstance(run, CrossReference) else None
        pieces.append(number.text if number is not None else run.text)
    return "".join(pieces)


class TextRenderer:
    """Renders text with formatting"""

    # Compiled regex — class-level cache
    # Delimiters must flank their text, as in `TextParser` (`2 * 3` stays literal).
    _BOLD_SPLIT_RE = re.compile(r"(\*\*(?!\s).*?(?<!\s)\*\*)")
    _ITALIC_SPLIT_RE = re.compile(r"(?<!\*)(\*(?![\s*])[^*]*?[^\s*]\*)(?!\*)")
    _SUPERSCRIPT_RE = re.compile(r"\^([^^]+?)\^")
    _SUBSCRIPT_PATTERN = r"(?<!~)~[A-Za-z0-9]{1,8}~(?!~)"
    _VERTICAL_ALIGN_SPLIT_RE = re.compile(r"(\^[^^]+?\^|" + _SUBSCRIPT_PATTERN + r")")
    _ESCAPE_RE = re.compile(r'\\([~.*"\'()\[\]{}|_-])')

    @staticmethod
    def render_runs(
        paragraph,
        runs: list[TextRun],
        default_color: RGBColor | None = None,
        font_name: str | None = None,
        font_size: Pt | None = None,
    ):
        """Render text runs to a paragraph.

        Term value runs are only registered here: post-processing then leaves
        them intact, and the orchestrator wraps them in content controls last
        when tags are on. Linked values stay plain.
        """
        font_name = font_name or STYLE.BODY_FONT
        font_size = font_size or STYLE.BODY_SIZE

        for index, run_data in enumerate(runs):
            if run_data.is_latex:
                TextRenderer._render_inline_latex(
                    paragraph,
                    run_data.text,
                    font_size=font_size,
                )
                continue
            if run_data.image is not None:
                if index + 1 < len(runs):
                    report_stray_figure_id(run_data.image, runs[index + 1].text)
                TextRenderer._render_inline_image(paragraph, run_data.image, font_size)
                continue
            target = (
                resolve_cross_reference(run_data) if isinstance(run_data, CrossReference) else None
            )
            if isinstance(run_data, CaptionNumber) or target is not None:
                style = run_data

                def apply(run, style: TextRun = style) -> None:
                    run.font.name = font_name
                    run.font.size = font_size
                    run.font.bold = style.bold
                    run.font.italic = style.italic
                    if style.superscript:
                        run.font.superscript = True
                    if style.subscript:
                        run.font.subscript = True
                    color = TextRenderer._rgb_from_hex(style.color_hex) or default_color
                    if color:
                        run.font.color.rgb = color
                    FontStyler.set_east_asian_font(run)

                if target is not None:
                    # CHARFORMAT: the result takes the instruction's (body) formatting,
                    # not the caption's, however long the number becomes after F9.
                    TextRenderer._field_runs(
                        paragraph, f" REF {target.bookmark} \\h \\* CHARFORMAT ", target.text, apply
                    )
                else:
                    TextRenderer.write_caption_number(
                        paragraph, cast(CaptionNumber, run_data), apply
                    )
                continue

            run = paragraph.add_run(run_data.text)
            # Inline code keeps its literal text in a monospace Latin font.
            run.font.name = STYLE.CODE_FONT if run_data.code else font_name
            run.font.size = font_size
            run.font.bold = run_data.bold
            run.font.italic = run_data.italic
            if run_data.superscript:
                run.font.superscript = True
            if run_data.subscript:
                run.font.subscript = True
            run_color = TextRenderer._rgb_from_hex(run_data.color_hex) or default_color
            if run_color:
                run.font.color.rgb = run_color
            FontStyler.set_east_asian_font(run)
            if run_data.footnote_id is not None:
                FootnoteRenderer._replace_with_native_reference(run, run_data.footnote_id)
            if run_data.hyperlink:
                link = OxmlElement("w:hyperlink")
                relation = paragraph.part.relate_to(
                    run_data.hyperlink, RT.HYPERLINK, is_external=True
                )
                link.set(qn("r:id"), relation)
                run.font.underline = True
                if not run_color:
                    run.font.color.rgb = STYLE.NAVY
                link.append(run._r)
                paragraph._p.append(link)
            elif run_data.term_key is not None:
                register_term_run(run._r, run_data.term_key, run_data.text)

    @staticmethod
    def write_caption_number(paragraph, number: CaptionNumber, style: Callable[..., None]) -> None:
        """Write `표 1` as Word's caption command does, styling every run alike.

        The label is text and the number is a `SEQ` field whose cached result is
        the number this render worked out, so the document reads correctly
        before any field update. A hidden bookmark spans label and number, which
        is what a reference to the caption shows.

        Args:
            paragraph: The caption paragraph.
            number: The caption number; an empty `bookmark` writes no bookmark.
            style: Applies the caption's run formatting to one python-docx run.
        """
        label = paragraph.add_run(number.label + " ")
        style(label)
        runs = TextRenderer._field_runs(
            paragraph, f" SEQ {number.label} \\* ARABIC ", str(number.number), style
        )
        numbers = _CAPTION_NUMBERS.get()
        if numbers is not None:
            numbers.written.append(number)
        if not number.bookmark:
            return
        bookmark_id = str(numbers.next_bookmark_id() if numbers is not None else 0)
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bookmark_id)
        start.set(qn("w:name"), number.bookmark)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bookmark_id)
        label._r.addprevious(start)
        runs[-1]._r.addnext(end)

    @staticmethod
    def _field_runs(paragraph, instruction: str, result: str, style: Callable[..., None]) -> list:
        """Append a complex field with a cached result; return its five runs.

        Args:
            paragraph: Paragraph receiving the field.
            instruction: Field instruction, with its surrounding spaces.
            result: Cached result shown until Word updates the field.
            style: Applies run formatting to each run.

        Returns:
            The begin, instruction, separate, result and end runs.
        """
        runs = []
        for kind, text in (
            ("begin", None),
            ("instr", instruction),
            ("separate", None),
            ("result", result),
            ("end", None),
        ):
            run = paragraph.add_run(result if kind == "result" else None)
            style(run)
            if kind == "instr":
                code = OxmlElement("w:instrText")
                code.set(qn("xml:space"), "preserve")
                code.text = text
                run._r.append(code)
            elif kind != "result":
                char = OxmlElement("w:fldChar")
                char.set(qn("w:fldCharType"), kind)
                run._r.append(char)
            runs.append(run)
        return runs

    @staticmethod
    def _render_inline_image(paragraph, image: Image, font_size: Pt) -> None:
        """Insert an inline picture, no wider than a body image.

        Relative paths resolve against the source folder the orchestrator put on
        the document part; table cells are fitted after their widths are final.
        A picture that cannot be loaded leaves a visible marker and a diagnostic
        that strict mode rejects.
        """
        message = "Image could not be rendered: " + (image.alt_text or image.path)
        try:
            source = ImageRenderer.image_source(
                image,
                getattr(paragraph.part, "_ib_source_dir", None),
                getattr(paragraph.part, "_ib_asset_roots", ()),
                limits=getattr(paragraph.part, "_ib_limits", DEFAULT_LIMITS),
            )
            if source is None:
                raise FileNotFoundError(image.path)
            shape = paragraph.add_run().add_picture(source)
            limit = Inches(ImageRenderer.MAX_WIDTH_INCHES)
            if shape.width > limit:
                shape.height = int(shape.height * limit / shape.width)
                shape.width = limit
            ImageRenderer._apply_alt_text(shape, image.alt_text)
            return
        except InputLimitError:
            raise
        except AssetOutsideRootError as err:
            message = str(err)
            logger.warning("%s", message)
        except Exception as err:
            logger.warning("Inline image could not be inserted (%s): %s", image.alt_text, err)
        FontStyler.apply_run_style(
            paragraph.add_run(f"[Image: {image.alt_text}]"),
            font_size=font_size,
            italic=True,
            color=STYLE.RED,
        )
        errors = getattr(paragraph.part, "_ib_render_errors", None)
        if errors is not None:
            errors.append(message)

    @staticmethod
    def _render_inline_latex(paragraph, expression: str, font_size: Pt):
        """Render inline LaTeX as an inline image when possible."""
        image_path = LaTeXRenderer.render_to_image(
            expression,
            fontsize=max(int(round(font_size.pt)), 12),
            dpi=200,
        )

        if image_path:
            try:
                run = paragraph.add_run()
                run.add_picture(
                    image_path,
                    height=Pt(max(font_size.pt * 1.4, 14)),
                )
                return
            except Exception as err:
                logger.warning("Inline LaTeX image insertion failed: %s", err)
            finally:
                try:
                    os.unlink(image_path)
                except OSError:
                    pass

        fallback_run = paragraph.add_run(f"[{expression}]")
        FontStyler.apply_run_style(
            fallback_run,
            font_name="Consolas",
            font_size=font_size,
            italic=True,
            color=STYLE.DARK_GRAY,
        )
        message = "Inline LaTeX could not be rendered: " + expression
        logger.warning("%s", message)
        # The document part owns this sink, so concurrent renders and reusable
        # static text helpers never share diagnostic state.
        errors = getattr(paragraph.part, "_ib_render_errors", None)
        if errors is not None:
            errors.append(message)

    @staticmethod
    def render_text_with_bold(
        paragraph,
        text: str,
        font_name: str | None = None,
        font_size: Pt | None = None,
    ):
        """Parse and render text with **bold** markers (legacy, delegates to render_text_with_formatting)."""
        TextRenderer.render_text_with_formatting(
            paragraph, text, font_name=font_name, font_size=font_size
        )

    @staticmethod
    def render_text_with_formatting(
        paragraph,
        text: str,
        font_name: str | None = None,
        font_size: Pt | None = None,
        default_color: RGBColor | None = None,
        source: str | None = None,
        escaped_ats: Sequence[int] = (),
    ):
        """Parse and render text with **bold**, *italic*, ^superscript^, and ~subscript~ markers.

        Handles nested markers in a multi-pass approach:
            1. Split on **bold** markers
            2. Within non-bold segments, split on *italic* markers
            3. Within all segments, detect ^superscript^ and ~subscript~ markers

        `source` is the text before it was stripped and `escaped_ats` the offsets
        there of each `@` escaped before the text was unescaped (a callout), so
        that such an `@` is never a reference.
        """
        font_name = font_name or STYLE.BODY_FONT
        font_size = font_size or STYLE.BODY_SIZE

        parsed_runs = (
            TextParser.parse_judged_runs(text, source, escaped_ats, plain=True)
            if source is not None and escaped_ats
            else TextParser.parse_runs_plain(text)
        )
        if any(
            run.bold
            or run.italic
            or run.superscript
            or run.subscript
            or run.color_hex
            or run.code
            or run.hyperlink
            or run.image is not None
            or isinstance(run, CrossReference)
            for run in parsed_runs
        ):
            TextRenderer.render_runs(
                paragraph,
                parsed_runs,
                default_color=default_color,
                font_name=font_name,
                font_size=font_size,
            )
            return

        # Split on bold markers first
        bold_parts = TextRenderer._BOLD_SPLIT_RE.split(text)
        for bold_part in bold_parts:
            if not bold_part:
                continue

            if TextParser.flanked(bold_part, "**"):
                # Bold segment — check for ^superscript^ / ~subscript~ inside
                inner = TextRenderer._cleanup(bold_part[2:-2])
                TextRenderer._render_with_vertical_align(
                    paragraph,
                    inner,
                    font_name,
                    font_size,
                    bold=True,
                    italic=False,
                    color=default_color,
                )
            else:
                # Non-bold segment — check for *italic* markers
                italic_parts = TextRenderer._ITALIC_SPLIT_RE.split(bold_part)
                for italic_part in italic_parts:
                    if not italic_part:
                        continue

                    if TextParser.flanked(italic_part, "*"):
                        inner = TextRenderer._cleanup(italic_part[1:-1])
                        TextRenderer._render_with_vertical_align(
                            paragraph,
                            inner,
                            font_name,
                            font_size,
                            bold=False,
                            italic=True,
                            color=default_color,
                        )
                    else:
                        cleaned = TextRenderer._cleanup(italic_part)
                        if cleaned:
                            TextRenderer._render_with_vertical_align(
                                paragraph,
                                cleaned,
                                font_name,
                                font_size,
                                bold=False,
                                italic=False,
                                color=default_color,
                            )

    @staticmethod
    def _render_with_vertical_align(
        paragraph,
        text: str,
        font_name: str,
        font_size: Pt,
        bold: bool,
        italic: bool,
        color: RGBColor | None = None,
    ):
        """Render text, detecting ^superscript^ and ~subscript~ markers."""
        if not text:
            return

        parts = TextRenderer._VERTICAL_ALIGN_SPLIT_RE.split(text)
        for part in parts:
            if not part:
                continue

            is_superscript = False
            is_subscript = False
            content = part

            if part.startswith("^") and part.endswith("^") and len(part) > 2:
                content = part[1:-1]
                is_superscript = True
            elif part.startswith("~") and part.endswith("~") and len(part) > 2:
                content = part[1:-1]
                is_subscript = True

            run = paragraph.add_run(content)
            run.font.name = font_name
            run.font.size = font_size
            run.font.bold = bold
            run.font.italic = italic
            if is_superscript:
                run.font.superscript = True
            if is_subscript:
                run.font.subscript = True
            if color:
                run.font.color.rgb = color
            FontStyler.set_east_asian_font(run)

    @staticmethod
    def _rgb_from_hex(color_hex: str | None) -> RGBColor | None:
        """Convert #RRGGBB strings to python-docx RGBColor."""
        if not color_hex:
            return None

        normalized = color_hex.lstrip("#")
        if len(normalized) != 6 or not re.fullmatch(r"[0-9A-Fa-f]{6}", normalized):
            return None
        return RGBColor.from_string(normalized.upper())

    @staticmethod
    def _cleanup(text: str) -> str:
        """Clean up escape characters (single-pass compiled regex)"""
        return TextRenderer._ESCAPE_RE.sub(r"\1", text).strip()
