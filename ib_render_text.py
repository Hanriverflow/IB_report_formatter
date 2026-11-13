"""Inline text rendering: runs, links, footnote references and inline images and equations."""

import logging
import os
import re

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from ib_render_fonts import FontStyler
from ib_render_footnotes import FootnoteRenderer
from ib_render_media import ImageRenderer, LaTeXRenderer
from input_safety import DEFAULT_LIMITS, AssetOutsideRootError, InputLimitError
from md_parser import Image, TextParser, TextRun
from render_styles import STYLE
from term_variables import register_term_run

logger = logging.getLogger("ib_renderer")


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

        for run_data in runs:
            if run_data.is_latex:
                TextRenderer._render_inline_latex(
                    paragraph,
                    run_data.text,
                    font_size=font_size,
                )
                continue
            if run_data.image is not None:
                TextRenderer._render_inline_image(paragraph, run_data.image, font_size)
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
    ):
        """Parse and render text with **bold**, *italic*, ^superscript^, and ~subscript~ markers.

        Handles nested markers in a multi-pass approach:
            1. Split on **bold** markers
            2. Within non-bold segments, split on *italic* markers
            3. Within all segments, detect ^superscript^ and ~subscript~ markers
        """
        font_name = font_name or STYLE.BODY_FONT
        font_size = font_size or STYLE.BODY_SIZE

        parsed_runs = TextParser.parse_runs_plain(text)
        if any(
            run.bold
            or run.italic
            or run.superscript
            or run.subscript
            or run.color_hex
            or run.code
            or run.hyperlink
            or run.image is not None
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
