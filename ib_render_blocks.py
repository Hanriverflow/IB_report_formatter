"""Block element rendering: table of contents, headings, paragraphs, lists and callouts.

Changelog (captions):
    - The static TOC preview shows a heading's `@tbl-…` references as `표 1`.
    - Review fixes: the preview reads the runs the heading renders
      (`HeadingRenderer.source_runs`), and headings and callouts unescaped at
      parse time keep an escaped `@` from becoming a reference.
    - Second review: a heading's references are judged on its source with the
      bold markers it is rendered without (`TextParser.parse_judged_runs`).
"""

import re
from dataclasses import replace

from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from document_model import MarkedBlockquote, MarkedHeading, TextRun
from ib_render_document_style import TableStyler
from ib_render_fonts import FontStyler
from ib_render_text import TextRenderer, reference_preview_text
from md_parser import (
    Blockquote,
    DocumentModel,
    ElementType,
    Heading,
    ListItem,
    Paragraph,
    TextParser,
)
from office_layout import NativeNumbering
from ooxml_order import insert_ordered
from render_styles import STYLE


class TOCRenderer:
    """Renders the Table of Contents"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, model: DocumentModel | None = None):
        """Insert an auto-updating TOC plus an immediate preview outline."""
        heading = self.doc.add_paragraph(STYLE.TOC_TITLE, style="TOC Heading")
        self._apply_toc_heading_style(heading)

        paragraph = self.doc.add_paragraph()
        run = paragraph.add_run()

        # TOC field
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")

        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = 'TOC \\o "1-4" \\h \\z \\u'

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "separate")

        fldChar3 = OxmlElement("w:fldChar")
        fldChar3.set(qn("w:fldCharType"), "end")

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        if model is not None:
            self._render_preview_entries(model)
        self.doc.paragraphs[-1].add_run()._r.append(fldChar3)

        self.doc.add_page_break()

    @staticmethod
    def _apply_toc_heading_style(paragraph) -> None:
        """Apply dedicated typography for the TOC title."""
        for run in paragraph.runs:
            FontStyler.apply_run_style(
                run,
                font_name=STYLE.TOC_FONT,
                font_size=STYLE.H1_SIZE,
                bold=True,
                color=STYLE.NAVY,
            )

    def _render_preview_entries(self, model: DocumentModel):
        """Render a static preview TOC so the document is useful before field update."""
        entries = []
        for element in model.elements:
            if element.element_type not in {
                ElementType.HEADING_1,
                ElementType.HEADING_2,
                ElementType.HEADING_3,
                ElementType.HEADING_4,
                ElementType.NUMBERED_HEADING,
            }:
                continue
            if not isinstance(element.content, Heading):
                continue

            level = max(1, min(element.content.level, 4))
            # A heading whose references resolve shows them as their cached
            # numbers, read from the runs the heading itself renders.
            preview = reference_preview_text(HeadingRenderer.source_runs(element.content))
            text = preview.strip() if preview is not None else element.content.text.strip()
            if text:
                entries.append((level, text))

        if not entries:
            return

        for level, text in entries:
            paragraph = self.doc.add_paragraph()
            paragraph.style = STYLE.STYLE_IB_BODY
            paragraph.paragraph_format.left_indent = Inches(0.2 * (level - 1))
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(2 if level <= 2 else 0)

            run = paragraph.add_run(text)
            FontStyler.apply_run_style(
                run,
                font_name=STYLE.TOC_FONT,
                font_size=STYLE.BODY_SIZE
                if level == 1
                else (Pt(10) if level == 2 else STYLE.SMALL_SIZE),
                bold=level == 1,
                color=STYLE.DARK_GRAY if level <= 2 else STYLE.MEDIUM_GRAY,
            )


class HeadingRenderer:
    """Renders headings"""

    # A code span, a backslash escape, or a bold marker. Only the bold marker is
    # dropped (every heading run is bold); the other two keep their characters.
    _BOLD_MARKER_RE = re.compile(TextParser._CODE_SPAN_RE.pattern + r"|\\.|\*\*", re.DOTALL)

    # Level → (font_size, color, bold)
    @property
    def _STYLE_CONFIG(self):
        return {
            1: (STYLE.H1_SIZE, STYLE.NAVY, True),
            2: (STYLE.H2_SIZE, STYLE.DARK_GRAY, True),
            3: (STYLE.H3_SIZE, STYLE.NAVY, True),
            4: (STYLE.H4_SIZE, STYLE.DARK_GRAY, True),
        }

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    @classmethod
    def source_runs(cls, heading: Heading) -> list[TextRun]:
        """The runs a heading renders; parse-time term runs are used as is."""
        if heading.runs:
            return heading.runs
        marked = isinstance(heading, MarkedHeading)
        source = heading.marked_text if isinstance(heading, MarkedHeading) else heading.text
        # Drop bold markers outside code spans and escapes, then parse once:
        # unescaping before the parse turned escaped punctuation into syntax.
        clean_text = cls._BOLD_MARKER_RE.sub(
            lambda match: "" if match.group(0) == "**" else match.group(0), source
        )
        # References are judged like body text, on the source with its markers.
        return TextParser.parse_judged_runs(clean_text.strip(), source.strip(), escaped_at=marked)

    def render(self, heading: Heading):
        """Render a heading with appropriate level; parse-time term runs are used as is."""
        source_runs = self.source_runs(heading)

        # Clamp level to 1-4 (Word supports Heading 1-9, but we style 1-4)
        level = max(1, min(heading.level, 4))
        p = self.doc.add_heading(level=level)

        size, color, bold = self._STYLE_CONFIG.get(level, (STYLE.BODY_SIZE, STYLE.DARK_GRAY, True))

        runs = [replace(run, bold=bold) for run in source_runs]
        TextRenderer.render_runs(
            p, runs, font_name=STYLE.HEADING_FONT, font_size=size, default_color=color
        )


class ParagraphRenderer:
    """Renders paragraphs"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, paragraph: Paragraph):
        """Render a paragraph with text runs"""
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
        if paragraph.runs:
            TextRenderer.render_runs(p, paragraph.runs)
        else:
            TextRenderer.render_text_with_bold(p, paragraph.text)


class ListRenderer:
    """Renders lists"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self.numbering = NativeNumbering(doc)

    def render_bullet(self, item: ListItem):
        """Render a bullet list item"""
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BULLET)
        indent = self._resolve_indent(item.indent_level)
        p.paragraph_format.left_indent = indent
        p.paragraph_format.first_line_indent = -STYLE.BULLET_INDENT

        if STYLE.NATIVE_NUMBERING:
            self.numbering.apply(p, item.indent_level, bullet=True)
            TextRenderer.render_runs(p, item.runs or TextParser.parse_runs(item.text))
            return

        # Bullet character
        bullet_run = p.add_run(f"{STYLE.BULLET_CHAR}  ")
        FontStyler.apply_run_style(
            bullet_run,
            font_name=STYLE.BODY_FONT,
            font_size=STYLE.BODY_SIZE,
            color=STYLE.NAVY,
        )

        # Content
        if item.runs:
            TextRenderer.render_runs(p, item.runs)
        else:
            TextRenderer.render_text_with_bold(p, item.text)

    def render_numbered(self, number: str, item: ListItem):
        """Render a numbered list item"""
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
        indent = self._resolve_indent(item.indent_level)
        p.paragraph_format.left_indent = indent
        p.paragraph_format.first_line_indent = -STYLE.BULLET_INDENT

        if STYLE.NATIVE_NUMBERING:
            self.numbering.apply(p, item.indent_level, start=int(number))
            TextRenderer.render_runs(p, item.runs or TextParser.parse_runs(item.text))
            return

        # Number
        num_run = p.add_run(f"{number}. ")
        num_run.font.bold = True
        num_run.font.name = STYLE.BODY_FONT
        num_run.font.size = STYLE.BODY_SIZE
        FontStyler.set_east_asian_font(num_run)

        # Content
        if item.runs:
            TextRenderer.render_runs(p, item.runs)
        else:
            TextRenderer.render_text_with_bold(p, item.text)

    @staticmethod
    def _resolve_indent(indent_level: int) -> Inches:
        """Compute a bounded hanging indent for nested lists.

        Levels 0-3 use the full 0.25in step. Deeper levels compress to a
        smaller increment so very deep lists remain readable within page
        margins instead of drifting off the page.
        """
        normalized_level = max(0, indent_level)
        full_levels = min(normalized_level + 1, STYLE.FULL_LIST_INDENT_LEVELS)
        extra_levels = max(0, normalized_level + 1 - STYLE.FULL_LIST_INDENT_LEVELS)

        indent_inches = (
            full_levels * STYLE.BULLET_INDENT.inches + extra_levels * STYLE.DEEP_LIST_INDENT.inches
        )
        indent_inches = min(indent_inches, STYLE.MAX_LIST_INDENT.inches)
        return Inches(indent_inches)


class CalloutRenderer:
    """
    Renders callout boxes (blockquotes) with professional IB styling.

    Supports different callout types with distinct visual styles:
        - KEY INSIGHT: Blue accent (default)
        - EXECUTIVE SUMMARY / 요약: Navy box with prominent styling
        - WARNING / 주의: Orange accent
        - NOTE / 참고: Gray accent
    """

    # Callout type configurations: (background_hex, border_color, title_color, icon)
    @property
    def _CALLOUT_STYLES(self):
        return {
            # Executive Summary / Important
            "EXECUTIVE SUMMARY": (STYLE.NAVY_HEX, STYLE.NAVY, STYLE.WHITE, "▶"),
            "요약": (STYLE.NAVY_HEX, STYLE.NAVY, STYLE.WHITE, "▶"),
            "핵심": (STYLE.NAVY_HEX, STYLE.NAVY, STYLE.WHITE, "▶"),
            "SUMMARY": (STYLE.NAVY_HEX, STYLE.NAVY, STYLE.WHITE, "▶"),
            # Insights
            "KEY INSIGHT": (STYLE.ACCENT_BLUE_HEX, STYLE.NAVY, STYLE.NAVY, "▌"),
            "시사점": (STYLE.ACCENT_BLUE_HEX, STYLE.NAVY, STYLE.NAVY, "▌"),
            "결론": (STYLE.ACCENT_BLUE_HEX, STYLE.NAVY, STYLE.NAVY, "▌"),
            # Warnings
            "WARNING": ("FFF3CD", STYLE.ORANGE, STYLE.ORANGE, "⚠"),
            "주의": ("FFF3CD", STYLE.ORANGE, STYLE.ORANGE, "⚠"),
            "RISK": ("FFF3CD", STYLE.ORANGE, STYLE.ORANGE, "⚠"),
            # Notes
            "NOTE": (STYLE.LIGHT_GRAY_HEX, STYLE.DARK_GRAY, STYLE.DARK_GRAY, "ℹ"),
            "참고": (STYLE.LIGHT_GRAY_HEX, STYLE.DARK_GRAY, STYLE.DARK_GRAY, "ℹ"),
        }

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, blockquote: Blockquote):
        """Render a callout box with style based on title."""
        title_upper = blockquote.title.upper()

        # Get style configuration (default to KEY INSIGHT style)
        style_config = self._CALLOUT_STYLES.get(
            title_upper,
            self._CALLOUT_STYLES.get(
                blockquote.title, (STYLE.ACCENT_BLUE_HEX, STYLE.NAVY, STYLE.NAVY, "▌")
            ),
        )
        bg_hex, border_color, title_color, icon = style_config

        # Handle RGBColor vs hex string for border
        if isinstance(border_color, str):
            border_hex = border_color
        else:
            border_hex = (
                f"{border_color[0]:02X}{border_color[1]:02X}{border_color[2]:02X}"
                if hasattr(border_color, "__getitem__")
                else STYLE.NAVY_HEX
            )

        table = self.doc.add_table(rows=1, cols=1)
        cell = table.rows[0].cells[0]
        TableStyler.set_cell_background(cell, bg_hex)

        # Title
        title_para = cell.paragraphs[0]
        title_run = title_para.add_run(f"{icon} {blockquote.title}")

        # For dark backgrounds (like Executive Summary), use white text
        if bg_hex == STYLE.NAVY_HEX:
            FontStyler.apply_run_style(
                title_run,
                font_name=STYLE.HEADING_FONT,
                font_size=Pt(12),
                bold=True,
                color=STYLE.WHITE,
            )
        else:
            FontStyler.apply_run_style(
                title_run,
                font_size=Pt(11),
                bold=True,
                color=title_color if isinstance(title_color, RGBColor) else STYLE.NAVY,
            )

        # Content — with inline formatting support (**bold**, *italic*, ^super^)
        content_text = blockquote.text.strip()
        # Determine text color based on background
        content_color = STYLE.WHITE if bg_hex == STYLE.NAVY_HEX else None
        if blockquote.runs:
            # Parse-time runs carry term values, which must not be parsed again.
            TextRenderer.render_runs(
                cell.add_paragraph(),
                blockquote.runs,
                font_name=STYLE.BODY_FONT,
                font_size=STYLE.BODY_SIZE,
                default_color=content_color,
            )
        elif content_text:
            content_para = cell.add_paragraph()
            marked = isinstance(blockquote, MarkedBlockquote)

            TextRenderer.render_text_with_formatting(
                content_para,
                blockquote.marked_text.strip()
                if isinstance(blockquote, MarkedBlockquote)
                else content_text,
                font_name=STYLE.BODY_FONT,
                font_size=STYLE.BODY_SIZE,
                default_color=content_color,
                escaped_at=marked,
            )

        # Apply border styling
        self._apply_callout_border(
            table, border_hex if isinstance(border_hex, str) else STYLE.NAVY_HEX
        )

        # Spacer
        self.doc.add_paragraph()

    def render_executive_summary(self, title: str, bullet_points: list):
        """
        Render a professional Executive Summary box.

        Args:
            title: Box title (e.g., "Executive Summary")
            bullet_points: List of key points to display
        """
        table = self.doc.add_table(rows=1, cols=1)
        cell = table.rows[0].cells[0]
        TableStyler.set_cell_background(cell, STYLE.NAVY_HEX)

        # Title
        title_para = cell.paragraphs[0]
        title_run = title_para.add_run(f"▶ {title}")
        FontStyler.apply_run_style(
            title_run,
            font_name=STYLE.HEADING_FONT,
            font_size=Pt(12),
            bold=True,
            color=STYLE.WHITE,
        )

        # Bullet points
        for point in bullet_points:
            point_para = cell.add_paragraph()
            bullet_run = point_para.add_run("  •  ")
            FontStyler.apply_run_style(
                bullet_run,
                font_name=STYLE.BODY_FONT,
                font_size=STYLE.BODY_SIZE,
                color=STYLE.WHITE,
            )
            text_run = point_para.add_run(point)
            FontStyler.apply_run_style(
                text_run,
                font_name=STYLE.BODY_FONT,
                font_size=STYLE.BODY_SIZE,
                color=STYLE.WHITE,
            )

        # Full border for executive summary
        self._apply_full_border(table, STYLE.NAVY_HEX)

        # Spacer
        self.doc.add_paragraph()

    @staticmethod
    def _apply_callout_border(table, border_hex: str):
        """Apply left accent border, hide other borders."""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tblBorders = OxmlElement("w:tblBorders")

        # Children follow the schema order: top, left, bottom, right.
        for border_name in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{border_name}")
            if border_name == "left":
                border.set(qn("w:val"), "single")
                border.set(qn("w:sz"), "32")
                border.set(qn("w:color"), border_hex)
            else:
                border.set(qn("w:val"), "nil")
            tblBorders.append(border)

        insert_ordered(tblPr, tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)

    @staticmethod
    def _apply_full_border(table, border_hex: str):
        """Apply full border around the callout box."""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tblBorders = OxmlElement("w:tblBorders")

        for border_name in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "12")
            border.set(qn("w:color"), border_hex)
            tblBorders.append(border)

        insert_ordered(tblPr, tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)
