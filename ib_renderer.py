"""
IB Renderer Module for Word Report Generation
Handles styling and rendering of document elements in IB Bank style.

Changelog (property-test regressions):
    - Heading bold-marker removal skips code spans and escaped characters, and
      escapes are resolved once; other headings render as before.
    - A whitespace-only code line keeps its own whitespace (an empty line keeps
      the one-space spacer).

Changelog (house style fixes):
    - A landscape table right after a bare page break (cover, TOC) turns the
      page once: the section break replaces the page break.

Changelog (memo rendering):
    - Render inline code runs in the code font; keep them out of number formatting.
    - General (non-IB) profiles print Korean next to Latin text and digits tight.

Changelog (legacy output fixes):
    - Insert borders, fills, cell margins and run styles in schema order.
    - Without a cover, office/memo openings precede a forced TOC, whose
      preview no longer lists the title heading.

Changelog (term-sheet foundation):
    - Validate term-sheet metadata and resolve house text before document output.
    - Preserve confirmation fences in code panels when confirmation text is missing.
    - Term-sheet composition in the shared path: document defaults, opening
      block, per-line paragraphs, accent headings, confirmation box and
      per-section header/footer (term_sheet.py; callbacks injected).

Changelog (table spans):
    - Merge validated span rectangles in every profile: size the empty table,
      merge, then fill each owner cell once; covered cells repeat vertical fills.
    - Render the optional table note after the source in every profile.
    - Term-sheet tables: fixed label grid for key-value tables, label tiers,
      per-line cell paragraphs with marker indents and estimated row splitting.

Changelog (term variables):
    - Render carried heading/callout runs; tag term values as content controls
      after all run post-processing, then snapshot them in custom properties.
    - Keep term values out of table numeric formatting.

Changelog (cover-free title):
    - Render the IB report title, subtitle and memo-style metadata without a cover.
    - Place the cover-free report opening before the TOC on the first page.
    - Keep the body title and inferred subtitle out of the report's TOC.

Changelog (hardening):
    - Diagnose CJK raster glyph loss without changing Word font declarations.
    - Preserve semantic numeric-cell runs and record equation/image failures.
    - Validate structure without the full audit observation pass.
    - Render equations on independent Agg figures with unconditional cleanup.
    - Omit inferred IB subtitle headings from the body and TOC only with a cover.

Changelog (v2):
    - Fixed header row styling (p.clear() + add_run pattern)
    - Compatible with md_parser v2 heading levels (1-4)
    - NUMBERED_HEADING rendered at correct level
    - Compiled regex for TextRenderer performance
    - Style name constants in IBStyle
    - Element-level error resilience with try-except
    - Table column count defensive checks
    - Blockquote multi-line content preserved
"""

import logging
import platform as platform
import re
import time
from collections.abc import Callable, Iterator, Sequence, Set
from contextlib import contextmanager
from copy import deepcopy
from dataclasses import replace
from pathlib import Path
from typing import (
    Any,
    cast,
)
from uuid import uuid4

import yaml
from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Mm, Pt, RGBColor

from document_model import Chart, ConfirmationBlock
from document_profiles import (
    RenderOptions,
    default_metadata,
    load_style,
    resolve_options,
    validate_office_metadata,
    validate_term_sheet_metadata,
)
from ib_render_blocks import CalloutRenderer as CalloutRenderer
from ib_render_blocks import HeadingRenderer as HeadingRenderer
from ib_render_blocks import ListRenderer as ListRenderer
from ib_render_blocks import ParagraphRenderer as ParagraphRenderer
from ib_render_blocks import TOCRenderer as TOCRenderer
from ib_render_document_style import DocumentStyler as DocumentStyler
from ib_render_document_style import GeneratorSignatureWriter as GeneratorSignatureWriter
from ib_render_document_style import TableStyler as TableStyler
from ib_render_fonts import FontPolicy as FontPolicy
from ib_render_fonts import FontStyler as FontStyler
from ib_render_footnotes import FootnoteRenderer as FootnoteRenderer
from ib_render_footnotes import NativeFootnotesPart as NativeFootnotesPart
from ib_render_footnotes import rebase_local_links as rebase_local_links
from ib_render_media import ChartRenderer as ChartRenderer
from ib_render_media import ImageRenderer as ImageRenderer
from ib_render_media import LaTeXRenderer as LaTeXRenderer
from ib_render_text import TextRenderer as TextRenderer
from input_safety import (
    ConversionRejection,
    InputLimitError,
    InputLimits,
    RejectionCode,
    StrictValidationError,
    document_validation,
    resolve_asset_roots,
    validate_xml_strings,
)
from md_parser import (
    Blockquote,
    CodeBlock,
    DocumentModel,
    Element,
    ElementType,
    Heading,
    Image,
    LaTeXEquation,
    ListItem,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TableType,
    TextParser,
    TextRun,
)
from numeric_checks import CHECKS_PROPERTY, checks_to_json, parse_checks
from office_layout import (
    add_text,
    letter_appendix_index,
    render_office_closing,
    render_office_opening,
    setup_korean_spacing,
    setup_letter_styles,
)
from ooxml_order import insert_ordered
from render_styles import STYLE, collect_raster_font_diagnostics, use_style
from render_styles import IBStyle as IBStyle
from term_sheet import (
    ROW_SPLIT_THRESHOLD,
    TermSheetStyle,
    TermSheetTexts,
    add_table_heading,
    add_table_note,
    add_table_source,
    add_table_spacer,
    apply_marker_layout,
    apply_table_frame,
    configure_cell_paragraph,
    confirmation_has_text,
    estimate_cell_lines,
    header_token_width,
    key_value_widths,
    label_widths,
    line_text,
    natural_cell_width,
    positive_printable_width,
    render_confirmation,
    render_term_sheet_heading,
    render_term_sheet_opening,
    render_term_sheet_paragraph,
    resolve_term_sheet_texts,
    set_cell_fill,
    set_row_pagination,
    setup_term_sheet_header_footer,
    setup_term_sheet_styles,
    split_run_lines,
)
from term_variables import (
    TERM_PROPERTY_PREFIX,
    TermRunCollector,
    collect_term_runs,
    term_cell_text,
    wrap_term_controls,
)

logger = logging.getLogger(__name__)

_BLACK = RGBColor(0, 0, 0)


# ═══════════════════════════════════════════════════════════════════════════════
# STYLE CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════════════
# ELEMENT RENDERERS
# ═══════════════════════════════════════════════════════════════════════════════


class CoverRenderer:
    """Renders the cover page"""

    include_disclaimer: bool = True

    _COVER_DISCLAIMER_TEXT = (
        "당행은 해당 문서에 최대한 정확하고 완전한 정보를 담고자 노력하였으나, 오류와 중요정보의 "
        "누락이 있을 수 있으며, 정보의 정확성, 완전성 및 적정성을 보장하지 않습니다. 이 문서는 "
        "고객의 이해를 돕기 위하여 작성된 설명자료에 불과하므로, 고객은 각자의 책임으로 개별 계약서나 "
        "공시된 정보를 통하여 거래의 내용을 숙지하여야 합니다. 이 문서는 확정적인 거래조건을 "
        "구성하지 않으며 법적인 책임을 위한 근거자료로 사용될 수 없습니다. 본 자료는 당행의 "
        "저작물로서 모든 저작권은 당행에게 있으며, 당행의 동의 없이 어떠한 경우에도 어떠한 형태로든 "
        "복제, 배포, 전송, 변경, 대여할 수 없으며, 당행의 요청 시에 즉시 반환, 파기하여 주시기 바랍니다."
    )

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, metadata):
        """Render upgraded professional IB cover page."""
        title = (metadata.title or "IB Report").strip()
        subtitle = metadata.subtitle.strip()
        ticker = metadata.ticker.strip()
        sector = metadata.sector.strip()
        analyst = metadata.analyst.strip()
        company = metadata.company.strip()
        subject_company = metadata.extra.get("subject_company", "").strip()
        date_text = metadata.extra.get("date", "").strip()
        recipient = metadata.extra.get("recipient", "").strip()

        report_type = metadata.extra.get("report_type", "").strip().upper()
        institution = subject_company or company
        identity = ticker if ticker else institution
        cover_title = title
        cover_identity = identity

        if (
            not ticker
            and self._is_meaningful_metadata_value(
                institution, style_default="Korea Development Bank"
            )
            and title.startswith(institution)
        ):
            stripped_title = title[len(institution) :].strip(" :-")
            if stripped_title:
                cover_identity = institution
                cover_title = stripped_title.strip(" —–-:")

        self._add_spacer(2)

        # Report classification block
        if report_type:
            type_para = self.doc.add_paragraph()
            type_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            type_run = type_para.add_run(report_type)
            FontStyler.apply_run_style(
                type_run,
                font_name=STYLE.COVER_FONT,
                font_size=Pt(11),
                bold=True,
                color=STYLE.DARK_GRAY,
            )

        if self._is_meaningful_metadata_value(sector, style_default="SECTOR"):
            sector_para = self.doc.add_paragraph()
            sector_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sector_run = sector_para.add_run(sector.upper())
            FontStyler.apply_run_style(
                sector_run,
                font_name=STYLE.COVER_FONT,
                font_size=Pt(10),
                color=STYLE.MEDIUM_GRAY,
            )

        self._add_spacer(1)
        self._add_horizontal_rule(STYLE.NAVY_HEX, "14")
        self._add_spacer(2)

        # Security / company identifier
        if self._is_meaningful_metadata_value(
            cover_identity, style_default="Korea Development Bank"
        ):
            id_para = self.doc.add_paragraph()
            id_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            id_run = id_para.add_run(cover_identity)
            FontStyler.apply_run_style(
                id_run,
                font_name=STYLE.COVER_FONT,
                font_size=Pt(15),
                bold=True,
                color=STYLE.NAVY,
            )

        # Main title
        title_para = self.doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(cover_title)
        FontStyler.apply_run_style(
            title_run,
            font_name=STYLE.COVER_FONT,
            font_size=Pt(24),
            bold=True,
            color=STYLE.NAVY,
        )

        # Subtitle
        if subtitle:
            subtitle_para = self.doc.add_paragraph()
            subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            subtitle_run = subtitle_para.add_run(subtitle)
            FontStyler.apply_run_style(
                subtitle_run,
                font_name=STYLE.COVER_FONT,
                font_size=Pt(12),
                italic=True,
                color=STYLE.DARK_GRAY,
            )

        self._add_spacer(2)
        self._add_horizontal_rule(STYLE.GRAY_BORDER_HEX, "6")
        self._add_spacer(2)

        # Metadata panel
        self._render_metadata_panel(
            report_date=date_text or time.strftime("%B %d, %Y"),
            analyst=analyst,
            company=institution,
            sector=sector,
            recipient=recipient,
            analysis_period=metadata.extra.get("analysis_period", "").strip(),
            analysis_basis=metadata.extra.get("analysis_basis", "").strip(),
        )

        self._add_spacer(1)
        if getattr(self, "include_disclaimer", True):
            self._render_cover_disclaimer_table()

        self.doc.add_page_break()

    def _render_cover_disclaimer_table(self):
        """Render cover disclaimer table."""
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = STYLE.STYLE_TABLE_GRID

        cell = table.rows[0].cells[0]
        TableStyler.set_cell_background(cell, STYLE.LIGHT_GRAY_HEX)

        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.clear()
        paragraph.paragraph_format.line_spacing = 1.15
        paragraph.paragraph_format.space_before = Pt(2)
        paragraph.paragraph_format.space_after = Pt(2)

        run = paragraph.add_run(self._COVER_DISCLAIMER_TEXT)
        FontStyler.apply_run_style(
            run,
            font_name=STYLE.COVER_FONT,
            font_size=Pt(8.5),
            color=STYLE.DARK_GRAY,
        )

        self._apply_cover_disclaimer_border(table)

    @staticmethod
    def _apply_cover_disclaimer_border(table):
        """Apply thin border around cover disclaimer table."""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tblBorders = OxmlElement("w:tblBorders")

        for border_name in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:color"), STYLE.GRAY_BORDER_HEX)
            tblBorders.append(border)

        for border_name in ("insideH", "insideV"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "nil")
            tblBorders.append(border)

        insert_ordered(tblPr, tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)

    def _render_metadata_panel(
        self,
        report_date: str,
        analyst: str,
        company: str,
        sector: str,
        recipient: str,
        analysis_period: str = "",
        analysis_basis: str = "",
    ):
        """Render a compact two-column cover metadata panel."""
        rows = []

        if report_date:
            rows.append(("Report Date", report_date))
        if analysis_period:
            rows.append(("Analysis Period", analysis_period))
        if analysis_basis:
            rows.append(("Analysis Basis", analysis_basis))

        if self._is_meaningful_metadata_value(analyst, style_default="DCM Team 1"):
            rows.append(("Prepared By", analyst))
        if self._is_meaningful_metadata_value(company, style_default="Korea Development Bank"):
            rows.append(("Institution", company))
        if self._is_meaningful_metadata_value(sector, style_default="SECTOR"):
            rows.append(("Sector", sector))

        if recipient:
            rows.append(("Prepared For", recipient))

        if not rows:
            return

        table = self.doc.add_table(rows=len(rows), cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = STYLE.STYLE_TABLE_GRID

        for idx, (label, value) in enumerate(rows):
            label_cell = table.rows[idx].cells[0]
            value_cell = table.rows[idx].cells[1]

            TableStyler.set_cell_background(label_cell, STYLE.LIGHT_GRAY_HEX)

            label_para = label_cell.paragraphs[0]
            label_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            label_para.clear()
            label_run = label_para.add_run(label.upper())
            FontStyler.apply_run_style(
                label_run,
                font_name=STYLE.COVER_FONT,
                font_size=Pt(9.5),
                bold=True,
                color=STYLE.DARK_GRAY,
            )

            value_para = value_cell.paragraphs[0]
            value_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            value_para.clear()
            value_run = value_para.add_run(value)
            FontStyler.apply_run_style(
                value_run,
                font_name=STYLE.COVER_FONT,
                font_size=STYLE.BODY_SIZE,
                color=STYLE.DARK_GRAY,
            )

        TableStyler.set_table_borders(table)

    @staticmethod
    def _is_meaningful_metadata_value(value: str, style_default: str = "") -> bool:
        """Return True when a cover metadata value is worth showing to a user."""
        normalized = (value or "").strip()
        if not normalized:
            return False
        if style_default and normalized == style_default:
            return False
        if normalized.upper() in {"SECTOR", "N/A", "IB REPORT"}:
            return False
        return True

    def _add_spacer(self, lines: int):
        """Add empty paragraphs as vertical spacers."""
        for _ in range(lines):
            self.doc.add_paragraph()

    def _add_horizontal_rule(self, color_hex: str, size: str):
        """Add a clean horizontal rule using paragraph border."""
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), size)
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), color_hex)
        pBdr.append(bottom)
        insert_ordered(pPr, pBdr)


class TableRenderer:
    """Renders tables with type-specific styling"""

    _EMUS_PER_INCH = 914400
    _COLUMN_SAMPLE_SIZE = 3
    _TOKEN_SPLIT_RE = re.compile(r"\s+")
    _NUMERIC_LIKE_RE = re.compile(
        r"^\(?[+-]?\d[\d,]*(?:\.\d+)?\)?\s*"
        r"(?:%|bp|bps|x|배|원|천원|만원|백만원|억원|억|조|주|개|명|건)?$",
        re.IGNORECASE,
    )
    _PURE_NUMBER_RE = re.compile(r"^\(?[+-]?\d[\d,]*(?:\.\d+)?\)?$")
    # A dash placeholder ("nothing this period") says nothing about a column's kind.
    _DASH_PLACEHOLDER_RE = re.compile(r"[-‒–—―－]+")
    _MIN_COLUMN_WIDTH_INCHES = 0.65
    _MIN_TEXT_COLUMN_WIDTH_INCHES = 1.15
    _MAX_NUMERIC_COLUMN_WIDTH_INCHES = 1.35
    _MAX_TEXT_COLUMN_SHARE = 0.55

    def __init__(
        self,
        doc: DocxDocument,
        term_sheet: bool = False,
        house_style: TermSheetStyle | None = None,
    ):
        self.doc = doc
        # Profile-only layout; merge emission itself is shared by every profile.
        self.term_sheet = term_sheet
        self.house_style = house_style or TermSheetStyle()

    def render(self, table: Table):
        """Render a table"""
        if not table.rows or table.col_count == 0:
            return

        row_count = len(table.rows)
        col_count = table.col_count

        previous_geometry = self._begin_landscape(table)
        if self.term_sheet:
            try:
                self._render_term_sheet_table(table)
            finally:
                # A diagnosed failure must not leave later content in landscape.
                self._end_landscape(previous_geometry)
            return
        for text in [
            table.caption,
            " · ".join(
                v
                for v in [
                    "단위: " + table.unit if table.unit else "",
                    "기준일: " + table.as_of if table.as_of else "",
                ]
                if v
            ),
        ]:
            if text:
                paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                paragraph.paragraph_format.keep_with_next = True
                TextRenderer.render_runs(
                    paragraph, TextParser.parse_runs(text), font_size=STYLE.SMALL_SIZE
                )
        word_table = self.doc.add_table(rows=row_count, cols=col_count)
        word_table.style = STYLE.STYLE_TABLE_GRID
        column_kinds = self._infer_column_kinds(table)
        self._apply_column_widths(word_table, table, column_kinds)
        # Merge the empty, sized grid so spanned widths add up and no content
        # is concatenated; each owner cell is then filled exactly once.
        rectangles = self._merge_cells(word_table, table)
        covered = self._covered_cells(rectangles)

        # Render header rows
        header_rows = min(table.header_rows, len(table.rows))
        for h_idx in range(header_rows):
            self._render_header_row(word_table, table.rows[h_idx], col_count, covered, h_idx)

        # Render data rows based on table type
        for r_idx, row in enumerate(table.rows[header_rows:], header_rows):
            self._render_data_row(
                word_table,
                row,
                r_idx,
                col_count,
                table.table_type,
                column_kinds,
                table.column_types,
                table.alignments,
                covered,
                header_rows,
            )
        self._mirror_vertical_fills(word_table, rectangles)

        for h_idx in range(header_rows):
            word_table.rows[h_idx]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for word_row in word_table.rows:
            word_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        # Apply borders
        TableStyler.set_table_borders(word_table)

        if table.source:
            paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
            TextRenderer.render_runs(
                paragraph,
                TextParser.parse_runs("출처: " + table.source),
                font_size=STYLE.SMALL_SIZE,
            )
        if table.note:
            paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            TextRenderer.render_runs(
                paragraph, TextParser.parse_runs(table.note), font_size=STYLE.SMALL_SIZE
            )
        # Spacer paragraph after table
        self.doc.add_paragraph()
        self._end_landscape(previous_geometry)

    def _begin_landscape(self, table: Table) -> tuple[Any, int, int] | None:
        """Open a landscape section for a landscape table; return the prior geometry.

        The new section already starts on a new page, so a lone page break
        right before it (a cover's or TOC's) is redundant and is dropped. Word
        ignores such a break, but other viewers may show it as a blank page.
        """
        if not table.landscape:
            return None
        section = self.doc.sections[-1]
        assert section.page_width is not None and section.page_height is not None
        previous_geometry = (section.orientation, section.page_width, section.page_height)
        self._drop_trailing_page_break()
        section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Emu(max(previous_geometry[1:]))
        section.page_height = Emu(min(previous_geometry[1:]))
        return previous_geometry

    def _drop_trailing_page_break(self) -> None:
        """Remove the last body paragraph when it holds nothing but one page break.

        Several breaks in one paragraph ask for blank pages and are kept.
        """
        body = self.doc.element.body
        last = body[-1] if len(body) else None
        if last is not None and last.tag == qn("w:sectPr"):
            last = last.getprevious()
        if (
            last is not None
            and last.tag == qn("w:p")
            and len(last.xpath('./w:r/w:br[@w:type="page"]')) == 1
            and not last.xpath("./*[not(self::w:pPr or self::w:r)] | ./w:pPr/w:sectPr")
            and not last.xpath('./w:r/*[not(self::w:rPr or self::w:br[@w:type="page"])]')
        ):
            body.remove(last)

    def _end_landscape(self, previous_geometry: tuple[Any, int, int] | None) -> None:
        """Restore the page geometry that preceded a landscape table."""
        if previous_geometry:
            section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
            section.orientation = previous_geometry[0]
            section.page_width = Emu(previous_geometry[1])
            section.page_height = Emu(previous_geometry[2])

    def _render_term_sheet_table(self, table: Table) -> None:
        """Render a term-sheet table through the shared merge emission.

        Key-value tables (one or two label columns plus one content column) use
        the fixed label grid so their first vertical line is common; other tables
        keep content-based widths. Each owner cell is filled once with one
        paragraph per line, and a row stays unsplit when its estimated height is
        at most `ROW_SPLIT_THRESHOLD` lines (the header row always).

        Args:
            table: Parsed table; `label_columns` is inferred by the parser.

        Raises:
            ValueError: The section has no positive printable width. Unlike the
                legacy estimate there is no three-inch floor, so a term-sheet table
                never exceeds its section.
        """
        row_count, col_count = len(table.rows), table.col_count
        available = positive_printable_width(self.doc.sections[-1], "Term-sheet table")
        add_table_heading(
            self.doc, table.caption, table.unit, table.as_of, TextRenderer.render_runs, available
        )
        word_table = self.doc.add_table(rows=row_count, cols=col_count)
        word_table.style = STYLE.STYLE_TABLE_GRID
        column_kinds = self._infer_column_kinds(table)
        label_columns = (
            table.label_columns
            if table.label_columns is not None
            else min(1, max(col_count - 1, 0))
        )
        key_value = label_columns in (1, 2) and col_count == label_columns + 1
        widths = key_value_widths(label_columns, available) if key_value else None
        if key_value and widths is None:
            needed_mm, section_mm = sum(label_widths(label_columns)) / 36000, available / 36000
            self._report(
                f"Term-sheet label columns need {needed_mm:.1f} mm but the section is only "
                f"{section_mm:.1f} mm wide; column widths estimated from content"
            )
        if widths is None:
            widths = self._term_sheet_grid_widths(table, available, column_kinds)
        self._set_column_widths(word_table, widths)
        rectangles = self._merge_cells(word_table, table)
        covered = self._covered_cells(rectangles)
        extents = {(top, left): (bottom, right) for top, left, bottom, right in rectangles}
        apply_table_frame(
            word_table, sum(widths), open_sides=self.house_style.table_sides == "open"
        )

        line_counts = [0] * row_count
        for row_index, row in enumerate(table.rows):
            word_cells = word_table.rows[row_index].cells
            for column_index in range(col_count):
                if (row_index, column_index) in covered:
                    continue
                bottom, right = extents.get((row_index, column_index), (row_index, column_index))
                cell_data = (
                    row.cells[column_index]
                    if column_index < len(row.cells)
                    else TableCell(content="")
                )
                role = self._term_sheet_role(
                    row_index,
                    column_index,
                    right,
                    label_columns,
                    key_value,
                    table.header_rows,
                )
                texts, size = self._fill_term_sheet_cell(
                    word_cells[column_index],
                    cell_data,
                    role,
                    row_index,
                    column_index,
                    table,
                    column_kinds,
                )
                if bottom == row_index:  # multi-row owners are excluded from row estimates
                    lines = estimate_cell_lines(
                        texts,
                        sum(widths[column_index : right + 1]),
                        size,
                        markers=role == "content",
                    )
                    line_counts[row_index] = max(line_counts[row_index], lines)
        self._mirror_vertical_fills(word_table, rectangles)
        for row_index, word_row in enumerate(word_table.rows):
            header = row_index < table.header_rows
            set_row_pagination(
                word_row,
                keep_together=header or line_counts[row_index] <= ROW_SPLIT_THRESHOLD,
                repeat_header=header,
            )
        if table.note:
            add_table_note(self.doc, table.note, TextRenderer.render_runs)
        if table.source:
            add_table_source(self.doc, table.source, TextRenderer.render_runs)
        add_table_spacer(self.doc)

    def _term_sheet_grid_widths(
        self, table: Table, available: int, column_kinds: list[str]
    ) -> list[int]:
        """Widths (EMU) for a term-sheet grid table, filling the section width.

        When the columns' content fits side by side on single lines, each column
        gets its natural width (see `_term_sheet_natural_widths`) plus an equal
        share of the spare width, so short amount columns are not starved while
        date or label columns balloon. Otherwise widths come from the content
        estimate, whose minimums keep header words on one line: each
        single-column header cell requires room for its longest unbreakable token
        in the bold header size plus the cell margins, and merged header cells
        impose no minimum (see `_required_minimums` for the graceful fallback).
        """
        natural = self._term_sheet_natural_widths(table, column_kinds)
        if natural is not None and sum(natural) <= available:
            spare = (available - sum(natural)) // table.col_count
            fitted = [width + spare for width in natural]
            fitted[-1] += available - sum(fitted)
            return fitted
        required = [0.0] * table.col_count
        for header in table.rows[: table.header_rows]:
            for index in range(min(len(header.cells), table.col_count)):
                if self._is_span_cell(header, index):
                    continue
                cell = header.cells[index]
                text = line_text(cell.runs or TextParser.parse_runs_plain(cell.content))
                width = header_token_width(text, STYLE.TABLE_HEADER_SIZE) / self._EMUS_PER_INCH
                required[index] = max(required[index], width)
        widths = self._estimate_column_widths(
            table, available / self._EMUS_PER_INCH, column_kinds, required
        )
        return [int(Inches(width)) for width in widths]

    def _term_sheet_natural_widths(self, table: Table, column_kinds: list[str]) -> list[int] | None:
        """Per-column widths (EMU) that keep every cell line on one line.

        Header lines are measured bold in the header size and body lines as
        displayed (formatted amounts, marker indents). A cell merged across
        columns spreads any extra width it needs evenly over them, and every
        column keeps at least the estimator's numeric minimum.

        Args:
            table: Term-sheet grid table.
            column_kinds: Inferred column kinds (unused by the measure itself,
                kept for symmetry with the content estimate).

        Returns:
            One width per column, or None when a cell holds an image, whose
            width text measurement cannot know.
        """
        natural = [0] * table.col_count
        spanning: list[tuple[int, int, int]] = []
        for row_index, row in enumerate(table.rows):
            header = row_index < table.header_rows
            for index in range(min(len(row.cells), table.col_count)):
                cell = row.cells[index]
                if cell.merge is not None:
                    continue
                right = index
                while (
                    right + 1 < min(len(row.cells), table.col_count)
                    and row.cells[right + 1].merge == "left"
                ):
                    right += 1
                if header:
                    runs = cell.runs or TextParser.parse_runs_plain(cell.content)
                    size, bold, markers = STYLE.TABLE_HEADER_SIZE, True, False
                else:
                    role = table.column_types[index] if table.column_types else None
                    content, runs = self._display_content(cell, role, table.table_type)
                    runs = runs or TextParser.parse_runs_plain(content)
                    size, bold, markers = STYLE.TABLE_BODY_SIZE, False, True
                if any(run.image is not None for run in runs):
                    return None
                width = natural_cell_width(
                    [line_text(line) for line in split_run_lines(runs)],
                    size,
                    bold=bold,
                    markers=markers,
                )
                if right == index:
                    natural[index] = max(natural[index], width)
                else:
                    spanning.append((index, right, width))
        floor = int(Inches(self._MIN_COLUMN_WIDTH_INCHES))
        natural = [max(width, floor) for width in natural]
        for left, right, width in spanning:
            missing = width - sum(natural[left : right + 1])
            if missing > 0:
                share = -(-missing // (right - left + 1))
                for index in range(left, right + 1):
                    natural[index] += share
        return natural

    @staticmethod
    def _term_sheet_role(
        row_index: int,
        column_index: int,
        right: int,
        label_columns: int,
        key_value: bool,
        header_rows: int = 1,
    ) -> str:
        """Classify an owner cell by where it starts and how far it spans.

        A cell that starts in a second-tier label column but reaches the content
        columns is content; a label in the first column stays a label however
        far it spans (for example a totals row).
        """
        if row_index < header_rows:
            return "header"
        if column_index >= label_columns or (column_index > 0 and right >= label_columns):
            return "content"
        if not key_value:
            return "grid-label"
        return "sublabel" if column_index == 1 else "label"

    def _fill_term_sheet_cell(
        self,
        word_cell,
        cell_data: TableCell,
        role: str,
        row_index: int,
        column_index: int,
        table: Table,
        column_kinds: list[str],
    ) -> tuple[list[str], Pt]:
        """Fill and style one owner cell exactly once, one paragraph per line.

        Header: header fill, bold accent, centred. Key-value label: label fill,
        bold black, centred; second label tier: unshaded, regular muted, centred.
        Grid-table labels: label fill, regular black, centred. Content cells keep
        column kinds, alignment markers and numeric roles, and indent marker lines.

        Returns:
            The cell's line texts and base text size, for row-split estimation.
        """
        column_role = table.column_types[column_index] if table.column_types else None
        fill: str | None = None
        color: RGBColor | None = None
        alignment = WD_ALIGN_PARAGRAPH.CENTER
        if role == "header":
            runs = [
                replace(run, bold=True, color_hex=None)
                for run in cell_data.runs or TextParser.parse_runs_plain(cell_data.content)
            ]
            if self.house_style.table_header == "dark":
                fill, color = STYLE.NAVY_HEX, RGBColor(0xFF, 0xFF, 0xFF)
            else:
                fill, color = STYLE.TABLE_HEADER_BG, STYLE.NAVY
            font_name, size = STYLE.HEADING_FONT, STYLE.TABLE_HEADER_SIZE
        else:
            content, runs = self._display_content(cell_data, column_role, table.table_type)
            runs = runs or TextParser.parse_runs_plain(content)
            font_name, size = STYLE.BODY_FONT, STYLE.TABLE_BODY_SIZE
            if role == "label":
                fill, color = STYLE.TS_LABEL_BG_HEX, _BLACK
                runs = [replace(run, bold=True) for run in runs]
            elif role == "grid-label":
                fill, color = STYLE.TS_LABEL_BG_HEX, _BLACK
            elif role == "sublabel":
                color = RGBColor.from_string(STYLE.TS_MUTED_HEX)
            else:
                alignment = self._cell_alignment(column_index, column_kinds, table.alignments)
                body_index = row_index - table.header_rows
                if STYLE.TABLE_ZEBRA and body_index % 2 == 0 and not cell_data.is_base_case:
                    fill = STYLE.LIGHT_GRAY_HEX
        if fill:
            set_cell_fill(word_cell._tc, fill)
        word_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        texts: list[str] = []
        for index, line in enumerate(split_run_lines(runs)):
            paragraph = word_cell.paragraphs[0] if index == 0 else word_cell.add_paragraph()
            configure_cell_paragraph(paragraph)
            paragraph.alignment = alignment
            if role == "header":
                # Never strand the header row alone at the foot of a page.
                paragraph.paragraph_format.keep_with_next = True
            text = line_text(line)
            line_size = apply_marker_layout(paragraph, text) if role == "content" else None
            TextRenderer.render_runs(
                paragraph,
                line,
                default_color=color,
                font_name=font_name,
                font_size=line_size or size,
            )
            texts.append(text)
        if role != "header" and not (
            table.table_type == TableType.FINANCIAL and column_role in {"text", "code", "date"}
        ):
            self._apply_type_styling(
                word_cell, cell_data, row_index, table.table_type, fill=self._ordered_fill
            )
        return texts, size

    @staticmethod
    def _ordered_fill(cell, hex_color: str) -> None:
        """Term-sheet cell fill placed in `w:tcPr` schema order (before `w:vAlign`)."""
        set_cell_fill(cell._tc, hex_color)

    @staticmethod
    def _set_column_widths(word_table, widths: list[int]) -> None:
        """Set exact grid and cell widths (EMU) on a table that is not merged yet."""
        word_table.autofit = False
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        for index, width in enumerate(widths):
            word_table.columns[index].width = Emu(width)
            for cell in word_table.columns[index].cells:
                cell.width = Emu(width)

    def _merge_cells(self, word_table, table: Table) -> list[tuple[int, int, int, int]]:
        """Merge validated span rectangles of an empty table whose widths are set.

        python-docx concatenates the paragraphs of merged cells that have content
        and adds the `tcW` of horizontally merged cells, so merging must happen
        after sizing and before any cell is filled.

        Args:
            word_table: Freshly created Word table with grid and cell widths applied.
            table: Parsed table whose `TableCell.merge` directives define the spans.

        Returns:
            Inclusive zero-based (top, left, bottom, right) rectangles that were merged.
        """
        rectangles, problems = self._span_rectangles(table)
        for problem in problems:
            self._report(problem)
        for top, left, bottom, right in rectangles:
            word_table.cell(top, left).merge(word_table.cell(bottom, right))
        return rectangles

    def _report(self, message: str) -> None:
        """Log a rendering diagnostic and record it for strict validation."""
        logger.warning("%s", message)
        errors = getattr(self.doc.part, "_ib_render_errors", None)
        if errors is not None:
            errors.append(message)

    @staticmethod
    def _span_rectangles(table: Table) -> tuple[list[tuple[int, int, int, int]], list[str]]:
        """Group span directives into rectangles, rejecting malformed groups.

        Parsed tables are already validated by `TableSpanResolver`; this guard keeps
        hand-built models from producing ragged or overlapping Word merges. Header
        and body are classified as the renderer draws them (the first
        `Table.header_rows` rows repeat as the header), not by
        `TableRow.is_header`, which hand-built rows may omit.

        Args:
            table: Table whose cells may carry "up"/"left" merge directives.

        Returns:
            Valid rectangles and one diagnostic per rejected group, whose cells
            are then rendered unmerged with their own content.
        """
        roots: dict[tuple[int, int], tuple[int, int]] = {}
        groups: dict[tuple[int, int], list[tuple[int, int]]] = {}
        outside: set[tuple[int, int]] = set()
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells[: table.col_count]):
                position = root = (row_index, column_index)
                if cell.merge in ("up", "left"):
                    target = (
                        (row_index - 1, column_index)
                        if cell.merge == "up"
                        else (row_index, column_index - 1)
                    )
                    if target in roots:
                        root = roots[target]
                    else:
                        outside.add(position)
                roots[position] = root
                groups.setdefault(root, []).append(position)
        rectangles: list[tuple[int, int, int, int]] = []
        problems: list[str] = []
        for (top, left), members in groups.items():
            if all(table.rows[row].cells[column].merge is None for row, column in members):
                continue
            bottom = max(row for row, _ in members)
            right = max(column for _, column in members)
            if (
                not outside.intersection(members)
                and min(row for row, _ in members) == top
                and min(column for _, column in members) == left
                and table.rows[top].cells[left].merge is None
                and len(members) == (bottom - top + 1) * (right - left + 1)
                and not (top < table.header_rows <= bottom)
            ):
                rectangles.append((top, left, bottom, right))
            else:
                problems.append(
                    f"Invalid table span group at row {top + 1}, column {left + 1}; "
                    "cells rendered unmerged"
                )
        return rectangles, problems

    @staticmethod
    def _covered_cells(rectangles: list[tuple[int, int, int, int]]) -> set[tuple[int, int]]:
        """Return grid positions represented by another cell's merged owner."""
        return {
            (row, column)
            for top, left, bottom, right in rectangles
            for row in range(top, bottom + 1)
            for column in range(left, right + 1)
            if (row, column) != (top, left)
        }

    @staticmethod
    def _mirror_vertical_fills(word_table, rectangles: list[tuple[int, int, int, int]]) -> None:
        """Repeat each vertical owner's fill on its continuation `w:tc` elements.

        Continuation cells keep their own properties in OOXML, so the owner's fill
        is repeated to keep the whole merged area shaded in every consumer.
        """
        for top, left, bottom, _ in rectangles:
            if bottom == top:
                continue
            owner = word_table.rows[top]._tr.tc_at_grid_offset(left)
            shading = owner.tcPr.find(qn("w:shd")) if owner.tcPr is not None else None
            if shading is None or not shading.get(qn("w:fill")):
                continue
            for row_index in range(top + 1, bottom + 1):
                continuation = word_table.rows[row_index]._tr.tc_at_grid_offset(left)
                set_cell_fill(continuation, shading.get(qn("w:fill")))

    @staticmethod
    def _is_span_cell(row: TableRow, col_idx: int) -> bool:
        """True for covered cells and owners spanning columns (excluded from sizing)."""
        cells = row.cells
        return cells[col_idx].merge is not None or (
            col_idx + 1 < len(cells) and cells[col_idx + 1].merge == "left"
        )

    def _apply_column_widths(self, word_table, table: Table, column_kinds: list[str]) -> None:
        """Apply content-aware column widths for more readable report tables."""
        widths = self._estimate_column_widths(
            table,
            self._get_available_table_width_inches(),
            column_kinds=column_kinds,
        )
        if not widths:
            return

        word_table.autofit = False
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT

        for c_idx, width_inches in enumerate(widths):
            width = Inches(width_inches)
            word_table.columns[c_idx].width = width
            for cell in word_table.columns[c_idx].cells:
                cell.width = width

    def _estimate_column_widths(
        self,
        table: Table,
        available_width_inches: float,
        column_kinds: list[str] | None = None,
        required_widths: list[float] | None = None,
    ) -> list[float]:
        """Estimate table column widths from content density and column semantics.

        Args:
            required_widths: Optional per-column widths (inches) that should not be
                undercut, such as a header's longest word; see `_required_minimums`.
        """
        if table.col_count <= 0:
            return []

        resolved_column_kinds = column_kinds or self._infer_column_kinds(table)
        column_infos = [
            self._build_column_info(table, col_idx, resolved_column_kinds[col_idx])
            for col_idx in range(table.col_count)
        ]
        min_widths = [
            self._minimum_column_width(column_info["kind"]) for column_info in column_infos
        ]
        max_widths = [
            self._maximum_column_width(column_info["kind"], available_width_inches)
            for column_info in column_infos
        ]
        if required_widths is not None:
            min_widths = self._required_minimums(
                min_widths, required_widths, available_width_inches
            )
            max_widths = [max(high, low) for high, low in zip(max_widths, min_widths, strict=False)]
        preferred = [column_info["score"] for column_info in column_infos]
        widths = self._fit_widths_to_available_space(
            preferred,
            min_widths,
            max_widths,
            available_width_inches,
        )

        return widths

    @classmethod
    def _required_minimums(
        cls, minimums: list[float], required: list[float], available_width_inches: float
    ) -> list[float]:
        """Raise column minimums to required widths when the section can hold them.

        First the heuristic minimums are raised; if that overflows, the required
        widths are kept over the numeric floor only (relaxing the heuristic text
        minimum); if even that overflows, the heuristic minimums are returned
        unchanged, so the table still fits without an error.
        """
        raised = [max(low, need) for low, need in zip(minimums, required, strict=False)]
        if sum(raised) <= available_width_inches:
            return raised
        relaxed = [max(cls._MIN_COLUMN_WIDTH_INCHES, need) for need in required]
        if sum(relaxed) <= available_width_inches:
            return relaxed
        return minimums

    def _build_column_info(self, table: Table, col_idx: int, column_kind: str) -> dict[str, Any]:
        """Summarize the content profile of a single column."""
        texts = []

        for row in table.rows:
            if col_idx >= len(row.cells) or self._is_span_cell(row, col_idx):
                continue

            cell = row.cells[col_idx]
            text = self._cell_display_text(cell, table.table_type).strip()
            if not text:
                continue

            texts.append(text)

        representative_length = self._representative_length(texts)
        longest_token = self._longest_token_length(texts)
        is_numeric_column = column_kind == "numeric"
        score = self._column_score(is_numeric_column, representative_length, longest_token)

        return {
            "kind": column_kind,
            "score": score,
        }

    def _infer_column_kinds(self, table: Table) -> list[str]:
        if table.column_types:
            return [
                "text" if role in {"text", "code", "date"} else "numeric"
                for role in table.column_types
            ]
        """Infer each column's semantic type from the first few data cells."""
        return [self._infer_column_kind(table, col_idx) for col_idx in range(table.col_count)]

    def _infer_column_kind(self, table: Table, col_idx: int) -> str:
        """Classify a column as text or numeric using the first 2-3 body cells.

        Empty cells and dash placeholders (`-`, `–`) are not samples.
        """
        sample_texts = []

        for row in table.rows[table.header_rows :]:
            if col_idx >= len(row.cells) or self._is_span_cell(row, col_idx):
                continue

            text = self._cell_display_text(row.cells[col_idx], table.table_type).strip()
            if not text or self._DASH_PLACEHOLDER_RE.fullmatch(text):
                continue

            sample_texts.append(text)
            if len(sample_texts) >= self._COLUMN_SAMPLE_SIZE:
                break

        if not sample_texts:
            return "text"

        numeric_like_count = sum(1 for text in sample_texts if self._is_numeric_like(text))
        return "numeric" if numeric_like_count >= (len(sample_texts) + 1) // 2 else "text"

    @classmethod
    def _is_numeric_like(cls, text: str) -> bool:
        """Return True for financial/count strings that should align right."""
        normalized = text.strip()
        if not normalized or not any(char.isdigit() for char in normalized):
            return False
        return bool(cls._NUMERIC_LIKE_RE.match(normalized))

    @classmethod
    def _representative_length(cls, texts: list[str]) -> int:
        """Return a robust content length that ignores a few extreme outliers."""
        if not texts:
            return 0

        lengths = sorted(len(text) for text in texts)
        percentile_index = int(round((len(lengths) - 1) * 0.75))
        return lengths[percentile_index]

    @classmethod
    def _longest_token_length(cls, texts: list[str]) -> int:
        """Return the length of the longest unbroken token in the sample."""
        longest = 0
        for text in texts:
            tokens = [token for token in cls._TOKEN_SPLIT_RE.split(text) if token]
            if not tokens:
                continue
            longest = max(longest, max(len(token) for token in tokens))
        return longest

    @classmethod
    def _column_score(
        cls,
        is_numeric_column: bool,
        representative_length: int,
        longest_token: int,
    ) -> float:
        """Assign a width score to a column based on how hard it is to wrap cleanly."""
        if is_numeric_column:
            return 1.0 + min(representative_length, 12) * 0.05

        return 1.8 + min(representative_length, 48) * 0.06 + min(longest_token, 24) * 0.03

    @classmethod
    def _minimum_column_width(cls, kind_marker: str) -> float:
        """Return a minimum readable width for a column kind."""
        if kind_marker == "numeric":
            return cls._MIN_COLUMN_WIDTH_INCHES
        return cls._MIN_TEXT_COLUMN_WIDTH_INCHES

    @classmethod
    def _maximum_column_width(cls, kind_marker: str, available_width_inches: float) -> float:
        """Return an upper width bound for a column kind."""
        if kind_marker == "numeric":
            return cls._MAX_NUMERIC_COLUMN_WIDTH_INCHES
        return max(
            cls._MIN_TEXT_COLUMN_WIDTH_INCHES,
            available_width_inches * cls._MAX_TEXT_COLUMN_SHARE,
        )

    @classmethod
    def _fit_widths_to_available_space(
        cls,
        preferred: list[float],
        min_widths: list[float],
        max_widths: list[float],
        available_width_inches: float,
    ) -> list[float]:
        """Fit preferred widths into the available width with min/max guards."""
        col_count = len(preferred)
        if col_count == 0:
            return []

        min_total = sum(min_widths)
        if min_total >= available_width_inches:
            scale = available_width_inches / min_total
            return [width * scale for width in min_widths]

        total_preferred = sum(preferred) or float(col_count)
        widths = [
            max(min_widths[idx], available_width_inches * preferred[idx] / total_preferred)
            for idx in range(col_count)
        ]

        widths = [min(widths[idx], max_widths[idx]) for idx in range(col_count)]
        widths = cls._rebalance_widths(
            widths, min_widths, max_widths, preferred, available_width_inches
        )
        return widths

    @classmethod
    def _rebalance_widths(
        cls,
        widths: list[float],
        min_widths: list[float],
        max_widths: list[float],
        preferred: list[float],
        available_width_inches: float,
    ) -> list[float]:
        """Rebalance widths after clamping so the table uses the full line width."""
        for _ in range(8):
            total_width = sum(widths)
            gap = available_width_inches - total_width

            if abs(gap) < 0.01:
                break

            if gap > 0:
                growable = [
                    idx for idx, width in enumerate(widths) if width < max_widths[idx] - 0.01
                ]
                if not growable:
                    break
                grow_weight = sum(preferred[idx] for idx in growable) or float(len(growable))
                for idx in growable:
                    share = preferred[idx] / grow_weight if grow_weight else 1.0 / len(growable)
                    widths[idx] = min(max_widths[idx], widths[idx] + gap * share)
            else:
                shrinkable = [
                    idx for idx, width in enumerate(widths) if width > min_widths[idx] + 0.01
                ]
                if not shrinkable:
                    break
                shrink_capacity = sum(widths[idx] - min_widths[idx] for idx in shrinkable)
                if shrink_capacity <= 0:
                    break
                overflow = -gap
                for idx in shrinkable:
                    capacity = widths[idx] - min_widths[idx]
                    share = capacity / shrink_capacity
                    widths[idx] = max(min_widths[idx], widths[idx] - overflow * share)

        return widths

    def _get_available_table_width_inches(self) -> float:
        """Return the horizontal space available for body tables."""
        return float(self._get_available_table_width_emu()) / float(self._EMUS_PER_INCH)

    def _get_available_table_width_emu(self) -> int:
        """Return the printable width of the current (last) section in EMU."""
        if not self.doc.sections:
            return int(Inches(6.0))

        section = self.doc.sections[-1]
        if (
            section.page_width is None
            or section.left_margin is None
            or section.right_margin is None
        ):
            return int(Inches(6.0))
        available_width = (
            int(section.page_width) - int(section.left_margin) - int(section.right_margin)
        )
        return max(available_width, int(Inches(3.0)))

    def _cell_display_text(self, cell_data: TableCell, table_type: TableType) -> str:
        """Return display text used for width estimation and column-kind inference."""
        shown = term_cell_text(cell_data)
        if shown is not None:
            return shown  # term values are shown exactly as written in `terms:`
        if table_type == TableType.FINANCIAL and cell_data.is_numeric:
            return self._format_financial_number(cell_data.content)
        return cell_data.content

    def _render_header_row(
        self,
        word_table,
        row: TableRow,
        col_count: int,
        covered: Set[tuple[int, int]] = frozenset(),
        row_index: int = 0,
    ):
        """Render one header row with Navy background, skipping merged-away cells."""
        word_cells = word_table.rows[row_index].cells

        for c_idx, cell_data in enumerate(row.cells):
            if c_idx >= col_count:
                break
            if (row_index, c_idx) in covered:
                continue

            cell = word_cells[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            TableStyler.set_cell_background(cell, STYLE.TABLE_HEADER_BG)

            # Clear default paragraph and add styled run
            p = cell.paragraphs[0]
            self._configure_cell_paragraph(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.clear()

            # Parse header text: strip markdown markers but preserve structure
            # Headers are always bold+white, so we parse for content only
            if cell_data.runs:
                # Use structured runs but override style for header
                TextRenderer.render_runs(
                    p,
                    [replace(run, bold=True, color_hex=None) for run in cell_data.runs],
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.TABLE_HEADER_SIZE,
                    default_color=STYLE.TABLE_HEADER_COLOR,
                )
            else:
                clean_text = cell_data.content.replace("**", "").strip()
                run = p.add_run(clean_text)
                FontStyler.apply_run_style(
                    run,
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.TABLE_HEADER_SIZE,
                    bold=True,
                    color=STYLE.TABLE_HEADER_COLOR,
                )

    def _render_data_row(
        self,
        word_table,
        row: TableRow,
        row_idx: int,
        col_count: int,
        table_type: TableType,
        column_kinds: list[str],
        column_types: list[str] | None = None,
        alignments: list[str] | None = None,
        covered: Set[tuple[int, int]] = frozenset(),
        header_rows: int = 1,
    ):
        """Render a data row with type-specific styling, skipping merged-away cells.

        Zebra shading counts body rows only, so it starts on the first body row
        whatever the number of header rows.
        """
        word_cells = word_table.rows[row_idx].cells

        for c_idx, cell_data in enumerate(row.cells):
            if c_idx >= col_count:
                break
            if (row_idx, c_idx) in covered:
                continue

            cell = word_cells[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            self._configure_cell_paragraph(p)

            # ── Alignment ───────────────────────────────────────────────────
            p.alignment = self._cell_alignment(c_idx, column_kinds, alignments)

            # ── Format content (financial number formatting) ───────────────
            role = column_types[c_idx] if column_types else None
            display_content, display_runs = self._display_content(cell_data, role, table_type)

            # ── Render content ──────────────────────────────────────────────
            if cell_data.runs:
                # Use structured runs for full formatting fidelity
                TextRenderer.render_runs(
                    p,
                    display_runs,
                    font_name=STYLE.BODY_FONT,
                    font_size=STYLE.TABLE_BODY_SIZE,
                )
            else:
                TextRenderer.render_text_with_formatting(
                    p,
                    display_content,
                    font_name=STYLE.BODY_FONT,
                    font_size=STYLE.TABLE_BODY_SIZE,
                )

            # ── Type-specific styling ───────────────────────────────────────
            if not (table_type == TableType.FINANCIAL and role in {"text", "code", "date"}):
                self._apply_type_styling(cell, cell_data, row_idx, table_type)

            # ── Alternating row colors (unless special styling applied) ─────
            if (
                STYLE.TABLE_ZEBRA
                and (row_idx - header_rows) % 2 == 0
                and not cell_data.is_base_case
            ):
                TableStyler.set_cell_background(cell, STYLE.LIGHT_GRAY_HEX)

    @staticmethod
    def _cell_alignment(
        c_idx: int, column_kinds: list[str], alignments: list[str] | None
    ) -> WD_ALIGN_PARAGRAPH:
        """Numeric columns align right; Markdown alignment markers take precedence."""
        if alignments and c_idx < len(alignments) and alignments[c_idx] in {"center", "right"}:
            return (
                WD_ALIGN_PARAGRAPH.CENTER
                if alignments[c_idx] == "center"
                else WD_ALIGN_PARAGRAPH.RIGHT
            )
        if c_idx < len(column_kinds) and column_kinds[c_idx] == "numeric":
            return WD_ALIGN_PARAGRAPH.RIGHT
        return WD_ALIGN_PARAGRAPH.LEFT

    def _display_content(
        self, cell_data: TableCell, role: str | None, table_type: TableType
    ) -> tuple[str, list[TextRun]]:
        """Return display text and runs, formatting amounts only for numeric roles."""
        format_numbers = role not in {None, "text", "code", "date"} or (
            not role and table_type == TableType.FINANCIAL and cell_data.is_numeric
        )
        display_content = cell_data.content
        display_runs = cell_data.runs
        if format_numbers:
            display_content = self._format_numeric_text(display_content, role)
            # Semantic runs are independent of the amount. In particular, a
            # footnote's displayed number must never enter numeric formatting,
            # and a term value keeps the exact text written in `terms:`.
            display_runs = [
                run
                if (
                    run.footnote_id is not None
                    or run.is_latex
                    or run.superscript
                    or run.subscript
                    or run.term_key is not None
                    or run.code
                )
                else replace(run, text=self._format_numeric_text(run.text, role))
                for run in display_runs
            ]
        return display_content, display_runs

    @classmethod
    def _format_numeric_text(cls, text: str, role: str | None) -> str:
        """Format a numeric text run without changing its semantic metadata.

        Args:
            text: One text run, excluding footnote and equation runs.
            role: Explicit column presentation role, if configured.

        Returns:
            Formatted number with its original scale, or unchanged nonnumeric text.
        """
        if not cls._NUMERIC_LIKE_RE.fullmatch(text.strip()):
            return text
        formatted = cls._format_financial_number(text)
        suffix = {"percent": "%", "bps": " bps", "multiple": "x"}.get(role or "", "")
        if suffix and cls._PURE_NUMBER_RE.fullmatch(text.strip()):
            formatted += suffix
        leading = text[: len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()) :]
        return leading + formatted + trailing

    @staticmethod
    def _configure_cell_paragraph(paragraph) -> None:
        """Normalize paragraph spacing inside table cells for a tighter, cleaner grid."""
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0

    @staticmethod
    def _format_financial_number(text: str) -> str:
        """
        Format financial numbers with thousand separators.

        Handles:
            - Plain numbers: 1234567 → 1,234,567
            - Percentages: 12.5% → 12.5%
            - Negative parentheses: (1234) → (1,234)
            - Korean 억/조 units preserved
            - Already formatted numbers: 1,234 → 1,234

        Args:
            text: Cell text content

        Returns:
            Formatted text with thousand separators
        """
        import re

        text = text.strip()

        # Skip if empty or non-numeric looking
        if not text or not any(c.isdigit() for c in text):
            return text

        # Skip if already has thousand separators
        if "," in text and re.search(r"\d{1,3}(,\d{3})+", text):
            return text

        # Handle negative in parentheses: (1234567) → (1,234,567)
        paren_match = re.match(r"^\((\d+(?:\.\d+)?)\)(.*)$", text)
        if paren_match:
            num_str = paren_match.group(1)
            suffix = paren_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"({formatted}){suffix}"

        # Handle negative with minus: -1234567 → -1,234,567
        neg_match = re.match(r"^-(\d+(?:\.\d+)?)(.*)$", text)
        if neg_match:
            num_str = neg_match.group(1)
            suffix = neg_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"-{formatted}{suffix}"

        # Handle positive numbers with optional suffix (%, 억, 조, 원, etc.)
        pos_match = re.match(r"^(\d+(?:\.\d+)?)(.*)$", text)
        if pos_match:
            num_str = pos_match.group(1)
            suffix = pos_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"{formatted}{suffix}"

        return text

    @staticmethod
    def _add_thousand_sep(num_str: str) -> str:
        """Add thousand separators to a numeric string."""
        if "." in num_str:
            integer_part, decimal_part = num_str.split(".", 1)
            formatted_int = f"{int(integer_part):,}"
            return f"{formatted_int}.{decimal_part}"
        else:
            return f"{int(num_str):,}"

    def _apply_type_styling(
        self,
        cell,
        cell_data: TableCell,
        row_idx: int,
        table_type: TableType,
        fill: Callable[[Any, str], None] = TableStyler.set_cell_background,
    ):
        """Apply table-type specific styling to every paragraph of one owner cell.

        Args:
            fill: Cell background setter. Legacy tables keep the historical append;
                term-sheet tables pass a schema-ordered setter.
        """
        runs = [run for paragraph in cell.paragraphs for run in paragraph.runs]
        if table_type == TableType.FINANCIAL:
            if cell_data.is_negative:
                for run in runs:
                    run.font.color.rgb = STYLE.RED

        elif table_type == TableType.BEP_SENSITIVITY:
            if cell_data.is_base_case:
                fill(cell, STYLE.YELLOW_HEX)
                for run in runs:
                    run.font.bold = True

        elif table_type == TableType.RISK_MATRIX:
            if cell_data.risk_level:
                color_map = {
                    "high": STYLE.RED,
                    "medium": STYLE.ORANGE,
                    "low": STYLE.GREEN,
                }
                color = color_map.get(cell_data.risk_level)
                if color:
                    for run in runs:
                        run.font.color.rgb = color
                        run.font.bold = True


class DisclaimerRenderer:
    """Renders disclaimer page"""

    _SECTIONS = [
        (
            "면책 조항",
            "본 자료는 해당 문서에 최대한 정확하고 완전한 정보를 담고자 노력하였으나, "
            "오류와 중요정보의 누락이 있을 수 있으며, 정보의 정확성, 완전성 및 적정성을 "
            "보장하지 않습니다. 이 문서는 고객의 이해를 돕기 위하여 작성된 설명자료에 "
            "불과하므로, 고객은 각자의 책임으로 개별 계약서나 공시된 정보를 통하여 "
            "거래의 내용을 숙지하여야 합니다. 이 문서는 확정적인 거래조건을 구성하지 "
            "않으며 법적인 책임을 위한 근거자료로 사용될 수 없습니다.",
        ),
        (
            "저작권 및 비밀유지",
            "본 자료는 당행의 저작물로서 모든 저작권은 당행에게 있으며, 당행의 동의 없이 "
            "어떠한 경우에도 어떠한 형태로든 복제, 배포, 전송, 변경, 대여할 수 없습니다. "
            "당행의 요청 시에 즉시 반환, 파기하여 주시기 바랍니다. 본 자료는 상기 제한에 "
            "대하여 동의하는 조건으로 제공되며 동의하지 않으시는 경우에는 즉시 파기하여 "
            "주시기 바랍니다.",
        ),
        (
            "조건부 제공",
            "본 제안서의 내용은 현재의 시장상황 및 발행구조에 대한 기초정보에 근거한 것으로 "
            "유동화대상 자산 등 구조에 대한 변경이나 기타 중대한 사유 발생시 변경될 수 있으며, "
            "당행의 내부 여신심의위원회 승인을 조건으로 합니다.",
        ),
    ]

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, company: str):
        """Render standard IB disclaimer page"""
        self.doc.add_page_break()
        self.doc.add_heading("면책 조항", level=1)

        for title, content in self._SECTIONS:
            title_para = self.doc.add_paragraph()
            title_run = title_para.add_run(title)
            FontStyler.apply_run_style(
                title_run,
                font_name=STYLE.HEADING_FONT,
                font_size=STYLE.SMALL_SIZE,
                bold=True,
                color=STYLE.MEDIUM_GRAY,
            )
            title_para.paragraph_format.space_before = Pt(4)
            title_para.paragraph_format.space_after = Pt(3)

            for line in self._split_content_lines(content):
                content_para = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                content_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                content_para.paragraph_format.line_spacing = 1.2
                content_para.paragraph_format.space_after = Pt(6)

                content_run = content_para.add_run(line)
                FontStyler.apply_run_style(
                    content_run,
                    font_name=STYLE.BODY_FONT,
                    font_size=STYLE.SMALL_SIZE,
                    color=STYLE.MEDIUM_GRAY,
                )

        # Copyright
        self.doc.add_paragraph()

        copy_para = self.doc.add_paragraph()
        copy_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        year = time.strftime("%Y")
        copy_run = copy_para.add_run(f"(c) {year} {company}. All rights reserved.")
        FontStyler.apply_run_style(
            copy_run,
            font_size=STYLE.SMALL_SIZE,
            italic=True,
            color=STYLE.MEDIUM_GRAY,
        )

    @staticmethod
    def _split_content_lines(content: str) -> list[str]:
        """Split disclaimer content into non-empty logical paragraphs."""
        lines = [line.strip() for line in content.split("\n")]
        return [line for line in lines if line]


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT RENDERER (ORCHESTRATOR)
# ═══════════════════════════════════════════════════════════════════════════════


@contextmanager
def _metadata_text(area: str) -> Iterator[None]:
    """Reject metadata of the wrong type at the composer that needs it as text.

    Nothing is checked ahead of use, so every value that renders today keeps
    rendering. Only the TypeError or AttributeError a value of the wrong type
    causes inside the named composer is reported, with that error as the cause.

    Args:
        area: The part of the document being composed, named in the message.

    Yields:
        Control to the composer.
    """
    try:
        yield
    except (TypeError, AttributeError) as error:
        raise StrictValidationError(
            "", (f"Metadata for the {area} has a value of the wrong type: {error}",)
        ) from error


class IBDocumentRenderer:
    """Main renderer that orchestrates all component renderers"""

    _TREE_PREFIX_RE = re.compile(r"^([\s│├└─]+)(.*)$")
    _TREE_VALUE_RE = re.compile(r"^(.*?)(\d+(?:\.\d+)?%|실질지배)(\s+──\s+)(.*)$")
    _SEMANTIC_BOOKMARK_PREFIX = "_ibrep_"
    _SEMANTIC_BOOKMARK_MAX_LEN = 40
    _SEMANTIC_BOOKMARK_EXTRA_RE = re.compile(r"[^a-z0-9]+")

    def __init__(
        self,
        separator_mode: str = "auto",
        options: RenderOptions | None = None,
        profile: str | None = None,
    ):
        self.options = options or RenderOptions(
            profile=profile, separator_mode=separator_mode if separator_mode != "auto" else None
        )
        self.separator_mode = separator_mode
        self.errors: list[str] = []
        self.charts = False
        self.term_sheet_texts: TermSheetTexts | None = None
        self._term_sheet = False
        self._reset_document()

    def _reset_document(self) -> None:
        self.doc: DocxDocument = Document()
        cast(Any, self.doc.part)._ib_render_errors = self.errors
        self._bookmark_id = 0
        self.styler = DocumentStyler(self.doc)
        self.cover_renderer = CoverRenderer(self.doc)
        self.toc_renderer = TOCRenderer(self.doc)
        self.heading_renderer = HeadingRenderer(self.doc)
        self.paragraph_renderer = ParagraphRenderer(self.doc)
        self.list_renderer = ListRenderer(self.doc)
        self.table_renderer = TableRenderer(
            self.doc,
            term_sheet=self._term_sheet,
            house_style=self.term_sheet_texts.style if self.term_sheet_texts is not None else None,
        )
        self.callout_renderer = CalloutRenderer(self.doc)
        self.image_renderer = ImageRenderer(self.doc)
        self.footnote_renderer = FootnoteRenderer(self.doc)
        self.disclaimer_renderer = DisclaimerRenderer(self.doc)
        self.chart_renderer = ChartRenderer(self.doc)

    def render(
        self,
        model: DocumentModel,
        *,
        asset_roots: Sequence[str | Path] = (),
        limits: InputLimits | None = None,
    ) -> DocxDocument:
        """Render through one composition path, using request-local settings.

        Args:
            model: Parsed or hand-built document with its carried allowed folders.
            asset_roots: Operator folders added to model and RenderOptions folders.
            limits: Explicit policy overriding options and the parsed model's policy.

        Returns:
            The rendered document; strict rejections occur before callers can save it.
        """
        self._composing = False
        try:
            return self._render_model(model, asset_roots=asset_roots, limits=limits)
        except ConversionRejection:
            raise
        except ValueError as error:
            # Non-strict rendering keeps per-element resilience, so a character the
            # XML writer cannot hold only gets this far from text outside it
            # (footnotes, outline headings): lxml raises ValueError for control
            # characters and UnicodeEncodeError, a ValueError, for lone surrogates.
            # Operator arguments are all validated before composition starts, so
            # an earlier error is never relabelled.
            if not self._composing:
                raise
            try:
                validate_xml_strings(model)
            except ValueError as invalid:
                raise StrictValidationError("", (str(invalid),)) from error
            raise

    def _render_model(
        self,
        model: DocumentModel,
        *,
        asset_roots: Sequence[str | Path],
        limits: InputLimits | None,
    ) -> DocxDocument:
        """Compose the document for render(), which names XML-writer failures.

        Args:
            model: Parsed or hand-built document with its carried allowed folders.
            asset_roots: Operator folders added to model and RenderOptions folders.
            limits: Explicit policy overriding options and the parsed model's policy.

        Returns:
            The rendered document.
        """
        self.term_sheet_texts = None
        limits = limits or self.options.limits or model.limits
        self._check_model_limits(model, limits)
        model = deepcopy(model)
        roots = resolve_asset_roots((*model.asset_roots, *self.options.asset_roots, *asset_roots))
        resolved = resolve_options(model.metadata, self.options)
        if model.parsed_profile is not None and model.parsed_profile != resolved.profile.name:
            raise ValueError(
                f"Document was parsed with profile {model.parsed_profile!r}; "
                f"reparse the Markdown with profile={resolved.profile.name!r} before rendering."
            )
        if model.metadata.profile != resolved.profile.name:
            old_defaults = default_metadata(model.metadata.profile)
            new_defaults = default_metadata(resolved.profile.name)
            for field_name in ("title", "company", "sector", "analyst"):
                if getattr(model.metadata, field_name) == getattr(old_defaults, field_name):
                    setattr(model.metadata, field_name, getattr(new_defaults, field_name))
            model.metadata.profile = resolved.profile.name
        with document_validation():
            if not resolved.profile.is_ib:
                validate_office_metadata(model.metadata)
        if resolved.profile.name == "term-sheet":
            with document_validation():
                validate_term_sheet_metadata(model.metadata)
            self.term_sheet_texts = resolve_term_sheet_texts(
                model.metadata, resolved.house, document_house=self.options.house is None
            )
        with _metadata_text("title"):
            report_title_block = (
                not resolved.cover
                and resolved.profile.name == "ib-report"
                and bool(model.metadata.title.strip())
            )
        term_sheet = resolved.profile.name == "term-sheet"
        # Without a cover, every other non-plain profile opens with its title too
        # (`render_office_opening`); it leads the page like the report title block.
        opening_title = (
            not resolved.cover
            and not report_title_block
            and not term_sheet
            and resolved.profile.name not in {"plain", "ib-report"}
        )
        if (resolved.cover and resolved.profile.is_ib) or report_title_block:
            # Filter the private render copy so the TOC and body share the same outline.
            model.elements = [
                element for element in model.elements if not element.inferred_subtitle
            ]
        if report_title_block or term_sheet or opening_title:
            # Transfer the first matching body H1 to the non-outline title block
            # (the term-sheet opening always renders the title).
            # Filtering before the TOC also prevents a stale preview title entry.
            for index, element in enumerate(model.elements):
                if (
                    element.element_type == ElementType.HEADING_1
                    and isinstance(element.content, Heading)
                    and element.content.text == model.metadata.title
                ):
                    model.elements.pop(index)
                    break
        with document_validation():
            appendix_index = letter_appendix_index(model)
        if resolved.strict and model.warnings:
            raise StrictValidationError("Input validation failed: ", model.warnings)
        # Strict mode rejects such a character anywhere in the model. Otherwise the
        # elements keep their per-element resilience and only metadata is checked here.
        with document_validation():
            validate_xml_strings(model if resolved.strict else model.metadata)
        self.errors = list(model.warnings)
        self._term_sheet = term_sheet
        self._reset_document()
        # Inline images (table cells, text) resolve relative paths against the source.
        cast(Any, self.doc.part)._ib_source_dir = model.source_dir
        cast(Any, self.doc.part)._ib_asset_roots = roots
        cast(Any, self.doc.part)._ib_limits = limits
        # Only an operator-selected house can exempt its own logo. A frontmatter
        # override remains document content even when it names the identical path.
        style = model.metadata.extra.get("style", {})
        cast(Any, self.doc.part)._ib_trusted_house_logo = (
            self.options.house is not None and isinstance(style, dict) and "logo" not in style
        )
        self.separator_mode = resolved.separator_mode
        self.charts = resolved.charts
        with document_validation(
            self.options.theme is None, errors=(ValueError, OSError, yaml.YAMLError)
        ):
            style_values = load_style(resolved.profile, resolved.theme)
        self._composing = True
        with (
            use_style(style_values),
            collect_raster_font_diagnostics() as font_warnings,
            collect_term_runs(resolved.term_tags) as term_runs,
        ):
            self.styler.setup_document()
            self.styler.create_styles()
            if resolved.profile.name == "office-letter":
                setup_letter_styles(self.doc)
            if self._term_sheet:
                setup_term_sheet_styles(self.doc)
            elif not resolved.profile.is_ib:
                setup_korean_spacing(self.doc)
            update_fields = OxmlElement("w:updateFields")
            update_fields.set(qn("w:val"), "true")
            insert_ordered(self.doc.settings.element, update_fields)
            if resolved.profile.a4:
                for section in self.doc.sections:
                    section.page_width, section.page_height = Mm(210), Mm(297)
            if resolved.cover:
                self.cover_renderer.include_disclaimer = resolved.disclaimer
                with _metadata_text("cover"):
                    self.cover_renderer.render(model.metadata)
            title_inserted = False
            if report_title_block:
                # The opening validates the attendee list itself for the report profiles.
                with document_validation(), _metadata_text("title block"):
                    title_inserted = render_office_opening(self.doc, model.metadata)
            if self._term_sheet:
                # Resolved in preflight; the opening leads the first page, before any TOC.
                with _metadata_text("term-sheet opening"):
                    title_inserted = render_term_sheet_opening(
                        self.doc,
                        model.metadata,
                        cast(TermSheetTexts, self.term_sheet_texts),
                        TextRenderer.render_runs,
                    )
            if opening_title:
                # Title before the TOC, whose preview no longer lists the title H1.
                with document_validation(), _metadata_text("title block"):
                    title_inserted = render_office_opening(self.doc, model.metadata)
            if resolved.toc:
                self.toc_renderer.render(model)
            # The title H1 was already removed from the render copy for these paths.
            skipped_title = report_title_block or term_sheet or opening_title
            office_closed = False
            for idx, element in enumerate(model.elements):
                if idx == appendix_index:
                    with _metadata_text("closing"):
                        render_office_closing(self.doc, model.metadata)
                    office_closed = True
                    self.doc.add_page_break()
                    label = model.metadata.extra["letter"].get("appendix_label")
                    if label:
                        add_text(self.doc, label, bold=True, keep_next=True)
                if (
                    title_inserted
                    and not skipped_title
                    and element.element_type == ElementType.HEADING_1
                    and isinstance(element.content, Heading)
                    and element.content.text == model.metadata.title
                ):
                    skipped_title = True
                    continue
                if element.element_type not in {ElementType.BULLET_LIST, ElementType.NUMBERED_LIST}:
                    self.list_renderer.numbering.reset()
                try:
                    self._render_element(element)
                except InputLimitError:
                    raise
                except Exception as exc:
                    message = f"Element {idx} ({element.element_type.name}): {exc}"
                    self.errors.append(message)
                    logger.warning("%s", message)
                    p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                    FontStyler.apply_run_style(
                        p.add_run(f"[Render Error: {element.element_type.name}]"),
                        italic=True,
                        color=STYLE.RED,
                    )
            if not office_closed:
                with _metadata_text("closing"):
                    render_office_closing(self.doc, model.metadata)
            if model.footnotes:
                self.footnote_renderer.render(
                    model.footnotes, allow_legacy_refs=resolved.profile.is_ib
                )
            if resolved.disclaimer:
                with _metadata_text("disclaimer (company)"):
                    self.disclaimer_renderer.render(model.metadata.company)
            sender = model.metadata.extra.get("sender", {})
            company = (
                sender.get("organization", model.metadata.company)
                if isinstance(sender, dict)
                else model.metadata.company
            )
            with _metadata_text("header (sender.organization or company)"):
                if self._term_sheet:
                    setup_term_sheet_header_footer(
                        self.doc,
                        model.metadata,
                        self.term_sheet_texts,
                        resolved.confidential,
                        TextRenderer.render_runs,
                    )
                else:
                    self.styler.setup_header_footer(
                        company="" if resolved.profile.name == "office-letter" else company,
                        confidential=resolved.confidential,
                        show_page_numbers=True,
                    )
            self.apply_generator_signature(
                "ib_generated" if resolved.profile.name == "ib-report" else resolved.profile.name
            )
            self._apply_term_controls(term_runs)
            if "checks" in model.metadata.extra:
                # Stored so docx-audit can re-evaluate the relations after edits in Word.
                with document_validation(errors=(ValueError, AttributeError)):
                    checks_json = checks_to_json(
                        parse_checks(model.metadata.extra["checks"]),
                        {
                            str(key): str(value)
                            for key, value in (model.metadata.extra.get("terms") or {}).items()
                        },
                    )
                GeneratorSignatureWriter.upsert_properties(
                    self.doc,
                    {CHECKS_PROPERTY: checks_json},
                )
            ImageRenderer.fit_table_cell_images(self.doc)
            from docx_audit import inspect_document_issues

            for warning in font_warnings:
                self.errors.append(warning)
                paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                FontStyler.apply_run_style(
                    paragraph.add_run(f"[Render Warning: {warning}]"),
                    italic=True,
                    color=STYLE.RED,
                )
            issues = inspect_document_issues(self.doc)
            self.errors.extend(issue for issue in issues if issue not in self.errors)
            if resolved.strict and self.errors:
                raise StrictValidationError("Document validation failed: ", self.errors)
        return self.doc

    @staticmethod
    def _check_model_limits(model: DocumentModel, limits: InputLimits) -> None:
        """Preflight tables and lists even when the caller bypassed parsing.

        Args:
            model: Parsed or hand-built document.
            limits: Effective caller policy for this render.
        """
        for element in model.elements:
            content = element.content
            if isinstance(content, Table):
                columns = max(
                    content.col_count, max((len(row.cells) for row in content.rows), default=0)
                )
                limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(content.rows) * columns)
            item = content[1] if isinstance(content, tuple) else content
            if isinstance(item, ListItem):
                limits.check(RejectionCode.LIMIT_LIST_DEPTH, item.indent_level + 1)

    def apply_generator_signature(self, profile: str = "ib_generated") -> None:
        """Stamp the DOCX package with a generator signature."""
        GeneratorSignatureWriter.apply(self.doc, profile)

    def _apply_term_controls(self, collector: TermRunCollector) -> None:
        """Tag term values as content controls last, then snapshot the tagged values.

        Run-based post-processing (negative and risk colours, base-case bold,
        native footnotes) has finished, because python-docx no longer lists a
        run once it is inside a content control.

        Args:
            collector: This render's registered value runs and tagging option.
        """
        tagged = wrap_term_controls(self.doc, collector)
        if tagged:
            GeneratorSignatureWriter.upsert_properties(
                self.doc,
                {TERM_PROPERTY_PREFIX + key: value for key, value in tagged.items()},
                owned_prefix=TERM_PROPERTY_PREFIX,
            )

    def _add_semantic_bookmark(self, paragraph, element_type: str, extra: str = "") -> None:
        """Wrap a paragraph's first run with a hidden semantic bookmark."""
        if paragraph is None:
            return

        bookmark_name = self._build_semantic_bookmark_name(element_type, extra)
        if not bookmark_name:
            return

        try:
            if not paragraph.runs:
                paragraph.add_run("")

            first_run = paragraph.runs[0]._r
            bookmark_id = str(self._next_bookmark_id())

            bookmark_start = OxmlElement("w:bookmarkStart")
            bookmark_start.set(qn("w:id"), bookmark_id)
            bookmark_start.set(qn("w:name"), bookmark_name)

            bookmark_end = OxmlElement("w:bookmarkEnd")
            bookmark_end.set(qn("w:id"), bookmark_id)

            first_run.addprevious(bookmark_start)
            first_run.addnext(bookmark_end)
        except Exception as exc:
            logger.warning("Failed to add semantic bookmark for %s: %s", element_type, exc)

    def _build_semantic_bookmark_name(self, element_type: str, extra: str = "") -> str:
        """Build a Word-safe semantic bookmark name within Word's length limit."""
        base = f"{self._SEMANTIC_BOOKMARK_PREFIX}{element_type}"
        suffix = uuid4().hex[:4]
        normalized_extra = self._normalize_semantic_bookmark_extra(extra)

        if not normalized_extra:
            return f"{base}_{suffix}"

        max_extra_len = self._SEMANTIC_BOOKMARK_MAX_LEN - len(base) - len(suffix) - 2
        if max_extra_len <= 0:
            return f"{base}_{suffix}"

        trimmed_extra = normalized_extra[:max_extra_len].strip("_")
        if not trimmed_extra:
            return f"{base}_{suffix}"
        return f"{base}_{trimmed_extra}_{suffix}"

    @classmethod
    def _normalize_semantic_bookmark_extra(cls, extra: str) -> str:
        """Normalize bookmark extra data to a compact lowercase slug."""
        normalized = cls._SEMANTIC_BOOKMARK_EXTRA_RE.sub("_", (extra or "").strip().lower())
        return normalized.strip("_")

    def _next_bookmark_id(self) -> int:
        """Return the next bookmark ID for the current document."""
        self._bookmark_id += 1
        return self._bookmark_id

    def _render_element(self, element: Element):
        """Render a single element based on its type"""

        etype = element.element_type

        if etype in (
            ElementType.HEADING_1,
            ElementType.HEADING_2,
            ElementType.HEADING_3,
            ElementType.HEADING_4,
            ElementType.NUMBERED_HEADING,
        ):
            if self._term_sheet:
                render_term_sheet_heading(
                    self.doc, cast(Heading, element.content), TextRenderer.render_runs
                )
            else:
                self.heading_renderer.render(cast(Heading, element.content))

        elif etype == ElementType.PARAGRAPH:
            if self._term_sheet:
                render_term_sheet_paragraph(
                    self.doc, cast(Paragraph, element.content), TextRenderer.render_runs
                )
            else:
                self.paragraph_renderer.render(cast(Paragraph, element.content))

        elif etype == ElementType.BULLET_LIST:
            self.list_renderer.render_bullet(cast(ListItem, element.content))

        elif etype == ElementType.NUMBERED_LIST:
            content = cast(tuple[str, ListItem], element.content)
            number, item = content
            self.list_renderer.render_numbered(number, item)

        elif etype == ElementType.TABLE:
            self.table_renderer.render(cast(Table, element.content))

        elif etype == ElementType.BLOCKQUOTE:
            self.callout_renderer.render(cast(Blockquote, element.content))

        elif etype == ElementType.IMAGE:
            self.image_renderer.render(cast(Image, element.content))

        elif etype == ElementType.LATEX_BLOCK:
            self._render_latex_block(cast(LaTeXEquation, element.content))

        elif etype == ElementType.LATEX_INLINE:
            self._render_latex_inline(cast(LaTeXEquation, element.content))

        elif etype == ElementType.SEPARATOR:
            self._render_separator(element)

        elif etype == ElementType.CODE_BLOCK:
            self._render_code_block(cast(CodeBlock, element.content))

        elif etype == ElementType.CONFIRMATION:
            texts = self.term_sheet_texts
            confirmation = texts.confirmation if self._term_sheet and texts is not None else None
            if confirmation is not None and confirmation_has_text(confirmation):
                render_confirmation(self.doc, confirmation, TextRenderer.render_runs)
            else:
                # Plan §2-9: without wording, keep the fence losslessly and report it.
                source = cast(ConfirmationBlock, element.content).source
                self._render_code_block(CodeBlock(source, "confirmation"))
                message = "Confirmation text is missing; original fence preserved as a code block"
                self.errors.append(message)
                logger.warning("%s", message)

        elif etype == ElementType.CHART:
            chart = cast(Chart, element.content)
            if self.charts:
                try:
                    self.chart_renderer.render(chart)
                    self._add_semantic_bookmark(self.doc.paragraphs[-1], ElementType.CHART.name)
                    return
                except Exception as exc:
                    message = f"{chart.label}: {exc}"
                    self.errors.append(message)
                    logger.warning("%s; rendering original code panel", message)
            self._render_code_block(
                CodeBlock(chart.code, "chart", CodeBlock.detect_ascii_art(chart.code))
            )

        elif etype == ElementType.DIAGRAM:
            from diagram_renderer import DiagramRenderer
            from md_parser import Diagram

            diagram = cast(Diagram, element.content)
            start_paragraph_count = len(self.doc.paragraphs)
            renderer = DiagramRenderer(
                self.doc,
                theme_colors={
                    "navy": f"#{STYLE.NAVY_HEX}",
                },
            )
            if not renderer.render(diagram):
                raise ValueError("Diagram could not be rendered")
            if len(self.doc.paragraphs) > start_paragraph_count:
                self._add_semantic_bookmark(
                    self.doc.paragraphs[start_paragraph_count],
                    ElementType.DIAGRAM.name,
                    diagram.diagram_type,
                )

        elif etype == ElementType.EMPTY:
            pass  # Intentionally skip empty elements

        else:
            raise ValueError("Unhandled element type: " + etype.name)

    def _render_separator(self, element: Element):
        """Render a separator as either a horizontal rule or a page break.

        Modes:
            - rule: always render a horizontal rule
            - page-break: always render a page break
            - auto: `## ---` becomes a page break, plain `---` stays a rule
        """
        if self._resolve_separator_mode(element) == "page-break":
            paragraph = self.doc.add_page_break()
            self._add_semantic_bookmark(paragraph, ElementType.SEPARATOR.name)
            return

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Create a horizontal rule via paragraph bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")  # 0.75pt line
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), STYLE.GRAY_BORDER_HEX)
        pBdr.append(bottom)
        insert_ordered(pPr, pBdr)

        # Minimal spacing
        pFmt = p.paragraph_format
        pFmt.space_before = Pt(6)
        pFmt.space_after = Pt(6)
        self._add_semantic_bookmark(p, ElementType.SEPARATOR.name)

    def _resolve_separator_mode(self, element: Element) -> str:
        """Resolve separator rendering mode for a specific element."""
        if self.separator_mode in {"rule", "page-break"}:
            return self.separator_mode

        raw_text = (element.raw_text or "").strip()
        if raw_text == "## ---":
            return "page-break"
        return "rule"

    def _render_code_block(self, code_block):
        """Render a fenced code block as a monospaced shaded block."""
        if not isinstance(code_block, CodeBlock):
            return

        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        cell = table.rows[0].cells[0]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        TableStyler.set_cell_background(cell, str(STYLE.CODE_BG))
        self._style_code_block_table(table)

        lines = code_block.code.splitlines() or [""]
        if code_block.code.endswith("\n"):
            # splitlines() drops the empty line after a final line break.
            lines.append("")
        for idx, line in enumerate(lines):
            paragraph = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
            paragraph.clear()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            self._render_code_block_line(paragraph, line)

        self._add_semantic_bookmark(
            cell.paragraphs[0],
            ElementType.CODE_BLOCK.name,
            code_block.language,
        )
        self.doc.add_paragraph()

    @staticmethod
    def _style_code_block_table(table) -> None:
        """Render code blocks as clean panels rather than visible grid tables."""
        tbl = table._tbl
        tbl_pr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tbl_borders = OxmlElement("w:tblBorders")

        for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "nil")
            tbl_borders.append(border)

        insert_ordered(tbl_pr, tbl_borders)
        if tbl.tblPr is None:
            tbl.insert(0, tbl_pr)

        cell = table.rows[0].cells[0]
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_mar = OxmlElement("w:tcMar")
        for edge in ("top", "left", "bottom", "right"):
            margin = OxmlElement(f"w:{edge}")
            margin.set(qn("w:w"), "180" if edge in {"left", "right"} else "140")
            margin.set(qn("w:type"), "dxa")
            tc_mar.append(margin)
        insert_ordered(tc_pr, tc_mar)

    def _render_code_block_line(self, paragraph, line: str) -> None:
        """Render one line of a code block with light semantic emphasis for tree diagrams."""
        if not line.strip():
            spacer = paragraph.add_run(line)
            FontStyler.apply_run_style(
                spacer,
                font_name="Consolas",
                font_size=Pt(9.5),
                color=STYLE.MEDIUM_GRAY,
            )
            # A blank line needs no spacer text: the paragraph mark, which keeps the
            # body font, sets the height of every line of the panel (checked in Word).
            return

        prefix_match = self._TREE_PREFIX_RE.match(line)
        if prefix_match:
            prefix, remainder = prefix_match.groups()
        else:
            prefix, remainder = "", line

        if prefix:
            prefix_run = paragraph.add_run(prefix)
            FontStyler.apply_run_style(
                prefix_run,
                font_name="Consolas",
                font_size=Pt(9.5),
                color=STYLE.MEDIUM_GRAY,
            )

        value_match = self._TREE_VALUE_RE.match(remainder)
        if value_match:
            before, value, divider, after = value_match.groups()
            if before:
                before_run = paragraph.add_run(before)
                FontStyler.apply_run_style(
                    before_run,
                    font_name="Consolas",
                    font_size=Pt(9.5),
                    color=STYLE.DARK_GRAY,
                )
            value_run = paragraph.add_run(value)
            FontStyler.apply_run_style(
                value_run,
                font_name="Consolas",
                font_size=Pt(9.5),
                bold=True,
                color=STYLE.NAVY,
            )
            divider_run = paragraph.add_run(divider)
            FontStyler.apply_run_style(
                divider_run,
                font_name="Consolas",
                font_size=Pt(9.5),
                color=STYLE.MEDIUM_GRAY,
            )
            tail_run = paragraph.add_run(after)
            FontStyler.apply_run_style(
                tail_run,
                font_name="Consolas",
                font_size=Pt(9.5),
                color=STYLE.DARK_GRAY,
            )
            return

        is_root = "사업지주회사" in remainder
        main_run = paragraph.add_run(remainder)
        FontStyler.apply_run_style(
            main_run,
            font_name="Consolas",
            font_size=Pt(9.8 if is_root else 9.5),
            bold=is_root,
            color=STYLE.NAVY if is_root else STYLE.DARK_GRAY,
        )

    def _render_latex_block(self, latex_eq):
        """
        Render a LaTeX block equation as an image.

        Uses matplotlib to render the LaTeX expression to a PNG,
        then inserts it into the document centered.
        """
        from md_parser import LaTeXEquation

        if not isinstance(latex_eq, LaTeXEquation):
            logger.warning("Invalid LaTeX equation object")
            return

        image_path = LaTeXRenderer.render_to_image(latex_eq.expression)

        if image_path:
            try:
                # Insert centered equation image
                paragraph = self.doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run()
                run.add_picture(image_path, width=Inches(4.0))

                # Spacer
                self.doc.add_paragraph()

            except Exception as e:
                logger.warning("Failed to insert LaTeX image: %s", e)
                self._render_latex_fallback(latex_eq.expression)
            finally:
                # Clean up temp file
                try:
                    import os

                    os.unlink(image_path)
                except OSError:
                    pass
        else:
            self._render_latex_fallback(latex_eq.expression)

    def _render_latex_inline(self, latex_eq):
        """Render a standalone inline equation through the common text path."""
        from md_parser import LaTeXEquation

        if isinstance(latex_eq, LaTeXEquation):
            paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
            TextRenderer._render_inline_latex(paragraph, latex_eq.expression, STYLE.BODY_SIZE)

    def _render_latex_fallback(self, expression: str, inline: bool = False):
        """Render LaTeX as styled text when image rendering fails."""
        message = "LaTeX could not be rendered: " + expression
        self.errors.append(message)
        logger.warning("%s", message)
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
        if not inline:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run(f"[{expression}]")
        FontStyler.apply_run_style(
            run,
            font_name="Consolas",
            font_size=STYLE.BODY_SIZE,
            italic=True,
            color=STYLE.DARK_GRAY,
        )
