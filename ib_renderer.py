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
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from copy import deepcopy
from pathlib import Path
from typing import (
    Any,
    cast,
)
from uuid import uuid4

import yaml
from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt

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
from ib_render_tables import _BLACK as _BLACK
from ib_render_tables import TableRenderer as TableRenderer
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
)
from numeric_checks import CHECKS_PROPERTY, checks_to_json, parse_checks
from office_layout import (
    add_text,
    letter_appendix_index,
    render_office_closing,
    render_office_opening,
    setup_korean_language,
    setup_korean_spacing,
    setup_letter_styles,
)
from ooxml_order import insert_ordered
from render_styles import STYLE, collect_raster_font_diagnostics, use_style
from render_styles import IBStyle as IBStyle
from term_sheet import (
    TermSheetTexts,
    confirmation_has_text,
    render_confirmation,
    render_term_sheet_heading,
    render_term_sheet_opening,
    render_term_sheet_paragraph,
    resolve_term_sheet_texts,
    setup_term_sheet_header_footer,
    setup_term_sheet_styles,
)
from term_variables import (
    TERM_PROPERTY_PREFIX,
    TermRunCollector,
    collect_term_runs,
    wrap_term_controls,
)

logger = logging.getLogger(__name__)


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
        # Rendering notes that never reject a document, strict or not.
        self.warnings: list[str] = []
        self.charts = False
        self.term_sheet_texts: TermSheetTexts | None = None
        self._term_sheet = False
        self._reset_document()

    def _reset_document(self) -> None:
        self.doc: DocxDocument = Document()
        cast(Any, self.doc.part)._ib_render_errors = self.errors
        cast(Any, self.doc.part)._ib_render_warnings = self.warnings
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
        self.warnings = []
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
            else:
                setup_korean_language(self.doc)
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
