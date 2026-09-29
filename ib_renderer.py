"""
IB Renderer Module for Word Report Generation
Handles styling and rendering of document elements in IB Bank style.

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
import os
import platform
import re
import time
from copy import deepcopy
from dataclasses import replace
from importlib.metadata import PackageNotFoundError, version
from io import BytesIO
from pathlib import Path
from typing import AbstractSet, Any, Dict, List, Optional, Set, Tuple, cast
from uuid import uuid4
from xml.sax.saxutils import escape

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart, serialize_part_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
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
from office_layout import (
    NativeNumbering,
    add_text,
    letter_appendix_index,
    render_office_closing,
    render_office_opening,
    setup_letter_styles,
)
from render_styles import STYLE, RasterFontPolicy, collect_raster_font_diagnostics, use_style
from render_styles import IBStyle as IBStyle
from term_sheet import (
    ROW_SPLIT_THRESHOLD,
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
    key_value_widths,
    line_text,
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

logger = logging.getLogger(__name__)

_BLACK = RGBColor(0, 0, 0)


# ═══════════════════════════════════════════════════════════════════════════════
# STYLE CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT SIGNATURE
# ═══════════════════════════════════════════════════════════════════════════════


class GeneratorSignatureWriter:
    """Embed an explicit generator signature in DOCX custom properties."""

    _CUSTOM_PROPS_PARTNAME = PackURI("/docProps/custom.xml")
    _CUSTOM_PROPS_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
    )
    _PROPERTY_TEMPLATE = (
        '<property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" '
        'fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" name="{name}">'
        "<vt:lpwstr>{value}</vt:lpwstr>"
        "</property>"
    )
    _PROJECT_NAME = "ib-report-formatter"
    _GENERATOR_NAME = "ib_report_formatter"
    _GENERATOR_PROFILE = "ib_generated"
    _PROPERTY_NAMES = ("generator", "generator_version", "generator_profile")
    _PYPROJECT_VERSION_RE = re.compile(r'^version\s*=\s*"([^"]+)"\s*$')
    _resolved_version: Optional[str] = None

    @classmethod
    def apply(cls, doc: DocxDocument, profile: str = "ib_generated") -> None:
        """Upsert generator signature custom properties on a DOCX package."""
        custom_part, custom_props = cls._get_or_add_custom_properties_part(doc)
        cls._upsert_signature_properties(custom_props, profile)
        if isinstance(custom_part, XmlPart):
            custom_part._element = custom_props
            return
        custom_part._blob = serialize_part_xml(custom_props)

    @classmethod
    def _get_or_add_custom_properties_part(cls, doc: DocxDocument):
        """Return the package custom-properties part plus a mutable XML root."""
        package = doc.part.package
        try:
            custom_part = package.part_related_by(RT.CUSTOM_PROPERTIES)
            try:
                custom_props = parse_xml(custom_part.blob)
            except Exception:
                custom_props = parse_xml(cls._CUSTOM_PROPS_XML)
            return custom_part, custom_props
        except KeyError:
            custom_props = parse_xml(cls._CUSTOM_PROPS_XML)
            custom_part = XmlPart(
                cls._CUSTOM_PROPS_PARTNAME,
                CT.OFC_CUSTOM_PROPERTIES,
                custom_props,
                package,
            )
            package.relate_to(custom_part, RT.CUSTOM_PROPERTIES)
            return custom_part, custom_props

    @classmethod
    def _upsert_signature_properties(cls, custom_props, profile: str = "ib_generated") -> None:
        """Replace only the generator signature properties, preserving others."""
        for prop in list(custom_props):
            if cls._local_name(prop.tag) != "property":
                continue
            if (prop.get("name") or "").strip() in cls._PROPERTY_NAMES:
                custom_props.remove(prop)

        used_pids = cls._used_property_ids(custom_props)
        for name, value in cls._signature_properties(profile).items():
            custom_props.append(cls._build_property(name, value, cls._next_pid(used_pids)))

    @classmethod
    def _signature_properties(cls, profile: str = "ib_generated") -> Dict[str, str]:
        """Return the ordered generator signature payload."""
        return {
            "generator": cls._GENERATOR_NAME,
            "generator_version": cls._resolve_version(),
            "generator_profile": profile,
        }

    @classmethod
    def _resolve_version(cls) -> str:
        """Resolve the package version, falling back to pyproject parsing."""
        if cls._resolved_version:
            return cls._resolved_version

        try:
            cls._resolved_version = version(cls._PROJECT_NAME)
            return cls._resolved_version
        except PackageNotFoundError:
            pass

        pyproject_path = Path(__file__).resolve().with_name("pyproject.toml")
        try:
            for line in pyproject_path.read_text(encoding="utf-8").splitlines():
                match = cls._PYPROJECT_VERSION_RE.match(line.strip())
                if match:
                    cls._resolved_version = match.group(1)
                    return cls._resolved_version
        except OSError:
            pass

        cls._resolved_version = "unknown"
        return cls._resolved_version

    @classmethod
    def _build_property(cls, name: str, value: str, pid: int):
        """Build a single string-valued custom property element."""
        property_xml = cls._PROPERTY_TEMPLATE.format(
            pid=pid,
            name=escape(name, {'"': "&quot;"}),
            value=escape(value),
        )
        return parse_xml(property_xml.encode("utf-8"))

    @staticmethod
    def _used_property_ids(custom_props) -> List[int]:
        """Collect all valid pid values already present in the custom-properties root."""
        used_pids = []
        for prop in list(custom_props):
            if not hasattr(prop, "get"):
                continue
            pid = (prop.get("pid") or "").strip()
            if pid.isdigit():
                used_pids.append(int(pid))
        return used_pids

    @staticmethod
    def _next_pid(used_pids: List[int]) -> int:
        """Return the next unused custom-property PID (starts at 2 per OPC convention)."""
        candidate = 2
        while candidate in used_pids:
            candidate += 1
        used_pids.append(candidate)
        return candidate

    @staticmethod
    def _local_name(tag: str) -> str:
        """Strip namespace from an XML tag."""
        return tag.split("}", 1)[-1] if "}" in tag else tag


# ═══════════════════════════════════════════════════════════════════════════════
# FONT STYLER
# ═══════════════════════════════════════════════════════════════════════════════


class FontStyler:
    """Handles font styling including East Asian fonts"""

    @staticmethod
    def set_east_asian_font(element, font_name: Optional[str] = None):
        """Set East Asian font (for Korean text) on a style or run element"""
        resolved_font = font_name or FontPolicy.resolve_korean_font()
        elm = element._element
        rPr = elm.get_or_add_rPr()
        if rPr.rFonts is None:
            rPr.get_or_add_rFonts()
        rPr.rFonts.set(qn("w:eastAsia"), resolved_font)
        # Theme attributes take precedence over explicit fonts in Word.
        for explicit, theme in (("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"), ("eastAsia", "eastAsiaTheme")):
            if rPr.rFonts.get(qn("w:" + explicit)):
                rPr.rFonts.attrib.pop(qn("w:" + theme), None)

    @staticmethod
    def apply_run_style(
        run,
        font_name: Optional[str] = None,
        font_size: Optional[Pt] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[RGBColor] = None,
        superscript: Optional[bool] = None,
    ):
        """Apply styling to a run"""
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if color:
            run.font.color.rgb = color
        if superscript is not None:
            run.font.superscript = superscript
        FontStyler.set_east_asian_font(run, font_name=font_name)


class FontPolicy:
    """Declare Word fonts for the reader's machine, independently of raster fonts."""

    @classmethod
    def resolve_korean_font(cls, system_name: Optional[str] = None) -> str:
        """Return the preferred Korean font for the current platform."""
        system = system_name or platform.system() or "Unknown"
        if system == "Darwin":
            candidates = ("Apple SD Gothic Neo", STYLE.KOREAN_FONT, "NanumGothic")
        elif system == "Windows":
            candidates = (STYLE.KOREAN_FONT, "NanumGothic", "Apple SD Gothic Neo")
        else:
            candidates = (STYLE.KOREAN_FONT, "NanumGothic", "Apple SD Gothic Neo")

        chosen = candidates[0]
        logger.debug(
            "Resolved Korean font for %s: %s (fallbacks: %s)",
            system,
            chosen,
            ", ".join(candidates[1:]) or "none",
        )
        return chosen


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT STYLER
# ═══════════════════════════════════════════════════════════════════════════════


class DocumentStyler:
    """Sets up document styles"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def setup_document(self):
        """Set up document margins and page settings"""
        for section in self.doc.sections:
            section.top_margin = STYLE.TOP_MARGIN
            section.bottom_margin = STYLE.BOTTOM_MARGIN
            section.left_margin = STYLE.LEFT_MARGIN
            section.right_margin = STYLE.RIGHT_MARGIN

    def create_styles(self):
        """Create all custom IB styles"""
        styles = self.doc.styles

        # Do not inherit an unrelated Word theme's title colour or borders.
        self._setup_heading_style(
            styles["Title"],
            font_size=STYLE.H1_SIZE,
            color=STYLE.NAVY,
            space_before=STYLE.H1_SPACE_BEFORE,
            space_after=STYLE.H1_SPACE_AFTER,
            add_border=STYLE.HEADING_BORDER,
        )

        # Heading 1
        self._setup_heading_style(
            styles["Heading 1"],
            font_size=STYLE.H1_SIZE,
            color=STYLE.NAVY,
            space_before=STYLE.H1_SPACE_BEFORE,
            space_after=STYLE.H1_SPACE_AFTER,
            add_border=STYLE.HEADING_BORDER,
        )

        # Heading 2
        self._setup_heading_style(
            styles["Heading 2"],
            font_size=STYLE.H2_SIZE,
            color=STYLE.DARK_GRAY,
            space_before=STYLE.H2_SPACE_BEFORE,
            space_after=STYLE.H2_SPACE_AFTER,
        )

        # Heading 3
        self._setup_heading_style(
            styles["Heading 3"],
            font_size=STYLE.H3_SIZE,
            color=STYLE.NAVY,
            space_before=STYLE.H3_SPACE_BEFORE,
            space_after=STYLE.H3_SPACE_AFTER,
        )

        # Heading 4
        self._setup_heading_style(
            styles["Heading 4"],
            font_size=STYLE.H4_SIZE,
            color=STYLE.DARK_GRAY,
            space_before=STYLE.H3_SPACE_BEFORE,
            space_after=STYLE.H3_SPACE_AFTER,
        )

        # IB Body
        body = self._get_or_create_style(STYLE.STYLE_IB_BODY, WD_STYLE_TYPE.PARAGRAPH)
        body.font.name = STYLE.BODY_FONT
        body.font.size = STYLE.BODY_SIZE
        body.paragraph_format.line_spacing = STYLE.BODY_LINE_SPACING
        body.paragraph_format.space_after = STYLE.BODY_SPACE_AFTER
        body.paragraph_format.alignment = (
            WD_ALIGN_PARAGRAPH.JUSTIFY if STYLE.BODY_JUSTIFY else WD_ALIGN_PARAGRAPH.LEFT
        )
        body.paragraph_format.widow_control = True
        FontStyler.set_east_asian_font(body)

        # IB Bullet
        bullet = self._get_or_create_style(STYLE.STYLE_IB_BULLET, WD_STYLE_TYPE.PARAGRAPH)
        bullet.font.name = STYLE.BODY_FONT
        bullet.font.size = STYLE.BODY_SIZE
        bullet.paragraph_format.left_indent = STYLE.BULLET_INDENT
        bullet.paragraph_format.first_line_indent = -STYLE.BULLET_INDENT
        bullet.paragraph_format.space_after = STYLE.BULLET_SPACE_AFTER
        FontStyler.set_east_asian_font(bullet)

        # Built-in TOC styles used by Word field updates
        self._setup_toc_style(
            self._get_or_create_style("TOC 1", WD_STYLE_TYPE.PARAGRAPH),
            STYLE.BODY_SIZE,
            bold=True,
        )
        self._setup_toc_style(
            self._get_or_create_style("TOC 2", WD_STYLE_TYPE.PARAGRAPH),
            Pt(10),
            left_indent=Inches(0.2),
        )
        self._setup_toc_style(
            self._get_or_create_style("TOC 3", WD_STYLE_TYPE.PARAGRAPH),
            STYLE.SMALL_SIZE,
            left_indent=Inches(0.4),
        )
        self._setup_toc_style(
            self._get_or_create_style("TOC 4", WD_STYLE_TYPE.PARAGRAPH),
            STYLE.SMALL_SIZE,
            left_indent=Inches(0.6),
        )

    def _setup_heading_style(
        self,
        style,
        font_size: Pt,
        color: RGBColor,
        space_before: Pt,
        space_after: Pt,
        add_border: bool = False,
    ):
        """Configure a heading style"""
        style.font.name = STYLE.HEADING_FONT
        style.font.size = font_size
        style.font.bold = True
        style.font.color.rgb = color
        style.paragraph_format.space_before = space_before
        style.paragraph_format.space_after = space_after
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True
        FontStyler.set_east_asian_font(style)
        for border in list(style.element.xpath("./w:pPr/w:pBdr")):
            border.getparent().remove(border)
        if add_border:
            self._add_bottom_border(style)

    def _get_or_create_style(self, name: str, style_type):
        """Get existing style or create new one"""
        try:
            return self.doc.styles.add_style(name, style_type)
        except ValueError:
            return self.doc.styles[name]

    def _setup_toc_style(
        self,
        style,
        font_size: Pt,
        bold: bool = False,
        left_indent: Inches = Inches(0),
    ) -> None:
        """Configure built-in TOC styles so Word-generated entries match Korean cover typography."""
        style.font.name = STYLE.TOC_FONT
        style.font.size = font_size
        style.font.bold = bold
        style.font.color.rgb = STYLE.DARK_GRAY
        style.paragraph_format.left_indent = left_indent
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(2 if bold else 0)
        FontStyler.set_east_asian_font(style, STYLE.TOC_FONT)

    @staticmethod
    def _add_bottom_border(style):
        """Add bottom border to a style"""
        pPr = style._element.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:color"), STYLE.NAVY_HEX)
        pBdr.append(bottom)
        pPr.append(pBdr)

    def setup_header_footer(
        self,
        company: str = "",
        confidential: bool = True,
        show_page_numbers: bool = True,
    ):
        """
        Set up professional IB-style header and footer.

        Args:
            company: Company name to display in header
            confidential: Whether to show "CONFIDENTIAL" mark
            show_page_numbers: Whether to show page numbers in footer
        """
        for section in self.doc.sections:
            # ── Header ─────────────────────────────────────────────────────────
            header = section.header
            header.is_linked_to_previous = False

            # Clear existing content
            for para in header.paragraphs:
                para.clear()

            # Add header content
            header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()

            # Left-aligned company name
            if company:
                company_run = header_para.add_run(company)
                FontStyler.apply_run_style(
                    company_run,
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.SMALL_SIZE,
                    color=STYLE.DARK_GRAY,
                )

            # Tab width follows this section, including landscape tables.
            width, left, right = section.page_width, section.left_margin, section.right_margin
            if width is None or left is None or right is None:
                raise ValueError("Section dimensions are required for header alignment")
            available_width = width - left - right
            header_style = header_para.style
            if header_style is not None:
                header_style.paragraph_format.tab_stops.clear_all()
            header_para.paragraph_format.tab_stops.clear_all()
            header_para.paragraph_format.tab_stops.add_tab_stop(available_width, WD_TAB_ALIGNMENT.RIGHT)
            if company or confidential:
                header_para.add_run("\t")

            # Right-aligned confidential mark
            if confidential:
                conf_run = header_para.add_run("CONFIDENTIAL")
                FontStyler.apply_run_style(
                    conf_run,
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.SMALL_SIZE,
                    bold=True,
                    color=STYLE.RED,
                )

            # Add separator line under header
            if company or confidential:
                self._add_header_border(header_para)

            # ── Footer ─────────────────────────────────────────────────────────
            footer = section.footer
            footer.is_linked_to_previous = False

            # Clear existing content
            for para in footer.paragraphs:
                para.clear()

            footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if show_page_numbers:
                # Add page number field
                self._add_page_number_field(footer_para)

    def _add_header_border(self, paragraph):
        """Add bottom border to header paragraph."""
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:color"), STYLE.GRAY_BORDER_HEX)
        pBdr.append(bottom)
        pPr.append(pBdr)

    def _add_page_number_field(self, paragraph):
        """Add page number field code to paragraph."""
        # "Page X of Y" format
        run1 = paragraph.add_run(STYLE.PAGE_LABEL)
        FontStyler.apply_run_style(
            run1,
            font_name=STYLE.BODY_FONT,
            font_size=STYLE.SMALL_SIZE,
            color=STYLE.DARK_GRAY,
        )

        # PAGE field
        run_page = paragraph.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")

        instrText = OxmlElement("w:instrText")
        instrText.text = "PAGE"

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")

        run_page._r.append(fldChar1)
        run_page._r.append(instrText)
        run_page._r.append(fldChar2)
        FontStyler.apply_run_style(
            run_page,
            font_name=STYLE.BODY_FONT,
            font_size=STYLE.SMALL_SIZE,
            color=STYLE.DARK_GRAY,
        )

        run2 = paragraph.add_run(STYLE.PAGE_OF_LABEL)
        FontStyler.apply_run_style(
            run2,
            font_name=STYLE.BODY_FONT,
            font_size=STYLE.SMALL_SIZE,
            color=STYLE.DARK_GRAY,
        )

        # NUMPAGES field
        run_total = paragraph.add_run()
        fldChar3 = OxmlElement("w:fldChar")
        fldChar3.set(qn("w:fldCharType"), "begin")

        instrText2 = OxmlElement("w:instrText")
        instrText2.text = "NUMPAGES"

        fldChar4 = OxmlElement("w:fldChar")
        fldChar4.set(qn("w:fldCharType"), "end")

        run_total._r.append(fldChar3)
        run_total._r.append(instrText2)
        run_total._r.append(fldChar4)
        FontStyler.apply_run_style(
            run_total,
            font_name=STYLE.BODY_FONT,
            font_size=STYLE.SMALL_SIZE,
            color=STYLE.DARK_GRAY,
        )


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE STYLER
# ═══════════════════════════════════════════════════════════════════════════════


class TableStyler:
    """Handles table styling"""

    @staticmethod
    def set_cell_background(cell, hex_color: str):
        """Set cell background color, replacing an earlier fill in place."""
        tcPr = cell._element.tcPr
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            cell._element.append(tcPr)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        existing = tcPr.find(qn("w:shd"))
        if existing is not None:
            # A second fill (e.g. a base case over a label) must not duplicate w:shd.
            tcPr.replace(existing, shd)
        else:
            tcPr.append(shd)

    @staticmethod
    def set_table_borders(table):
        """Apply IB-style borders to table"""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tblBorders = OxmlElement("w:tblBorders")

        # Outer borders: Navy, thick
        for border_name in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "12")
            border.set(qn("w:color"), STYLE.NAVY_HEX)
            tblBorders.append(border)

        # Inner horizontal: Gray, thin solid
        insideH = OxmlElement("w:insideH")
        insideH.set(qn("w:val"), "single")
        insideH.set(qn("w:sz"), "4")
        insideH.set(qn("w:color"), STYLE.GRAY_BORDER_HEX)
        tblBorders.append(insideH)

        # Inner vertical: Gray, dotted for readability
        insideV = OxmlElement("w:insideV")
        insideV.set(qn("w:val"), "dotted")
        insideV.set(qn("w:sz"), "4")
        insideV.set(qn("w:color"), STYLE.GRAY_BORDER_HEX)
        tblBorders.append(insideV)

        tblPr.append(tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)


# ═══════════════════════════════════════════════════════════════════════════════
# ELEMENT RENDERERS
# ═══════════════════════════════════════════════════════════════════════════════


class TextRenderer:
    """Renders text with formatting"""

    # Compiled regex — class-level cache
    _BOLD_SPLIT_RE = re.compile(r"(\*\*.*?\*\*)")
    _ITALIC_SPLIT_RE = re.compile(r"(?<!\*)(\*[^*]+?\*)(?!\*)")
    _SUPERSCRIPT_RE = re.compile(r"\^([^^]+?)\^")
    _SUBSCRIPT_PATTERN = r"(?<!~)~[A-Za-z0-9]{1,8}~(?!~)"
    _VERTICAL_ALIGN_SPLIT_RE = re.compile(r"(\^[^^]+?\^|" + _SUBSCRIPT_PATTERN + r")")
    _ESCAPE_RE = re.compile(r'\\([~.*"\'()\[\]{}|_-])')

    @staticmethod
    def render_runs(
        paragraph,
        runs: List[TextRun],
        default_color: Optional[RGBColor] = None,
        font_name: Optional[str] = None,
        font_size: Optional[Pt] = None,
    ):
        """Render text runs to a paragraph"""
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

            run = paragraph.add_run(run_data.text)
            run.font.name = font_name
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
        font_name: Optional[str] = None,
        font_size: Optional[Pt] = None,
    ):
        """Parse and render text with **bold** markers (legacy, delegates to render_text_with_formatting)."""
        TextRenderer.render_text_with_formatting(
            paragraph, text, font_name=font_name, font_size=font_size
        )

    @staticmethod
    def render_text_with_formatting(
        paragraph,
        text: str,
        font_name: Optional[str] = None,
        font_size: Optional[Pt] = None,
        default_color: Optional[RGBColor] = None,
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
            run.bold or run.italic or run.superscript or run.subscript or run.color_hex
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

            if bold_part.startswith("**") and bold_part.endswith("**") and len(bold_part) > 4:
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

                    if (
                        italic_part.startswith("*")
                        and italic_part.endswith("*")
                        and len(italic_part) > 2
                        and not italic_part.startswith("**")
                    ):
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
        color: Optional[RGBColor] = None,
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
    def _rgb_from_hex(color_hex: Optional[str]) -> Optional[RGBColor]:
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

        tblPr.append(tblBorders)
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
        pPr.append(pBdr)


class TOCRenderer:
    """Renders the Table of Contents"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, model: Optional[DocumentModel] = None):
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
            text = element.content.text.strip()
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

    def render(self, heading: Heading):
        """Render a heading with appropriate level"""
        # Clean markdown bold markers from heading text
        clean_text = heading.text.replace("**", "").strip()
        clean_text = TextRenderer._cleanup(clean_text)

        # Clamp level to 1-4 (Word supports Heading 1-9, but we style 1-4)
        level = max(1, min(heading.level, 4))
        p = self.doc.add_heading(level=level)

        size, color, bold = self._STYLE_CONFIG.get(level, (STYLE.BODY_SIZE, STYLE.DARK_GRAY, True))

        runs = [replace(run, bold=bold) for run in TextParser.parse_runs(clean_text)]
        TextRenderer.render_runs(p, runs, font_name=STYLE.HEADING_FONT, font_size=size, default_color=color)


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
    _MIN_COLUMN_WIDTH_INCHES = 0.65
    _MIN_TEXT_COLUMN_WIDTH_INCHES = 1.15
    _MAX_NUMERIC_COLUMN_WIDTH_INCHES = 1.35
    _MAX_TEXT_COLUMN_SHARE = 0.55

    def __init__(self, doc: DocxDocument, term_sheet: bool = False):
        self.doc = doc
        # Profile-only layout; merge emission itself is shared by every profile.
        self.term_sheet = term_sheet

    def render(self, table: Table):
        """Render a table"""
        if not table.rows or table.col_count == 0:
            return

        row_count = len(table.rows)
        col_count = table.col_count

        previous_geometry = self._begin_landscape(table)
        if self.term_sheet:
            self._render_term_sheet_table(table)
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

        # Render header row
        if table.rows:
            self._render_header_row(word_table, table.rows[0], col_count, covered)

        # Render data rows based on table type
        for r_idx, row in enumerate(table.rows[1:], 1):
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
            )
        self._mirror_vertical_fills(word_table, rectangles)

        word_table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
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

    def _begin_landscape(self, table: Table) -> Optional[Tuple[Any, int, int]]:
        """Open a landscape section for a landscape table; return the prior geometry."""
        if not table.landscape:
            return None
        section = self.doc.sections[-1]
        assert section.page_width is not None and section.page_height is not None
        previous_geometry = (section.orientation, section.page_width, section.page_height)
        section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Emu(max(previous_geometry[1:]))
        section.page_height = Emu(min(previous_geometry[1:]))
        return previous_geometry

    def _end_landscape(self, previous_geometry: Optional[Tuple[Any, int, int]]) -> None:
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
        """
        row_count, col_count = len(table.rows), table.col_count
        available = self._get_available_table_width_emu()
        add_table_heading(
            self.doc, table.caption, table.unit, table.as_of, TextRenderer.render_runs, available
        )
        word_table = self.doc.add_table(rows=row_count, cols=col_count)
        word_table.style = STYLE.STYLE_TABLE_GRID
        column_kinds = self._infer_column_kinds(table)
        label_columns = (
            table.label_columns if table.label_columns is not None else min(1, max(col_count - 1, 0))
        )
        key_value = label_columns in (1, 2) and col_count == label_columns + 1
        widths = key_value_widths(label_columns, available) if key_value else None
        if widths is None:
            widths = [
                int(Inches(width))
                for width in self._estimate_column_widths(
                    table, available / self._EMUS_PER_INCH, column_kinds
                )
            ]
        self._set_column_widths(word_table, widths)
        rectangles = self._merge_cells(word_table, table)
        covered = self._covered_cells(rectangles)
        extents = {(top, left): (bottom, right) for top, left, bottom, right in rectangles}
        apply_table_frame(word_table, sum(widths))

        line_counts = [0] * row_count
        for row_index, row in enumerate(table.rows):
            word_cells = word_table.rows[row_index].cells
            for column_index in range(col_count):
                if (row_index, column_index) in covered:
                    continue
                bottom, right = extents.get((row_index, column_index), (row_index, column_index))
                cell_data = (
                    row.cells[column_index] if column_index < len(row.cells) else TableCell(content="")
                )
                role = self._term_sheet_role(row_index, column_index, label_columns, key_value)
                texts, size = self._fill_term_sheet_cell(
                    word_cells[column_index], cell_data, role, row_index, column_index,
                    table, column_kinds,
                )
                if bottom == row_index:  # multi-row owners are excluded from row estimates
                    lines = estimate_cell_lines(
                        texts, sum(widths[column_index:right + 1]), size, markers=role == "content"
                    )
                    line_counts[row_index] = max(line_counts[row_index], lines)
        self._mirror_vertical_fills(word_table, rectangles)
        for row_index, word_row in enumerate(word_table.rows):
            header = row_index == 0
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

    @staticmethod
    def _term_sheet_role(row_index: int, column_index: int, label_columns: int, key_value: bool) -> str:
        """Classify an owner cell by the grid position where it starts."""
        if row_index == 0:
            return "header"
        if column_index >= label_columns:
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
        column_kinds: List[str],
    ) -> Tuple[List[str], Pt]:
        """Fill and style one owner cell exactly once, one paragraph per line.

        Header: header fill, bold accent, centred. Key-value label: label fill,
        bold black, centred; second label tier: unshaded, regular muted, centred.
        Grid-table labels: label fill, regular black, centred. Content cells keep
        column kinds, alignment markers and numeric roles, and indent marker lines.

        Returns:
            The cell's line texts and base text size, for row-split estimation.
        """
        column_role = table.column_types[column_index] if table.column_types else None
        fill: Optional[str] = None
        color: Optional[RGBColor] = None
        alignment = WD_ALIGN_PARAGRAPH.CENTER
        if role == "header":
            runs = [
                replace(run, bold=True, color_hex=None)
                for run in cell_data.runs or TextParser.parse_runs_plain(cell_data.content)
            ]
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
                if STYLE.TABLE_ZEBRA and row_index % 2 == 1 and not cell_data.is_base_case:
                    fill = STYLE.LIGHT_GRAY_HEX
        if fill:
            set_cell_fill(word_cell._tc, fill)
        word_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        texts: List[str] = []
        for index, line in enumerate(split_run_lines(runs)):
            paragraph = word_cell.paragraphs[0] if index == 0 else word_cell.add_paragraph()
            configure_cell_paragraph(paragraph)
            paragraph.alignment = alignment
            text = line_text(line)
            line_size = apply_marker_layout(paragraph, text) if role == "content" else None
            TextRenderer.render_runs(
                paragraph, line, default_color=color, font_name=font_name, font_size=line_size or size
            )
            texts.append(text)
        if role != "header" and not (
            table.table_type == TableType.FINANCIAL and column_role in {"text", "code", "date"}
        ):
            self._apply_type_styling(word_cell, cell_data, row_index, table.table_type)
        return texts, size

    @staticmethod
    def _set_column_widths(word_table, widths: List[int]) -> None:
        """Set exact grid and cell widths (EMU) on a table that is not merged yet."""
        word_table.autofit = False
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        for index, width in enumerate(widths):
            word_table.columns[index].width = Emu(width)
            for cell in word_table.columns[index].cells:
                cell.width = Emu(width)

    def _merge_cells(self, word_table, table: Table) -> List[Tuple[int, int, int, int]]:
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
        errors = getattr(self.doc.part, "_ib_render_errors", None)
        for problem in problems:
            logger.warning("%s", problem)
            if errors is not None:
                errors.append(problem)
        for top, left, bottom, right in rectangles:
            word_table.cell(top, left).merge(word_table.cell(bottom, right))
        return rectangles

    @staticmethod
    def _span_rectangles(table: Table) -> Tuple[List[Tuple[int, int, int, int]], List[str]]:
        """Group span directives into rectangles, rejecting malformed groups.

        Parsed tables are already validated by `TableSpanResolver`; this guard keeps
        hand-built models from producing ragged or overlapping Word merges. Header
        and body are classified as the renderer draws them (row 0 is the repeating
        header row), not by `TableRow.is_header`, which hand-built rows may omit.

        Args:
            table: Table whose cells may carry "up"/"left" merge directives.

        Returns:
            Valid rectangles and one diagnostic per rejected group, whose cells
            are then rendered unmerged with their own content.
        """
        roots: Dict[Tuple[int, int], Tuple[int, int]] = {}
        groups: Dict[Tuple[int, int], List[Tuple[int, int]]] = {}
        outside: Set[Tuple[int, int]] = set()
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
        rectangles: List[Tuple[int, int, int, int]] = []
        problems: List[str] = []
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
                and not (top == 0 and bottom > 0)
            ):
                rectangles.append((top, left, bottom, right))
            else:
                problems.append(
                    f"Invalid table span group at row {top + 1}, column {left + 1}; "
                    "cells rendered unmerged"
                )
        return rectangles, problems

    @staticmethod
    def _covered_cells(rectangles: List[Tuple[int, int, int, int]]) -> Set[Tuple[int, int]]:
        """Return grid positions represented by another cell's merged owner."""
        return {
            (row, column)
            for top, left, bottom, right in rectangles
            for row in range(top, bottom + 1)
            for column in range(left, right + 1)
            if (row, column) != (top, left)
        }

    @staticmethod
    def _mirror_vertical_fills(word_table, rectangles: List[Tuple[int, int, int, int]]) -> None:
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

    def _apply_column_widths(self, word_table, table: Table, column_kinds: List[str]) -> None:
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
        column_kinds: Optional[List[str]] = None,
    ) -> List[float]:
        """Estimate table column widths from content density and column semantics."""
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
        preferred = [column_info["score"] for column_info in column_infos]
        widths = self._fit_widths_to_available_space(
            preferred,
            min_widths,
            max_widths,
            available_width_inches,
        )

        return widths

    def _build_column_info(self, table: Table, col_idx: int, column_kind: str) -> Dict[str, Any]:
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

    def _infer_column_kinds(self, table: Table) -> List[str]:
        if table.column_types:
            return [
                "text" if role in {"text", "code", "date"} else "numeric"
                for role in table.column_types
            ]
        """Infer each column's semantic type from the first few data cells."""
        return [self._infer_column_kind(table, col_idx) for col_idx in range(table.col_count)]

    def _infer_column_kind(self, table: Table, col_idx: int) -> str:
        """Classify a column as text or numeric using the first 2-3 body cells."""
        sample_texts = []

        for row in table.rows[1:]:
            if col_idx >= len(row.cells) or self._is_span_cell(row, col_idx):
                continue

            text = self._cell_display_text(row.cells[col_idx], table.table_type).strip()
            if not text:
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
    def _representative_length(cls, texts: List[str]) -> int:
        """Return a robust content length that ignores a few extreme outliers."""
        if not texts:
            return 0

        lengths = sorted(len(text) for text in texts)
        percentile_index = int(round((len(lengths) - 1) * 0.75))
        return lengths[percentile_index]

    @classmethod
    def _longest_token_length(cls, texts: List[str]) -> int:
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
        preferred: List[float],
        min_widths: List[float],
        max_widths: List[float],
        available_width_inches: float,
    ) -> List[float]:
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
        widths: List[float],
        min_widths: List[float],
        max_widths: List[float],
        preferred: List[float],
        available_width_inches: float,
    ) -> List[float]:
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
        """Return display text used for width estimation."""
        if table_type == TableType.FINANCIAL and cell_data.is_numeric:
            return self._format_financial_number(cell_data.content)
        return cell_data.content

    def _render_header_row(
        self,
        word_table,
        row: TableRow,
        col_count: int,
        covered: AbstractSet[Tuple[int, int]] = frozenset(),
    ):
        """Render header row with Navy background, skipping merged-away cells."""
        word_cells = word_table.rows[0].cells

        for c_idx, cell_data in enumerate(row.cells):
            if c_idx >= col_count:
                break
            if (0, c_idx) in covered:
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
                    p, [replace(run, bold=True, color_hex=None) for run in cell_data.runs],
                    font_name=STYLE.HEADING_FONT, font_size=STYLE.TABLE_HEADER_SIZE,
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
        column_kinds: List[str],
        column_types: Optional[List[str]] = None,
        alignments: Optional[List[str]] = None,
        covered: AbstractSet[Tuple[int, int]] = frozenset(),
    ):
        """Render a data row with type-specific styling, skipping merged-away cells."""
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
            if STYLE.TABLE_ZEBRA and row_idx % 2 == 1 and not cell_data.is_base_case:
                TableStyler.set_cell_background(cell, STYLE.LIGHT_GRAY_HEX)

    @staticmethod
    def _cell_alignment(
        c_idx: int, column_kinds: List[str], alignments: Optional[List[str]]
    ) -> WD_ALIGN_PARAGRAPH:
        """Numeric columns align right; Markdown alignment markers take precedence."""
        if alignments and c_idx < len(alignments) and alignments[c_idx] in {"center", "right"}:
            return (
                WD_ALIGN_PARAGRAPH.CENTER if alignments[c_idx] == "center" else WD_ALIGN_PARAGRAPH.RIGHT
            )
        if c_idx < len(column_kinds) and column_kinds[c_idx] == "numeric":
            return WD_ALIGN_PARAGRAPH.RIGHT
        return WD_ALIGN_PARAGRAPH.LEFT

    def _display_content(
        self, cell_data: TableCell, role: Optional[str], table_type: TableType
    ) -> Tuple[str, List[TextRun]]:
        """Return display text and runs, formatting amounts only for numeric roles."""
        format_numbers = (
            role not in {None, "text", "code", "date"}
            or (not role and table_type == TableType.FINANCIAL and cell_data.is_numeric)
        )
        display_content = cell_data.content
        display_runs = cell_data.runs
        if format_numbers:
            display_content = self._format_numeric_text(display_content, role)
            # Semantic runs are independent of the amount. In particular, a
            # footnote's displayed number must never enter numeric formatting.
            display_runs = [
                run if (
                    run.footnote_id is not None or run.is_latex
                    or run.superscript or run.subscript
                ) else replace(run, text=self._format_numeric_text(run.text, role))
                for run in display_runs
            ]
        return display_content, display_runs

    @classmethod
    def _format_numeric_text(cls, text: str, role: Optional[str]) -> str:
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
        leading = text[:len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()):]
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
    ):
        """Apply table-type specific styling to every paragraph of one owner cell."""
        runs = [run for paragraph in cell.paragraphs for run in paragraph.runs]
        if table_type == TableType.FINANCIAL:
            if cell_data.is_negative:
                for run in runs:
                    run.font.color.rgb = STYLE.RED

        elif table_type == TableType.BEP_SENSITIVITY:
            if cell_data.is_base_case:
                TableStyler.set_cell_background(cell, STYLE.YELLOW_HEX)
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
        if content_text:
            content_para = cell.add_paragraph()

            # Determine text color based on background
            content_color = STYLE.WHITE if bg_hex == STYLE.NAVY_HEX else None

            TextRenderer.render_text_with_formatting(
                content_para,
                content_text,
                font_name=STYLE.BODY_FONT,
                font_size=STYLE.BODY_SIZE,
                default_color=content_color,
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

        left_border = OxmlElement("w:left")
        left_border.set(qn("w:val"), "single")
        left_border.set(qn("w:sz"), "32")
        left_border.set(qn("w:color"), border_hex)
        tblBorders.append(left_border)

        for border_name in ("top", "bottom", "right"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "nil")
            tblBorders.append(border)

        tblPr.append(tblBorders)
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

        tblPr.append(tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)


class ImageRenderer:
    """
    Renders images in Word documents.

    Supports:
        - Base64 embedded images (data:image/... URI)
        - File path images (local files)
        - Fallback to placeholder if image cannot be loaded
    """

    # Maximum image width in inches (fits within typical IB report margins)
    MAX_WIDTH_INCHES: float = 5.5

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, image: Image) -> bool:
        """
        Render an image to the Word document.

        Args:
            image: Image object with either base64_data or path

        Attempts to insert actual image; falls back to placeholder on failure.

        Returns:
            Whether an image was inserted successfully.
        """
        import base64
        import os
        import tempfile
        from pathlib import Path

        inserted = False
        temp_file_path = None

        try:
            # ── Case 1: Base64 embedded image ──────────────────────────────────
            if image.base64_data:
                # Decode Base64 to bytes
                img_bytes = base64.b64decode(image.base64_data)

                # Determine file extension from MIME type
                ext = self._mime_to_extension(image.mime_type)

                # Write to temporary file
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False, mode="wb") as f:
                    f.write(img_bytes)
                    temp_file_path = f.name

                # Insert into document
                self._insert_image(temp_file_path, image.alt_text)
                inserted = True
                logger.debug("Inserted Base64 image: %s", image.alt_text)

            # ── Case 2: File path image ────────────────────────────────────────
            elif image.path:
                img_path = Path(image.path)

                # Handle relative paths
                if not img_path.is_absolute():
                    # Try relative to current working directory
                    if not img_path.exists():
                        logger.warning(
                            "Image file not found: %s — inserting placeholder",
                            image.path,
                        )
                    else:
                        self._insert_image(str(img_path), image.alt_text)
                        inserted = True
                        logger.debug("Inserted file image: %s", image.path)
                elif img_path.exists():
                    self._insert_image(str(img_path), image.alt_text)
                    inserted = True
                    logger.debug("Inserted file image: %s", image.path)
                else:
                    logger.warning(
                        "Image file not found: %s — inserting placeholder",
                        image.path,
                    )

        except Exception as e:
            logger.warning(
                "Failed to insert image '%s': %s — inserting placeholder",
                image.alt_text,
                e,
            )

        finally:
            # Clean up temporary file
            if temp_file_path:
                try:
                    os.unlink(temp_file_path)
                except OSError:
                    pass

        # ── Fallback: Placeholder ──────────────────────────────────────────────
        if not inserted:
            self._render_placeholder(image.alt_text)
        return inserted

    def _insert_image(self, file_path: str, alt_text: str):
        """
        Insert image file into document with proper sizing.

        Args:
            file_path: Path to image file
            alt_text: Alt text for caption
        """
        # Add image with max width constraint
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        inline_shape = run.add_picture(file_path, width=Inches(self.MAX_WIDTH_INCHES))
        self._apply_alt_text(inline_shape, alt_text)

        # Add caption below image
        if alt_text:
            caption_para = self.doc.add_paragraph()
            caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_run = caption_para.add_run(alt_text)
            FontStyler.apply_run_style(
                caption_run,
                font_size=STYLE.SMALL_SIZE,
                italic=True,
                color=STYLE.DARK_GRAY,
            )

        # Spacer
        self.doc.add_paragraph()

    @staticmethod
    def _apply_alt_text(inline_shape, alt_text: str):
        """Attach descriptive alt text to a Word inline image when available."""
        if not alt_text:
            return

        doc_pr = inline_shape._inline.docPr
        doc_pr.set("descr", alt_text)
        doc_pr.set("title", alt_text)

    def _render_placeholder(self, alt_text: str):
        """Render a placeholder when image cannot be loaded."""
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
        run = p.add_run(f"[Image: {alt_text}]")
        FontStyler.apply_run_style(run, italic=True, color=STYLE.RED)

    @staticmethod
    def _mime_to_extension(mime_type: str) -> str:
        """Convert MIME type to file extension."""
        mime_map = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/jpg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/webp": ".webp",
            "image/svg+xml": ".svg",
            "image/tiff": ".tiff",
        }
        return mime_map.get(mime_type.lower(), ".png")


class FootnoteRenderer:
    """Renders footnotes section"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    def render(self, footnotes: dict, allow_legacy_refs: bool = True):
        """Render footnotes natively when possible, else fall back to an ENDNOTES section."""
        if not footnotes:
            return

        if self._render_native(footnotes, allow_legacy_refs):
            return

        self._render_endnotes(footnotes)

    def _render_native(self, footnotes: dict, allow_legacy_refs: bool = True) -> bool:
        """Replace superscript markers with native Word footnote references."""
        run_refs = self._collect_reference_runs(footnotes) if allow_legacy_refs else []
        explicit_ids = {int(node.get(qn("w:id"))) for node in self.doc.element.xpath(".//w:footnoteReference")}
        if not run_refs and not explicit_ids:
            return False

        footnotes_part = NativeFootnotesPart.get_or_add(self.doc.part)
        referenced_numbers = sorted({number for _, number in run_refs} | explicit_ids)
        footnotes_part.set_footnotes({number: footnotes[number] for number in referenced_numbers if number in footnotes})

        for run, number in run_refs:
            self._replace_with_native_reference(run, number)

        return True

    def _render_endnotes(self, footnotes: dict):
        """Render footnotes as a plain ENDNOTES section when native refs are unavailable."""
        self.doc.add_page_break()
        self.doc.add_heading("ENDNOTES", level=1)

        for num, text in sorted(footnotes.items()):
            p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)

            # Superscript number
            num_run = p.add_run(str(num))
            FontStyler.apply_run_style(
                num_run,
                font_size=STYLE.SMALL_SIZE,
                color=STYLE.NAVY,
                superscript=True,
            )

            # Space and text
            p.add_run(" ")
            text_run = p.add_run(text)
            FontStyler.apply_run_style(text_run, font_size=STYLE.SMALL_SIZE)

    def _collect_reference_runs(self, footnotes: dict) -> List[Tuple[object, int]]:
        """Find superscript numeric runs that correspond to known footnotes."""
        references: List[Tuple[object, int]] = []
        for paragraph in self._iter_document_paragraphs():
            for run in paragraph.runs:
                text = run.text.strip()
                if not text.isdigit():
                    continue
                if not run.font.superscript:
                    continue
                number = int(text)
                if number in footnotes:
                    references.append((run, number))
        return references

    def _iter_document_paragraphs(self):
        """Yield paragraphs from the document body and all top-level tables."""
        for paragraph in self.doc.paragraphs:
            yield paragraph
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        yield paragraph

    @staticmethod
    def _replace_with_native_reference(run, number: int) -> None:
        """Replace a rendered superscript number run with a native footnoteReference."""
        run_element = run._r
        r_pr = run_element.get_or_add_rPr()
        for child in list(run_element):
            if child is not r_pr:
                run_element.remove(child)

        r_style = OxmlElement("w:rStyle")
        r_style.set(qn("w:val"), "FootnoteReference")
        r_pr.append(r_style)

        footnote_ref = OxmlElement("w:footnoteReference")
        footnote_ref.set(qn("w:id"), str(number))
        run_element.append(footnote_ref)


class NativeFootnotesPart(XmlPart):
    """Minimal native footnotes part used for MD→Word footnote rendering."""

    _DEFAULT_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        b'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        b"</w:footnotes>"
    )

    @classmethod
    def default(cls, package) -> "NativeFootnotesPart":
        """Create a new empty footnotes part with required separator entries."""
        return cls(
            PackURI("/word/footnotes.xml"),
            CT.WML_FOOTNOTES,
            parse_xml(cls._DEFAULT_XML),
            package,
        )

    @classmethod
    def get_or_add(cls, document_part) -> "NativeFootnotesPart":
        """Get the existing footnotes part or create one and relate it to the document."""
        try:
            existing = document_part.part_related_by(RT.FOOTNOTES)
            return cast("NativeFootnotesPart", existing)
        except KeyError:
            footnotes_part = cls.default(document_part.package)
            document_part.relate_to(footnotes_part, RT.FOOTNOTES)
            return footnotes_part

    def set_footnotes(self, footnotes: Dict[int, str]) -> None:
        """Replace dynamic footnotes with the provided note mapping."""
        for footnote in list(self._element):
            if not str(footnote.tag).endswith("footnote"):
                continue
            footnote_id = footnote.get(qn("w:id"))
            if footnote_id not in {"-1", "0"}:
                self._element.remove(footnote)

        for number, text in sorted(footnotes.items()):
            self._element.append(self._build_footnote(number, text))

    @staticmethod
    def _build_footnote(number: int, text: str):
        """Build a single w:footnote element with simple paragraph content."""
        xml_space = "{http://www.w3.org/XML/1998/namespace}space"

        footnote = OxmlElement("w:footnote")
        footnote.set(qn("w:id"), str(number))

        paragraph = OxmlElement("w:p")
        paragraph_props = OxmlElement("w:pPr")
        paragraph_style = OxmlElement("w:pStyle")
        paragraph_style.set(qn("w:val"), "FootnoteText")
        paragraph_props.append(paragraph_style)
        paragraph.append(paragraph_props)

        ref_run = OxmlElement("w:r")
        ref_run_props = OxmlElement("w:rPr")
        ref_run_style = OxmlElement("w:rStyle")
        ref_run_style.set(qn("w:val"), "FootnoteReference")
        ref_run_props.append(ref_run_style)
        ref_run.append(ref_run_props)
        ref_marker = OxmlElement("w:footnoteRef")
        ref_run.append(ref_marker)
        paragraph.append(ref_run)

        text_run = OxmlElement("w:r")
        text_element = OxmlElement("w:t")
        text_element.set(xml_space, "preserve")
        text_element.text = f" {text}"
        text_run.append(text_element)
        paragraph.append(text_run)

        footnote.append(paragraph)
        return footnote


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
    def _split_content_lines(content: str) -> List[str]:
        """Split disclaimer content into non-empty logical paragraphs."""
        lines = [line.strip() for line in content.split("\n")]
        return [line for line in lines if line]


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT RENDERER (ORCHESTRATOR)
# ═══════════════════════════════════════════════════════════════════════════════


class ChartRenderer:
    """Insert a validated chart from memory, respecting the active section width."""

    def __init__(self, doc: DocxDocument) -> None:
        self.doc = doc

    def render(self, chart: Chart) -> None:
        """Render an enabled chart; the orchestrator handles diagnostic fallbacks.

        Args:
            chart: Parsed specification and original fence source.

        Raises:
            ValueError: Invalid chart data or unusable page geometry.
        """
        from chart_renderer import render_chart_png

        if chart.error or chart.spec is None:
            raise ValueError(chart.error or "chart specification is missing")
        section = self.doc.sections[-1]
        page_width = section.page_width
        left_margin, right_margin = section.left_margin, section.right_margin
        if page_width is None or left_margin is None or right_margin is None:
            raise ValueError("chart requires explicit page width and margins")
        width = min(6.0, (page_width - left_margin - right_margin) / 914400)
        if width <= 0:
            raise ValueError("chart requires a positive content width")
        png = render_chart_png(chart.spec, width_inches=width)
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            with BytesIO(png) as buffer:
                shape = paragraph.add_run().add_picture(buffer, width=Inches(width))
            ImageRenderer._apply_alt_text(shape, chart.spec.title or chart.label)
        except Exception:
            paragraph._p.getparent().remove(paragraph._p)
            raise


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
        options: Optional[RenderOptions] = None,
        profile: Optional[str] = None,
    ):
        self.options = options or RenderOptions(
            profile=profile, separator_mode=separator_mode if separator_mode != "auto" else None
        )
        self.separator_mode = separator_mode
        self.errors: List[str] = []
        self.charts = False
        self.term_sheet_texts: Optional[TermSheetTexts] = None
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
        self.table_renderer = TableRenderer(self.doc, term_sheet=self._term_sheet)
        self.callout_renderer = CalloutRenderer(self.doc)
        self.image_renderer = ImageRenderer(self.doc)
        self.footnote_renderer = FootnoteRenderer(self.doc)
        self.disclaimer_renderer = DisclaimerRenderer(self.doc)
        self.chart_renderer = ChartRenderer(self.doc)

    def render(self, model: DocumentModel) -> DocxDocument:
        """Render through one composition path, using request-local settings."""
        self.term_sheet_texts = None
        model = deepcopy(model)
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
        if not resolved.profile.is_ib:
            validate_office_metadata(model.metadata)
        if resolved.profile.name == "term-sheet":
            validate_term_sheet_metadata(model.metadata)
            self.term_sheet_texts = resolve_term_sheet_texts(model.metadata, resolved.house)
        report_title_block = (
            not resolved.cover and resolved.profile.name == "ib-report"
            and bool(model.metadata.title.strip())
        )
        if (resolved.cover and resolved.profile.is_ib) or report_title_block:
            # Filter the private render copy so the TOC and body share the same outline.
            model.elements = [element for element in model.elements if not element.inferred_subtitle]
        if report_title_block:
            # Transfer the first matching body H1 to the non-outline title block.
            # Filtering before the TOC also prevents a stale preview title entry.
            for index, element in enumerate(model.elements):
                if (
                    element.element_type == ElementType.HEADING_1
                    and isinstance(element.content, Heading)
                    and element.content.text == model.metadata.title
                ):
                    model.elements.pop(index)
                    break
        appendix_index = letter_appendix_index(model)
        if resolved.strict and model.warnings:
            raise ValueError("Input validation failed: " + "; ".join(model.warnings))
        self.errors = list(model.warnings)
        self._term_sheet = resolved.profile.name == "term-sheet"
        self._reset_document()
        self.separator_mode = resolved.separator_mode
        self.charts = resolved.charts
        with use_style(load_style(resolved.profile, resolved.theme)), collect_raster_font_diagnostics() as font_warnings:
            self.styler.setup_document()
            self.styler.create_styles()
            if resolved.profile.name == "office-letter":
                setup_letter_styles(self.doc)
            if self._term_sheet:
                setup_term_sheet_styles(self.doc)
            update_fields = OxmlElement("w:updateFields")
            update_fields.set(qn("w:val"), "true")
            self.doc.settings.element.append(update_fields)
            if resolved.profile.a4:
                for section in self.doc.sections:
                    section.page_width, section.page_height = Mm(210), Mm(297)
            if resolved.cover:
                self.cover_renderer.include_disclaimer = resolved.disclaimer
                self.cover_renderer.render(model.metadata)
            title_inserted = (
                render_office_opening(self.doc, model.metadata) if report_title_block else False
            )
            if self._term_sheet:
                # Resolved in preflight; the opening leads the first page, before any TOC.
                title_inserted = render_term_sheet_opening(
                    self.doc, model.metadata, cast(TermSheetTexts, self.term_sheet_texts),
                    TextRenderer.render_runs,
                )
            if resolved.toc:
                self.toc_renderer.render(model)
            if not resolved.cover and not report_title_block and not self._term_sheet:
                title_inserted = render_office_opening(self.doc, model.metadata)
            skipped_title = report_title_block
            office_closed = False
            for idx, element in enumerate(model.elements):
                if idx == appendix_index:
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
                render_office_closing(self.doc, model.metadata)
            if model.footnotes:
                self.footnote_renderer.render(model.footnotes, allow_legacy_refs=resolved.profile.is_ib)
            if resolved.disclaimer:
                self.disclaimer_renderer.render(model.metadata.company)
            sender = model.metadata.extra.get("sender", {})
            company = (
                sender.get("organization", model.metadata.company)
                if isinstance(sender, dict)
                else model.metadata.company
            )
            if self._term_sheet:
                setup_term_sheet_header_footer(
                    self.doc, model.metadata, self.term_sheet_texts, resolved.confidential,
                    TextRenderer.render_runs,
                )
            else:
                self.styler.setup_header_footer(
                    company="" if resolved.profile.name == "office-letter" else company,
                    confidential=resolved.confidential, show_page_numbers=True
                )
            self.apply_generator_signature(
                "ib_generated" if resolved.profile.name == "ib-report" else resolved.profile.name
            )
            from docx_audit import inspect_document_issues

            for warning in font_warnings:
                self.errors.append(warning)
                paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                FontStyler.apply_run_style(
                    paragraph.add_run(f"[Render Warning: {warning}]"), italic=True, color=STYLE.RED,
                )
            issues = inspect_document_issues(self.doc)
            self.errors.extend(issue for issue in issues if issue not in self.errors)
            if resolved.strict and self.errors:
                raise ValueError("Document validation failed: " + "; ".join(self.errors))
        return self.doc

    def apply_generator_signature(self, profile: str = "ib_generated") -> None:
        """Stamp the DOCX package with a generator signature."""
        GeneratorSignatureWriter.apply(self.doc, profile)

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
            content = cast(Tuple[str, ListItem], element.content)
            number, item = content
            self.list_renderer.render_numbered(number, item)

        elif etype == ElementType.TABLE:
            self.table_renderer.render(cast(Table, element.content))

        elif etype == ElementType.BLOCKQUOTE:
            self.callout_renderer.render(cast(Blockquote, element.content))

        elif etype == ElementType.IMAGE:
            if not self.image_renderer.render(cast(Image, element.content)):
                self.errors.append("Image could not be rendered")

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
            self._render_code_block(CodeBlock(chart.code, "chart", CodeBlock.detect_ascii_art(chart.code)))

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
        pPr.append(pBdr)

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

        tbl_pr.append(tbl_borders)
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
        tc_pr.append(tc_mar)

    def _render_code_block_line(self, paragraph, line: str) -> None:
        """Render one line of a code block with light semantic emphasis for tree diagrams."""
        if not line.strip():
            spacer = paragraph.add_run(" ")
            FontStyler.apply_run_style(
                spacer,
                font_name="Consolas",
                font_size=Pt(9.5),
                color=STYLE.MEDIUM_GRAY,
            )
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


# ═══════════════════════════════════════════════════════════════════════════════
# LATEX RENDERER (NEW)
# ═══════════════════════════════════════════════════════════════════════════════


class LaTeXRenderer:
    """
    Renders LaTeX expressions to PNG images using matplotlib.

    This class provides static methods to convert LaTeX math expressions
    into image files suitable for embedding in Word documents.

    Dependencies:
        - matplotlib (optional, graceful degradation if unavailable)
    """

    # Flag to track if matplotlib is available
    _matplotlib_available: Optional[bool] = None
    _NON_ASCII_RE = re.compile(r"[^\x00-\x7F]")
    _TEXT_COMMAND_RE = re.compile(r"\\(?:text|mathrm|operatorname)\{([^{}]*)\}")
    _FRAC_RE = re.compile(r"\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}")
    _MULTISPACE_RE = re.compile(r"\s+")
    _UNRESOLVED_COMMAND_RE = re.compile(r"\\[A-Za-z]+")
    _LATEX_SYMBOL_REPLACEMENTS: Tuple[Tuple[str, str], ...] = (
        (r"\Leftrightarrow", "⇔"),
        (r"\Leftarrow", "⇐"),
        (r"\Rightarrow", "⇒"),
        (r"\leftrightarrow", "↔"),
        (r"\leftarrow", "←"),
        (r"\rightarrow", "→"),
        (r"\approx", "≈"),
        (r"\alpha", "α"),
        (r"\beta", "β"),
        (r"\gamma", "γ"),
        (r"\delta", "δ"),
        (r"\epsilon", "ε"),
        (r"\varepsilon", "ε"),
        (r"\theta", "θ"),
        (r"\lambda", "λ"),
        (r"\mu", "μ"),
        (r"\pi", "π"),
        (r"\rho", "ρ"),
        (r"\sigma", "σ"),
        (r"\tau", "τ"),
        (r"\phi", "φ"),
        (r"\varphi", "φ"),
        (r"\omega", "ω"),
        (r"\Gamma", "Γ"),
        (r"\Delta", "Δ"),
        (r"\Theta", "Θ"),
        (r"\Lambda", "Λ"),
        (r"\Pi", "Π"),
        (r"\Sigma", "Σ"),
        (r"\Phi", "Φ"),
        (r"\Omega", "Ω"),
        (r"\cdot", "·"),
        (r"\times", "×"),
        (r"\sqrt", "√"),
        (r"\infty", "∞"),
        (r"\prod", "∏"),
        (r"\sum", "∑"),
        (r"\int", "∫"),
        (r"\geq", "≥"),
        (r"\leq", "≤"),
        (r"\neq", "≠"),
        (r"\pm", "±"),
        (r"\to", "→"),
        (r"\left", ""),
        (r"\right", ""),
        (r"\%", "%"),
    )
    _LATEX_SYMBOL_MAP = dict(_LATEX_SYMBOL_REPLACEMENTS)
    _LATEX_SYMBOL_RE = re.compile(
        "|".join(
            re.escape(source) + r"(?![A-Za-z])"
            for source, _ in sorted(
                _LATEX_SYMBOL_REPLACEMENTS,
                key=lambda item: len(item[0]),
                reverse=True,
            )
        )
    )

    @classmethod
    def is_available(cls) -> bool:
        """Check if matplotlib is available for LaTeX rendering."""
        if cls._matplotlib_available is None:
            try:
                from importlib.util import find_spec

                cls._matplotlib_available = find_spec("matplotlib") is not None
            except ImportError:
                cls._matplotlib_available = False
            if not cls._matplotlib_available:
                logger.info(
                    "matplotlib not installed — LaTeX will render as text. "
                    "Install with: pip install matplotlib"
                )
        return cls._matplotlib_available

    @classmethod
    def render_to_image(
        cls,
        expression: str,
        fontsize: int = 14,
        dpi: int = 150,
    ) -> Optional[str]:
        """
        Render a LaTeX expression to a PNG image file.

        Args:
            expression: LaTeX math expression (without $ delimiters)
            fontsize: Font size for the equation
            dpi: Resolution of the output image

        Returns:
            Path to the temporary PNG file, or None if rendering fails
        """
        if not cls.is_available():
            return None

        if cls._NON_ASCII_RE.search(expression):
            display_text = cls.to_display_text(expression)
            if cls._UNRESOLVED_COMMAND_RE.search(display_text):
                return None
            image_path = cls._render_plain_text_to_image(display_text, fontsize=fontsize, dpi=dpi)
            if image_path:
                return image_path

        image_path = cls._render_mathtext_to_image(expression, fontsize=fontsize, dpi=dpi)
        if image_path:
            return image_path

        # A failed mathtext parse is a rendering failure. Removing braces and
        # drawing the remaining command as text cannot establish success.
        return None

    @classmethod
    def to_display_text(cls, expression: str) -> str:
        """Convert a LaTeX expression into a readable unicode text fallback."""
        display = expression.strip()

        previous = None
        while previous != display:
            previous = display
            display = cls._TEXT_COMMAND_RE.sub(lambda match: match.group(1), display)
            display = cls._FRAC_RE.sub(
                lambda match: f"({match.group(1)} / {match.group(2)})",
                display,
            )

        display = cls._LATEX_SYMBOL_RE.sub(
            lambda match: cls._LATEX_SYMBOL_MAP[match.group(0)],
            display,
        )

        display = display.replace("{", "").replace("}", "")
        display = cls._MULTISPACE_RE.sub(" ", display)
        return display.strip()

    @classmethod
    def _render_mathtext_to_image(
        cls,
        expression: str,
        fontsize: int,
        dpi: int,
    ) -> Optional[str]:
        """Render mathtext-compatible LaTeX to a PNG image."""
        fig = None
        temp_path = None
        succeeded = False
        try:
            import tempfile

            from matplotlib.backends.backend_agg import FigureCanvasAgg
            from matplotlib.figure import Figure

            fig = Figure(figsize=(0.01, 0.01))
            FigureCanvasAgg(fig)
            ax = fig.subplots()
            fig.patch.set_alpha(0)
            ax.set_axis_off()

            latex_text = f"${expression}$"

            ax.text(
                0.5,
                0.5,
                latex_text,
                fontsize=fontsize,
                ha="center",
                va="center",
                transform=ax.transAxes,
            )

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False, mode="wb") as temp_file:
                temp_path = temp_file.name

            fig.savefig(
                temp_path,
                dpi=dpi,
                bbox_inches="tight",
                pad_inches=0.1,
                transparent=False,
                facecolor="white",
            )
            succeeded = True
            logger.debug("Rendered LaTeX to: %s", temp_path)
            return temp_path

        except Exception as err:
            logger.warning("LaTeX rendering failed: %s", err)
            return None
        finally:
            if fig is not None:
                fig.clear()
            if temp_path is not None and not succeeded:
                try:
                    os.unlink(temp_path)
                except OSError as err:
                    logger.warning("Cannot remove unsuccessful equation image: %s", err)

    @classmethod
    def _render_plain_text_to_image(
        cls,
        display_text: str,
        fontsize: int,
        dpi: int,
    ) -> Optional[str]:
        """Render a readable unicode fallback image for non-ASCII equations."""
        fig = None
        temp_path = None
        succeeded = False
        try:
            import tempfile

            from matplotlib.backends.backend_agg import FigureCanvasAgg
            from matplotlib.figure import Figure
            from matplotlib.font_manager import FontProperties

            fig = Figure(figsize=(0.01, 0.01))
            FigureCanvasAgg(fig)
            fig.patch.set_facecolor("white")
            fig.patch.set_alpha(1)
            font_props = FontProperties(family=RasterFontPolicy.resolve(STYLE.KOREAN_FONT, display_text))

            fig.text(
                0.5,
                0.5,
                display_text,
                fontsize=fontsize,
                ha="center",
                va="center",
                fontproperties=font_props,
            )

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False, mode="wb") as temp_file:
                temp_path = temp_file.name

            fig.savefig(
                temp_path,
                dpi=dpi,
                bbox_inches="tight",
                pad_inches=0.12,
                transparent=False,
                facecolor="white",
            )
            succeeded = True
            logger.debug("Rendered unicode equation to: %s", temp_path)
            return temp_path

        except Exception as err:
            logger.warning("Unicode equation rendering failed: %s", err)
            return None
        finally:
            if fig is not None:
                fig.clear()
            if temp_path is not None and not succeeded:
                try:
                    os.unlink(temp_path)
                except OSError as err:
                    logger.warning("Cannot remove unsuccessful equation image: %s", err)
