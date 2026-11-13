"""Document styles, the generator signature and table styling for rendered documents."""

import re
from collections.abc import Mapping
from copy import deepcopy
from datetime import UTC, datetime
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from xml.sax.saxutils import escape

from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart, serialize_part_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import Inches, Pt, RGBColor

from ib_render_fonts import FontStyler
from ooxml_order import insert_ordered
from render_styles import STYLE

# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT SIGNATURE
# ═══════════════════════════════════════════════════════════════════════════════


class GeneratorSignatureWriter:
    """Embed an explicit generator signature and other string DOCX custom properties."""

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
    _PYPROJECT_VERSION_RE = re.compile(r'^version\s*=\s*"([^"]+)"\s*$')
    _resolved_version: str | None = None

    @classmethod
    def apply(cls, doc: DocxDocument, profile: str = "ib_generated") -> None:
        """Upsert generator signature custom properties on a DOCX package."""
        cls.upsert_properties(doc, cls._signature_properties(profile))

    @classmethod
    def upsert_properties(
        cls,
        doc: DocxDocument,
        properties: Mapping[str, str],
        owned_prefix: str = "",
    ) -> None:
        """Replace named string properties, preserving all others and their PIDs.

        Args:
            doc: Document whose package receives the custom-properties part.
            properties: Ordered property names and string values; XML is escaped.
            owned_prefix: Also remove existing properties with this prefix, for a
                family written as a whole (such as the term snapshot).
        """
        custom_part, custom_props = cls._get_or_add_custom_properties_part(doc)
        cls._upsert_properties(custom_props, properties, owned_prefix)
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
    def _upsert_properties(
        cls,
        custom_props,
        properties: Mapping[str, str],
        owned_prefix: str = "",
    ) -> None:
        """Replace only the given (and owned-prefix) properties, preserving others."""
        for prop in list(custom_props):
            if cls._local_name(prop.tag) != "property":
                continue
            name = (prop.get("name") or "").strip()
            if name in properties or (owned_prefix and name.startswith(owned_prefix)):
                custom_props.remove(prop)

        used_pids = cls._used_property_ids(custom_props)
        for name, value in properties.items():
            custom_props.append(cls._build_property(name, value, cls._next_pid(used_pids)))

    @classmethod
    def _signature_properties(cls, profile: str = "ib_generated") -> dict[str, str]:
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
    def _used_property_ids(custom_props) -> list[int]:
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
    def _next_pid(used_pids: list[int]) -> int:
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
# CORE PROPERTIES
# ═══════════════════════════════════════════════════════════════════════════════


def write_core_properties(doc: DocxDocument, title: str) -> None:
    """Replace python-docx's template file information with this document's.

    Without this, Word's File > Info shows the template's author "python-docx",
    the comment "generated by python-docx" and 2013 dates. The author is left
    empty rather than guessed; Word records who saves the file next.

    Args:
        doc: Document being composed.
        title: Document title for the file properties; "" leaves it empty.
    """
    properties = doc.core_properties
    now = datetime.now(UTC).replace(tzinfo=None, microsecond=0)
    properties.author = ""
    properties.comments = ""
    properties.last_modified_by = ""
    properties.created = now
    properties.modified = now
    properties.title = title


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT PROPERTY FIELDS
# ═══════════════════════════════════════════════════════════════════════════════

# Custom document property that holds the frontmatter `version`.
VERSION_PROPERTY = "Version"
_PROPERTY_NAME_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]*")


def make_property_field(run, name: str) -> None:
    """Turn a rendered run into the stored result of a `DOCPROPERTY` field.

    Word shows the stored result until the field is updated (F9 inside the
    header or footer, the print preview, printing or saving as PDF), so the run
    must already hold the property's current value. The field code matches what
    Word itself writes, and every field-character run copies the result run's
    formatting so an update keeps the look.

    Args:
        run: Run already added to its paragraph and styled.
        name: Custom property name; letters, digits and underscores only.
    """
    if not _PROPERTY_NAME_RE.fullmatch(name):
        raise ValueError(f"Unsupported document property name: {name!r}")
    result = run._r
    rpr = result.find(qn("w:rPr"))

    def field_run(kind: str | None, instruction: str | None = None):
        element = OxmlElement("w:r")
        if rpr is not None:
            element.append(deepcopy(rpr))
        if kind is None:
            code = OxmlElement("w:instrText")
            code.set(qn("xml:space"), "preserve")
            code.text = instruction
            element.append(code)
        else:
            char = OxmlElement("w:fldChar")
            char.set(qn("w:fldCharType"), kind)
            element.append(char)
        return element

    result.addprevious(field_run("begin"))
    result.addprevious(field_run(None, f" DOCPROPERTY {name} \\* MERGEFORMAT "))
    result.addprevious(field_run("separate"))
    result.addnext(field_run("end"))


# ═══════════════════════════════════════════════════════════════════════════════
# DRAFT WATERMARK
# ═══════════════════════════════════════════════════════════════════════════════

# Word's Design > Watermark > Remove Watermark deletes header shapes whose VML id
# starts with this prefix, in every header of every section (Word 16.0.20527).
WATERMARK_SHAPE_PREFIX = "PowerPlusWaterMarkObject"
DRAFT_WATERMARK_TEXT = "DRAFT"
_VML_NAMESPACES = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word"'
)
# The markup Word itself saves for a custom text watermark: a VML text-path
# shape behind the text, centred on the margins, silver at 50% opacity.
_WATERMARK_RUN = (
    "<w:r {ns}><w:rPr><w:noProof/></w:rPr><w:pict>"
    '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
    'path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/>'
    '<v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
    '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/>'
    '<v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/>'
    '<v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/>'
    '<v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" '
    'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
    '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" '
    'xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/>'
    "</v:shapetype>"
    '<v:shape id="{shape_id}" o:spid="_x0000_s{spid}" type="#_x0000_t136" '
    'style="position:absolute;margin-left:0;margin-top:0;width:456pt;height:180.25pt;'
    "rotation:315;z-index:-251653120;mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
    'mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f">'
    '<v:fill opacity=".5"/>'
    '<v:textpath style="font-family:&quot;Malgun Gothic&quot;;font-size:1pt" string="{text}"/>'
    '<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>'
)


def add_draft_watermark(doc: DocxDocument) -> int:
    """Put Word's own DRAFT watermark into every header of every section.

    Word's watermark gallery fills only the header where the cursor is, so a
    document with one header per section (landscape tables) would be marked on
    its first section alone. Each header part gets its own shape with a unique
    `PowerPlusWaterMarkObject` id, which Word's Remove Watermark recognises.

    Args:
        doc: Document whose sections and headers are final.

    Returns:
        Number of header parts that received a watermark.
    """
    seen: set[int] = set()
    count = 0
    for section in doc.sections:
        stories = [section.header]
        if section.different_first_page_header_footer:
            stories.append(section.first_page_header)
        if doc.settings.odd_and_even_pages_header_footer:
            stories.append(section.even_page_header)
        for story in stories:
            # A linked header shows the previous section's part, already marked.
            if story.is_linked_to_previous or id(story.part) in seen:
                continue
            seen.add(id(story.part))
            count += 1
            paragraph = story.paragraphs[0] if story.paragraphs else story.add_paragraph()
            run = parse_xml(
                _WATERMARK_RUN.format(
                    ns=_VML_NAMESPACES,
                    shape_id=f"{WATERMARK_SHAPE_PREFIX}{count}",
                    spid=1024 + count,
                    text=DRAFT_WATERMARK_TEXT,
                ).encode("utf-8")
            )
            p = paragraph._p
            if p.pPr is not None:
                p.pPr.addnext(run)
            else:
                p.insert(0, run)
    return count


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
        insert_ordered(pPr, pBdr)

    def setup_header_footer(
        self,
        company: str = "",
        confidential: bool = True,
        show_page_numbers: bool = True,
        version: str = "",
    ):
        """
        Set up professional IB-style header and footer.

        Args:
            company: Company name to display in header
            confidential: Whether to show "CONFIDENTIAL" mark
            show_page_numbers: Whether to show page numbers in footer
            version: Version shown after the company as a `DOCPROPERTY Version`
                field whose stored result is this text; empty shows nothing. The
                caller writes the property itself.
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

            # Version after the company, kept current by a document-property field.
            if version:
                labels = [" | ", version] if company else [version]
                runs = [header_para.add_run(label) for label in labels]
                for run in runs:
                    FontStyler.apply_run_style(
                        run,
                        font_name=STYLE.HEADING_FONT,
                        font_size=STYLE.SMALL_SIZE,
                        color=STYLE.DARK_GRAY,
                    )
                make_property_field(runs[-1], VERSION_PROPERTY)

            # Tab width follows this section, including landscape tables.
            width, left, right = section.page_width, section.left_margin, section.right_margin
            if width is None or left is None or right is None:
                raise ValueError("Section dimensions are required for header alignment")
            available_width = width - left - right
            header_style = header_para.style
            if header_style is not None:
                header_style.paragraph_format.tab_stops.clear_all()
            header_para.paragraph_format.tab_stops.clear_all()
            header_para.paragraph_format.tab_stops.add_tab_stop(
                available_width, WD_TAB_ALIGNMENT.RIGHT
            )
            if company or version or confidential:
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
            if company or version or confidential:
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
        insert_ordered(pPr, pBdr)

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
        # One fill, in schema order (before w:vAlign); a second fill replaces the first.
        insert_ordered(tcPr, shd)

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

        insert_ordered(tblPr, tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)
