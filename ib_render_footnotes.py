"""Footnotes, the native footnotes part and local link rebasing for rendered documents."""

from pathlib import Path
from typing import cast

from docx.document import Document as DocxDocument
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml

from ib_render_fonts import FontStyler
from md_parser import rebase_link_target
from ooxml_order import insert_ordered
from render_styles import STYLE
from term_variables import is_term_run


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
        explicit_ids = {
            int(node.get(qn("w:id"))) for node in self.doc.element.xpath(".//w:footnoteReference")
        }
        if not run_refs and not explicit_ids:
            return False

        footnotes_part = NativeFootnotesPart.get_or_add(self.doc.part)
        referenced_numbers = sorted({number for _, number in run_refs} | explicit_ids)
        footnotes_part.set_footnotes(
            {number: footnotes[number] for number in referenced_numbers if number in footnotes}
        )

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

    def _collect_reference_runs(self, footnotes: dict) -> list[tuple[object, int]]:
        """Find superscript numeric runs that correspond to known footnotes.

        A term value is text the author wrote in `terms:`, never a footnote marker.
        """
        references: list[tuple[object, int]] = []
        for paragraph in self._iter_document_paragraphs():
            for run in paragraph.runs:
                if is_term_run(run._r):
                    continue
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
        insert_ordered(r_pr, r_style)

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

    def set_footnotes(self, footnotes: dict[int, str]) -> None:
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


# ═══════════════════════════════════════════════════════════════════════════════
# SAVING
# ═══════════════════════════════════════════════════════════════════════════════


def rebase_local_links(document: DocxDocument, source_dir: Path | None, output_path: Path) -> None:
    """Point local hyperlinks of a rendered document at its output folder.

    Relative links written against the Markdown folder then open from wherever
    the DOCX is saved, and files inside the output folder are linked relatively
    (see `rebase_link_target`). Call once per rendered document, right before it
    is saved; a document without a source folder is left unchanged.

    Args:
        document: Rendered document, not yet saved.
        source_dir: `DocumentModel.source_dir` of the rendered model.
        output_path: Requested DOCX path; a lock fallback stays in its folder.
    """
    if source_dir is None:
        return
    # Resolve the folder, not the file: saving replaces an existing output symlink.
    output_dir = Path(output_path).parent.resolve()
    for part in document.part.package.iter_parts():
        for relationship in part.rels.values():
            if relationship.reltype == RT.HYPERLINK and relationship.is_external:
                relationship._target = rebase_link_target(
                    relationship.target_ref, source_dir, output_dir
                )
