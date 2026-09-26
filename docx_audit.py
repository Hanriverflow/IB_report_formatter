"""Inspect generated DOCX structure without an inverse Word-to-Markdown parser."""

import argparse
import json
import logging
from dataclasses import asdict, dataclass, field
from typing import List

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

logger = logging.getLogger(__name__)


@dataclass
class DocumentAudit:
    """Structural observations and actionable errors, not a visual QA verdict."""

    paragraphs: int = 0
    tables: int = 0
    headings: int = 0
    images: int = 0
    numbered_paragraphs: int = 0
    issues: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    pagination_marked_paragraphs: int = 0
    visual_review: str = "not_performed"


def _effective_pagination_flag(paragraph, attribute: str, xml_tag: str) -> bool:
    """Resolve local > paragraph style ancestry > document default controls."""
    value = getattr(paragraph.paragraph_format, attribute)
    style = paragraph.style
    visited = set()
    while value is None and style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        value = getattr(style.paragraph_format, attribute)
        style = style.base_style
    if value is not None:
        return bool(value)
    default = paragraph.part.document.styles.element.xpath(
        "./w:docDefaults/w:pPrDefault/w:pPr/w:" + xml_tag
    )
    return bool(default and default[0].get(qn("w:val"), "1") not in {"0", "false", "off"})


def inspect_document(doc) -> DocumentAudit:
    """Inspect an open python-docx document, including text inside tables."""
    result = DocumentAudit(
        paragraphs=len(doc.paragraphs),
        tables=len(doc.tables),
        headings=sum(p.style.name.startswith("Heading ") for p in doc.paragraphs),
        images=len(doc.inline_shapes),
        numbered_paragraphs=len(doc.element.xpath(".//w:pPr/w:numPr")),
    )
    for element in doc.element.xpath(".//w:p"):
        paragraph = Paragraph(element, doc)
        if any(
            _effective_pagination_flag(paragraph, attribute, tag)
            for attribute, tag in (
                ("keep_together", "keepLines"),
                ("keep_with_next", "keepNext"),
                ("page_break_before", "pageBreakBefore"),
            )
        ):
            result.pagination_marked_paragraphs += 1
    normal = doc.styles["Normal"].paragraph_format
    if normal.keep_together or normal.keep_with_next or normal.page_break_before:
        result.warnings.append(
            "Normal style has pagination constraints: these can show nonprinting square "
            "marks and create large page gaps. Keep such controls on necessary headings only."
        )
    for text in doc.element.xpath(".//w:t/text()"):
        if "[Render Error:" in text or text.startswith(("[Image:", "[Diagram:")):
            result.issues.append("Unresolved rendering placeholder: " + text)
    for index, section in enumerate(doc.sections, 1):
        if section.page_width <= section.left_margin + section.right_margin:
            result.issues.append(f"Section {index} has no printable width")
    references = doc.element.xpath(".//w:footnoteReference")
    if references:
        try:
            part = doc.part.part_related_by(RT.FOOTNOTES)
            notes = etree.fromstring(part.blob)
            ids = {node.get(qn("w:id")) for node in notes}
            for reference in references:
                if reference.get(qn("w:id")) not in ids:
                    result.issues.append("Footnote reference has no definition")
        except KeyError:
            result.issues.append("Footnote definitions part is missing")
    for hyperlink in doc.element.xpath(".//w:hyperlink"):
        relation = hyperlink.get(qn("r:id"))
        if relation and relation not in doc.part.rels:
            result.issues.append("Hyperlink relationship is missing")
    return result


def audit_file(path: str) -> DocumentAudit:
    """Open an output file and inspect its structure."""
    return inspect_document(Document(path))


def main() -> None:
    """Print a JSON audit result and return nonzero on structural errors."""
    parser = argparse.ArgumentParser(description="Inspect generated Word document structure")
    parser.add_argument("input_file")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    try:
        result = audit_file(args.input_file)
    except Exception as exc:
        logger.error("Cannot inspect DOCX: %s", exc)
        raise SystemExit(1) from exc
    logger.info(json.dumps(asdict(result), ensure_ascii=False, indent=2))
    raise SystemExit(1 if result.issues else 0)


if __name__ == "__main__":
    main()
