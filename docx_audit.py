"""Inspect generated DOCX structure without an inverse Word-to-Markdown parser.

Changelog (hardening):
    - Index styles once per audit and separate render validation from observations.
    - Treat placeholder-like text as warnings and validate native numbering references.
"""

import argparse
import json
import logging
from dataclasses import asdict, dataclass, field
from typing import Dict, List, Optional, Tuple, cast

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.parts.numbering import NumberingPart
from docx.styles.style import ParagraphStyle
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


_PAGINATION_FLAGS = (
    ("keep_together", "keepLines"),
    ("keep_with_next", "keepNext"),
    ("page_break_before", "pageBreakBefore"),
)


class _StyleLookup:
    """Audit-local style index and inherited pagination cache."""

    def __init__(self, doc: DocxDocument) -> None:
        self.styles = {style.style_id: style for style in doc.styles}
        self.default = next((
            style for style in self.styles.values()
            if style.type == WD_STYLE_TYPE.PARAGRAPH
            and style.element.get(qn("w:default")) in {"1", "true", "on"}
        ), None)
        self.defaults: Dict[str, bool] = {}
        for _, tag in _PAGINATION_FLAGS:
            nodes = doc.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr/w:" + tag)
            self.defaults[tag] = bool(
                nodes and nodes[0].get(qn("w:val"), "1") not in {"0", "false", "off"}
            )
        self.flags: Dict[Tuple[Optional[str], str], bool] = {}

    def paragraph_style(self, paragraph: Paragraph) -> Optional[ParagraphStyle]:
        """Resolve a paragraph style, including the default for unknown IDs."""
        style = self.styles.get(paragraph._p.style)
        if style is None or style.type != WD_STYLE_TYPE.PARAGRAPH:
            style = self.default
        return style

    def inherited_flag(self, paragraph: Paragraph, attribute: str, tag: str) -> bool:
        """Resolve each style's ancestry only once for each pagination control."""
        style = self.paragraph_style(paragraph)
        key = (style.style_id if style is not None else None, attribute)
        if key not in self.flags:
            value = None
            visited = set()
            while value is None and style is not None and style.style_id not in visited:
                visited.add(style.style_id)
                value = getattr(style.paragraph_format, attribute)
                style = self.styles.get(style.element.basedOn_val)
            self.flags[key] = bool(value) if value is not None else self.defaults[tag]
        return self.flags[key]


def _effective_pagination_flag(
    paragraph: Paragraph, attribute: str, xml_tag: str, styles: _StyleLookup,
) -> bool:
    """Resolve local > paragraph style ancestry > document default controls."""
    value = getattr(paragraph.paragraph_format, attribute)
    return bool(value) if value is not None else styles.inherited_flag(paragraph, attribute, xml_tag)


def inspect_document(doc: DocxDocument) -> DocumentAudit:
    """Inspect structure and observations, including text inside tables.

    Args:
        doc: Open python-docx document to inspect without modifying it.

    Returns:
        Structural issues and descriptive counts, warnings and review status.
    """
    styles = _StyleLookup(doc)
    paragraph_styles = (styles.paragraph_style(p) for p in doc.paragraphs)
    result = DocumentAudit(
        paragraphs=len(doc.paragraphs),
        tables=len(doc.tables),
        headings=sum(style is not None and style.name.startswith("Heading ") for style in paragraph_styles),
        images=len(doc.inline_shapes),
        numbered_paragraphs=len(doc.element.xpath(".//w:pPr/w:numPr")),
        issues=inspect_document_issues(doc),
    )
    for element in doc.element.xpath(".//w:p"):
        paragraph = Paragraph(element, doc)
        if any(
            _effective_pagination_flag(paragraph, attribute, tag, styles)
            for attribute, tag in _PAGINATION_FLAGS
        ):
            result.pagination_marked_paragraphs += 1
    normal = doc.styles["Normal"].paragraph_format
    if normal.keep_together or normal.keep_with_next or normal.page_break_before:
        result.warnings.append(
            "Normal style has pagination constraints: these can show nonprinting square "
            "marks and create large page gaps. Keep such controls on necessary headings only."
        )
    # Text alone cannot distinguish renderer fallback markers from quoted user
    # examples. Real rendering failures are recorded by the renderer itself.
    for text in doc.element.xpath(".//w:t/text()"):
        if "[Render Error:" in text or text.startswith(("[Image:", "[Diagram:")):
            result.warnings.append("Possible rendering placeholder (may be literal user text): " + text)
    return result


def inspect_document_issues(doc: DocxDocument) -> List[str]:
    """Validate structure without computing descriptive audit observations.

    Args:
        doc: Open python-docx document to validate before saving.

    Returns:
        Structural errors suitable for strict render validation.
    """
    issues: List[str] = []
    for index, section in enumerate(doc.sections, 1):
        width, left, right = section.page_width, section.left_margin, section.right_margin
        if width is not None and left is not None and right is not None and width <= left + right:
            issues.append(f"Section {index} has no printable width")
    references = doc.element.xpath(".//w:footnoteReference")
    if references:
        try:
            part = doc.part.part_related_by(RT.FOOTNOTES)
            notes = etree.fromstring(part.blob)
            ids = {node.get(qn("w:id")) for node in notes}
            for reference in references:
                if reference.get(qn("w:id")) not in ids:
                    issues.append("Footnote reference has no definition")
        except KeyError:
            issues.append("Footnote definitions part is missing")
    for hyperlink in doc.element.xpath(".//w:hyperlink"):
        relation = hyperlink.get(qn("r:id"))
        if relation and relation not in doc.part.rels:
            issues.append("Hyperlink relationship is missing")
    issues.extend(_numbering_issues(doc))
    return issues


def _numbering_issues(doc: DocxDocument) -> List[str]:
    """Validate numbering definitions and numPr references across XML parts."""
    issues: List[str] = []
    definitions = {}
    try:
        numbering = cast(NumberingPart, doc.part.part_related_by(RT.NUMBERING)).element
    except KeyError:
        numbering = None
    if numbering is not None:
        abstracts = {node.get(qn("w:abstractNumId")) for node in numbering.findall(qn("w:abstractNum"))}
        definitions = {node.get(qn("w:numId")): node for node in numbering.findall(qn("w:num"))}
        for num_id, node in definitions.items():
            abstract = node.find(qn("w:abstractNumId"))
            abstract_id = abstract.get(qn("w:val")) if abstract is not None else None
            if abstract_id is None or abstract_id not in abstracts:
                issues.append(f"Numbering numId {num_id} has no abstract definition ({abstract_id})")
    for part in doc.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            if not part.content_type.endswith("+xml"):
                continue
            root = etree.fromstring(part.blob)
        for properties in root.iter(qn("w:numPr")):
            for reference in properties.findall(qn("w:numId")):
                num_id = reference.get(qn("w:val"))
                # Word uses zero to explicitly suppress inherited numbering.
                if num_id == "0":
                    continue
                if numbering is None:
                    issue = "Numbering definitions part is missing"
                elif num_id not in definitions:
                    issue = f"Numbering reference numId {num_id} has no definition"
                else:
                    continue
                if issue not in issues:
                    issues.append(issue)
    return issues


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
