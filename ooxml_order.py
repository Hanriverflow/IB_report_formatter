"""Insert WordprocessingML property children in ECMA-376 schema order.

Word opens most out-of-order property elements, but the schema defines a fixed
sequence and strict validators (and some other editors) reject or repair
documents that break it. Engine code that builds property elements by hand
inserts them through `insert_ordered` instead of appending.

Changelog (schema order):
    - NEW: shared tblPr/tcPr/pPr/rPr sequences and an ordered insert.
"""

from typing import Dict, Optional, Tuple

from lxml import etree

# Child sequences of CT_TblPr, CT_TcPr, CT_PPr and CT_RPr (ECMA-376 Part 1, §17).
SEQUENCES: Dict[str, Tuple[str, ...]] = {
    "tblPr": (
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
        "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
    ),
    "tcPr": (
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
        "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns",
        "cellDel", "cellMerge", "tcPrChange",
    ),
    "pPr": (
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
        "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
        "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    ),
    "rPr": (
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
        "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
        "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
        "eastAsianLayout", "specVanish", "oMath", "rPrChange",
    ),
}


def _local_name(element: etree._Element) -> Optional[str]:
    """Return an element's local name, or None for comments and processing instructions."""
    return etree.QName(element).localname if isinstance(element.tag, str) else None


def insert_ordered(parent: etree._Element, child: etree._Element, replace: bool = True) -> None:
    """Insert a property child before the first sibling that must follow it.

    Args:
        parent: A `w:tblPr`, `w:tcPr`, `w:pPr` or `w:rPr` element.
        child: The new child element (for example `w:shd` or `w:pBdr`).
        replace: Remove existing children with the same name first, so the
            element occurs once as the schema requires.

    Raises:
        KeyError: The parent or child is not in a known sequence.
    """
    sequence = SEQUENCES[etree.QName(parent).localname]
    name = etree.QName(child).localname
    rank = sequence.index(name)
    if replace:
        for existing in list(parent):
            if existing is not child and _local_name(existing) == name:
                parent.remove(existing)
    for sibling in parent:
        sibling_name = _local_name(sibling)
        if sibling is not child and sibling_name in sequence and sequence.index(sibling_name) > rank:
            sibling.addprevious(child)
            return
    parent.append(child)
