"""Insert WordprocessingML property children in ECMA-376 schema order.

Word opens most out-of-order property elements, but the schema defines a fixed
sequence and strict validators (and some other editors) reject or repair
documents that break it. Engine code that builds property elements by hand
inserts them through `insert_ordered` instead of appending.

Changelog (schema order):
    - NEW: shared tblPr/tcPr/pPr/rPr sequences and an ordered insert.
    - FIXED: settings sequence keeps updateFields before its schema successors.
    - NEW: tcBorders sequence for cell borders on explicit totals rows.
"""

from lxml import etree

# Child sequences of ECMA-376 Part 1, §17 property and settings elements.
SEQUENCES: dict[str, tuple[str, ...]] = {
    # Full CT_Settings._tag_seq from python-docx 1.2.0, docx/oxml/settings.py.
    # As below, use local names (including m:mathPr and sl:schemaLibrary).
    "settings": (
        "writeProtection",
        "view",
        "zoom",
        "removePersonalInformation",
        "removeDateAndTime",
        "doNotDisplayPageBoundaries",
        "displayBackgroundShape",
        "printPostScriptOverText",
        "printFractionalCharacterWidth",
        "printFormsData",
        "embedTrueTypeFonts",
        "embedSystemFonts",
        "saveSubsetFonts",
        "saveFormsData",
        "mirrorMargins",
        "alignBordersAndEdges",
        "bordersDoNotSurroundHeader",
        "bordersDoNotSurroundFooter",
        "gutterAtTop",
        "hideSpellingErrors",
        "hideGrammaticalErrors",
        "activeWritingStyle",
        "proofState",
        "formsDesign",
        "attachedTemplate",
        "linkStyles",
        "stylePaneFormatFilter",
        "stylePaneSortMethod",
        "documentType",
        "mailMerge",
        "revisionView",
        "trackRevisions",
        "doNotTrackMoves",
        "doNotTrackFormatting",
        "documentProtection",
        "autoFormatOverride",
        "styleLockTheme",
        "styleLockQFSet",
        "defaultTabStop",
        "autoHyphenation",
        "consecutiveHyphenLimit",
        "hyphenationZone",
        "doNotHyphenateCaps",
        "showEnvelope",
        "summaryLength",
        "clickAndTypeStyle",
        "defaultTableStyle",
        "evenAndOddHeaders",
        "bookFoldRevPrinting",
        "bookFoldPrinting",
        "bookFoldPrintingSheets",
        "drawingGridHorizontalSpacing",
        "drawingGridVerticalSpacing",
        "displayHorizontalDrawingGridEvery",
        "displayVerticalDrawingGridEvery",
        "doNotUseMarginsForDrawingGridOrigin",
        "drawingGridHorizontalOrigin",
        "drawingGridVerticalOrigin",
        "doNotShadeFormData",
        "noPunctuationKerning",
        "characterSpacingControl",
        "printTwoOnOne",
        "strictFirstAndLastChars",
        "noLineBreaksAfter",
        "noLineBreaksBefore",
        "savePreviewPicture",
        "doNotValidateAgainstSchema",
        "saveInvalidXml",
        "ignoreMixedContent",
        "alwaysShowPlaceholderText",
        "doNotDemarcateInvalidXml",
        "saveXmlDataOnly",
        "useXSLTWhenSaving",
        "saveThroughXslt",
        "showXMLTags",
        "alwaysMergeEmptyNamespace",
        "updateFields",
        "hdrShapeDefaults",
        "footnotePr",
        "endnotePr",
        "compat",
        "docVars",
        "rsids",
        "mathPr",
        "attachedSchema",
        "themeFontLang",
        "clrSchemeMapping",
        "doNotIncludeSubdocsInStats",
        "doNotAutoCompressPictures",
        "forceUpgrade",
        "captions",
        "readModeInkLockDown",
        "smartTagType",
        "schemaLibrary",
        "shapeDefaults",
        "doNotEmbedSmartTags",
        "decimalSymbol",
        "listSeparator",
    ),
    "tblPr": (
        "tblStyle",
        "tblpPr",
        "tblOverlap",
        "bidiVisual",
        "tblStyleRowBandSize",
        "tblStyleColBandSize",
        "tblW",
        "jc",
        "tblCellSpacing",
        "tblInd",
        "tblBorders",
        "shd",
        "tblLayout",
        "tblCellMar",
        "tblLook",
        "tblCaption",
        "tblDescription",
        "tblPrChange",
    ),
    "tcPr": (
        "cnfStyle",
        "tcW",
        "gridSpan",
        "hMerge",
        "vMerge",
        "tcBorders",
        "shd",
        "noWrap",
        "tcMar",
        "textDirection",
        "tcFitText",
        "vAlign",
        "hideMark",
        "headers",
        "cellIns",
        "cellDel",
        "cellMerge",
        "tcPrChange",
    ),
    # CT_TcBorders child sequence; transitional left/right sit beside strict start/end.
    "tcBorders": (
        "top",
        "left",
        "start",
        "bottom",
        "right",
        "end",
        "insideH",
        "insideV",
        "tl2br",
        "tr2bl",
    ),
    "pPr": (
        "pStyle",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "framePr",
        "widowControl",
        "numPr",
        "suppressLineNumbers",
        "pBdr",
        "shd",
        "tabs",
        "suppressAutoHyphens",
        "kinsoku",
        "wordWrap",
        "overflowPunct",
        "topLinePunct",
        "autoSpaceDE",
        "autoSpaceDN",
        "bidi",
        "adjustRightInd",
        "snapToGrid",
        "spacing",
        "ind",
        "contextualSpacing",
        "mirrorIndents",
        "suppressOverlap",
        "jc",
        "textDirection",
        "textAlignment",
        "textboxTightWrap",
        "outlineLvl",
        "divId",
        "cnfStyle",
        "rPr",
        "sectPr",
        "pPrChange",
    ),
    "rPr": (
        "rStyle",
        "rFonts",
        "b",
        "bCs",
        "i",
        "iCs",
        "caps",
        "smallCaps",
        "strike",
        "dstrike",
        "outline",
        "shadow",
        "emboss",
        "imprint",
        "noProof",
        "snapToGrid",
        "vanish",
        "webHidden",
        "color",
        "spacing",
        "w",
        "kern",
        "position",
        "sz",
        "szCs",
        "highlight",
        "u",
        "effect",
        "bdr",
        "shd",
        "fitText",
        "vertAlign",
        "rtl",
        "cs",
        "em",
        "lang",
        "eastAsianLayout",
        "specVanish",
        "oMath",
        "rPrChange",
    ),
}


def _local_name(element: etree._Element) -> str | None:
    """Return an element's local name, or None for comments and processing instructions."""
    return etree.QName(element).localname if isinstance(element.tag, str) else None


def insert_ordered(parent: etree._Element, child: etree._Element, replace: bool = True) -> None:
    """Insert a property child before the first sibling that must follow it.

    Args:
        parent: A `w:settings`, `w:tblPr`, `w:tcPr`, `w:tcBorders`, `w:pPr` or `w:rPr`
            element.
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
        if (
            sibling is not child
            and sibling_name in sequence
            and sequence.index(sibling_name) > rank
        ):
            sibling.addprevious(child)
            return
    parent.append(child)
