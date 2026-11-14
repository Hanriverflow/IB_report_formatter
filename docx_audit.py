"""Inspect generated DOCX structure without an inverse Word-to-Markdown parser.

Changelog (XML safety):
    - Disable entity resolution, DTD loading and network access for raw XML parts.
    - Report unreadable or DTD-bearing parts without consuming their values.

Changelog (output encoding):
    - Write the JSON report as UTF-8 regardless of the console code page.

Changelog (term variables):
    - Compare tagged term values with each other and with the generation
      snapshot (`terms`); mismatches and missing controls are warnings only.
    - Serialize reports explicitly so documents without terms keep their JSON.

Changelog (hardening):
    - Index styles once per audit and separate render validation from observations.
    - Treat placeholder-like text as warnings and validate native numbering references.

Changelog (captions):
    - A REF/PAGEREF field whose bookmark is missing is an issue; a REF result or
      SEQ number that no longer matches (fields not updated) is a warning.
    - Review fixes: quoted bookmark names are unquoted, a field nested in another
      field's instruction keeps its own result, and SEQ order is judged only for
      sequences with no switch or `\\* ARABIC` alone. Only the main body is read.
    - Second review: a `w:fldSimple` result skips the instruction text of a
      complex field nested in it, as complex fields do.

Changelog (repeated footnote references):
    - A NOTEREF field whose bookmark is missing is an issue, as for REF; a
      NOTEREF result that is not the shown number of the footnote its bookmark
      marks is a warning (default footnote numbering only).
"""

import argparse
import codecs
import json
import logging
import posixpath
import re
import sys
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from typing import Any, cast
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.parts.numbering import NumberingPart
from docx.styles.style import ParagraphStyle
from docx.text.paragraph import Paragraph
from lxml import etree

from numeric_checks import CHECKS_PROPERTY, checks_from_json, evaluate_checks
from term_variables import TERM_PROPERTY_PREFIX, TERM_TAG_PREFIX

logger = logging.getLogger(__name__)


@dataclass
class DocumentAudit:
    """Structural observations and actionable errors, not a visual QA verdict."""

    paragraphs: int = 0
    tables: int = 0
    headings: int = 0
    images: int = 0
    numbered_paragraphs: int = 0
    issues: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    pagination_marked_paragraphs: int = 0
    visual_review: str = "not_performed"


@dataclass
class TermAudit:
    """Consistency of tagged term values after editing, not their financial validity.

    Attributes:
        mismatched: Key -> distinct current values in document order, when the
            key's controls disagree (for example, only one place was edited).
        changed: Key -> generated snapshot value and the first current value
            that differs from it, for copying back into the `terms:` YAML.
        missing: Snapshot keys none of whose controls remain in the document.
        indicative: Keys with a current value containing a bracketed `[...]` part.
        failed_checks: Stored `checks:` relations that fail (or cannot be
            evaluated) with the current values, falling back to the generated
            values; None when the document stores no checks.
    """

    mismatched: dict[str, list[str]] = field(default_factory=dict)
    changed: dict[str, dict[str, str]] = field(default_factory=dict)
    missing: list[str] = field(default_factory=list)
    indicative: list[str] = field(default_factory=list)
    failed_checks: list[str] | None = None


# Visible run content of a control's value other than text (`w:t`) and symbols
# (`w:sym`, decoded from `w:char`). Objects and note marks are visible without
# text, so they read as U+FFFC. An optional hyphen is invisible unless a line
# breaks there: `w:softHyphen` reads as nothing and U+00AD is dropped from both
# current and generated values, because Word may convert one into the other.
_VISIBLE_RUN_CONTENT: dict[str, str] = {
    qn("w:tab"): "\t",
    qn("w:ptab"): "\t",
    qn("w:br"): "\n",
    qn("w:cr"): "\n",
    qn("w:noBreakHyphen"): "\u2011",
    qn("w:softHyphen"): "",
    qn("w:drawing"): "\ufffc",
    qn("w:pict"): "\ufffc",
    qn("w:object"): "\ufffc",
    qn("w:footnoteReference"): "\ufffc",
    qn("w:endnoteReference"): "\ufffc",
}
_TEXT_TAG = qn("w:t")
_SYMBOL_TAG = qn("w:sym")
_RUN_CONTENT_TAGS = (_TEXT_TAG, _SYMBOL_TAG, *_VISIBLE_RUN_CONTENT)
_SOFT_HYPHEN = "\u00ad"
_REMOVED_CONTENT = frozenset({qn("w:del"), qn("w:moveFrom")})
_TERM_TAG_PATH = qn("w:sdtPr") + "/" + qn("w:tag")
_PLACEHOLDER_PATH = qn("w:sdtPr") + "/" + qn("w:showingPlcHdr")
_INDICATIVE_RE = re.compile(r"\[[^\[\]]*\]")


_PAGINATION_FLAGS = (
    ("keep_together", "keepLines"),
    ("keep_with_next", "keepNext"),
    ("page_break_before", "pageBreakBefore"),
)


class _StyleLookup:
    """Audit-local style index and inherited pagination cache."""

    def __init__(self, doc: DocxDocument) -> None:
        self.styles = {style.style_id: style for style in doc.styles}
        self.default = next(
            (
                style
                for style in self.styles.values()
                if style.type == WD_STYLE_TYPE.PARAGRAPH
                and style.element.get(qn("w:default")) in {"1", "true", "on"}
            ),
            None,
        )
        self.defaults: dict[str, bool] = {}
        for _, tag in _PAGINATION_FLAGS:
            nodes = doc.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr/w:" + tag)
            self.defaults[tag] = bool(
                nodes and nodes[0].get(qn("w:val"), "1") not in {"0", "false", "off"}
            )
        self.flags: dict[tuple[str | None, str], bool] = {}

    def paragraph_style(self, paragraph: Paragraph) -> ParagraphStyle | None:
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
    paragraph: Paragraph,
    attribute: str,
    xml_tag: str,
    styles: _StyleLookup,
) -> bool:
    """Resolve local > paragraph style ancestry > document default controls."""
    value = getattr(paragraph.paragraph_format, attribute)
    return (
        bool(value) if value is not None else styles.inherited_flag(paragraph, attribute, xml_tag)
    )


def inspect_document(doc: DocxDocument) -> DocumentAudit:
    """Inspect structure and observations, including text inside tables.

    Args:
        doc: Open python-docx document to inspect without modifying it.

    Returns:
        Structural issues and descriptive counts, warnings and review status.
        Inconsistent or missing term values are warnings, never issues.
    """
    styles = _StyleLookup(doc)
    paragraph_styles = (styles.paragraph_style(p) for p in doc.paragraphs)
    result = DocumentAudit(
        paragraphs=len(doc.paragraphs),
        tables=len(doc.tables),
        headings=sum(
            style is not None and style.name.startswith("Heading ") for style in paragraph_styles
        ),
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
            result.warnings.append(
                "Possible rendering placeholder (may be literal user text): " + text
            )
    result.warnings.extend(_cross_reference_findings(doc)[1])
    terms = inspect_terms(doc)
    if terms is not None:
        result.warnings.extend(_term_warnings(terms))
    return result


def inspect_terms(doc: DocxDocument) -> TermAudit | None:
    """Compare tagged term values with each other and with the generation snapshot.

    Only the tagged values are compared; numbers typed outside a control and
    the financial validity of the terms are not checked.

    Args:
        doc: Open python-docx document to inspect without modifying it.

    Returns:
        Term results when the body has term controls or the package has a term
        snapshot; None otherwise, so a report without terms is unchanged.
    """
    current = _current_term_values(doc)
    generated = _generated_term_values(doc)
    stored = _custom_property(doc, CHECKS_PROPERTY)
    if not current and not generated and stored is None:
        return None
    result = TermAudit()
    if stored is not None:
        try:
            checks, latest = checks_from_json(stored)
        except ValueError as error:
            result.failed_checks = [f"Stored checks cannot be read: {error}"]
        else:
            latest.update(generated)
            latest.update({key: texts[0] for key, texts in current.items() if texts})
            result.failed_checks = evaluate_checks(checks, latest)
    for key, values in sorted(current.items()):
        distinct = list(dict.fromkeys(values))
        if len(distinct) > 1:
            result.mismatched[key] = distinct
        if any(_INDICATIVE_RE.search(value) for value in values):
            result.indicative.append(key)
    for key, value in sorted(generated.items()):
        if key not in current:
            result.missing.append(key)
            continue
        differing = [text for text in current[key] if text != value]
        if differing:
            result.changed[key] = {"generated": value, "current": differing[0]}
    return result


def _current_term_values(doc: DocxDocument) -> dict[str, list[str]]:
    """Read each remaining term control's current text in document order."""
    values: dict[str, list[str]] = {}
    for control in doc.element.body.iter(qn("w:sdt")):
        tag = control.find(_TERM_TAG_PATH)
        name = tag.get(qn("w:val"), "") if tag is not None else ""
        if not name.startswith(TERM_TAG_PREFIX) or _is_removed(control):
            continue
        content = control.find(qn("w:sdtContent"))
        text = ""
        # Word refills a cleared control with placeholder text and flags it.
        if content is not None and not _shows_placeholder(control):
            text = _current_text(content, control)
        values.setdefault(name[len(TERM_TAG_PREFIX) :], []).append(text)
    return values


def _shows_placeholder(control: Any) -> bool:
    """Whether a content control displays placeholder text instead of a value."""
    flag = control.find(_PLACEHOLDER_PATH)
    return flag is not None and flag.get(qn("w:val"), "true") not in {"0", "false", "off"}


def _current_text(content: Any, control: Any) -> str:
    """Join visible run content, including insertions but not deletions or moves away."""
    parts: list[str] = []
    for node in content.iter(*_RUN_CONTENT_TAGS):
        if _is_removed(node, stop=control):
            continue
        if node.tag == _TEXT_TAG:
            parts.append(node.text or "")
        elif node.tag == _SYMBOL_TAG:
            parts.append(_symbol_text(node))
        else:
            parts.append(_VISIBLE_RUN_CONTENT[node.tag])
    return _visible("".join(parts))


def _symbol_text(node: Any) -> str:
    """Decode a `w:sym` character code; symbol fonts use the F0xx private-use range."""
    try:
        return chr(int(node.get(qn("w:char"), ""), 16))
    except (ValueError, OverflowError):
        return "\ufffd"


def _visible(text: str) -> str:
    """Drop optional (soft) hyphens, which are not visible within a line."""
    return text.replace(_SOFT_HYPHEN, "")


def _is_removed(node: Any, stop: Any = None) -> bool:
    """Whether a deletion or move-away ancestor (below `stop`) hides the node."""
    for ancestor in node.iterancestors():
        if ancestor is stop:
            return False
        if ancestor.tag in _REMOVED_CONTENT:
            return True
    return False


def _parse_xml_part(blob: bytes) -> Any | None:
    """Read an untrusted XML part without resolving entities or loading a DTD.

    Args:
        blob: Raw XML bytes from the document package.

    Returns:
        The root element, or None for malformed XML or a DTD-bearing part.
        Callers report rejected parts without exposing parser error contents.
    """
    parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True)
    try:
        root = etree.fromstring(blob, parser=parser)
    except etree.XMLSyntaxError:
        return None
    # Even with entity resolution disabled, internal entities in attributes
    # can be substituted. Do not consume any values from a DTD-bearing tree.
    if root.getroottree().docinfo.doctype:
        return None
    return root


def _custom_property(doc: DocxDocument, name: str) -> str | None:
    """Text of one custom document property, or None when it is absent."""
    try:
        part = doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
    except KeyError:
        return None
    root = _parse_xml_part(part.blob)
    if root is None:
        return None
    for prop in root:
        if isinstance(prop.tag, str) and prop.get("name") == name:
            value = next((child for child in prop if isinstance(child.tag, str)), None)
            return (value.text or "") if value is not None else ""
    return None


def _generated_term_values(doc: DocxDocument) -> dict[str, str]:
    """Read the `ibrep.term.` snapshot written when the document was generated."""
    try:
        part = doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
    except KeyError:
        return {}
    root = _parse_xml_part(part.blob)
    if root is None:
        return {}
    values: dict[str, str] = {}
    for prop in root:
        if not isinstance(prop.tag, str):
            continue  # comments and processing instructions
        name = prop.get("name") or ""
        if not name.startswith(TERM_PROPERTY_PREFIX):
            continue
        # The value is the text of the typed value element (vt:lpwstr); the
        # property's own whitespace is only XML indentation.
        value = next((child for child in prop if isinstance(child.tag, str)), None)
        text = (value.text or "") if value is not None else ""
        values[name[len(TERM_PROPERTY_PREFIX) :]] = _visible(text)
    return values


def _term_warnings(terms: TermAudit) -> list[str]:
    """Human-readable warnings for inconsistent and missing term values."""
    warnings = [
        f"Term {key!r} has inconsistent values in the document: "
        + ", ".join(repr(value) for value in values)
        for key, values in terms.mismatched.items()
    ]
    warnings.extend(
        f"Term {key!r} was generated but none of its content controls remain in the document"
        for key in terms.missing
    )
    warnings.extend(terms.failed_checks or [])
    return warnings


def audit_to_dict(result: DocumentAudit, terms: TermAudit | None = None) -> dict[str, Any]:
    """Serialize an audit report; documents without terms keep the previous schema.

    Args:
        result: Structural audit from inspect_document.
        terms: Term results from inspect_terms, if any.

    Returns:
        The audit fields, plus `terms` only when term results exist.
    """
    data = asdict(result)
    if terms is not None:
        data["terms"] = asdict(terms)
        if terms.failed_checks is None:  # documents without checks keep the previous schema
            del data["terms"]["failed_checks"]
    return data


def inspect_document_issues(doc: DocxDocument) -> list[str]:
    """Validate structure without computing descriptive audit observations.

    Args:
        doc: Open python-docx document to validate before saving.

    Returns:
        Structural errors suitable for strict render validation.
    """
    issues: list[str] = []
    for index, section in enumerate(doc.sections, 1):
        width, left, right = section.page_width, section.left_margin, section.right_margin
        if width is not None and left is not None and right is not None and width <= left + right:
            issues.append(f"Section {index} has no printable width")
    references = doc.element.xpath(".//w:footnoteReference")
    if references:
        try:
            part = doc.part.part_related_by(RT.FOOTNOTES)
            notes = _parse_xml_part(part.blob)
            if notes is None:
                issues.append("Footnote definitions part cannot be read safely")
            else:
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
    issues.extend(_cross_reference_findings(doc)[0])
    return issues


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
# The bookmark argument may be quoted (`REF "_Ref123" \h`).
_REFERENCE_FIELD_RE = re.compile(
    r'^(REF|PAGEREF|NOTEREF)\s+(?:"([^"]*)"|(\S+))(.*)$', re.IGNORECASE | re.DOTALL
)
# A NOTEREF switch that adds text to the number (`above`/`below`).
_NOTEREF_SHOWN_SWITCH_RE = re.compile(r"\\p\b", re.IGNORECASE)
_SEQ_FIELD_RE = re.compile(r'^SEQ\s+(?:"([^"]*)"|(\S+))(.*)$', re.IGNORECASE | re.DOTALL)
# A REF switch that changes what is shown (paragraph number, position, ...).
_REF_SHOWN_SWITCH_RE = re.compile(r"\\[nprwtfd]\b", re.IGNORECASE)
# The SEQ forms whose order is checked: no switch, or Arabic numbers as the
# engine and Word's Insert Caption write them. A sequence with any other
# switch (restart, repeat, hide, chapter, another number format) is not judged.
_SEQ_CHECKED_SWITCHES = frozenset({"", "\\* arabic"})


def _body_fields(
    doc: DocxDocument,
) -> tuple[dict[str, str], list[tuple[str, str]], dict[str, int]]:
    """Read bookmarks and fields of the main document in order.

    Text counts as visible unless it is a field instruction, so a bookmark
    around `표 ` and a `SEQ` field reads `표 1`. A field nested in another
    field's instruction still has its own cached result.

    Args:
        doc: Open document.

    Returns:
        Visible text of each bookmark by casefolded name, each field's
        instruction and cached result in document order, and the shown number
        of the first footnote reference inside each bookmark that holds one.
    """
    marks: dict[str, list[str]] = {}
    open_marks: dict[str, str] = {}
    stack: list[dict[str, Any]] = []
    found: list[tuple[str, str]] = []
    notes: dict[str, int] = {}
    note_number = 0
    for node in doc.element.body.iter():
        tag = node.tag
        if tag == _W + "footnoteReference":
            # Word numbers notes by the order of their references; a custom
            # mark takes no number.
            if node.get(qn("w:customMarkFollows")) not in ("1", "true", "on"):
                note_number += 1
                for name in open_marks.values():
                    notes.setdefault(name, note_number)
        elif tag == _W + "bookmarkStart":
            name = (node.get(qn("w:name")) or "").casefold()
            open_marks[node.get(qn("w:id")) or ""] = name
            marks.setdefault(name, [])
        elif tag == _W + "bookmarkEnd":
            open_marks.pop(node.get(qn("w:id")) or "", None)
        elif tag == _W + "fldChar":
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"code": [], "result": [], "shown": False})
            elif kind == "separate" and stack:
                stack[-1]["shown"] = True
            elif kind == "end" and stack:
                field_state = stack.pop()
                found.append(("".join(field_state["code"]).strip(), "".join(field_state["result"])))
        elif tag == _W + "instrText" and stack and not stack[-1]["shown"]:
            stack[-1]["code"].append(node.text or "")
        elif tag == _W + "fldSimple":
            found.append(((node.get(qn("w:instr")) or "").strip(), _simple_field_result(node)))
        elif tag == _W + "t":
            # Text belongs to the result of every open field from the innermost
            # outwards while each of them is past its separator.
            for state in reversed(stack):
                if not state["shown"]:
                    break
                state["result"].append(node.text or "")
            if all(state["shown"] for state in stack):
                for name in open_marks.values():
                    marks[name].append(node.text or "")
    return {name: "".join(parts) for name, parts in marks.items()}, found, notes


def _simple_field_result(field: Any) -> str:
    """Visible text of a `w:fldSimple` result, by the rule for complex fields.

    Text inside the instruction of a complex field nested in the result (such
    as the result of a field nested in that instruction) is not shown.

    Args:
        field: The `w:fldSimple` element.

    Returns:
        The text Word shows for the field.
    """
    shown: list[bool] = []
    parts: list[str] = []
    for node in field.iter(_W + "fldChar", _W + "t"):
        if node.tag == _W + "t":
            if all(shown):
                parts.append(node.text or "")
            continue
        kind = node.get(qn("w:fldCharType"))
        if kind == "begin":
            shown.append(False)
        elif kind == "separate" and shown:
            shown[-1] = True
        elif kind == "end" and shown:
            shown.pop()
    return "".join(parts)


def _cross_reference_findings(doc: DocxDocument) -> tuple[list[str], list[str]]:
    """Check REF/PAGEREF/NOTEREF targets and cached results, and SEQ numbering order.

    A reference whose bookmark is missing (its caption or footnote was deleted
    in Word) is an issue. A reference or caption number whose cached result no
    longer matches (fields not updated after an edit) is a warning. A NOTEREF
    result is judged only under Word's default footnote numbering (no
    `w:footnotePr` in the body or settings).

    Args:
        doc: Open document.

    Returns:
        Issues and warnings.
    """
    marks, found, notes = _body_fields(doc)
    settings = doc.settings.element
    default_note_numbers = (
        not doc.element.body.xpath(".//w:footnotePr") and settings.find(qn("w:footnotePr")) is None
    )
    issues: list[str] = []
    warnings: list[str] = []
    sequences: dict[str, list[tuple[str, str]]] = {}
    for code, result in found:
        reference = _REFERENCE_FIELD_RE.match(code)
        if reference is not None:
            kind, quoted, bare, switches = reference.groups()
            name = quoted if quoted is not None else bare
            target = marks.get(name.casefold())
            if target is None:
                issue = f"Cross-reference field {kind.upper()} {name} has no bookmark named {name}"
                if issue not in issues:
                    issues.append(issue)
            elif (
                kind.upper() == "REF"
                and not _REF_SHOWN_SWITCH_RE.search(switches)
                and result.strip() != target.strip()
            ):
                warnings.append(
                    f"Cross-reference REF {name} shows {result.strip()!r} but its target reads "
                    f"{target.strip()!r}; update fields in Word (select all, F9)"
                )
            elif (
                kind.upper() == "NOTEREF"
                and default_note_numbers
                and name.casefold() in notes
                and not _NOTEREF_SHOWN_SWITCH_RE.search(switches)
                and result.strip() != str(notes[name.casefold()])
            ):
                warnings.append(
                    f"Footnote cross-reference NOTEREF {name} shows {result.strip()!r} but its "
                    f"footnote is number {notes[name.casefold()]}; update fields in Word "
                    "(select all, F9)"
                )
            continue
        sequence = _SEQ_FIELD_RE.match(code)
        if sequence is None:
            continue
        quoted, bare, switches = sequence.groups()
        name = quoted if quoted is not None else bare
        sequences.setdefault(name, []).append((" ".join(switches.split()).casefold(), result))
    for name, entries in sequences.items():
        if any(switches not in _SEQ_CHECKED_SWITCHES for switches, _ in entries):
            continue
        if [result.strip() for _, result in entries] != [
            str(number) for number in range(1, len(entries) + 1)
        ]:
            warnings.append(
                f"Caption numbers SEQ {name} are not 1, 2, 3... in document order; "
                "update fields in Word (select all, F9)"
            )
    return issues, warnings


def _numbering_issues(doc: DocxDocument) -> list[str]:
    """Validate numbering definitions and numPr references across XML parts."""
    issues: list[str] = []
    definitions = {}
    try:
        numbering = cast(NumberingPart, doc.part.part_related_by(RT.NUMBERING)).element
    except KeyError:
        numbering = None
    if numbering is not None:
        abstracts = {
            node.get(qn("w:abstractNumId")) for node in numbering.findall(qn("w:abstractNum"))
        }
        definitions = {node.get(qn("w:numId")): node for node in numbering.findall(qn("w:num"))}
        for num_id, node in definitions.items():
            abstract = node.find(qn("w:abstractNumId"))
            abstract_id = abstract.get(qn("w:val")) if abstract is not None else None
            if abstract_id is None or abstract_id not in abstracts:
                issues.append(
                    f"Numbering numId {num_id} has no abstract definition ({abstract_id})"
                )
    for part in doc.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            if not part.content_type.endswith("+xml"):
                continue
            root = _parse_xml_part(part.blob)
            if root is None:
                issues.append(f"XML part cannot be read safely: {part.partname}")
                continue
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
    return _audit_package(path)[0]


def _read_package_xml(archive: ZipFile, name: str) -> Any | None:
    """Read a package XML member without exposing unreadable or unsafe contents."""
    try:
        return _parse_xml_part(archive.read(name))
    except (KeyError, OSError, BadZipFile, RuntimeError, NotImplementedError) as error:
        logger.debug("Cannot read XML part %s: %s", name, error)
        return None


def _audit_package(path: str) -> tuple[DocumentAudit, TermAudit | None]:
    """Check every XML part before allowing python-docx to consume the package.

    Args:
        path: Untrusted DOCX package to inspect.

    Returns:
        A rejection naming unsafe parts, or the ordinary document and term audit.
    """
    with open(path, "rb") as stream:
        with ZipFile(stream) as archive:
            types_name = "[Content_Types].xml"
            types = _read_package_xml(archive, types_name)
            namespace = "{http://schemas.openxmlformats.org/package/2006/content-types}"
            if types is None or types.tag != namespace + "Types":
                return DocumentAudit(
                    issues=[f"XML part cannot be read safely: /{types_name}"]
                ), None
            defaults: dict[str, str] = {}
            overrides: dict[str, str] = {}
            for node in types:
                if node.tag not in {namespace + "Default", namespace + "Override"}:
                    continue
                key = "Extension" if node.tag == namespace + "Default" else "PartName"
                name, content_type = node.get(key), node.get("ContentType")
                if name is None or content_type is None:
                    return DocumentAudit(
                        issues=[f"XML part cannot be read safely: /{types_name}"]
                    ), None
                mapping = defaults if key == "Extension" else overrides
                mapping[name.lower()] = content_type.lower()
            # Match python-docx's case-insensitive content-type lookup, including
            # override precedence. PartFactory's registered XML classes all use
            # XML content types; generic binary Parts retain their bytes unchanged.
            issues = []
            for part in archive.infolist():
                if part.filename == types_name:
                    continue
                name = part.filename.lower()
                # The extension as python-docx's PackURI.ext derives it: empty for a
                # name without one and for a dotfile, which a Default may also declare.
                extension = posixpath.splitext(name)[1].removeprefix(".")
                content_type = overrides.get("/" + name, defaults.get(extension, ""))
                is_xml = (
                    name.endswith(".rels")
                    or content_type in {"application/xml", "text/xml"}
                    or content_type.endswith("+xml")
                )
                if is_xml and _read_package_xml(archive, part.filename) is None:
                    issues.append(f"XML part cannot be read safely: /{part.filename}")
        if issues:
            return DocumentAudit(issues=issues), None
        stream.seek(0)
        document = Document(stream)
        return inspect_document(document), inspect_terms(document)


@contextmanager
def _utf8_log_stream() -> Iterator[None]:
    """Encode the report as UTF-8 even when the console code page differs.

    Redirected JSON would otherwise use the locale code page (cp949 on Korean
    Windows) and fail to parse as UTF-8. Only stderr, the logging stream, is
    changed, only while `main` runs, and its previous settings are restored for
    in-process callers. A stream that cannot be reconfigured is left as is.
    """
    stream = sys.stderr
    reconfigure = getattr(stream, "reconfigure", None)
    encoding = getattr(stream, "encoding", None)
    errors = getattr(stream, "errors", None)
    restore: Callable[..., Any] | None = None
    if callable(reconfigure) and codecs.lookup(encoding or "ascii").name != "utf-8":
        try:
            reconfigure(encoding="utf-8")
            restore = reconfigure
        except (OSError, ValueError) as exc:  # includes io.UnsupportedOperation
            logger.debug("Keeping the stderr encoding: %s", exc)
    try:
        yield
    finally:
        if restore is not None:
            try:
                restore(encoding=encoding, errors=errors)
            except (OSError, ValueError) as exc:
                logger.debug("Cannot restore the stderr encoding: %s", exc)


def main() -> None:
    """Print a JSON audit result and return nonzero on structural errors only."""
    parser = argparse.ArgumentParser(description="Inspect generated Word document structure")
    parser.add_argument("input_file")
    args = parser.parse_args()
    with _utf8_log_stream():
        logging.basicConfig(level=logging.INFO, format="%(message)s")
        try:
            result, terms = _audit_package(args.input_file)
        except Exception as exc:
            logger.error("Cannot inspect DOCX: %s", exc)
            raise SystemExit(1) from exc
        logger.info(json.dumps(audit_to_dict(result, terms), ensure_ascii=False, indent=2))
    raise SystemExit(1 if result.issues else 0)


if __name__ == "__main__":
    main()
