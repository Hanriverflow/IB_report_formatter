"""Office document composition and editable Word numbering.

Changelog (letter closing):
    - `letter.table_end: partial` writes 이하 빈칸 in the first blank row after the
      last written row of the table that ends the letter body, instead of "끝."
      (`letter_partial_table_row`). Without the setting the output is unchanged.
    - Attachments follow the administrative handbook: one attachment is a single
      unnumbered line `붙임⇥항목.  끝.`; several start on the `붙임⇥1.` line of a
      native list whose other numbers align under `1.` (`render_attachments`).

Changelog (Korean language):
    - Every profile other than term-sheet declares `ko-KR` as the East Asian
      language (`setup_korean_language`), so Word keeps `㈜는` and `」에` whole.

Changelog (Korean spacing):
    - General profiles print Korean next to Latin text and digits without
      Word's automatic spacing (`setup_korean_spacing`).
    - The IB profiles do the same, so `2025년` and `300억원` print tight there too.

Changelog (hardening):
    - Scope native list instances to their parent item and override the actual level.
    - Reuse memo title/metadata typography for cover-free IB reports, adding subtitles.
"""

from copy import deepcopy

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from document_model import DocumentMetadata, DocumentModel, ElementType, Heading, Table, TableRow
from document_profiles import string_list
from ooxml_order import insert_ordered
from render_styles import STYLE


def add_text(
    doc, text: str, bold: bool = False, keep_next: bool = False, style_name: str | None = None
):
    """Add a consistently styled office paragraph."""
    paragraph = doc.add_paragraph(style=style_name or STYLE.STYLE_IB_BODY)
    paragraph.paragraph_format.keep_with_next = keep_next
    run = paragraph.add_run(text)
    run.font.name = STYLE.BODY_FONT
    if style_name is None:
        run.font.size = STYLE.BODY_SIZE
    run.bold = bold
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), STYLE.KOREAN_FONT)
    return paragraph


def setup_letter_styles(doc) -> None:
    """Create reusable letter styles without blanket pagination constraints."""
    for name, size, before, after in [
        ("Office Metadata", STYLE.BODY_SIZE.pt, 0, 2),
        ("Office Subject", STYLE.BODY_SIZE.pt, 0, 16),
        ("Office Attachment", 10, 0, 1),
        ("Office Signatory", 15, 18, 12),
        ("Office Contact", 9, 0, 1),
    ]:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles[STYLE.STYLE_IB_BODY]
        style.font.name, style.font.size = STYLE.BODY_FONT, Pt(size)
        style.font.color.rgb = STYLE.NAVY
        style._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), STYLE.KOREAN_FONT)
        fmt = style.paragraph_format
        fmt.space_before, fmt.space_after = Pt(before), Pt(after)
        fmt.keep_together, fmt.keep_with_next = False, False
        fmt.tab_stops.add_tab_stop(Pt(42))
        if name in {"Office Metadata", "Office Subject", "Office Contact"}:
            fmt.left_indent, fmt.first_line_indent = Pt(42), Pt(-42)


def setup_korean_spacing(doc) -> None:
    """Print Korean next to Latin letters and digits without Word's automatic gap.

    Word adds space between East Asian text and Latin text (`w:autoSpaceDE`) or
    digits (`w:autoSpaceDN`) by default, printing `SPC 는`, `제 2 종`, `300 억원`.
    Korean business documents print these tight, so every profile turns both
    off in the document defaults (the term-sheet profile does so itself).

    Args:
        doc: Document whose request-scoped styles were created already.
    """
    styles = doc.styles.element
    defaults = styles.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles.insert(0, defaults)
    paragraph_defaults = defaults.find(qn("w:pPrDefault"))
    if paragraph_defaults is None:
        paragraph_defaults = OxmlElement("w:pPrDefault")
        defaults.append(paragraph_defaults)
    properties = paragraph_defaults.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        paragraph_defaults.append(properties)
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
        element = OxmlElement(tag)
        element.set(qn("w:val"), "0")
        insert_ordered(properties, element)


def setup_korean_language(doc) -> None:
    """Declare Korean as the document's East Asian language.

    With the East Asian language left at the template's `en-US`, Word breaks a
    line inside a Korean word after `㈜` or a closing `」` (`가나다㈜⏎는`,
    `「신탁법」⏎에`). `ko-KR` keeps such words whole; Latin text keeps its own
    language. The term-sheet profile sets the same value itself.

    Args:
        doc: Document whose request-scoped styles were created already.
    """
    styles = doc.styles.element
    defaults = styles.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles.insert(0, defaults)
    run_defaults = defaults.find(qn("w:rPrDefault"))
    if run_defaults is None:
        run_defaults = OxmlElement("w:rPrDefault")
        defaults.insert(0, run_defaults)
    properties = run_defaults.find(qn("w:rPr"))
    if properties is None:
        properties = OxmlElement("w:rPr")
        run_defaults.append(properties)
    language = properties.find(qn("w:lang"))
    if language is None:
        language = OxmlElement("w:lang")
        insert_ordered(properties, language)
    language.set(qn("w:eastAsia"), "ko-KR")


def letter_appendix_index(model: DocumentModel) -> int | None:
    """Resolve a unique explicit H1 boundary before composing any output."""
    if model.metadata.profile != "office-letter":
        return None
    heading = model.metadata.extra.get("letter", {}).get("appendix_heading")
    if heading is None:
        return None
    matches = [
        index
        for index, element in enumerate(model.elements)
        if element.element_type == ElementType.HEADING_1
        and isinstance(element.content, Heading)
        and element.content.text == heading
    ]
    if len(matches) != 1 or heading == model.metadata.title:
        raise ValueError("letter.appendix_heading must match one distinct body H1")
    if not any(
        element.element_type
        not in {ElementType.HEADING_1, ElementType.HEADING_2, ElementType.HEADING_3}
        for element in model.elements[: matches[0]]
    ):
        raise ValueError("letter.appendix_heading must follow the letter body")
    return matches[0]


PARTIAL_TABLE_MARK = "이하 빈칸"


def _blank_row(row: TableRow) -> bool:
    """A row with nothing written: empty cells, or cells merged into one to the left.

    A cell merged into the row above (`^^`) continues written content, so its row
    is not blank.
    """
    return all(
        cell.merge == "left" or (cell.merge is None and not cell.content.strip())
        for cell in row.cells
    )


def letter_partial_table_row(model: DocumentModel, appendix_index: int | None) -> int | None:
    """Resolve where `letter.table_end: partial` writes 이하 빈칸, before any output.

    The setting declares that the table ending the letter body is filled only
    partway. The mark goes in the first cell of the first blank row after the last
    written row. Blank rows alone never trigger it: a reply form may keep blank
    rows for the recipient.

    Args:
        model: Render copy of the document model, title H1 already removed.
        appendix_index: Index of the explicit appendix H1; the body ends before it.

    Returns:
        Zero-based row index in the body's last table, or None without the setting.

    Raises:
        ValueError: The body does not end with a table, or that table has no blank
            row after its last written row.
    """
    if model.metadata.profile != "office-letter":
        return None
    if model.metadata.extra.get("letter", {}).get("table_end") != "partial":
        return None
    body = model.elements if appendix_index is None else model.elements[:appendix_index]
    last = body[-1] if body else None
    if (
        last is None
        or last.element_type != ElementType.TABLE
        or not isinstance(last.content, Table)
    ):
        raise ValueError("letter.table_end requires the letter body to end with a table")
    rows = last.content.rows
    written = [index for index, row in enumerate(rows) if not _blank_row(row)]
    target = written[-1] + 1 if written else 0
    if target >= len(rows):
        raise ValueError(
            "letter.table_end requires a blank row after the last written row of the final table"
        )
    return target


def _mark_partial_table(doc, row_index: int) -> None:
    """Write 이하 빈칸 into the first cell of a row of the last rendered table.

    The table renderer creates one Word row per model row, so the model index
    addresses the Word row. The mark is left-aligned even in a number column.
    """
    table = doc.element.body.findall(qn("w:tbl"))[-1]
    cell = table.findall(qn("w:tr"))[row_index].findall(qn("w:tc"))[0]
    paragraph = Paragraph(cell.findall(qn("w:p"))[0], doc)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(PARTIAL_TABLE_MARK)
    run.font.name = STYLE.BODY_FONT
    run.font.size = STYLE.TABLE_BODY_SIZE
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), STYLE.KOREAN_FONT)


def render_office_opening(doc, metadata: DocumentMetadata) -> bool:
    """Render profile metadata at the start of the body when there is no cover.

    Args:
        doc: Document with request-scoped styles already initialized.
        metadata: Profile metadata, including an optional IB report subtitle.

    Returns:
        Whether a title was inserted and a matching body H1 should be skipped.
    """
    name = metadata.profile
    if name == "plain" or (name == "ib-report" and not metadata.title.strip()):
        return False
    extra = metadata.extra
    sender = extra.get("sender", {})
    if name == "office-letter":
        organization = add_text(doc, sender["organization"], bold=True, keep_next=True)
        organization.style = doc.styles["Title"]
        organization.alignment = WD_ALIGN_PARAGRAPH.CENTER
        organization.runs[0].font.size = Pt(18)
        organization.paragraph_format.space_after = Pt(24)
        for label, key in [("수신", "recipients"), ("참조", "cc")]:
            values = string_list(extra, key)
            if values:
                add_text(
                    doc,
                    "{}\t{}".format(label, ", ".join(values)),
                    keep_next=True,
                    style_name="Office Metadata",
                )
        add_text(
            doc, "제목\t" + metadata.title, bold=True, keep_next=True, style_name="Office Subject"
        )
        return True
    title = add_text(doc, metadata.title, bold=True, keep_next=True)
    title.style = doc.styles["Title"]
    title.runs[0].font.size = STYLE.H1_SIZE
    title.paragraph_format.space_before = Pt(12)
    title.paragraph_format.space_after = Pt(12)
    if name == "ib-report" and metadata.subtitle.strip():
        add_text(doc, metadata.subtitle, keep_next=True)
    rows = []
    for label, value in [
        ("작성일", extra.get("date")),
        ("작성부서", sender.get("department")),
        ("작성자", metadata.analyst),
        ("일시", extra.get("meeting_time")),
        ("장소", extra.get("location")),
    ]:
        if value:
            rows.append(f"{label}: {value}")
    attendees = string_list(extra, "attendees")
    if attendees:
        rows.append("참석자: " + ", ".join(attendees))
    for row in rows:
        add_text(doc, row, keep_next=True)
    return True


# "붙임" (two characters of the 10pt attachment style) plus a one-character gap,
# the 2타 of the administrative handbook. Attachment numbers start here.
ATTACHMENT_NUMBER_PT = 30
# Attachment text starts at the attachment style's own tab stop.
ATTACHMENT_TEXT_PT = 42


def render_attachments(doc, attachments: list[str]) -> None:
    """Write the attachment lines as in the administrative handbook, ending in "끝.".

    One attachment is unnumbered: `붙임⇥발행조건 요약표 1부.  끝.`. Several are a
    native Word list whose first item carries the label (`붙임⇥1.`) and whose
    other numbers align under `1.`; Word keeps counting across them, so the
    numbers stay editable. Wrapped lines align with the attachment text.

    Args:
        doc: Document after the letter body.
        attachments: Declared attachment names; each gets a closing period.
    """
    lines = [item.rstrip().rstrip(".") + "." for item in attachments]
    lines[-1] += "  끝."
    if len(lines) == 1:
        paragraph = add_text(doc, "붙임\t" + lines[0], style_name="Office Attachment")
        paragraph.paragraph_format.space_before = Pt(8)
        paragraph.paragraph_format.left_indent = Pt(ATTACHMENT_NUMBER_PT)
        paragraph.paragraph_format.first_line_indent = Pt(-ATTACHMENT_NUMBER_PT)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(ATTACHMENT_NUMBER_PT))
        return
    numbering = NativeNumbering(doc)
    for index, line in enumerate(lines):
        paragraph = add_text(doc, line, style_name="Office Attachment")
        numbering.apply(paragraph, 0)
        fmt = paragraph.paragraph_format
        fmt.left_indent = Pt(ATTACHMENT_TEXT_PT)
        fmt.tab_stops.add_tab_stop(Pt(ATTACHMENT_NUMBER_PT))
        if index == 0:
            numbering.label_first_item(paragraph, "붙임\t")
            fmt.space_before = Pt(8)
            fmt.first_line_indent = Pt(-ATTACHMENT_TEXT_PT)
        else:
            fmt.first_line_indent = Pt(ATTACHMENT_NUMBER_PT - ATTACHMENT_TEXT_PT)


def render_office_closing(
    doc, metadata: DocumentMetadata, partial_table_row: int | None = None
) -> None:
    """Render attachments and sender without manufacturing a signature.

    Args:
        doc: Document whose letter body was rendered last.
        metadata: Letter metadata (attachments, sender, issue details).
        partial_table_row: Row of the body's last table that receives 이하 빈칸
            (`letter_partial_table_row`). That mark replaces "끝." unless
            attachments follow, which still end with "끝.".
    """
    if metadata.profile != "office-letter":
        return
    if partial_table_row is not None:
        _mark_partial_table(doc, partial_table_row)
    attachments = string_list(metadata.extra, "attachments")
    if attachments:
        render_attachments(doc, attachments)
    elif partial_table_row is None:
        # Attach the closing marker to the last actual body paragraph when possible.
        last = doc.paragraphs[-1] if doc.paragraphs else None
        if (
            last is not None
            and last._p.getnext() is not None
            and last._p.getnext().tag == qn("w:sectPr")
        ):
            if not last.text.rstrip().endswith("끝."):
                last.add_run("  끝.")
        else:
            add_text(doc, "끝.")
    sender = metadata.extra["sender"]
    paragraph = add_text(
        doc,
        sender.get("signatory") or sender["organization"],
        bold=True,
        style_name="Office Signatory",
    )
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    issued = " ".join(
        str(metadata.extra[k]) for k in ("document_no", "date") if metadata.extra.get(k)
    )
    if issued:
        add_text(doc, "시행\t" + issued, style_name="Office Contact")
    contact = " · ".join(sender[k] for k in ("department", "contact") if sender.get(k))
    if contact:
        add_text(doc, "담당\t" + contact, style_name="Office Contact")


class NativeNumbering:
    """One numbering state per document; supports Korean multilevel lists."""

    def __init__(self, doc) -> None:
        self.doc = doc
        self.ids: dict[tuple[int, str], int] = {}

    def reset(self) -> None:
        """Start a fresh list after a non-list block."""
        self.ids.clear()

    def apply(self, paragraph, level: int, bullet: bool = False, start: int = 1) -> None:
        """Attach native numbering, restarting descendants under each new parent.

        Args:
            paragraph: Destination Word paragraph.
            level: Zero-based nesting level, clamped to Word's supported range.
            bullet: Whether the list uses bullets instead of numbers.
            start: First value for a newly encountered ordered-list scope.
        """
        level = min(max(level, 0), 8)
        # A new item at this level ends any child list owned by the previous item.
        self.ids = {key: value for key, value in self.ids.items() if key[0] <= level}
        kind = "bullet" if bullet else "number"
        key = (level, kind)
        if key not in self.ids:
            self.ids[key] = self._create(bullet, start, level)
        num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        num_pr.get_or_add_ilvl().val = level
        num_pr.get_or_add_numId().val = self.ids[key]

    def label_first_item(self, paragraph, label: str) -> None:
        """Put `label` before the number of a top-level list's first item.

        A second list instance shares the abstract definition and overrides only
        the level-0 text. Word counts items of one abstract definition together,
        so the next items still number 2, 3, ... and stay editable.

        Args:
            paragraph: First item, already numbered at level 0 by `apply`.
            label: Text before the number, such as "붙임	".
        """
        numbering = self.doc.part.numbering_part.element
        num_id = self.ids[(0, "number")]
        source = next(
            n for n in numbering.findall(qn("w:num")) if int(n.get(qn("w:numId"))) == num_id
        )
        abstract_id = source.find(qn("w:abstractNumId")).get(qn("w:val"))
        abstract = next(
            a
            for a in numbering.findall(qn("w:abstractNum"))
            if a.get(qn("w:abstractNumId")) == abstract_id
        )
        level = deepcopy(abstract.find(qn("w:lvl")))
        level.find(qn("w:lvlText")).set(qn("w:val"), label + "%1.")
        override = OxmlElement("w:lvlOverride")
        override.set(qn("w:ilvl"), "0")
        override.append(level)
        labelled = OxmlElement("w:num")
        labelled_id = max(int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))) + 1
        labelled.set(qn("w:numId"), str(labelled_id))
        reference = OxmlElement("w:abstractNumId")
        reference.set(qn("w:val"), abstract_id)
        labelled.append(reference)
        labelled.append(override)
        numbering.append(labelled)
        paragraph._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = labelled_id

    def _create(self, bullet: bool, start: int, start_level: int) -> int:
        numbering = self.doc.part.numbering_part.element
        abstract_ids = [
            int(e.get(qn("w:abstractNumId"))) for e in numbering.findall(qn("w:abstractNum"))
        ]
        num_ids = [int(e.get(qn("w:numId"))) for e in numbering.findall(qn("w:num"))]
        abstract_id = max(abstract_ids + [-1]) + 1
        num_id = max(num_ids + [0]) + 1
        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "multilevel")
        abstract.append(multi)
        for level in range(9):
            lvl = OxmlElement("w:lvl")
            lvl.set(qn("w:ilvl"), str(level))
            number_format = "bullet" if bullet else ("ganada" if level % 3 == 1 else "decimal")
            label = "•" if bullet else ("(%{})" if level % 3 == 2 else "%{}.").format(level + 1)
            for tag, value in [
                ("start", "1"),
                ("numFmt", number_format),
                ("lvlText", label),
                ("lvlJc", "left"),
            ]:
                element = OxmlElement("w:" + tag)
                element.set(qn("w:val"), value)
                lvl.append(element)
            p_pr = OxmlElement("w:pPr")
            indent = OxmlElement("w:ind")
            indent.set(qn("w:left"), str(360 * (level + 1)))
            indent.set(qn("w:hanging"), "360")
            p_pr.append(indent)
            lvl.append(p_pr)
            abstract.append(lvl)
        first_num = numbering.find(qn("w:num"))
        if first_num is not None:
            numbering.insert(list(numbering).index(first_num), abstract)
        else:
            numbering.append(abstract)
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_ref = OxmlElement("w:abstractNumId")
        abstract_ref.set(qn("w:val"), str(abstract_id))
        num.append(abstract_ref)
        if not bullet and (start != 1 or start_level > 0):
            override = OxmlElement("w:lvlOverride")
            override.set(qn("w:ilvl"), str(start_level))
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), str(start))
            override.append(start_override)
            num.append(override)
        numbering.append(num)
        return num_id
