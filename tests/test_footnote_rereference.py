"""A footnote referenced more than once (2026-11-14).

Word does not open a document in which two `w:footnoteReference` elements name
the same note ("파일이 손상되었습니다" in Word 16.0.20527). The first reference
stays native inside a hidden `_Ref_fn_<n>` bookmark and every later one is
`NOTEREF _Ref_fn_<n> \\f \\h` with the shown number cached, as Word's own
Insert Cross-reference > Footnote number (formatted) writes it. Every test goes
through the actual parser and renderer. Record:
docs/verification-footnote-rereference-20261114.md.
"""

import io
from pathlib import Path

import pytest
from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn

from document_profiles import PROFILES, RenderOptions
from docx_audit import inspect_document, inspect_document_issues
from ib_renderer import IBDocumentRenderer
from md_parser import parse_markdown_file

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Required document information per profile, beyond a title.
PROFILE_FRONTMATTER = {
    "office-letter": "recipients:\n  - 가상 수신처\nsender:\n  organization: 주식회사 예시\n",
    "term-sheet": "prepared_by: 예시증권\ndisclaimer: 시험용 가상 자료입니다.\n",
}
NOTES = "\n\n[^1]: 가상 각주 하나.\n[^2]: 가상 각주 둘.\n[^3]: 가상 각주 셋.\n"


def write(tmp_path: Path, body: str, profile: str = "business-report", tables: str = "") -> Path:
    """Write a Markdown file with frontmatter, the body and three footnotes."""
    front = f"---\nprofile: {profile}\ntitle: 각주 시험\n{PROFILE_FRONTMATTER.get(profile, '')}"
    if tables:
        front += "tables:\n" + tables
    path = tmp_path / "doc.md"
    path.write_text(front + "---\n\n# 각주 시험\n\n" + body + NOTES, encoding="utf-8")
    return path


def render(path: Path, strict: bool = True) -> tuple[DocxDocument, list[str]]:
    """Parse and render through the single composition path."""
    renderer = IBDocumentRenderer(options=RenderOptions(strict=strict))
    document = renderer.render(parse_markdown_file(str(path)))
    return document, renderer.errors


def saved(document: DocxDocument) -> DocxDocument:
    """Save and reopen, as Word and the strict save path read the package."""
    stream = io.BytesIO()
    document.save(stream)
    stream.seek(0)
    return Document(stream)


def note_ids(document: DocxDocument) -> list[str]:
    """IDs of the native footnote references of the body, in order."""
    return [node.get(qn("w:id")) for node in document.element.body.iter(f"{W}footnoteReference")]


def note_fields(document: DocxDocument) -> list[tuple[str, str]]:
    """Each NOTEREF field of the body in order: (instruction, cached result)."""
    found: list[tuple[str, str]] = []
    code: list[str] = []
    result: list[str] = []
    state = None
    for node in document.element.body.iter(f"{W}fldChar", f"{W}instrText", f"{W}t"):
        if node.tag == f"{W}fldChar":
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                state, code, result = "code", [], []
            elif kind == "separate":
                state = "result"
            else:
                instruction = "".join(code).strip()
                if instruction.startswith("NOTEREF"):
                    found.append((instruction, "".join(result)))
                state = None
        elif node.tag == f"{W}instrText" and state == "code":
            code.append(node.text or "")
        elif node.tag == f"{W}t" and state == "result":
            result.append(node.text or "")
    return found


def marked_notes(document: DocxDocument) -> dict[str, list[str]]:
    """Bookmark name -> IDs of the footnote references between its start and end."""
    names: dict[str, str] = {}
    held: dict[str, list[str]] = {}
    open_ids: set[str] = set()
    for node in document.element.body.iter(
        f"{W}bookmarkStart", f"{W}bookmarkEnd", f"{W}footnoteReference"
    ):
        if node.tag == f"{W}bookmarkStart":
            names[node.get(qn("w:id"))] = node.get(qn("w:name"))
            held[node.get(qn("w:id"))] = []
            open_ids.add(node.get(qn("w:id")))
        elif node.tag == f"{W}bookmarkEnd":
            open_ids.discard(node.get(qn("w:id")))
        else:
            for bookmark_id in open_ids:
                held[bookmark_id].append(node.get(qn("w:id")))
    return {name: held[bookmark_id] for bookmark_id, name in names.items()}


def note_marks(document: DocxDocument) -> dict[str, list[str]]:
    """Only the `_Ref_fn_` bookmarks of `marked_notes`."""
    return {
        name: notes for name, notes in marked_notes(document).items() if name.startswith("_Ref_fn_")
    }


def assert_one_reference_per_note(document: DocxDocument) -> None:
    """The condition Word needs to open the document, and no audit finding."""
    ids = note_ids(document)
    assert len(ids) == len(set(ids))
    starts = [node.get(qn("w:id")) for node in document.element.body.iter(f"{W}bookmarkStart")]
    assert len(starts) == len(set(starts))
    reopened = saved(document)
    assert inspect_document_issues(reopened) == []
    assert inspect_document(reopened).warnings == []


# ═══════════════════════════════════════════════════════════════════════════════
# REPEATED REFERENCES
# ═══════════════════════════════════════════════════════════════════════════════


def test_second_reference_is_a_noteref_field_to_the_first(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, "본문[^1] 문장이다. 다시[^1] 참조한다."))
    assert errors == []
    assert note_ids(document) == ["1"]
    assert note_marks(document) == {"_Ref_fn_1": ["1"]}
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")]
    assert_one_reference_per_note(document)


def test_bookmark_wraps_exactly_the_first_reference_run(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "본문[^1] 다시[^1]."))
    [start] = [
        node
        for node in document.element.body.iter(f"{W}bookmarkStart")
        if node.get(qn("w:name")) == "_Ref_fn_1"
    ]
    run = start.getnext()
    assert run.tag == f"{W}r" and run.find(f"{W}footnoteReference") is not None
    end = run.getnext()
    assert end.tag == f"{W}bookmarkEnd" and end.get(qn("w:id")) == start.get(qn("w:id"))


def test_field_runs_keep_the_reference_formatting(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "본문[^1] 다시[^1]."))
    [native] = [node.getparent() for node in document.element.body.iter(f"{W}footnoteReference")]
    [instruction] = [
        node for node in document.element.body.iter(f"{W}instrText") if "NOTEREF" in node.text
    ]
    field_runs = [
        instruction.getparent().getprevious(),
        instruction.getparent(),
        *list(instruction.getparent().itersiblings())[:3],
    ]
    expected = native.find(f"{W}rPr")
    for run in field_runs:
        r_pr = run.find(f"{W}rPr")
        assert r_pr is not None and r_pr is not expected
        assert [child.tag for child in r_pr] == [child.tag for child in expected]
        assert r_pr.find(f"{W}rStyle").get(qn("w:val")) == "FootnoteReference"
        assert r_pr.find(f"{W}vertAlign").get(qn("w:val")) == "superscript"


def test_three_references_write_one_note_and_two_fields(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, "가[^1], 나[^1], 다[^1]."))
    assert errors == []
    assert note_ids(document) == ["1"]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")] * 2
    assert_one_reference_per_note(document)


def test_cached_numbers_follow_the_order_of_first_references(tmp_path: Path) -> None:
    # Word numbers notes by reference order: note 2 is shown as 1 here.
    body = "둘[^2], 하나[^1], 다시 하나[^1], 다시 둘[^2], 셋[^3], 또 둘[^2]."
    document, errors = render(write(tmp_path, body))
    assert errors == []
    assert note_ids(document) == ["2", "1", "3"]
    assert note_marks(document) == {"_Ref_fn_1": ["2"], "_Ref_fn_2": ["1"]}
    assert note_fields(document) == [
        ("NOTEREF _Ref_fn_2 \\f \\h", "2"),
        ("NOTEREF _Ref_fn_1 \\f \\h", "1"),
        ("NOTEREF _Ref_fn_1 \\f \\h", "1"),
    ]
    # Bookmark IDs are given in document order.
    marks = [
        node.get(qn("w:name"))
        for node in document.element.body.iter(f"{W}bookmarkStart")
        if node.get(qn("w:name")).startswith("_Ref_fn_")
    ]
    assert marks == ["_Ref_fn_1", "_Ref_fn_2"]
    assert_one_reference_per_note(document)


SECOND_PLACES = {
    "heading": "## 둘째 제목[^1]\n",
    "table cell": "| 항목 | 값 |\n|---|---|\n| 가상 항목[^1] | 100 |\n",
    "list": "- 목록 항목[^1]\n- 다른 항목\n",
    "numbered list": "1. 번호 항목[^1]\n2. 다른 항목\n",
    "callout": "> **참고:** 콜아웃 안의 참조[^1]이다.\n",
    "quote": "> 인용 안의 참조[^1]이다.\n",
}


@pytest.mark.parametrize("place", list(SECOND_PLACES))
@pytest.mark.parametrize("first_in_place", [False, True])
def test_repeated_reference_in_every_block(
    tmp_path: Path, place: str, first_in_place: bool
) -> None:
    paragraph = "본문 참조[^1]이다.\n"
    blocks = (
        [SECOND_PLACES[place], paragraph] if first_in_place else [paragraph, SECOND_PLACES[place]]
    )
    document, errors = render(write(tmp_path, "\n".join(blocks)))
    assert errors == []
    assert note_ids(document) == ["1"]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")]
    assert note_marks(document) == {"_Ref_fn_1": ["1"]}
    assert_one_reference_per_note(document)


def test_report_with_toc_heading_and_cell_references(tmp_path: Path) -> None:
    body = "## 개요[^1]\n\n본문[^1].\n\n| 항목 | 값 |\n|---|---|\n| 가[^2] | 나[^1] |\n\n문장[^2]."
    document, errors = render(write(tmp_path, body, profile="ib-report"))
    assert errors == []
    assert note_ids(document) == ["1", "2"]
    assert note_fields(document) == [
        ("NOTEREF _Ref_fn_1 \\f \\h", "1"),
        ("NOTEREF _Ref_fn_1 \\f \\h", "1"),
        ("NOTEREF _Ref_fn_2 \\f \\h", "2"),
    ]
    assert_one_reference_per_note(document)


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_every_profile_opens_with_a_repeated_reference(tmp_path: Path, profile: str) -> None:
    document, errors = render(write(tmp_path, "본문[^1] 다시[^1].", profile=profile))
    assert errors == []
    assert note_ids(document) == ["1"]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")]
    assert_one_reference_per_note(document)


def test_default_mode_writes_the_same_fields(tmp_path: Path) -> None:
    path = write(tmp_path, "본문[^1] 다시[^1].")
    strict, _ = render(path)
    lenient, errors = render(path, strict=False)
    assert errors == []
    assert note_fields(lenient) == note_fields(strict)
    assert note_ids(lenient) == note_ids(strict)


def test_legacy_superscript_markers_are_linked_too(tmp_path: Path) -> None:
    # IB profiles read a typed `^1^` that names a defined note as its marker.
    document, errors = render(write(tmp_path, "본문^1^ 다시^1^ 둘째^2^.", profile="ib-report"))
    assert errors == []
    assert note_ids(document) == ["1", "2"]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")]
    assert_one_reference_per_note(document)


def test_notes_referenced_once_get_no_bookmark_or_field(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, "하나[^1], 둘[^2], 셋[^3]."))
    assert errors == []
    assert note_ids(document) == ["1", "2", "3"]
    assert note_fields(document) == []
    assert note_marks(document) == {}


def test_with_a_table_reference_and_a_negative_style_cell(tmp_path: Path) -> None:
    tables = (
        "  - id: tbl-pl\n    caption: 가상 손익\n    type: financial\n"
        "    columns: [text, money]\n    negative: paren\n"
    )
    body = (
        "본문[^1]은 @tbl-pl 참조.\n\n"
        "| 구분 | 금액 |\n|---|---:|\n| 손익[^2] | -1,200[^1] |\n| 법인세 | -300 |\n\n"
        "다시[^2]."
    )
    document, errors = render(write(tmp_path, body, tables=tables))
    assert errors == []
    assert note_ids(document) == ["1", "2"]
    assert note_fields(document) == [
        ("NOTEREF _Ref_fn_1 \\f \\h", "1"),
        ("NOTEREF _Ref_fn_2 \\f \\h", "2"),
    ]
    [table] = document.tables
    # The amount keeps its paren style; the marker after it is a field.
    assert table.cell(1, 1).text == "(1,200)1"
    assert table.cell(2, 1).text == "(300)"
    refs = [
        node.text.strip()
        for node in document.element.body.iter(f"{W}instrText")
        if node.text.strip().startswith("REF ")
    ]
    assert refs == ["REF _Ref_tbl_pl \\h \\* CHARFORMAT"]
    assert_one_reference_per_note(document)


# ═══════════════════════════════════════════════════════════════════════════════
# AUDIT
# ═══════════════════════════════════════════════════════════════════════════════


def test_audit_reports_a_noteref_whose_bookmark_is_missing(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "본문[^1] 다시[^1]."))
    for node in list(document.element.body.iter(f"{W}bookmarkStart", f"{W}bookmarkEnd")):
        node.getparent().remove(node)
    expected = "Cross-reference field NOTEREF _Ref_fn_1 has no bookmark named _Ref_fn_1"
    assert inspect_document_issues(document) == [expected]
    assert expected in inspect_document(document).issues


def test_audit_warns_about_a_stale_noteref_number(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "본문[^1] 다시[^1]."))
    [instruction] = [
        node for node in document.element.body.iter(f"{W}instrText") if "NOTEREF" in node.text
    ]
    result_run = instruction.getparent().getnext().getnext()  # -> separate -> result
    result_run.find(f"{W}t").text = "2"
    assert inspect_document_issues(document) == []
    assert inspect_document(document).warnings == [
        "Footnote cross-reference NOTEREF _Ref_fn_1 shows '2' but its footnote is number 1; "
        "update fields in Word (select all, F9)"
    ]


def test_audit_does_not_judge_numbers_under_custom_footnote_numbering(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "본문[^1] 다시[^1]."))
    [instruction] = [
        node for node in document.element.body.iter(f"{W}instrText") if "NOTEREF" in node.text
    ]
    instruction.getparent().getnext().getnext().find(f"{W}t").text = "i"
    from docx.oxml import OxmlElement

    footnote_pr = OxmlElement("w:footnotePr")
    number_format = OxmlElement("w:numFmt")
    number_format.set(qn("w:val"), "lowerRoman")
    footnote_pr.append(number_format)
    document.sections[0]._sectPr.insert(0, footnote_pr)
    assert inspect_document(document).warnings == []


def test_audit_reads_a_custom_mark_as_unnumbered(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, "하나[^1], 둘[^2] 다시 둘[^2]."))
    first = next(document.element.body.iter(f"{W}footnoteReference"))
    first.set(qn("w:customMarkFollows"), "1")
    # Note 2 is now Word's note 1, so the cached 2 is stale.
    assert inspect_document(document).warnings == [
        "Footnote cross-reference NOTEREF _Ref_fn_2 shows '2' but its footnote is number 1; "
        "update fields in Word (select all, F9)"
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# EARLIER TEST INPUTS THAT REFERENCED ONE NOTE SEVERAL TIMES
# ═══════════════════════════════════════════════════════════════════════════════
#
# `test_numeric_runs_keep_hyperlinks_spacing_and_explicit_value_scale`
# (test_render_hardening.py) and `test_review_five_footnote_then_line_break`
# (test_table_rules.py) pinned these inputs while they produced documents Word
# does not open. Those tests now use one note per reference and keep their
# assertions; the original inputs are pinned here with the fixed output.


def _render_model(markdown: str, profile: str) -> tuple[DocxDocument, IBDocumentRenderer]:
    from md_parser import MarkdownParser

    renderer = IBDocumentRenderer(
        options=RenderOptions(
            profile=profile,
            strict=True,
            include_cover=False,
            include_toc=False,
            include_disclaimer=False,
        )
    )
    document = renderer.render(MarkdownParser(profile=profile).parse(markdown))
    return saved(document), renderer


def test_numeric_cells_referencing_one_note_four_times() -> None:
    document, renderer = _render_model(
        "---\ntables:\n  - type: financial\n    columns: [text, money, percent, bps, multiple]\n"
        "---\n| Metric | Amount | Rate | Spread | Multiple |\n|---|---|---|---|---|\n"
        "| Revenue | 1234 [source](https://example.com)[^1] | **12.50**[^1] | 100.00[^1] "
        "| 2.50[^1] |\n\n[^1]: Audited amount.",
        "ib-report",
    )
    assert renderer.errors == []
    cells = document.tables[0].rows[1].cells
    # The cached number of each NOTEREF is visible text after the value.
    assert ["".join(cell._tc.xpath(".//w:t/text()")) for cell in cells] == [
        "Revenue",
        "1,234 source",
        "12.50%1",
        "100.00 bps1",
        "2.50x1",
    ]
    assert [len(cell._tc.xpath(".//w:footnoteReference")) for cell in cells] == [0, 1, 0, 0, 0]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")] * 3
    assert len(cells[1]._tc.xpath(".//w:hyperlink")) == 1
    assert cells[2].paragraphs[0].runs[0].bold
    assert_one_reference_per_note(document)


@pytest.mark.parametrize("style", ["paren", "minus", "triangle"])
def test_negative_style_cells_referencing_one_note_twice(style: str) -> None:
    table = "| 구분 | 값 |\n|---|---|\n| 다섯0 | -1200**.50**[^1]<br> |\n| 다섯1 | **-1300** [^1]<br> |\n"
    document, renderer = _render_model(
        "---\ntables:\n  - type: financial\n    columns: [text, money]\n"
        f"    negative: {style}\n---\n" + table + "\n[^1]: 가상 주석입니다.\n",
        "ib-memo",
    )
    assert renderer.errors == []
    assert renderer.warnings == [
        "Table negative style not applied to '-1200.50': the amount is split across "
        "differently formatted text, so it keeps its written form"
    ]
    shown = {"paren": "(1,300)", "minus": "-1,300", "triangle": "△1,300"}[style]
    [table_element] = document.tables
    assert table_element.cell(1, 1).text.startswith("-1,200.50")
    assert table_element.cell(2, 1).text.startswith(shown + " ")
    assert note_ids(document) == ["1"]
    assert note_fields(document) == [("NOTEREF _Ref_fn_1 \\f \\h", "1")]
    assert_one_reference_per_note(document)
