"""Office-letter closing rules for bodies that end with a table (synthetic letters).

The enforcement rule of the Korean administrative-document regulation (Art. 4(5)2)
writes "끝" one character in from the left margin below a table that is filled
to its last row. A table filled only partway gets "이하 빈칸" in the cell after
the last written row instead; the letter declares that with
`letter.table_end: partial`, and the engine never infers it. These tests pin both
through parsing, rendering and a saved Word package.
"""

import sys
from io import BytesIO

import pytest
from docx import Document
from docx.oxml.ns import qn

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser

HEADER = """---
profile: office-letter
title: 사채 발행 일정 안내
document_no: 가상-2026-101
date: "2026. 11. 14."
sender:
  organization: 가상금융 주식회사
  department: 채권발행지원팀
  signatory: 가상금융 주식회사 대표이사
  contact: desk@example.com
recipients: [가상공사 재무처장]
{extra}---
# 사채 발행 일정 안내

1. 제5회 무보증사채 발행 일정을 아래와 같이 안내드립니다.

"""

FULL_TABLE = """| 구분 | 일자 | 비고 |
|---|---|---|
| 수요예측 | 2026. 12. 1. | 기관투자자 대상 |
| 납입 | 2026. 12. 8. | 대금 납입 |
"""


def letter(body: str, extra: str = "") -> str:
    """A fictional office letter whose body ends with `body`."""
    return HEADER.format(extra=extra) + body


def render(source: str, strict: bool = True):
    """Parse, render, save and reopen the letter as Word would receive it."""
    model = MarkdownParser(profile="office-letter").parse(source)
    doc = IBDocumentRenderer(options=RenderOptions(strict=strict)).render(model)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return Document(buffer)


def body_children(doc) -> list:
    """Top-level body elements without the section properties."""
    return [child for child in doc.element.body if child.tag != qn("w:sectPr")]


def text_of(element) -> str:
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def all_text(doc) -> str:
    return "\n".join(text_of(child) for child in body_children(doc))


def test_full_table_ending_writes_closing_mark_below_table():
    doc = render(letter(FULL_TABLE))
    children = body_children(doc)
    table_index = max(i for i, child in enumerate(children) if child.tag == qn("w:tbl"))
    closing = children[table_index + 1]
    assert closing.tag == qn("w:p")
    # One character (two half-width spaces) in from the left margin, below the table.
    assert text_of(closing) == "  끝."
    assert closing.find(qn("w:pPr") + "/" + qn("w:numPr")) is None
    assert all_text(doc).count("끝.") == 1
    assert "이하 빈칸" not in all_text(doc)
    assert text_of(children[table_index + 2]) == "가상금융 주식회사 대표이사"


def test_full_table_crossing_a_page_still_closes_below_its_last_row():
    rows = "".join(f"| {n} | 가상기관 {n:02d} | {n * 10} |\n" for n in range(1, 46))
    doc = render(letter("| 번호 | 기관 | 배정액 |\n|---|---|---|\n" + rows))
    children = body_children(doc)
    table_index = max(i for i, child in enumerate(children) if child.tag == qn("w:tbl"))
    assert len(children[table_index].findall(qn("w:tr"))) == 46
    assert text_of(children[table_index + 1]) == "  끝."
    assert all_text(doc).count("끝.") == 1


def test_table_with_blank_rows_keeps_closing_mark_without_a_setting():
    """Blank rows alone never make the engine guess that the table ended early."""
    doc = render(letter(FULL_TABLE + "| | | |\n| | | |\n"))
    children = body_children(doc)
    table_index = max(i for i, child in enumerate(children) if child.tag == qn("w:tbl"))
    assert text_of(children[table_index + 1]) == "  끝."
    assert "이하 빈칸" not in all_text(doc)


# ═══════════════════════════════════════════════════════════════════════════════
# A TABLE FILLED ONLY PARTWAY: letter.table_end: partial
# ═══════════════════════════════════════════════════════════════════════════════

PARTIAL = "letter:\n  table_end: partial\n"


def last_table(doc):
    return doc.element.body.findall(qn("w:tbl"))[-1]


def row_texts(table) -> list[list[str]]:
    return [[text_of(tc) for tc in tr.findall(qn("w:tc"))] for tr in table.findall(qn("w:tr"))]


def cell_alignment(table, row: int, column: int) -> str | None:
    tc = table.findall(qn("w:tr"))[row].findall(qn("w:tc"))[column]
    jc = tc.find(qn("w:p") + "/" + qn("w:pPr") + "/" + qn("w:jc"))
    return None if jc is None else jc.get(qn("w:val"))


def test_partial_table_marks_first_blank_row_and_omits_closing_mark():
    doc = render(letter(FULL_TABLE + "| | | |\n| | | |\n", PARTIAL))
    rows = row_texts(last_table(doc))
    assert rows[3] == ["이하 빈칸", "", ""]
    assert rows[4] == ["", "", ""]
    assert len(rows) == 5
    assert "끝." not in all_text(doc)
    assert all_text(doc).count("이하 빈칸") == 1


def test_mark_is_left_aligned_and_number_column_keeps_its_alignment():
    rows = "".join(f"| {n} | 가상기관 {n:02d} | {n * 10} |\n" for n in range(1, 41))
    doc = render(
        letter("| 번호 | 기관 | 배정액 |\n|---|---|---|\n" + rows + "| | | |\n" * 5, PARTIAL)
    )
    table = last_table(doc)
    assert row_texts(table)[41][0] == "이하 빈칸"
    assert cell_alignment(table, 41, 0) == "left"
    # The mark does not change how the number column is classified.
    assert cell_alignment(table, 40, 0) == "right"
    assert cell_alignment(table, 40, 2) == "right"


def test_row_merged_into_written_row_above_is_not_blank():
    source = (
        "| 단계 | 구분 | 일자 |\n|---|---|---|\n"
        "| 발행 전 | 수요예측 | 2026. 12. 1. |\n"
        "| ^^ | | |\n"
        "| | | |\n"
    )
    doc = render(letter(source, "tables:\n  - spans: true\n" + PARTIAL))
    rows = row_texts(last_table(doc))
    assert rows[3][0] == "이하 빈칸"
    assert "이하 빈칸" not in "".join(rows[2])


def test_header_only_table_marks_its_first_body_row():
    doc = render(letter("| 구분 | 일자 |\n|---|---|\n| | |\n| | |\n", PARTIAL))
    assert row_texts(last_table(doc))[1][0] == "이하 빈칸"


def test_partial_table_with_attachments_keeps_closing_mark_after_attachments():
    extra = PARTIAL + "attachments: [발행조건 요약표 1부]\n"
    doc = render(letter(FULL_TABLE + "| | | |\n", extra))
    text = all_text(doc)
    assert row_texts(last_table(doc))[3][0] == "이하 빈칸"
    assert text.count("끝.") == 1
    assert text.index("이하 빈칸") < text.index("발행조건 요약표 1부") < text.index("끝.")


def test_partial_table_before_explicit_appendix():
    extra = PARTIAL + "  appendix_heading: 상세 내역\n"
    source = letter(FULL_TABLE + "| | | |\n", extra) + (
        "\n# 상세 내역\n\n| 항목 | 값 |\n|---|---|\n| 코드 | 001234 |\n| | |\n"
    )
    doc = render(source)
    tables = doc.element.body.findall(qn("w:tbl"))
    assert row_texts(tables[0])[3][0] == "이하 빈칸"
    assert "이하 빈칸" not in "".join(sum(row_texts(tables[1]), []))
    assert "끝." not in all_text(doc)
    assert len(doc.element.xpath('.//w:br[@w:type="page"]')) == 1


@pytest.mark.parametrize("strict", [True, False])
@pytest.mark.parametrize(
    "body, extra, message",
    [
        (FULL_TABLE + "\n본문을 마칩니다.\n", PARTIAL, "end with a table"),
        ("본문을 마칩니다.\n", PARTIAL, "end with a table"),
        (FULL_TABLE, PARTIAL, "blank row"),
        ("| 구분 | 일자 |\n|---|---|\n| | |\n| 납입 | 2026. 12. 8. |\n", PARTIAL, "blank row"),
        (FULL_TABLE + "| | | |\n", "letter:\n  table_end: full\n", "must be 'partial'"),
        (FULL_TABLE + "| | | |\n", "letter:\n  table_end: true\n", "non-empty text"),
    ],
)
def test_table_end_setting_fails_closed(body, extra, message, strict):
    with pytest.raises(StrictValidationError, match=message):
        render(letter(body, extra), strict=strict)


def test_table_end_in_appendix_only_does_not_count_as_letter_body():
    extra = PARTIAL + "  appendix_heading: 상세 내역\n"
    source = letter("본문을 마칩니다.\n", extra) + "\n# 상세 내역\n\n" + FULL_TABLE + "| | | |\n"
    with pytest.raises(StrictValidationError, match="end with a table"):
        render(source)


def test_invalid_table_end_cli_does_not_save(tmp_path, monkeypatch):
    from md_to_word import main

    source, output = tmp_path / "letter.md", tmp_path / "letter.docx"
    source.write_text(letter(FULL_TABLE, PARTIAL), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["md-to-word", str(source), str(output), "--strict"])
    with pytest.raises(SystemExit) as result:
        main()
    assert result.value.code != 0
    assert not output.exists()
