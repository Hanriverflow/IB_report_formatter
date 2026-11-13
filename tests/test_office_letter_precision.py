"""Office-letter closing rules for bodies that end with a table (synthetic letters).

The enforcement rule of the Korean administrative-document regulation (Art. 4(5)2)
writes "끝" one character in from the left margin below a table that is filled
to its last row. These tests pin that output through parsing, rendering and a
saved Word package.
"""

from io import BytesIO

from docx import Document
from docx.oxml.ns import qn

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
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
