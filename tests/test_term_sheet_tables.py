"""A2 table contracts: generic merge emission and term-sheet table layout.

Every test parses synthetic Markdown with the real parser, renders through
`IBDocumentRenderer.render`, saves to memory and inspects the reopened XML.
All names and figures are fictional.
"""

import hashlib
from io import BytesIO
from typing import List

import pytest
import yaml
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

import ib_renderer
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser


def profile_markdown(profile: str, body: str, tables=None) -> str:
    """Build fictional input for a non-term-sheet profile."""
    data = {"profile": profile, "title": "가상 검토 문서"}
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def render(content: str, **options):
    """Parse, render, save and reopen one synthetic document."""
    model = MarkdownParser().parse(content)
    renderer = IBDocumentRenderer(options=RenderOptions(**options))
    document = renderer.render(model)
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return renderer, Document(buffer)


def tables(saved) -> list:
    return saved.element.body.xpath("./w:tbl")


def rows(table) -> list:
    return table.xpath("./w:tr")


def cells(row) -> list:
    return row.xpath("./w:tc")


def text(element) -> str:
    return "".join(element.xpath(".//w:t/text()"))


def attr(element, path: str, name: str = "w:val") -> List[str]:
    return [node.get(qn(name)) for node in element.xpath(path)]


def fill(tc) -> List[str]:
    return attr(tc, "./w:tcPr/w:shd", "w:fill")


def following_paragraphs(table, count: int) -> list:
    siblings = []
    node = table.getnext()
    while node is not None and len(siblings) < count:
        if node.tag == qn("w:p"):
            siblings.append(node)
        node = node.getnext()
    return siblings


# ═══════════════════════════════════════════════════════════════════════════════
# GENERIC MERGE EMISSION (EVERY PROFILE)
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("profile", ["plain", "business-report", "ib-report"])
def test_explicit_spans_merge_owner_cells_once_in_other_profiles(profile: str) -> None:
    body = (
        "| 구분 | << | 값 |\n|---|---|---|\n"
        "| 항목[^1] | 설명 | 1234 |\n"
        "| ^^ | 추가 | 2345 |\n\n[^1]: 가상 각주입니다.\n"
    )
    content = profile_markdown(profile, body, [{"spans": True, "columns": ["text", "text", "money"]}])
    _, saved = render(content, strict=True, include_cover=False, include_toc=False, include_disclaimer=False)
    table = tables(saved)[0]
    header, first, second = rows(table)
    grid = [int(width) for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")]
    assert attr(cells(header)[0], "./w:tcPr/w:gridSpan") == ["2"]
    assert int(attr(cells(header)[0], "./w:tcPr/w:tcW", "w:w")[0]) == pytest.approx(grid[0] + grid[1], abs=1)
    assert attr(cells(first)[0], "./w:tcPr/w:vMerge") == ["restart"]
    assert attr(cells(second)[0], "./w:tcPr/w:vMerge") in ([None], ["continue"])
    assert text(table).count("항목") == 1 and text(table).count("1,234") == 1
    assert "^^" not in text(table) and "<<" not in text(table)
    assert len(table.xpath(".//w:footnoteReference")) == 1
    assert all(len(tc.xpath("./w:p")) == 1 for tc in table.xpath(".//w:tc"))
    assert all(len(tc_pr.xpath("./w:shd")) <= 1 for tc_pr in table.xpath(".//w:tcPr"))
    assert fill(cells(second)[0]) == fill(cells(first)[0])


def test_invalid_hand_built_merge_directive_renders_unmerged_with_diagnostic() -> None:
    model = MarkdownParser(profile="plain").parse("| 가 | 나 |\n|---|---|\n| 다 | 라 |")
    table_model = model.elements[0].content
    table_model.rows[1].cells[0].merge = "left"
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain"))
    document = renderer.render(model)
    table = document.element.body.xpath("./w:tbl")[0]
    assert not table.xpath(".//w:gridSpan") and not table.xpath(".//w:vMerge")
    assert text(table) == "가나다라"
    assert any("span" in error.lower() for error in renderer.errors)


# ═══════════════════════════════════════════════════════════════════════════════
# LEGACY OUTPUT LOCK (NO MERGES)
# ═══════════════════════════════════════════════════════════════════════════════

LEGACY_BODY = (
    "| 구분 | 내용 | 금액 |\n|---|:---:|---:|\n"
    "| 첫 항목 | 설명<br>둘째 줄 | 1234567 |\n"
    "| 둘째 | **굵게** [링크](https://example.com)[^1] | (1234) |\n"
    "| | 빈 칸 | |\n\n"
    "| ^^ | 내용 |\n|---|---|\n| 항목 | \\<< |\n\n"
    "| A | B |\n|---|---|\n| 1 | 2 |\n\n"
    "[^1]: 가상 각주입니다.\n"
)
LEGACY_SPECS = [
    {
        "caption": "가상 표", "unit": "억원", "as_of": "2026-09-29", "source": "가상 출처",
        "columns": ["text", "text", "money"], "spans": True,
    },
    {"spans": True},
]
# SHA-256 of the saved body XML produced before A2 (commit 26b9c74). Tables
# without merge groups must keep rendering byte-identically in every profile.
LEGACY_DIGESTS = {
    "plain": "03d688af5424686e3d0cedb10212af3b5f0de95bb3fb332b6c335eb7cedd8662",
    "ib-report": "4e5942dd846f5e1ceaf1c10f8ac1fd1142a3c8c07834ea53d5d023d33b12ca11",
}


@pytest.mark.parametrize("profile", sorted(LEGACY_DIGESTS))
def test_tables_without_merges_render_byte_identically(profile: str, monkeypatch) -> None:
    monkeypatch.setattr(ib_renderer.platform, "system", lambda: "Windows")
    content = profile_markdown(profile, LEGACY_BODY, LEGACY_SPECS)
    _, saved = render(content, include_cover=False, include_toc=False, include_disclaimer=False)
    body = etree.tostring(saved.element.body, encoding="unicode")
    assert hashlib.sha256(body.encode("utf-8")).hexdigest() == LEGACY_DIGESTS[profile]


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE NOTES (EVERY PROFILE)
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("profile", ["plain", "ib-memo"])
def test_note_is_rendered_after_source_in_other_profiles(profile: str) -> None:
    content = profile_markdown(profile, "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n", [{"note": "가상 주석", "source": "가상 자료"}])
    _, saved = render(content, strict=True)
    source, note, spacer = following_paragraphs(tables(saved)[0], 3)
    assert text(source) == "출처: 가상 자료"
    assert text(note) == "가상 주석" and attr(note, "./w:pPr/w:jc") == ["right"]
    assert not spacer.xpath(".//w:t")


@pytest.mark.parametrize("value", [["가상"], {"가": "나"}])
def test_note_specification_must_be_text(value) -> None:
    content = profile_markdown("plain", "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n", [{"note": value}])
    with pytest.raises(ValueError, match="Table note must be text"):
        MarkdownParser().parse(content)


def test_note_specification_is_kept_on_the_model() -> None:
    model = MarkdownParser().parse(profile_markdown("plain", "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n", [{"note": 3}]))
    assert model.elements[0].content.note == "3"
