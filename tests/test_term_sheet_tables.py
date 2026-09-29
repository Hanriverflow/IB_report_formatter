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
from docx.text.paragraph import Paragraph as DocxParagraph
from lxml import etree

import ib_renderer
import term_sheet
from document_model import (
    DocumentMetadata,
    DocumentModel,
    Element,
    ElementType,
    Table,
    TableCell,
    TableRow,
    TextRun,
)
from document_profiles import RenderOptions
from docx_audit import inspect_document
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TITLE = "가나다머티리얼즈㈜ 구조화금융"
FIELDS = {
    "profile": "term-sheet",
    "title": TITLE,
    "prepared_by": "라마바은행 자본시장부",
    "disclaimer": "본 자료는 가상 조건 검토용이며 거래 확약이 아닙니다.",
}

KEY_VALUE = (
    "| 구 분 | << | 내 용 |\n|---|---|---|\n"
    "| 신탁원본 | << | • 향후 3년간 발생하는 장래매출채권<br>- 집금계좌 예금반환채권 |\n"
    "| ABCP | 금액 | • 500억원 |\n"
    "| ^^ | CAP | • CD(3개월) + [1.10]% (변동)<br>※ 기준일 CD(3M) 기준 |\n"
)
ONE_LABEL = "| 구 분 | 내 용 |\n|---|---|\n| 발행금액 | • 500억원 |\n"
GRID = (
    "| 구 분 | 1차 | 2차 | 합계 |\n|---|---:|---:|---:|\n"
    "| 선순위 | 100 | 200 | 300 |\n"
    "| 합 계 | << | << | 300 |\n"
)


def ts_markdown(body: str, tables=None, **fields) -> str:
    """Build fictional term-sheet input through normal YAML frontmatter."""
    data = dict(FIELDS)
    data.update(fields)
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


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


def run_values(paragraph, path: str, name: str = "w:val") -> List[str]:
    return [node.get(qn(name)) for node in paragraph.xpath(".//w:r/w:rPr/" + path)]


def bold_runs(paragraph) -> list:
    """Runs whose own properties switch bold on (python-docx writes <w:b/>)."""
    return paragraph.xpath(".//w:r/w:rPr/w:b[not(@w:val) or @w:val='1' or @w:val='true']")


def twips(millimetres: float) -> int:
    return round(millimetres / 25.4 * 1440)


def printable_twips(section) -> int:
    return round((section.page_width - section.left_margin - section.right_margin) / 635)


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


def test_term_sheet_merge_emits_spans_once_without_markers() -> None:
    _, saved = render(ts_markdown(KEY_VALUE), strict=True)
    table = tables(saved)[0]
    grid = [int(width) for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")]
    header, first, abcp, cap = rows(table)
    assert len(cells(header)) == 2 and len(cells(first)) == 2
    assert attr(cells(header)[0], "./w:tcPr/w:gridSpan") == ["2"]
    assert attr(cells(first)[0], "./w:tcPr/w:gridSpan") == ["2"]
    merged_width = int(attr(cells(header)[0], "./w:tcPr/w:tcW", "w:w")[0])
    assert merged_width == pytest.approx(grid[0] + grid[1], abs=1)
    owner, continuation = cells(abcp)[0], cells(cap)[0]
    assert attr(owner, "./w:tcPr/w:vMerge") == ["restart"]
    assert attr(continuation, "./w:tcPr/w:vMerge") in ([None], ["continue"])
    body_text = text(table)
    assert body_text.count("ABCP") == 1 and body_text.count("신탁원본") == 1
    assert "^^" not in body_text and "<<" not in body_text
    assert all(len(tc_pr.xpath("./w:shd")) <= 1 for tc_pr in table.xpath(".//w:tcPr"))
    assert fill(continuation) == fill(owner) == ["F2F5FC"]
    assert len(continuation.xpath("./w:p")) == 1 and not continuation.xpath(".//w:t")


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


@pytest.mark.parametrize("profile", ["plain", "term-sheet"])
def test_hand_built_merge_never_joins_the_rendered_header_row_to_the_body(profile: str) -> None:
    # Default TableRow.is_header is False, but the renderer always repeats row 0 as the header.
    table_model = Table(
        rows=[
            TableRow(cells=[TableCell("Header"), TableCell("Value")]),
            TableRow(cells=[TableCell("^^", merge="up"), TableCell("Body")]),
        ],
        col_count=2,
    )
    extra = {"prepared_by": FIELDS["prepared_by"], "disclaimer": FIELDS["disclaimer"]}
    metadata = DocumentMetadata(title=TITLE, company="", sector="", analyst="", profile=profile, extra=extra)
    model = DocumentModel(metadata=metadata, elements=[Element(ElementType.TABLE, table_model)])
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile))
    table = renderer.render(model).element.body.xpath("./w:tbl")[0]
    assert not table.xpath(".//w:vMerge") and text(table) == "HeaderValue^^Body"
    assert any("span" in error.lower() for error in renderer.errors)
    with pytest.raises(ValueError, match="validation"):
        IBDocumentRenderer(options=RenderOptions(profile=profile, strict=True)).render(model)


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


# ═══════════════════════════════════════════════════════════════════════════════
# TERM-SHEET TABLE LAYOUT
# ═══════════════════════════════════════════════════════════════════════════════


def test_key_value_tables_share_the_fixed_label_grid() -> None:
    _, saved = render(ts_markdown(ONE_LABEL + "\n" + KEY_VALUE), strict=True)
    grids = [[int(width) for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")] for table in tables(saved)]
    total = printable_twips(saved.sections[0])
    assert grids[0][0] == grids[1][0] == twips(33.5)
    assert grids[1][1] == twips(30)
    assert sum(grids[0]) == pytest.approx(total, abs=1)
    assert sum(grids[1]) == pytest.approx(total, abs=1)
    for table, grid in zip(tables(saved), grids):
        assert attr(table, "./w:tblPr/w:tblLayout", "w:type") == ["fixed"]
        assert attr(table, "./w:tblPr/w:tblW", "w:type") == ["dxa"]
        assert int(attr(table, "./w:tblPr/w:tblW", "w:w")[0]) == pytest.approx(sum(grid), abs=1)
        assert attr(table, "./w:tblPr/w:tblInd", "w:w") == [str(term_sheet.CELL_MARGIN_HORIZONTAL)]
        first_cells = [cells(row)[0] for row in rows(table)[1:]]
        assert {attr(tc, "./w:tcPr/w:tcW", "w:w")[0] for tc in first_cells if not tc.xpath("./w:tcPr/w:gridSpan")} == {str(grid[0])}


def test_landscape_key_value_table_takes_the_remaining_section_width() -> None:
    _, saved = render(ts_markdown("앞 문단.\n\n" + ONE_LABEL, tables=[{"landscape": True}]), strict=True)
    table = tables(saved)[0]
    grid = [int(width) for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")]
    landscape = saved.sections[1]
    assert landscape.page_width > landscape.page_height
    assert grid[0] == twips(33.5)
    assert sum(grid) == pytest.approx(printable_twips(landscape), abs=1)


def test_term_sheet_frame_borders_margins_and_header_repeat() -> None:
    _, saved = render(ts_markdown(KEY_VALUE), strict=True)
    table = tables(saved)[0]
    borders = table.xpath("./w:tblPr/w:tblBorders/*")
    assert [etree.QName(node).localname for node in borders] == [
        "top", "left", "bottom", "right", "insideH", "insideV",
    ]
    assert {(node.get(qn("w:val")), node.get(qn("w:sz")), node.get(qn("w:color"))) for node in borders} == {
        ("single", "4", "9AA5C4"),
    }
    margins = {
        etree.QName(node).localname: (node.get(qn("w:w")), node.get(qn("w:type")))
        for node in table.xpath("./w:tblPr/w:tblCellMar/*")
    }
    assert margins == {"top": ("70", "dxa"), "left": ("100", "dxa"), "bottom": ("70", "dxa"), "right": ("100", "dxa")}
    header = rows(table)[0]
    assert header.xpath("./w:trPr/w:tblHeader") and header.xpath("./w:trPr/w:cantSplit")
    assert [etree.QName(node).localname for node in header.xpath("./w:trPr/*")] == ["cantSplit", "tblHeader"]
    for tc in cells(header):
        assert fill(tc) == ["DCE3F5"]
        paragraph = tc.xpath("./w:p")[0]
        assert attr(paragraph, "./w:pPr/w:jc") == ["center"]
        assert set(run_values(paragraph, "w:color")) == {"1A2270"}
        assert len(bold_runs(paragraph)) == len(paragraph.xpath(".//w:r"))
    owners = table.xpath(".//w:tc[not(w:tcPr/w:vMerge) or w:tcPr/w:vMerge/@w:val='restart']")
    assert owners and all(attr(tc, "./w:tcPr/w:vAlign") == ["center"] for tc in owners)


def test_key_value_label_hierarchy_and_content_cells() -> None:
    _, saved = render(ts_markdown(KEY_VALUE), strict=True)
    abcp_row = rows(tables(saved)[0])[2]
    label, sublabel, content = cells(abcp_row)
    assert fill(label) == ["F2F5FC"] and fill(sublabel) == [] and fill(content) == []
    label_p, sublabel_p, content_p = (tc.xpath("./w:p")[0] for tc in (label, sublabel, content))
    assert attr(label_p, "./w:pPr/w:jc") == attr(sublabel_p, "./w:pPr/w:jc") == ["center"]
    assert bold_runs(label_p) and set(run_values(label_p, "w:color")) == {"000000"}
    assert not bold_runs(sublabel_p)
    assert set(run_values(sublabel_p, "w:color")) == {"555555"}
    assert attr(content_p, "./w:pPr/w:jc") in ([], ["left"])


def test_grid_table_labels_are_shaded_regular_and_totals_keep_label_fill() -> None:
    _, saved = render(ts_markdown(GRID), strict=True)
    table = tables(saved)[0]
    header, senior, total = rows(table)
    label = cells(senior)[0]
    assert fill(label) == ["F2F5FC"]
    label_p = label.xpath("./w:p")[0]
    assert attr(label_p, "./w:pPr/w:jc") == ["center"]
    assert not bold_runs(label_p) and set(run_values(label_p, "w:color")) == {"000000"}
    total_owner = cells(total)[0]
    assert attr(total_owner, "./w:tcPr/w:gridSpan") == ["3"]
    assert fill(total_owner) == ["F2F5FC"] and text(total_owner) == "합 계"
    assert [attr(tc.xpath("./w:p")[0], "./w:pPr/w:jc") for tc in cells(senior)[1:]] == [["right"]] * 3
    grid = [int(width) for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")]
    assert sum(grid) == pytest.approx(printable_twips(saved.sections[0]), abs=2)


def test_type_styling_replaces_a_label_fill_instead_of_duplicating_it() -> None:
    body = "| 금리 | 1% | 2% |\n|---|---|---|\n| 3.0% | 10 | 11 |\n| 3.5% | 9 | 10 |\n"
    spec = [{"type": "sensitivity", "base_case": {"row": 1, "column": 1}}]
    _, saved = render(ts_markdown(body, tables=spec), strict=True)
    table = tables(saved)[0]
    base, neighbour = cells(rows(table)[1])[0], cells(rows(table)[2])[0]
    assert fill(base) == ["FFFF00"] and fill(neighbour) == ["F2F5FC"]
    assert bold_runs(base.xpath("./w:p")[0])
    assert all(len(tc_pr.xpath("./w:shd")) <= 1 for tc_pr in table.xpath(".//w:tcPr"))


def test_cell_lines_become_paragraphs_with_marker_hanging_indents() -> None:
    _, saved = render(ts_markdown(KEY_VALUE), strict=True)
    table = tables(saved)[0]
    first_content = cells(rows(table)[1])[1]
    bullet, dash = first_content.xpath("./w:p")
    assert text(bullet).startswith("• 향후") and text(dash).startswith("- 집금계좌")
    assert (attr(bullet, "./w:pPr/w:ind", "w:left"), attr(bullet, "./w:pPr/w:ind", "w:hanging")) == (["170"], ["170"])
    assert (attr(dash, "./w:pPr/w:ind", "w:left"), attr(dash, "./w:pPr/w:ind", "w:hanging")) == (["340"], ["170"])
    cap_content = cells(rows(table)[3])[2]
    variable, note = cap_content.xpath("./w:p")
    assert text(note) == "※ 기준일 CD(3M) 기준"
    assert (attr(note, "./w:pPr/w:ind", "w:left"), attr(note, "./w:pPr/w:ind", "w:hanging")) == (["227"], ["227"])
    assert set(run_values(note, "w:sz")) == {"16"}
    assert set(run_values(variable, "w:sz")) == {"18"}
    for paragraph in first_content.xpath("./w:p"):
        assert attr(paragraph, "./w:pPr/w:spacing", "w:after") == ["30"]
        assert attr(paragraph, "./w:pPr/w:spacing", "w:line") == ["252"]


@pytest.mark.parametrize(
    ("line", "left", "hanging"),
    [("· 3수준", "510", "170"), ("① 번호", "255", "255"), ("⑳ 번호", "255", "255"), ("-5%", None, None), ("-", None, None)],
)
def test_marker_table_matches_the_plan(line: str, left, hanging) -> None:
    _, saved = render(ts_markdown(ONE_LABEL.replace("• 500억원", "첫 줄<br>" + line)), strict=True)
    paragraph = cells(rows(tables(saved)[0])[1])[1].xpath("./w:p")[1]
    assert text(paragraph) == line
    assert attr(paragraph, "./w:pPr/w:ind", "w:left") == ([left] if left else [])
    assert attr(paragraph, "./w:pPr/w:ind", "w:hanging") == ([hanging] if hanging else [])


def test_row_cant_split_depends_on_the_estimated_line_count() -> None:
    long_lines = "<br>".join(f"• 가상 조건 {index}" for index in range(14))
    body = (
        "| 구 분 | 내 용 |\n|---|---|\n"
        "| 짧은 행 | • 한 줄 |\n"
        f"| 긴 행 | {long_lines} |\n"
        f"| 긴 문장 | {'가' * 700} |\n"
    )
    _, saved = render(ts_markdown(body), strict=True)
    header, short, lines, sentence = rows(tables(saved)[0])
    assert header.xpath("./w:trPr/w:cantSplit") and short.xpath("./w:trPr/w:cantSplit")
    assert not lines.xpath("./w:trPr/w:cantSplit")
    assert not sentence.xpath("./w:trPr/w:cantSplit")


def test_multi_row_owner_is_excluded_from_row_estimates() -> None:
    long_lines = "<br>".join(f"가상 {index}" for index in range(20))
    body = (
        "| 구 분 | 항 목 | 내 용 |\n|---|---|---|\n"
        f"| {long_lines} | 가 | 짧음 |\n"
        "| ^^ | 나 | 짧음 |\n"
    )
    _, saved = render(ts_markdown(body), strict=True)
    _, first, second = rows(tables(saved)[0])
    assert first.xpath("./w:trPr/w:cantSplit") and second.xpath("./w:trPr/w:cantSplit")


def test_caption_unit_line_and_trailing_note_source_spacer() -> None:
    spec = {"caption": "주요 금융조건", "unit": "억원", "as_of": "2026-09-29", "note": "(VAT 별도)", "source": "가상 자료"}
    _, saved = render(ts_markdown(ONE_LABEL, tables=[spec]), strict=True)
    table = tables(saved)[0]
    heading = table.getprevious()
    paragraph = DocxParagraph(heading, saved._body)
    assert paragraph.text == "주요 금융조건\t(단위 : 억원, 기준일 : 2026-09-29)"
    assert paragraph.paragraph_format.keep_with_next
    stops = heading.xpath("./w:pPr/w:tabs/w:tab")
    assert [(stop.get(qn("w:val")), int(stop.get(qn("w:pos")))) for stop in stops] == [
        ("right", printable_twips(saved.sections[0])),
    ]
    caption_runs = [run for run in paragraph.runs if run.text == "주요 금융조건"]
    assert caption_runs and caption_runs[0].bold and caption_runs[0].font.size.pt == 9
    assert str(caption_runs[0].font.color.rgb) == "1A2270"
    unit_runs = [run for run in paragraph.runs if "단위" in run.text]
    assert unit_runs[0].font.size.pt == 8 and str(unit_runs[0].font.color.rgb) == "555555"
    note, source, spacer = following_paragraphs(table, 3)
    assert text(note) == "(VAT 별도)" and attr(note, "./w:pPr/w:jc") == ["right"]
    assert set(run_values(note, "w:sz")) == {"16"} and set(run_values(note, "w:color")) == {"555555"}
    assert text(source) == "출처 : 가상 자료" and set(run_values(source, "w:sz")) == {"16"}
    assert not spacer.xpath(".//w:t")
    spacing = spacer.xpath("./w:pPr/w:spacing")[0]
    assert (spacing.get(qn("w:line")), spacing.get(qn("w:lineRule"))) == ("80", "exact")
    assert (spacing.get(qn("w:before")), spacing.get(qn("w:after"))) == ("0", "0")


def test_unit_only_heading_is_right_aligned() -> None:
    _, saved = render(ts_markdown(ONE_LABEL, tables=[{"unit": "억원"}]), strict=True)
    heading = tables(saved)[0].getprevious()
    assert text(heading) == "(단위 : 억원)" and attr(heading, "./w:pPr/w:jc") == ["right"]


def test_split_run_lines_preserves_formatting_links_footnotes_and_blank_lines() -> None:
    runs = [
        TextRun("첫 줄 \n 둘째", bold=True),
        TextRun("링크", hyperlink="https://example.com"),
        TextRun("1", superscript=True, footnote_id=1),
        TextRun("\n\n셋째 "),
    ]
    lines = term_sheet.split_run_lines(runs)
    assert [[run.text for run in line] for line in lines] == [["첫 줄"], ["둘째", "링크", "1"], [], ["셋째 "]]
    assert lines[0][0].bold and lines[1][0].bold
    assert lines[1][1].hyperlink == "https://example.com"
    assert lines[1][2].footnote_id == 1 and lines[1][2].superscript
    assert term_sheet.split_run_lines([]) == [[]]
    term = TextRun(" 500억원 ", term_key="amount")
    assert term_sheet.split_run_lines([TextRun("앞 \n"), term]) == [[TextRun("앞")], [term]]


TCPR_ORDER = (
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
    "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel",
    "cellMerge", "tcPrChange",
)
TBLPR_ORDER = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
    "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
)
PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
    "rPr", "sectPr", "pPrChange",
)


def assert_schema_order(element, sequence) -> None:
    names = [etree.QName(child).localname for child in element]
    assert set(names) <= set(sequence) and len(set(names)) == len(names), names
    positions = [sequence.index(name) for name in names]
    assert positions == sorted(positions), names


@pytest.mark.parametrize("zebra", [False, True])
def test_term_sheet_properties_follow_the_schema_order(zebra: bool, tmp_path) -> None:
    options = {}
    if zebra:
        theme = tmp_path / "zebra.yaml"
        theme.write_text("TABLE_ZEBRA: true\n", encoding="utf-8")
        options["theme"] = str(theme)
    body = (
        ONE_LABEL + "\n" + KEY_VALUE + "\n" + GRID + "\n"
        "| 금리 | 1% | 2% |\n|---|---|---|\n| 3.0% | 10 | 11 |\n| 3.5% | 9 | 10 |\n\n"
        "① 가상 설명<br>※ 가상 주석\n\n```confirmation\n```\n"
    )
    spec = [
        {"caption": "개요", "unit": "억원", "note": "(VAT 별도)", "source": "가상 자료"},
        {"landscape": True},
        {"type": "risk"},
        {"type": "sensitivity", "base_case": {"row": 1, "column": 2}},
    ]
    confirmation = {"intro": "가상 안내", "items": ["□ 가상 항목"], "signature": "가상 서명"}
    _, saved = render(ts_markdown(body, tables=spec, confirmation=confirmation), strict=True, **options)
    body_element = saved.element.body
    assert body_element.xpath(".//w:tcPr/w:shd[@w:fill='FFFF00']")
    for tbl_pr in body_element.xpath(".//w:tblPr"):
        assert_schema_order(tbl_pr, TBLPR_ORDER)
    for tc_pr in body_element.xpath(".//w:tcPr"):
        assert_schema_order(tc_pr, TCPR_ORDER)
    for p_pr in body_element.xpath(".//w:pPr"):
        assert_schema_order(p_pr, PPR_ORDER)
    defaults = etree.XPath("./w:docDefaults/w:pPrDefault/w:pPr", namespaces=W_NS)
    for p_pr in defaults(saved.styles.element):
        assert_schema_order(p_pr, PPR_ORDER)
    assert_schema_order(saved.styles["Heading 2"].element.pPr, PPR_ORDER)


def test_strict_term_sheet_tables_have_no_audit_issues() -> None:
    body = ONE_LABEL + "\n" + KEY_VALUE + "\n" + GRID
    spec = [{"caption": "개요", "unit": "억원"}, {"note": "가상 주석"}, {"landscape": True, "source": "가상 자료"}]
    renderer, saved = render(ts_markdown(body, tables=spec), strict=True)
    assert not renderer.errors
    assert inspect_document(saved).issues == []
