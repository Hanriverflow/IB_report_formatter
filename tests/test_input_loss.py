"""Input-loss fixes found on a converted real term sheet (synthetic data only)."""

from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from PIL import Image as PILImage

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, TextParser, parse_markdown_file

TERM_SHEET = (
    "---\nprofile: term-sheet\ntitle: 가나다머티리얼즈㈜ 조건 검토\n"
    "prepared_by: 라마바은행 자본시장부\ndisclaimer: 가상 조건 검토용입니다.\n---\n\n"
)


def _render_model(model, strict: bool = True):
    doc = IBDocumentRenderer(options=RenderOptions(strict=strict)).render(model)
    payload = BytesIO()
    doc.save(payload)
    return Document(payload)


def _render(markdown: str, profile: str = "business-report", strict: bool = True):
    return _render_model(MarkdownParser(profile=profile).parse(markdown), strict)


def _source(tmp_path: Path, body: str, front: str = "") -> Path:
    path = tmp_path / "memo.md"
    front = front or "---\nprofile: business-report\ntitle: 가상 문서\n---\n\n"
    path.write_text(front + body + "\n", encoding="utf-8")
    return path


def _png(path: Path, size=(40, 20)) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    PILImage.new("RGB", size, "navy").save(path)
    return path


def _texts(doc) -> str:
    return "".join(doc.element.body.xpath(".//w:t/text()"))


def _cell_text(cell) -> str:
    return "".join(run.text for run in cell.runs)


# ═══════════════════════════════════════════════════════════════════════════════
# EMPHASIS DELIMITERS
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("source,expected", [
    ("금액 2 * 3 * 4", [("금액 2 * 3 * 4", False, False)]),
    ("매출처* 앞 채권<br>* 주요 매출처 :", [("매출처* 앞 채권\n* 주요 매출처 :", False, False)]),
    ("값 ** 2 ** 3", [("값 ** 2 ** 3", False, False)]),
    ("a _ b _ c", [("a _ b _ c", False, False)]),
    ("*기울임* 과 _밑줄_ 과 **굵게**", [
        ("기울임", True, False), (" 과 ", False, False), ("밑줄", True, False),
        (" 과 ", False, False), ("굵게", False, True),
    ]),
    ("*a*", [("a", True, False)]),
])
def test_emphasis_needs_flanking_delimiters(source: str, expected: list) -> None:
    for parse in (TextParser.parse_runs, TextParser.parse_runs_plain):
        assert [(run.text, run.italic, run.bold) for run in parse(source)] == expected


def test_callout_keeps_spaced_asterisks_literal() -> None:
    doc = _render("# 메모\n\n> 금액 2 * 3 * 4 확인\n")
    assert "금액 2 * 3 * 4 확인" in _texts(doc)
    italic = [
        "".join(run.xpath("./w:t/text()"))
        for run in doc.element.body.xpath(".//w:r[w:rPr/w:i[not(@w:val) or @w:val!='0']]")
    ]
    assert not any("3" in text for text in italic)


# ═══════════════════════════════════════════════════════════════════════════════
# IMAGES: PERCENT-ENCODED PATHS, TABLE CELLS, INLINE
# ═══════════════════════════════════════════════════════════════════════════════


def test_percent_encoded_image_path_finds_the_file(tmp_path: Path) -> None:
    _png(tmp_path / "img dir" / "a.png")
    _png(tmp_path / "50%.png")
    source = _source(tmp_path, "![그림](img%20dir/a.png)\n\n![비율](50%.png)")
    doc = _render_model(parse_markdown_file(str(source)))
    assert len(doc.element.body.xpath(".//w:drawing")) == 2


def test_image_in_a_table_cell_is_inserted_and_fits_the_cell(tmp_path: Path) -> None:
    _png(tmp_path / "img" / "구조도.png", size=(2400, 900))
    source = _source(tmp_path, "| 구분 | 내용 |\n|---|---|\n| 구조 | ![구조도](<img/구조도.png>)<br>① 설명 |")
    doc = _render_model(parse_markdown_file(str(source)))
    table = doc.tables[0]
    cell = table.cell(1, 1)
    assert cell._tc.xpath(".//w:drawing")
    assert cell._tc.xpath(".//wp:docPr/@descr") == ["구조도"]
    assert "![" not in cell.text and "① 설명" in cell.text
    grid = [int(width) for width in table._tbl.xpath("./w:tblGrid/w:gridCol/@w:w")]
    assert int(cell._tc.xpath(".//wp:extent/@cx")[0]) <= grid[1] * 635


def test_term_sheet_cell_image_keeps_the_following_marker_line(tmp_path: Path) -> None:
    _png(tmp_path / "구조도.png", size=(1200, 500))
    source = _source(
        tmp_path, "| 구 분 | 내 용 |\n|---|---|\n| 금융구조 | ![구조도](구조도.png)<br>① 매출대금 집금 |",
        front=TERM_SHEET,
    )
    doc = _render_model(parse_markdown_file(str(source)))
    paragraphs = doc.tables[0].cell(1, 1).paragraphs
    assert paragraphs[0]._p.xpath(".//w:drawing")
    assert paragraphs[1].text == "① 매출대금 집금"


def test_missing_cell_image_is_rejected_in_strict_and_marked_otherwise(tmp_path: Path) -> None:
    source = _source(tmp_path, "| 구분 | 내용 |\n|---|---|\n| 구조 | ![없음](missing.png) |")
    with pytest.raises(ValueError, match="[Ii]mage"):
        _render_model(parse_markdown_file(str(source)))
    doc = _render_model(parse_markdown_file(str(source)), strict=False)
    assert doc.tables[0].cell(1, 1).text == "[Image: 없음]"


def test_inline_image_in_a_paragraph_and_escaped_image_syntax(tmp_path: Path) -> None:
    _png(tmp_path / "icon.png")
    source = _source(tmp_path, "앞 ![아이콘](icon.png) 뒤\n\n글자 !\\[x](icon.png) 유지")
    doc = _render_model(parse_markdown_file(str(source)))
    paragraph = next(p for p in doc.paragraphs if p.text.startswith("앞"))
    assert paragraph._p.xpath(".//w:drawing") and paragraph.text == "앞  뒤"
    assert "글자 ![x](icon.png) 유지" in _texts(doc)


def test_standalone_html_img_is_an_image(tmp_path: Path) -> None:
    _png(tmp_path / "logo.png")
    doc = _render_model(parse_markdown_file(str(_source(tmp_path, '<img src="logo.png" alt="로고">'))))
    assert len(doc.element.body.xpath(".//w:drawing")) == 1
    assert "<img" not in _texts(doc)


# ═══════════════════════════════════════════════════════════════════════════════
# HTML TABLES
# ═══════════════════════════════════════════════════════════════════════════════

HTML_TABLE = """<table>
<tr><th>구 분</th><th colspan="2">내 용</th></tr>
<tr><td rowspan="2">참여기관</td><td>• 차주</td><td>: 가나다머티리얼즈㈜</td></tr>
<tr><td>• 대주</td><td>: 라마바은행 &amp; 신탁부</td></tr>
<tr><td>비고</td><td colspan="2">첫 줄<br>둘째 줄 2 * 3 | <b>굵게</b> &lt;br&gt; ^^</td></tr>
</table>"""


@pytest.mark.parametrize("profile", ["plain", "business-report", "ib-report"])
def test_html_table_becomes_a_word_table_with_merges(profile: str) -> None:
    doc = _render("# 표\n\n" + HTML_TABLE + "\n\n본문 계속\n", profile=profile, strict=profile != "ib-report")
    table = next(table for table in doc.tables if "참여기관" in table._tbl.xml)
    assert (len(table.rows), len(table.columns)) == (4, 3)
    assert table._tbl.xpath("./w:tr[1]/w:tc[2]/w:tcPr/w:gridSpan/@w:val") == ["2"]
    assert table._tbl.xpath("./w:tr[4]/w:tc[2]/w:tcPr/w:gridSpan/@w:val") == ["2"]
    assert len(table._tbl.xpath(".//w:vMerge")) == 2
    assert table.cell(2, 2).text == ": 라마바은행 & 신탁부"
    assert table.cell(3, 1).text == "첫 줄\n둘째 줄 2 * 3 | 굵게 <br> ^^"
    assert table._tbl.xpath(".//w:r[w:rPr/w:b][w:t='굵게']")
    assert "<t" not in _texts(doc) and "본문 계속" in _texts(doc)


SCHEDULE = """<table>
<tr><th rowspan="2">회 차</th><th>기 간</th><th colspan="2">대출</th></tr>
<tr><td>(개월)</td><td>상환액</td><td>잔액</td></tr>
<tr><td>1</td><td>3</td><td>-</td><td>300</td></tr>
<tr><td colspan="2">합 계</td><td>300</td><td>-</td></tr>
</table>"""


def test_multi_row_html_header_is_flattened_into_one_header_row() -> None:
    table = MarkdownParser(profile="plain").parse(SCHEDULE).elements[0].content
    assert [_cell_text(cell) for cell in table.rows[0].cells] == ["회 차", "기 간\n(개월)", "대출\n상환액", "대출\n잔액"]
    assert [row.is_header for row in table.rows] == [True, False, False]
    doc = _render(SCHEDULE, profile="plain")
    assert doc.tables[0]._tbl.xpath("./w:tr[3]/w:tc[1]/w:tcPr/w:gridSpan/@w:val") == ["2"]


def test_nested_single_cell_table_with_an_image_becomes_a_cell_image(tmp_path: Path) -> None:
    _png(tmp_path / "images" / "a b.png")
    body = (
        "<table>\n<tr><th>구 분</th><th>내 용</th></tr>\n"
        '<tr><td>구조</td><td><table>\n<tr><th><img src="images/a%20b.png" alt="구조도"></th></tr>\n'
        "</table><br>① 설명</td></tr>\n</table>"
    )
    doc = _render_model(parse_markdown_file(str(_source(tmp_path, body))))
    assert len(doc.tables) == 1
    cell = doc.tables[0].cell(1, 1)
    assert cell._tc.xpath(".//w:drawing") and cell.text.strip() == "① 설명"


def test_html_table_in_term_sheet_uses_label_columns() -> None:
    doc = _render(TERM_SHEET + "## 1. 개요\n\n" + HTML_TABLE + "\n", profile="term-sheet")
    assert doc.tables[0]._tbl.xpath("./w:tr[2]/w:tc[1]/w:tcPr/w:shd/@w:fill")


@pytest.mark.parametrize("source,message", [
    ("<table><tr><td>가</td></tr>\n</table><table><tr><td>나</td></tr></table>\n", "further HTML table"),
    ("<table>캡션<tr><td>가</td></tr></table>\n", "outside the HTML table cells"),
])
def test_html_text_that_cannot_be_placed_is_reported(source: str, message: str) -> None:
    model = MarkdownParser(profile="plain").parse(source)
    assert any(message in warning for warning in model.warnings)
    assert _cell_text(model.elements[0].content.rows[0].cells[0]) == "가"


def test_unclosed_html_table_is_reported_and_strict_rejects() -> None:
    model = MarkdownParser(profile="plain").parse("<table>\n<tr><td>가</td></tr>\n")
    assert any("HTML table" in warning for warning in model.warnings)
    with pytest.raises(ValueError):
        _render_model(model)
