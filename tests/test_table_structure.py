"""Table structure found on a second real term sheet (synthetic data only).

Merged cells that reach the content columns, several header rows (table spec
`header_rows`, HTML tables) and content-sized term-sheet grid columns. Every
test parses Markdown with the real parser and renders through
`IBDocumentRenderer.render`.
"""

from io import BytesIO
from typing import List

import pytest
import yaml
from docx import Document
from docx.oxml.ns import qn

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

TERM_SHEET = {
    "profile": "term-sheet",
    "title": "가나다머티리얼즈㈜ 구조화금융",
    "prepared_by": "라마바은행 자본시장부",
    "disclaimer": "본 자료는 가상 조건 검토용이며 거래 확약이 아닙니다.",
}


def markdown(body: str, profile: str = "term-sheet", tables=None) -> str:
    data = dict(TERM_SHEET) if profile == "term-sheet" else {"profile": profile, "title": "가상 검토 문서"}
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def render(content: str, strict: bool = True):
    doc = IBDocumentRenderer(options=RenderOptions(strict=strict)).render(MarkdownParser().parse(content))
    buffer = BytesIO()
    doc.save(buffer)
    return Document(buffer)


def body_table(doc, marker: str):
    return next(table for table in doc.tables if marker in table._tbl.xml)._tbl


def cells(row) -> list:
    return row.xpath("./w:tc")


def attr(element, path: str, name: str = "w:val") -> List[str]:
    return [node.get(qn(name)) for node in element.xpath(path)]


def text(element) -> str:
    return "".join(element.xpath(".//w:t/text()"))


def grid_mm(table) -> List[float]:
    return [int(width) / 1440 * 25.4 for width in attr(table, "./w:tblGrid/w:gridCol", "w:w")]


# ═══════════════════════════════════════════════════════════════════════════════
# MERGED CELLS REACHING THE CONTENT COLUMNS
# ═══════════════════════════════════════════════════════════════════════════════

PARTICIPANTS = (
    "| 구 분 | 내 용 | << | << |\n|---|---|---|---|\n"
    "| 거래개요 | • 가나다머티리얼즈㈜의 장래매출채권을 기초로 한 가상 구조화금융 | << | << |\n"
    "| 참여기관 | 라마바은행 | • 대주 | : 라마바은행 |\n"
    "| ^^ | ^^ | • 수탁 | : 라마바은행 신탁부 |\n"
    "| 합 계 | << | << | 2개 기관 |\n"
)


def test_a_cell_from_a_second_label_column_into_content_is_content() -> None:
    table = body_table(render(markdown(PARTICIPANTS, tables=[{"label_columns": 2}])), "거래개요")
    overview = cells(table.xpath("./w:tr")[1])[1]
    assert attr(overview, "./w:tcPr/w:gridSpan") == ["3"]
    assert not attr(overview, "./w:tcPr/w:shd", "w:fill")
    assert attr(overview.xpath("./w:p")[0], "./w:pPr/w:jc") in ([], ["left"])
    # A second-tier label that stays in its column keeps the label layout.
    bank = cells(table.xpath("./w:tr")[2])[1]
    assert attr(bank, "./w:tcPr/w:shd", "w:fill") and attr(bank.xpath("./w:p")[0], "./w:pPr/w:jc") == ["center"]
    # A label that starts in the first column stays a label however far it spans.
    total = cells(table.xpath("./w:tr")[4])[0]
    assert text(total) == "합 계" and attr(total, "./w:tcPr/w:shd", "w:fill")


# ═══════════════════════════════════════════════════════════════════════════════
# SEVERAL HEADER ROWS
# ═══════════════════════════════════════════════════════════════════════════════

SCHEDULE = (
    "| 회 차 | 기 간 | 대출 | << |\n|---|---|---|---|\n"
    "| ^^ | (개월) | 상환액 | 잔액 |\n"
    "| 1 | 3 | - | 300 |\n"
    "| 2 | 6 | 100 | 200 |\n"
)


@pytest.mark.parametrize("profile", ["term-sheet", "business-report", "ib-report"])
def test_header_rows_are_drawn_merged_and_repeated(profile: str) -> None:
    # Span markers are the term-sheet default and opt-in elsewhere.
    spec = {"header_rows": 2} if profile == "term-sheet" else {"header_rows": 2, "spans": True}
    doc = render(markdown(SCHEDULE, profile, tables=[spec]), strict=profile != "ib-report")
    table = body_table(doc, "상환액")
    word_rows = table.xpath("./w:tr")
    assert len(word_rows) == 4
    assert [bool(row.xpath("./w:trPr/w:tblHeader")) for row in word_rows] == [True, True, False, False]
    first, second = word_rows[0], word_rows[1]
    assert attr(cells(first)[0], "./w:tcPr/w:vMerge") == ["restart"]
    assert attr(cells(second)[0], "./w:tcPr/w:vMerge") in ([None], ["continue"])
    assert attr(cells(first)[2], "./w:tcPr/w:gridSpan") == ["2"]
    header_fills = {tuple(attr(tc, "./w:tcPr/w:shd", "w:fill")) for row in (first, second) for tc in cells(row)}
    assert len(header_fills) == 1 and header_fills != {()}
    assert "^^" not in text(table) and "<<" not in text(table)


def test_header_rows_shift_the_body_for_base_case_and_zebra() -> None:
    body = (
        "| 금리 | 할인율 | << |\n|---|---|---|\n| ^^ | 1% | 2% |\n"
        "| 3.0% | 10 | 11 |\n| 3.5% | 9 | 10 |\n"
    )
    spec = [{"type": "sensitivity", "header_rows": 2, "base_case": {"row": 1, "column": 2}}]
    model = MarkdownParser().parse(markdown(body, "business-report", tables=spec))
    table = model.elements[-1].content
    assert table.header_rows == 2 and [row.is_header for row in table.rows] == [True, True, False, False]
    assert table.rows[2].cells[1].is_base_case
    word = body_table(render(markdown(body, "business-report", tables=spec)), "할인율")
    # The first body row has the base-case fill; zebra shading counts body rows only.
    assert attr(cells(word.xpath("./w:tr")[2])[1], "./w:tcPr/w:shd", "w:fill") == ["FFFF00"]


@pytest.mark.parametrize("value", [0, "2", 4, True])
def test_header_rows_must_leave_a_body_row(value) -> None:
    with pytest.raises(ValueError, match="header_rows"):
        MarkdownParser().parse(markdown(SCHEDULE, "business-report", tables=[{"header_rows": value}]))


def test_html_header_rows_stay_separate_rows() -> None:
    source = (
        "<table>\n<tr><th rowspan=\"2\">회 차</th><th>기 간</th><th colspan=\"2\">대출</th></tr>\n"
        "<tr><td>(개월)</td><td>상환액</td><td>잔액</td></tr>\n"
        "<tr><td>1</td><td>3</td><td>-</td><td>300</td></tr>\n</table>\n"
    )
    table = MarkdownParser(profile="plain").parse(source).elements[0].content
    assert table.header_rows == 2 and [row.is_header for row in table.rows] == [True, True, False]
    word = body_table(render(markdown(source, "plain")), "상환액")
    assert [bool(row.xpath("./w:trPr/w:tblHeader")) for row in word.xpath("./w:tr")] == [True, True, False]
    assert attr(cells(word.xpath("./w:tr")[0])[2], "./w:tcPr/w:gridSpan") == ["2"]


# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN WIDTHS
# ═══════════════════════════════════════════════════════════════════════════════


def _schedule(rows: int = 21) -> str:
    lines = [
        f"| {index} | 2027년 {index % 12 + 1}월 | {3 * index} | {'-' if index < 9 else 14} | {500 - max(0, index - 8) * 14} |"
        for index in range(rows)
    ]
    return (
        "| 회 차 | 일 자 | 기 간<br>(개 월) | 상환액 | 잔액 |\n|---|---|---|---|---|\n"
        + "\n".join(lines) + "\n| 합 계 | << | - | 500 | - |\n"
    )


def test_term_sheet_grid_columns_fit_their_content_and_share_the_rest() -> None:
    doc = render(markdown(_schedule()))
    table = body_table(doc, "상환액")
    widths = grid_mm(table)
    section = doc.sections[0]
    printable = (section.page_width - section.left_margin - section.right_margin) / 36000
    assert sum(widths) == pytest.approx(printable, abs=0.2)
    # Short columns share the spare width: none is starved, none balloons, and
    # the column with the longest content (일 자) stays the widest.
    assert max(widths) / min(widths) < 2.2
    # Content decides the order: the date column is strictly the widest.
    assert all(widths[1] > width + 2 for index, width in enumerate(widths) if index != 1)


def test_text_heavy_grid_tables_keep_the_content_estimate() -> None:
    body = (
        "| 구 분 | 조 건 | 비 고 |\n|---|---|---|\n"
        "| 신용공여 | " + "가상 조건 문구가 길게 이어지는 설명입니다. " * 6 + "| 가상 비고 |\n"
    )
    widths = grid_mm(body_table(render(markdown(body)), "신용공여"))
    assert widths[1] > widths[0] * 2  # the long text column still takes the room


def test_grid_columns_with_images_keep_a_usable_width() -> None:
    import base64
    from io import BytesIO as Buffer

    from PIL import Image as PILImage

    buffer = Buffer()
    PILImage.new("RGB", (100, 100), "navy").save(buffer, format="PNG")
    uri = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()
    body = f"| A | | C |\n|---|---|---|\n| {'X' * 97} | ![]({uri}) | Y |\n"
    table = body_table(render(markdown(body)), "XXXX")
    widths = grid_mm(table)
    assert min(widths) >= 16  # the estimator's numeric minimum (0.65 in)
    extent = int(table.xpath(".//wp:extent/@cx")[0]) / 36000
    assert extent <= widths[1]


def test_zebra_shading_starts_on_the_first_body_row_after_header_rows() -> None:
    from document_profiles import get_profile, load_style

    style = load_style(get_profile("ib-report"), None)
    assert style.TABLE_ZEBRA
    body = "| 구분 | 금액 | << |\n|---|---|---|\n| ^^ | 1차 | 2차 |\n| 가 | 10 | 11 |\n| 나 | 12 | 13 |\n"
    table = body_table(render(markdown(body, "ib-report", tables=[{"header_rows": 2, "spans": True}]), strict=False), "1차")
    first, second = table.xpath("./w:tr")[2:4]
    assert attr(cells(first)[2], "./w:tcPr/w:shd", "w:fill") == [style.LIGHT_GRAY_HEX]
    assert attr(cells(second)[2], "./w:tcPr/w:shd", "w:fill") != [style.LIGHT_GRAY_HEX]


def test_type_detection_reads_every_header_row_and_merged_labels() -> None:
    from document_model import TableType

    html = (
        "<table><tr><th rowspan=\"2\">Risk</th><th>Assessment</th></tr><tr><th>Probability</th></tr>"
        "<tr><td>Delay</td><td>High</td></tr></table>\n"
    )
    table = MarkdownParser(profile="ib-report").parse(html).elements[-1].content
    assert table.table_type == TableType.RISK_MATRIX and table.rows[2].cells[1].risk_level == "high"
    body = "| Risk | Impact | << |\n|---|---|---|\n| ^^ | A | B |\n| X | High | Low |\n"
    spec = [{"type": "risk", "header_rows": 2, "spans": True}]
    model = MarkdownParser().parse(markdown(body, "business-report", tables=spec))
    risk_row = model.elements[-1].content.rows[2]
    assert [cell.risk_level for cell in risk_row.cells] == [None, "high", "low"]


def test_dash_placeholders_do_not_make_an_amount_column_text() -> None:
    body = "| 회차 | 상환액 |\n|---|---|\n| 1 | - |\n| 2 | – |\n| 3 | - |\n| 4 | 14 |\n"
    table = body_table(render(markdown(body, "business-report")), "상환액")
    amounts = [cells(row)[1] for row in table.xpath("./w:tr")[1:]]
    assert {tuple(attr(tc.xpath("./w:p")[0], "./w:pPr/w:jc")) for tc in amounts} == {("right",)}
