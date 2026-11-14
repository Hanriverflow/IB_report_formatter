"""Explicit IB table rules: totals and subtotal rows (synthetic data only).

Rows are named in the table specification as one-based body rows and are never
inferred from their text. Every test parses Markdown with the real parser and
renders through `IBDocumentRenderer.render`.
"""

from io import BytesIO

import pytest
import yaml
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser
from ooxml_order import SEQUENCES

PROFILES = [
    "ib-report",
    "ib-memo",
    "plain",
    "office-letter",
    "business-report",
    "meeting-minutes",
    "term-sheet",
]
# The rule colour is each profile's table accent (`STYLE.NAVY_HEX`).
ACCENT = {"ib-report": "003366", "ib-memo": "003366", "term-sheet": "1A2270"}

STATEMENT = (
    "| 구분 | 전기 | 당기 |\n|---|---|---|\n"
    "| 매출액 | 1200 | 1350 |\n"
    "| 매출총이익 | 300 | 340 |\n"
    "| 판매비 | (120) | (130) |\n"
    "| 영업이익 | 180 | 210 |\n"
    "| 합계 | 180 | 210 |\n"
)


def markdown(body: str, profile: str, tables=None) -> str:
    data: dict = {"profile": profile, "title": "가상 검토 문서"}
    if profile == "term-sheet":
        data.update(prepared_by="가상은행 자본시장부", disclaimer="가상 조건 검토용 문서입니다.")
    if profile == "office-letter":
        data.update(sender={"organization": "가상회사"}, recipients=["가상 수신부서"])
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def render(content: str, strict: bool = True):
    doc = IBDocumentRenderer(options=RenderOptions(strict=strict)).render(
        MarkdownParser().parse(content)
    )
    buffer = BytesIO()
    doc.save(buffer)
    return Document(buffer)


def body_table(doc, marker: str):
    return next(table for table in doc.tables if marker in table._tbl.xml)._tbl


def rows(table) -> list:
    return table.xpath("./w:tr")


def cells(row) -> list:
    return row.xpath("./w:tc")


def edge(tc, name: str) -> tuple[str, ...] | None:
    """(val, sz, color) of one tcBorders edge, or None."""
    found = tc.xpath(f"./w:tcPr/w:tcBorders/w:{name}")
    if not found:
        return None
    return tuple(found[0].get(qn(f"w:{key}")) for key in ("val", "sz", "color"))


def style(tc, name: str) -> str | None:
    """Border style of one tcBorders edge, or None."""
    found = edge(tc, name)
    return found[0] if found else None


def bold_runs(tc) -> list[bool]:
    """Effective direct bold of each text run (`<w:b w:val="0"/>` is not bold)."""
    return [
        any(b.get(qn("w:val"), "1") not in ("0", "false") for b in run.xpath("./w:rPr/w:b"))
        for run in tc.xpath(".//w:r[w:t]")
    ]


def order_violations(root) -> list:
    problems = []
    for kind in ("tcPr", "tcBorders"):
        sequence = SEQUENCES[kind]
        for element in root.iter(qn(f"w:{kind}")):
            names = [etree.QName(c).localname for c in element if isinstance(c.tag, str)]
            ranks = [sequence.index(name) for name in names if name in sequence]
            if ranks != sorted(ranks) or len(set(names)) != len(names):
                problems.append((kind, names))
    return problems


# ═══════════════════════════════════════════════════════════════════════════════
# NAMED ROWS ONLY
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("profile", PROFILES)
def test_only_named_rows_are_ruled_and_bold(profile: str) -> None:
    spec = {"subtotal_rows": [2], "total_rows": [4]}
    table = body_table(render(markdown(STATEMENT, profile, tables=[spec])), "매출총이익")
    color = ACCENT.get(profile, "202020")
    single, double = ("single", "6", color), ("double", "4", color)
    expected = {  # Word row: (top, bottom, bold)
        0: (None, None, None),
        1: (None, single, False),
        2: (single, None, True),
        3: (None, single, False),
        4: (single, double, True),
        5: (None, None, False),
    }
    for index, row in enumerate(rows(table)):
        top, bottom, bold = expected[index]
        for tc in cells(row):
            assert (edge(tc, "top"), edge(tc, "bottom")) == (top, bottom), index
            assert not tc.xpath("./w:tcPr/w:tcBorders/*[not(self::w:top or self::w:bottom)]")
            if bold is not None:
                assert set(bold_runs(tc)) == {bold}, index
    assert not order_violations(table)


@pytest.mark.parametrize("profile", PROFILES)
def test_a_row_reading_total_is_not_inferred(profile: str) -> None:
    # No list, no rule: the `합계` row stays an ordinary row in every profile.
    table = body_table(render(markdown(STATEMENT, profile, tables=[{}])), "매출총이익")
    assert not table.xpath(".//w:tcBorders")
    total = next(
        tc for tc in table.xpath(".//w:tc") if "합계" in "".join(tc.xpath(".//w:t/text()"))
    )
    assert set(bold_runs(total)) == {False}


@pytest.mark.parametrize("profile", PROFILES)
def test_unnamed_tables_and_rows_keep_their_xml(profile: str) -> None:
    other = "| 별도항목 | 금액 |\n|---|---|\n| 가 | 10 |\n| 나 | 20 |\n"
    body = other + "\n" + STATEMENT
    marked = render(markdown(body, profile, tables=[{}, {"total_rows": [4]}]))
    plain = render(markdown(body, profile, tables=[{}, {}]))
    empty = render(markdown(body, profile, tables=[{}, {"total_rows": [], "subtotal_rows": []}]))
    # The table without lists is identical; so is a whole document whose lists are empty.
    first = [etree.tostring(body_table(doc, "별도항목")) for doc in (marked, plain)]
    assert first[0] == first[1]
    assert empty.element.xml == plain.element.xml
    # In the named table only the named row and the row above it change.
    pairs = zip(rows(body_table(marked, "판매비")), rows(body_table(plain, "판매비")), strict=True)
    changed = [
        index for index, (a, b) in enumerate(pairs) if etree.tostring(a) != etree.tostring(b)
    ]
    assert changed == [3, 4]


def test_schedule_total_does_not_draw_rules() -> None:
    body = (
        "| 회차 | 원금상환액 | 잔액 |\n|---|---|---|\n"
        "| 1 | 50 | 50 |\n| 2 | 50 | 0 |\n| 합계 | 100 | 0 |\n"
    )
    spec = {
        "columns": ["text", "number", "number"],
        "unit": "억원",
        "schedule": {"repayment": 2, "balance": 3, "principal": "100억원", "total": True},
    }
    table = body_table(render(markdown(body, "term-sheet", tables=[spec])), "원금상환액")
    assert not table.xpath(".//w:tcBorders")


# ═══════════════════════════════════════════════════════════════════════════════
# HEADER ROWS AND MERGED CELLS
# ═══════════════════════════════════════════════════════════════════════════════

TWO_HEADERS = (
    "| 구분 | 실적 | << | 추정 |\n|---|---|---|---|\n"
    "| ^^ | FY24A | FY25A | FY26E |\n"
    "| 매출액 | 100 | 110 | 120 |\n"
    "| 영업이익 | 10 | 12 | 14 |\n"
    "| 당기순이익 | 7 | 8 | 9 |\n"
)


@pytest.mark.parametrize("profile", ["ib-memo", "business-report", "term-sheet"])
def test_rows_are_counted_from_the_first_body_row(profile: str) -> None:
    spec = {"header_rows": 2, "spans": True, "total_rows": [3]}
    table = body_table(render(markdown(TWO_HEADERS, profile, tables=[spec])), "FY24A")
    word_rows = rows(table)
    assert [bool(row.xpath(".//w:tcBorders")) for row in word_rows] == [
        False,
        False,
        False,
        True,  # bottom rule of the row above
        True,
    ]
    assert style(cells(word_rows[4])[0], "bottom") == "double"
    assert "당기순이익" in "".join(word_rows[4].xpath(".//w:t/text()"))


def test_a_merged_totals_label_is_ruled_across_its_span() -> None:
    body = (
        "| 회차 | 지급일 | 원금 | 잔액 |\n|---|---|---|---|\n"
        "| 1 | 2027-01-15 | 50 | 50 |\n"
        "| 2 | 2027-04-15 | 50 | 0 |\n"
        "| 합 계 | << | 100 | – |\n"
    )
    table = body_table(render(markdown(body, "term-sheet", tables=[{"total_rows": [3]}])), "합 계")
    total = cells(rows(table)[3])
    assert total[0].xpath("./w:tcPr/w:gridSpan/@w:val") == ["2"]
    assert [style(tc, "top") for tc in total] == ["single"] * 3
    assert [style(tc, "bottom") for tc in total] == ["double"] * 3
    assert set(bold_runs(total[0])) == {True}
    # The label shading of the merged label is kept.
    assert total[0].xpath("./w:tcPr/w:shd/@w:fill")
    assert not order_violations(table)


def test_a_vertically_merged_cell_crossing_a_named_row_keeps_its_merge() -> None:
    # The `^^` cell of the named row is a covered `w:tc`; it is ruled like every
    # other cell, and Word draws no rule inside the merged area (Word page check
    # in docs/verification-ib-table-rules-20261114.md).
    body = (
        "| 구분 | 항목 | 금액 |\n|---|---|---|\n"
        "| 영업 | 매출 | 100 |\n"
        "| ^^ | 원가 | (60) |\n"
        "| ^^ | 영업 소계 | 40 |\n"
    )
    spec = {"spans": True, "subtotal_rows": [3]}
    table = body_table(render(markdown(body, "ib-memo", tables=[spec])), "영업 소계")
    subtotal = cells(rows(table)[3])
    assert subtotal[0].xpath("./w:tcPr/w:vMerge") and not subtotal[0].xpath(".//w:t")
    assert [style(tc, "top") for tc in subtotal] == ["single"] * 3
    assert set(bold_runs(subtotal[1])) == {True}


# ═══════════════════════════════════════════════════════════════════════════════
# REJECTED SPECIFICATIONS
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        ({"total_rows": 3}, "list of body row numbers from 1 to 5"),
        ({"total_rows": ["3"]}, "list of body row numbers"),
        ({"total_rows": [True]}, "list of body row numbers"),
        ({"total_rows": [2.0]}, "list of body row numbers"),
        ({"subtotal_rows": [0]}, "list of body row numbers"),
        ({"subtotal_rows": [6]}, "list of body row numbers from 1 to 5"),
        ({"total_rows": [2, 2]}, "lists body row 2 twice"),
        ({"total_rows": [4], "subtotal_rows": [2, 4]}, "row 4 is in both"),
        ({"header_rows": 2, "total_rows": [5]}, "from 1 to 4"),
    ],
)
@pytest.mark.parametrize("strict", [True, False])
def test_invalid_row_lists_are_rejected_in_both_modes(spec, message, strict: bool) -> None:
    content = markdown(STATEMENT, "ib-memo", tables=[spec])
    with pytest.raises(ValueError, match=message):
        render(content, strict=strict)


# ═══════════════════════════════════════════════════════════════════════════════
# NEGATIVE-NUMBER STYLE (table spec `negative`)
# ═══════════════════════════════════════════════════════════════════════════════

# Negative amounts written three ways, a percent role, a unit suffix, a bold
# amount, positive and token cells, and a text column with arrows and △.
MIXED = (
    "| 구분 | 전기 | 당기 | 증감률 | 증감 |\n|---|---|---|---|---|\n"
    "| 금융손익 | (2100) | -2400 | -3.5 | ▲12800 |\n"
    "| 법인세 | △2600 | **-1800** | (1.2) | ▼900 |\n"
    "| 순이익 | 7100 | 1200억원 | NM | △500 |\n"
)
MIXED_SPEC = {"type": "financial", "columns": ["text", "money", "money", "percent", "text"]}
RED = "C00000"


def cell_texts(table) -> list[list[str]]:
    return [["".join(tc.xpath(".//w:t/text()")) for tc in cells(row)] for row in rows(table)]


def colors(tc) -> set[str]:
    return set(tc.xpath(".//w:r/w:rPr/w:color/@w:val"))


def digits(texts: list[list[str]]) -> list[str]:
    return ["".join(ch for ch in "".join(row) if ch.isdigit()) for row in texts]


@pytest.mark.parametrize("profile", ["ib-memo", "business-report", "term-sheet"])
def test_without_the_option_negatives_keep_their_written_form_and_red(profile: str) -> None:
    table = body_table(render(markdown(MIXED, profile, tables=[MIXED_SPEC])), "금융손익")
    assert cell_texts(table)[1:] == [
        ["금융손익", "(2,100)", "-2,400", "-3.5%", "▲12800"],
        ["법인세", "△2600", "-1,800", "(1.2)%", "▼900"],
        ["순이익", "7,100", "1,200억원", "NM", "△500"],
    ]
    first, second = cells(rows(table)[1]), cells(rows(table)[2])
    assert [RED in colors(tc) for tc in first[1:4]] == [True, True, True]
    # A written △ is not read as negative, and the text column is never red.
    assert RED not in colors(second[1]) and RED in colors(second[2])
    assert not any(RED in colors(tc) for row in rows(table) for tc in cells(row)[4:])


EXPECTED = {
    "paren": [
        ["금융손익", "(2,100)", "(2,400)", "(3.5%)", "▲12800"],
        ["법인세", "(2,600)", "(1,800)", "(1.2%)", "▼900"],
    ],
    "minus": [
        ["금융손익", "-2,100", "-2,400", "-3.5%", "▲12800"],
        ["법인세", "-2,600", "-1,800", "-1.2%", "▼900"],
    ],
    "triangle": [
        ["금융손익", "△2,100", "△2,400", "△3.5%", "▲12800"],
        ["법인세", "△2,600", "△1,800", "△1.2%", "▼900"],
    ],
}


@pytest.mark.parametrize("profile", ["ib-memo", "business-report", "term-sheet"])
@pytest.mark.parametrize("style", ["paren", "minus", "triangle"])
def test_the_option_restyles_only_negative_amounts_in_black(profile: str, style: str) -> None:
    spec = {**MIXED_SPEC, "negative": style}
    table = body_table(render(markdown(MIXED, profile, tables=[spec])), "금융손익")
    texts = cell_texts(table)
    assert texts[1:3] == EXPECTED[style]
    # Positive amounts, tokens and the text column (arrows, △500) are unchanged.
    assert texts[3] == ["순이익", "7,100", "1,200억원", "NM", "△500"]
    # The amounts are the same digits as without the option.
    plain = cell_texts(
        body_table(render(markdown(MIXED, profile, tables=[MIXED_SPEC])), "금융손익")
    )
    assert digits(texts) == digits(plain)
    assert not any(RED in colors(tc) for row in rows(table) for tc in cells(row))
    # Emphasis written in Markdown is kept.
    assert set(bold_runs(cells(rows(table)[2])[2])) == {True}


def test_the_option_leaves_other_tables_of_the_document_unchanged() -> None:
    body = MIXED + "\n" + MIXED.replace("금융손익", "영업외손익")
    spec = {**MIXED_SPEC, "negative": "paren"}
    with_option = render(markdown(body, "ib-memo", tables=[spec, MIXED_SPEC]))
    without = render(markdown(body, "ib-memo", tables=[MIXED_SPEC, MIXED_SPEC]))
    assert etree.tostring(body_table(with_option, "영업외손익")) == etree.tostring(
        body_table(without, "영업외손익")
    )


def test_a_suffix_after_parentheses_moves_inside_and_a_unit_amount_is_restyled() -> None:
    body = "| 구분 | 금액 |\n|---|---|\n| 가 | (1200)억원 |\n| 나 | -300억원 |\n"
    spec = {"type": "financial", "columns": ["text", "money"], "negative": "paren"}
    table = body_table(render(markdown(body, "ib-memo", tables=[spec])), "억원")
    assert [row[1] for row in cell_texts(table)[1:]] == ["(1,200억원)", "(300억원)"]


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        ({"type": "financial", "negative": "red"}, "paren, minus or triangle"),
        ({"type": "financial", "negative": "Paren"}, "paren, minus or triangle"),
        ({"type": "financial", "negative": True}, "paren, minus or triangle"),
        ({"type": "financial", "negative": ["paren"]}, "paren, minus or triangle"),
        ({"negative": "paren"}, "requires type: financial"),
        ({"type": "generic", "negative": "paren"}, "requires type: financial"),
        ({"type": "sensitivity", "negative": "minus"}, "requires type: financial"),
    ],
)
@pytest.mark.parametrize("strict", [True, False])
def test_invalid_negative_settings_are_rejected_in_both_modes(spec, message, strict) -> None:
    # The ib-memo header row reads as financial, but the option needs the spec to say so.
    with pytest.raises(ValueError, match=message):
        render(markdown(MIXED, "ib-memo", tables=[spec]), strict=strict)
