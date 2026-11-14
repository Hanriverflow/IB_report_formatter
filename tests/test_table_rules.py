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


# ═══════════════════════════════════════════════════════════════════════════════
# NEGATIVE STYLE: ONE AMOUNT PER CELL (independent review, 2026-11-14)
# ═══════════════════════════════════════════════════════════════════════════════

STYLES = ["paren", "minus", "triangle"]
THREE = ["ib-memo", "business-report", "term-sheet"]

# (row label, column role, written forms, magnitude as shown). Every written form of
# one row is the same negative amount, so every form must give the same text.
FORMS = [
    ("금액형", "money", ["-1200", "−1200", "(1200)", "△1200", "-1,200", "(1,200)"], "1,200"),
    (
        "억원형",
        "money",
        ["-1200억원", "−1200억원", "(1200억원)", "(1200)억원", "△1200억원"],
        "1,200억원",
    ),
    (
        "비율형",
        "percent",
        ["-3.5", "−3.5", "(3.5)", "△3.5", "-3.5%", "(3.5%)", "(3.5)%", "△3.5%"],
        "3.5%",
    ),
    (
        "베이시스형",
        "bps",
        ["-120", "−120", "(120)", "△120", "-120 bps", "(120 bps)", "(120) bps", "△120 bps"],
        "120 bps",
    ),
    (
        "배수형",
        "multiple",
        ["-1.5", "−1.5", "(1.5)", "△1.5", "-1.5x", "(1.5x)", "(1.5)x", "△1.5x"],
        "1.5x",
    ),
]


def shown(style: str, magnitude: str) -> str:
    return {"paren": f"({magnitude})", "minus": f"-{magnitude}", "triangle": f"△{magnitude}"}[style]


def one_column_table(label: str, values: list[str]) -> str:
    rows_md = "".join(f"| {label}{index} | {value} |\n" for index, value in enumerate(values))
    return f"| 구분 | 값 |\n|---|---|\n{rows_md}\n"


def forms_document(profile: str, style: str | None, inputs=None) -> str:
    body, specs = "", []
    for (label, role, _, _), values in zip(FORMS, inputs or [f[2] for f in FORMS], strict=True):
        body += one_column_table(label, values)
        spec = {"type": "financial", "columns": ["text", role]}
        if style:
            spec["negative"] = style
        specs.append(spec)
    return markdown(body, profile, tables=specs)


def render_noting(content: str, strict: bool = True):
    """Render and also return the renderer, whose `warnings` never reject a document."""
    renderer = IBDocumentRenderer(options=RenderOptions(strict=strict))
    doc = renderer.render(MarkdownParser().parse(content))
    buffer = BytesIO()
    doc.save(buffer)
    return Document(buffer), renderer


def value_column(doc, label: str) -> list[str]:
    return [row[1] for row in cell_texts(body_table(doc, label + "0"))[1:]]


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_every_written_form_of_one_amount_gives_the_same_text(profile: str, style: str) -> None:
    # Findings [3] and [4]: a unit inside the parentheses and a △ amount are read like
    # `-…`; a bare number in a percent, bps or multiple column gets the column's unit.
    doc = render(forms_document(profile, style))
    for label, _, written, magnitude in FORMS:
        assert value_column(doc, label) == [shown(style, magnitude)] * len(written), label


@pytest.mark.parametrize("source", STYLES)
@pytest.mark.parametrize("target", STYLES)
def test_each_style_reads_back_the_text_of_every_other_style(source: str, target: str) -> None:
    first = render(forms_document("ib-memo", source))
    written_back = [value_column(first, label) for label, *_ in FORMS]
    again = render(forms_document("ib-memo", target, inputs=written_back))
    for label, _, written, magnitude in FORMS:
        assert value_column(again, label) == [shown(target, magnitude)] * len(written), label


def test_without_the_option_a_triangle_is_still_not_read_as_negative() -> None:
    doc = render(forms_document("ib-memo", None))
    assert value_column(doc, "비율형")[3] == "△3.5"  # no % added, unlike -3.5
    assert value_column(doc, "비율형")[0] == "-3.5%"
    assert value_column(doc, "금액형")[3] == "△1200"
    table = body_table(doc, "금액형0")
    assert RED not in colors(cells(rows(table)[4])[1]) and RED in colors(cells(rows(table)[1])[1])


UNCHANGED = [
    "1-2",
    "2026-11-14",
    "NM",
    "-",
    "▲1200",
    "▼1200",
    "-0",
    "(0)",
    "−0.0",
    "△0",
    "-1200 / -300",
    "약 -1200",
    "+1200",
    "(3.5%)%",
    "-1,2",
    "`-1200`",
]


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_a_cell_that_is_not_exactly_one_negative_amount_is_as_without_the_option(
    profile: str, style: str
) -> None:
    body = one_column_table("불변", UNCHANGED) + (
        "| 문자 | 코드 | 일자 | 금액 |\n|---|---|---|---|\n"
        "| -1200 | (1200) | △1200 | 1 |\n| (1200) | △1200 | -1200 | 2 |\n"
        "| △1200 | -1200 | (1200) | 3 |\n"
    )
    columns = [
        {"type": "financial", "columns": ["text", "money"]},
        {"type": "financial", "columns": ["text", "code", "date", "money"]},
    ]
    with_option = render(
        markdown(body, profile, tables=[{**spec, "negative": style} for spec in columns])
    )
    without = render(markdown(body, profile, tables=columns))
    for marker in ("불변0", "코드"):
        assert cell_texts(body_table(with_option, marker)) == cell_texts(
            body_table(without, marker)
        ), marker
    assert value_column(with_option, "불변") == [
        "1-2",
        "2026-11-14",
        "NM",
        "-",
        "▲1200",
        "▼1200",
        "-0",
        "(0)",
        "−0.0",
        "△0",
        "-1200 / -300",
        "약 -1200",
        "+1200",
        "(3.5%)%",
        "-1,2",
        "-1200",
    ]


SPLIT = [
    # (cell, text as without the option, warns)
    ("-1200**.50**", "-1,200.50", True),
    ("-**1200**", "-1,200", True),
    ("1**-2**", "1-2", False),
]


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_an_amount_split_across_formatting_keeps_its_text_and_warns(
    profile: str, style: str
) -> None:
    # Finding [1]: never rewrite one formatted piece of an amount. A split amount keeps
    # its written form (as without the option) and leaves a warning that strict mode
    # does not reject; a range such as 1-2 is not an amount and gives no warning.
    body = one_column_table("분할", [cell for cell, _, _ in SPLIT])
    spec = {"type": "financial", "columns": ["text", "money"]}
    doc, renderer = render_noting(markdown(body, profile, tables=[{**spec, "negative": style}]))
    without = render(markdown(body, profile, tables=[spec]))
    assert value_column(doc, "분할") == value_column(without, "분할")
    assert value_column(doc, "분할") == [text for _, text, _ in SPLIT]
    table = body_table(doc, "분할0")
    assert not any(RED in colors(tc) for row in rows(table) for tc in cells(row))
    assert len(renderer.warnings) == 2
    assert all("split across differently formatted text" in w for w in renderer.warnings)
    assert "'-1200.50'" in renderer.warnings[0] and "'-1200'" in renderer.warnings[1]
    assert not any("split" in error for error in renderer.errors)
    # Formatting is kept: the bold piece stays bold, the rest does not.
    assert bold_runs(cells(rows(table)[1])[1]) == [False, True]


def test_a_split_amount_does_not_warn_in_a_table_without_the_option() -> None:
    body = one_column_table("분할", [cell for cell, _, _ in SPLIT])
    spec = {"type": "financial", "columns": ["text", "money"]}
    _, renderer = render_noting(markdown(body, "ib-memo", tables=[spec]))
    assert renderer.warnings == []


@pytest.mark.parametrize("style", STYLES)
def test_a_repayment_schedule_reads_the_shown_amounts_as_the_same_numbers(style: str) -> None:
    from numeric_checks import read_cell_number

    body = (
        "| 회차 | 원금상환액 | 잔액 | 이자 |\n|---|---|---|---|\n"
        "| 1 | 50 | 50 | -20**.50** |\n| 2 | 50 | 0 | (30) |\n"
        "| 합계 | 100 | 0 | △40 |\n"
    )
    spec = {
        "type": "financial",
        "columns": ["text", "number", "number", "number"],
        "unit": "억원",
        "schedule": {"repayment": 2, "balance": 3, "principal": "100억원", "total": True},
    }
    with_option = cell_texts(
        body_table(
            render(markdown(body, "term-sheet", tables=[{**spec, "negative": style}])), "이자"
        )
    )
    without = cell_texts(body_table(render(markdown(body, "term-sheet", tables=[spec])), "이자"))
    assert [row[3] for row in with_option[1:]] == ["-20.50", shown(style, "30"), shown(style, "40")]
    for shown_row, plain_row in zip(with_option[1:], without[1:], strict=True):
        for shown_text, plain_text in zip(shown_row[1:], plain_row[1:], strict=True):
            assert read_cell_number(shown_text) == read_cell_number(plain_text)


COLORED = (
    "| 구분 | 값 |\n|---|---|\n"
    '| 색0 | <span style="color:#C00000">-1200</span> |\n'
    "| 색1 | [-1300](https://example.invalid) |\n"
    "| 색2 | -1400 |\n"
)


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_the_option_turns_off_only_the_automatic_red(profile: str, style: str) -> None:
    # Finding [2], kept by decision: a colour the author wrote and a link's colour stay;
    # only the engine's automatic red is turned off.
    spec = {"type": "financial", "columns": ["text", "money"]}
    doc = render(markdown(COLORED, profile, tables=[{**spec, "negative": style}]))
    without = render(markdown(COLORED, profile, tables=[spec]))
    table, plain = body_table(doc, "색0"), body_table(without, "색0")
    assert value_column(doc, "색") == [shown(style, n) for n in ("1,200", "1,300", "1,400")]
    span, link, ordinary = (
        cells(rows(table)[1])[1],
        cells(rows(table)[2])[1],
        cells(rows(table)[3])[1],
    )
    assert RED in colors(span)
    assert colors(link) == colors(cells(rows(plain)[2])[1]) and colors(link)
    assert table.xpath(".//w:hyperlink") and RED not in colors(link)
    assert colors(ordinary) <= {"auto"}
    assert RED in colors(cells(rows(plain)[3])[1])


def layout(table) -> tuple:
    grid = tuple(table.xpath("./w:tblGrid/w:gridCol/@w:w"))
    widths = tuple(tuple(row.xpath("./w:tc/w:tcPr/w:tcW/@w:w")) for row in rows(table))
    aligns = tuple(tuple(row.xpath("./w:tc/w:p[1]/w:pPr/w:jc/@w:val")) for row in rows(table))
    return grid, widths, aligns


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_the_option_keeps_alignment_and_widths(profile: str, style: str) -> None:
    # Finding [5]: alignment, column kinds and widths are judged on the written text.
    body = (
        "| 구분 | 금액 | 비율 |\n|---|---|---|\n"
        "| 정렬가 | -1200 | -3.5% |\n| 정렬나 | -1300 | -4.2% |\n"
    )
    doc = render(markdown(body, profile, tables=[{"type": "financial", "negative": style}]))
    without = render(markdown(body, profile, tables=[{"type": "financial"}]))
    table = body_table(doc, "정렬가")
    assert layout(table) == layout(body_table(without, "정렬가"))
    assert [row[1:] for row in layout(table)[2][1:]] == [("right", "right")] * 2
    assert [row[1:] for row in cell_texts(table)[1:]] == [
        [shown(style, "1,200"), shown(style, "3.5%")],
        [shown(style, "1,300"), shown(style, "4.2%")],
    ]


@pytest.mark.parametrize("style", STYLES)
def test_a_footnote_marker_after_the_amount_stays_after_it(style: str) -> None:
    body = "| 구분 | 값 |\n|---|---|\n| 각주0 | -1200[^1] |\n\n[^1]: 가상 주석입니다.\n"
    spec = {"type": "financial", "columns": ["text", "money"], "negative": style}
    table = body_table(render(markdown(body, "ib-memo", tables=[spec])), "각주0")
    cell = cells(rows(table)[1])[1]
    assert "".join(cell.xpath(".//w:r[not(w:footnoteReference)]/w:t/text()")) == shown(
        style, "1,200"
    )
    assert cell.xpath(".//w:r[last()]/w:footnoteReference")


# ═══════════════════════════════════════════════════════════════════════════════
# NEGATIVE STYLE: SECOND INDEPENDENT REVIEW (2026-11-14)
# ═══════════════════════════════════════════════════════════════════════════════


def option_and_plain(values: list[str], style: str, profile: str = "ib-memo"):
    """Value column of one money table with and without `negative`, and the warnings."""
    body = one_column_table("검토", values)
    spec = {"type": "financial", "columns": ["text", "money"]}
    doc, renderer = render_noting(markdown(body, profile, tables=[{**spec, "negative": style}]))
    without = render(markdown(body, profile, tables=[spec]))
    return value_column(doc, "검토"), value_column(without, "검토"), renderer


# Finding [1]: an integer part with a leading zero is never restyled, so no digit
# is dropped by the option; the cell shows what a table without the option shows.
LEADING_ZERO = ["(0001.20%)", "(0001200억원)", "−0001200"]


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_a_leading_zero_amount_shows_the_same_text_as_without_the_option(
    profile: str, style: str
) -> None:
    shown_texts, plain, _ = option_and_plain(LEADING_ZERO, style, profile)
    assert shown_texts == plain
    assert shown_texts == LEADING_ZERO


@pytest.mark.parametrize("style", STYLES)
def test_other_leading_zero_forms_follow_the_table_without_the_option(style: str) -> None:
    values = ["-0001200", "(0001.20)%", "-01,200", "-00.5", "△0012"]
    shown_texts, plain, _ = option_and_plain(values, style)
    assert shown_texts == plain


@pytest.mark.parametrize("style", STYLES)
def test_plain_amounts_keep_every_digit_of_the_fraction_and_the_unit(style: str) -> None:
    values = ["-1,200.50", "-1.20%", "(1200.50억원)", "-0.50", "−3.10백만원", "△12.0bp"]
    shown_texts, _, _ = option_and_plain(values, style)
    assert shown_texts == [
        shown(style, "1,200.50"),
        shown(style, "1.20%"),
        shown(style, "1,200.50억원"),
        shown(style, "0.50"),
        shown(style, "3.10백만원"),
        shown(style, "12.0bp"),
    ]


SCRIPT_SPLIT = ["-1200^.50^", "-12~00~", "^-1200^"]


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_an_amount_split_by_a_superscript_or_subscript_keeps_its_text_and_warns(
    profile: str, style: str
) -> None:
    # Finding [2]: a superscript or subscript piece splits an amount like bold does.
    # An amount written wholly in superscript is not split: no warning, written form.
    shown_texts, plain, renderer = option_and_plain(SCRIPT_SPLIT, style, profile)
    assert shown_texts == plain
    assert shown_texts == ["-1,200.50", "-1200", "-1200"]
    assert renderer.warnings == [
        f"Table negative style not applied to {text!r}: the amount is split across "
        "differently formatted text, so it keeps its written form"
        for text in ("-1200.50", "-1200")
    ]
    assert not any("split" in error for error in renderer.errors)


def test_a_footnote_marker_and_a_code_span_stay_protected_without_a_warning() -> None:
    body = "| 구분 | 값 |\n|---|---|\n| 검토0 | -1200[^1] |\n| 검토1 | `-1200` |\n\n[^1]: 가상.\n"
    spec = {"type": "financial", "columns": ["text", "money"], "negative": "paren"}
    doc, renderer = render_noting(markdown(body, "ib-memo", tables=[spec]))
    assert value_column(doc, "검토") == ["(1,200)", "-1200"]
    assert renderer.warnings == []


@pytest.mark.parametrize("style", STYLES)
def test_full_width_digits_are_read_as_the_formatting_without_the_option_reads_them(
    style: str,
) -> None:
    # Finding [3]: the zero test uses the same digits the amount pattern accepts.
    values = ["-１２３４", "(１２３４)", "-１２３４1", "-０", "(０.０)", "-０１２３"]
    shown_texts, plain, _ = option_and_plain(values, style)
    assert plain[:2] == ["-1,234", "(1,234)"]
    assert shown_texts[:3] == [shown(style, "1,234"), shown(style, "1,234"), shown(style, "12,341")]
    # Negative zero and a leading zero keep the text of a table without the option.
    assert shown_texts[3:] == plain[3:]


def test_a_render_rejected_before_composition_does_not_keep_earlier_warnings() -> None:
    # Finding [4]: success, then a strict pre-validation failure, then success.
    from input_safety import StrictValidationError

    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    spec = {"type": "financial", "columns": ["text", "money"], "negative": "paren"}
    split = MarkdownParser().parse(
        markdown(one_column_table("검토", ["-1200**.50**"]), "ib-memo", tables=[spec])
    )
    renderer.render(split)
    assert len(renderer.warnings) == 1
    rejected = MarkdownParser(profile="plain").parse("가상 본문")
    rejected.warnings = ["가상 검증 실패"]
    with pytest.raises(StrictValidationError):
        renderer.render(rejected)
    assert renderer.warnings == []
    clean = MarkdownParser().parse(
        markdown(one_column_table("검토", ["-1200"]), "ib-memo", tables=[spec])
    )
    renderer.render(clean)
    assert renderer.warnings == []
    renderer.render(split)
    assert len(renderer.warnings) == 1


AUTHOR_BLUE = (
    "| 구분 | 값 |\n|---|---|\n"
    '| 청색0 | <span style="color:#0000FF">-1200</span> |\n'
    "| 청색1 | -1300 |\n"
)


@pytest.mark.parametrize("profile", THREE)
@pytest.mark.parametrize("style", STYLES)
def test_an_author_colour_other_than_red_stays_only_in_a_table_with_the_option(
    profile: str, style: str
) -> None:
    # Finding [6]: without the option the automatic red covers the author's colour
    # (existing behaviour, unchanged); with it the author's colour is shown.
    spec = {"type": "financial", "columns": ["text", "money"]}
    table = body_table(
        render(markdown(AUTHOR_BLUE, profile, tables=[{**spec, "negative": style}])), "청색0"
    )
    plain = body_table(render(markdown(AUTHOR_BLUE, profile, tables=[spec])), "청색0")
    assert colors(cells(rows(table)[1])[1]) == {"0000FF"}
    assert colors(cells(rows(table)[2])[1]) <= {"auto"}
    assert colors(cells(rows(plain)[1])[1]) == {RED}
    assert RED in colors(cells(rows(plain)[2])[1])
