"""Numeric consistency checks: `checks:` relations and repayment schedules (fictional data)."""

from decimal import Decimal
from io import BytesIO

import pytest
import yaml
from docx import Document
from lxml import etree

from document_profiles import RenderOptions
from docx_audit import inspect_terms
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser
from numeric_checks import CHECKS_PROPERTY, Quantity, evaluate_checks, parse_checks, read_quantity

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


# ═══════════════════════════════════════════════════════════════════════════════
# READING DISPLAYED VALUES
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("text,value,dimension,step", [
    ("1,076,175,000원", "1076175000", "money", "1"),
    ("10억원", "1000000000", "money", "100000000"),
    ("3,330만원", "33300000", "money", "10000"),
    ("12.66억원", "1266000000", "money", "1000000"),
    ("5억 3,000만원", "530000000", "money", "10000"),
    ("약 12.66억원 수준", "1266000000", "money", "1000000"),
    ("300억", "30000000000", "money", "100000000"),
    ("[4.90]%", "4.90", "percent", "0.01"),
    ("[ 1.10 ]%p", "1.10", "percent", "0.01"),
    ("29bp", "29", "bp", "1"),
    ("36개월", "36", "months", "1"),
    ("[4.54]년", "4.54", "years", "0.01"),
    ("(12)", "-12", "plain", "1"),
    ("△3.5", "-3.5", "plain", "0.1"),
])
def test_read_quantity(text: str, value: str, dimension: str, step: str) -> None:
    assert read_quantity(text) == Quantity(Decimal(value), dimension, Decimal(step))


@pytest.mark.parametrize("text", ["CD + 1.28%", "가나다", "", "5 3억원", "3천"])
def test_unreadable_values_are_rejected(text: str) -> None:
    with pytest.raises(ValueError):
        read_quantity(text)


# ═══════════════════════════════════════════════════════════════════════════════
# RELATIONS BETWEEN TERM VALUES
# ═══════════════════════════════════════════════════════════════════════════════

FEES = {
    "upfront": "1,086,175,000원",
    "advisory": "10억원",
    "legal": "4,330만원",
    "audit": "3,300만원",
    "trust": "1,000만원",
    "stamp": "175,000원",
    "structuring": "약 12.76억원",
    "amount": "300억원",
    "facility": "315억원",
    "base": "[2.70]%",
    "spread": "[1.10]%p",
    "issue": "[3.80]%",
}


def check(source, values=FEES):
    return evaluate_checks(parse_checks([source]), values)


def test_a_fee_total_that_does_not_add_up_is_reported() -> None:
    assert check("upfront = advisory + legal + audit + trust + stamp") == [
        "Check failed: upfront = advisory + legal + audit + trust + stamp "
        "(left 1,086,175,000원, right 1,086,475,000원, difference -300,000원)"
    ]


@pytest.mark.parametrize("source", [
    "advisory - trust * 100 = 0 * amount",  # subtraction, and a plain factor on money
    "(legal + audit) * 2 - legal - audit = legal + audit",
    "structuring = upfront + 19 * trust",  # 12.76억원 is shown rounded: 12.7617억원 agrees
    "facility = amount * 1.05",
    "issue = base + spread",
    "amount / facility = 1 / 1.05",
])
def test_relations_that_hold_pass(source: str) -> None:
    assert check(source) == []


def test_explicit_tolerance_and_units() -> None:
    relation = {"check": "upfront = advisory + legal + audit + trust + stamp", "tolerance": "30만원"}
    assert evaluate_checks(parse_checks([relation]), FEES) == []
    assert check("amount = issue")[0].startswith("Check cannot be evaluated: amount = issue (the sides are money and percent)")
    assert check("missing = amount") == [
        "Check cannot be evaluated: missing = amount (undefined term {{missing}})"
    ]
    assert "cannot read {{bad}}" in check("bad = amount", {**FEES, "bad": "CD + 1%"})[0]


@pytest.mark.parametrize("raw", [
    "a = b", ["a == b"], ["a"], ["a = "], ["a = b ="], ["a = b $"],
    [{"check": "a = b", "extra": 1}], [{"tolerance": "1원"}], [{"check": "a = b", "tolerance": True}],
    [{"check": "a = b", "tolerance": "많이"}],
])
def test_malformed_checks_are_rejected(raw) -> None:
    with pytest.raises(ValueError, match="checks|cannot read"):
        parse_checks(raw)


def _markdown(terms, checks, body="검토 {{upfront}}.\n", tables=None, profile="business-report") -> str:
    data = {"profile": profile, "title": "가상 조건 검토", "terms": terms, "checks": checks}
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def _render(markdown: str, strict: bool = True):
    doc = IBDocumentRenderer(options=RenderOptions(strict=strict)).render(MarkdownParser().parse(markdown))
    payload = BytesIO()
    doc.save(payload)
    payload.seek(0)
    return Document(payload)


def test_failing_checks_are_model_warnings_that_strict_rejects() -> None:
    source = _markdown(FEES, ["upfront = advisory + legal + audit + trust + stamp", "facility = amount * 1.05"])
    model = MarkdownParser().parse(source)
    assert [warning for warning in model.warnings if warning.startswith("Check")] == [
        "Check failed: upfront = advisory + legal + audit + trust + stamp "
        "(left 1,086,175,000원, right 1,086,475,000원, difference -300,000원)"
    ]
    with pytest.raises(ValueError, match="Check failed"):
        _render(source)


def test_malformed_checks_fail_the_parse() -> None:
    with pytest.raises(ValueError, match="checks"):
        MarkdownParser().parse(_markdown(FEES, "upfront = legal"))


def test_docx_audit_rechecks_relations_after_editing_a_value() -> None:
    values = {**FEES, "upfront": "1,086,475,000원"}
    doc = _render(_markdown(values, ["upfront = advisory + legal + audit + trust + stamp"], "총 {{upfront}}, 법률 {{legal}}.\n"))
    assert inspect_terms(doc).failed_checks == []
    properties = etree.fromstring(doc.part.package.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
    ).blob)
    assert any(prop.get("name") == CHECKS_PROPERTY for prop in properties)
    # Edit one tagged value as a Word user would, then audit again.
    control = doc.element.body.xpath(".//w:sdt[w:sdtPr/w:tag/@w:val='ibrep:term:legal']")[0]
    control.xpath(".//w:t", namespaces=NS)[0].text = "4,360만원"
    failed = inspect_terms(doc).failed_checks
    assert failed and failed[0].startswith("Check failed: upfront = advisory + legal")


# ═══════════════════════════════════════════════════════════════════════════════
# REPAYMENT SCHEDULES
# ═══════════════════════════════════════════════════════════════════════════════

SCHEDULE_TERMS = {"amount": "300억원", "life": "[2.13]년"}
GOOD_ROWS = [
    ("0", "대출실행", "300"), ("3", "-", "300"), ("6", "-", "300"), ("9", "75", "225"),
    ("12", "75", "150"), ("15", "75", "75"), ("18", "75", "-"),
]


def _schedule(rows=GOOD_ROWS, total="300", spec=None, terms=None, unit="억원") -> str:
    body = "| 회차 | 경과(개월) | 상환액 | 잔액 |\n|---|---|---|---|\n"
    body += "".join(f"| {index} | {months} | {repaid} | {balance} |\n" for index, (months, repaid, balance) in enumerate(rows))
    body += f"| 합 계 | << | {total} | - |\n"
    schedule = {"repayment": 3, "balance": 4, "months": 2, "principal": "{{amount}}", "total": True, "average_life": "{{life}}"}
    schedule.update(spec or {})
    data = {
        "profile": "term-sheet", "title": "가상 조건", "prepared_by": "라마바은행 자본시장부",
        "disclaimer": "가상 조건 검토용입니다.", "terms": terms or SCHEDULE_TERMS,
        "tables": [{"schedule": schedule, **({"unit": unit} if unit else {})}],
    }
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n## 별첨\n\n" + body


def _schedule_warnings(markdown: str) -> list:
    return [warning for warning in MarkdownParser().parse(markdown).warnings if "schedule" in warning]


def test_a_consistent_schedule_passes_and_renders_strictly() -> None:
    # Weighted average life: 75 x (9 + 12 + 15 + 18) / 300 / 12 = 1.125 years.
    source = _schedule(terms={"amount": "300억원", "life": "1.13년"})
    assert _schedule_warnings(source) == []
    _render(source)


@pytest.mark.parametrize("rows,total,terms,expected", [
    ([*GOOD_ROWS[:4], ("12", "75", "160"), *GOOD_ROWS[5:]], "300", None,
     "Table 1 schedule: row 5: balance 160 should be 225 − 75 = 150"),
    ([*GOOD_ROWS[:-1], ("18", "70", "5")], "295", None,
     "Table 1 schedule: the last balance is 5, not 0"),
    (GOOD_ROWS, "290", None, "Table 1 schedule: the totals row shows 290, but the repayments add up to 300"),
    (GOOD_ROWS, "300", {"amount": "310억원", "life": "1.13년"},
     "Table 1 schedule: row 1: balance 300 should be 310 − 0 = 310"),
    (GOOD_ROWS, "300", {"amount": "300억원", "life": "1.25년"},
     "Table 1 schedule: the weighted average life is 1.125년, not the stated 1.25년"),
])
def test_schedule_disagreements_are_reported(rows, total, terms, expected) -> None:
    assert expected in _schedule_warnings(_schedule(rows, total, terms=terms))


def test_schedule_figures_that_cannot_be_read_are_reported() -> None:
    rows = [*GOOD_ROWS[:3], ("9", "7x5", "225"), *GOOD_ROWS[4:]]
    assert "Table 1 schedule: row 4: cannot read the repayment '7x5'" in _schedule_warnings(_schedule(rows))
    no_unit = _schedule(unit=None)
    assert "Table 1 schedule: a money principal needs the table `unit` (for example 억원)" in _schedule_warnings(no_unit)


@pytest.mark.parametrize("spec", [
    {"repayment": 9}, {"balance": 3}, {"total": "yes"}, {"months": None, "average_life": 2},
    {"principal": [1]}, {"extra": 1},
])
def test_invalid_schedule_specs_are_rejected(spec) -> None:
    source = _schedule(spec=spec)
    if spec == {"months": None, "average_life": 2}:
        source = source.replace("  months: 2\n", "")
    with pytest.raises(ValueError, match="schedule"):
        MarkdownParser().parse(source)
