"""Audit follow-up (2026-11-15), group C: schedule checks and chart arithmetic.

Every case goes through the real parser and renderer with fictional figures.
"""

import pytest
import yaml

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser


def _schedule_document(rows: list[tuple[str, ...]], schedule: dict) -> str:
    """A business report holding one schedule table (first column: months or period)."""
    data = {
        "profile": "business-report",
        "title": "가상 상환 일정",
        "tables": [{"schedule": schedule}],
    }
    body = "| 개월 | 상환액 | 잔액 |\n|---|---|---|\n" + "".join(
        "| " + " | ".join(row) + " |\n" for row in rows
    )
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def _schedule_warnings(markdown: str) -> list[str]:
    return [w for w in MarkdownParser().parse(markdown).warnings if "schedule" in w]


def _strict_render(markdown: str) -> None:
    IBDocumentRenderer(options=RenderOptions(strict=True)).render(MarkdownParser().parse(markdown))


# ═══════════════════════════════════════════════════════════════════════════════
# C1: A STATED AVERAGE LIFE WITH NO REPAYMENTS
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_stated_average_life_without_repayments_is_reported_and_rejected() -> None:
    schedule = {"months": 1, "repayment": 2, "balance": 3, "principal": 0, "average_life": "7년"}
    source = _schedule_document([("12", "0", "0")], schedule)
    message = (
        "Table 1 schedule: the weighted average life cannot be worked out: "
        "the repayments add up to 0"
    )
    assert _schedule_warnings(source) == [message]
    with pytest.raises(StrictValidationError, match="average life cannot be worked out"):
        _strict_render(source)


def test_a_schedule_without_repayments_and_without_average_life_stays_silent() -> None:
    schedule = {"months": 1, "repayment": 2, "balance": 3, "principal": 0}
    assert _schedule_warnings(_schedule_document([("12", "0", "0")], schedule)) == []


# ═══════════════════════════════════════════════════════════════════════════════
# C2: A BLANK LAST BALANCE CARRIES THE EXPECTED BALANCE
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize(
    "rows",
    [
        [("1", "50", "50"), ("2", "50", "")],
        [("1", "50", "50"), ("2", "25", ""), ("3", "25", "")],
        [("1", "50", "50"), ("2", "50", "0")],
    ],
)
def test_a_complete_schedule_with_blank_last_balances_passes(rows) -> None:
    source = _schedule_document(rows, {"repayment": 2, "balance": 3, "principal": 100})
    assert _schedule_warnings(source) == []
    _strict_render(source)


LAST = "Table 1 schedule: the last balance is {}, not 0"
SUM = "Table 1 schedule: repayments add up to {}, not the principal 100"


@pytest.mark.parametrize(
    ("rows", "messages"),
    [
        # Printed above zero: both facts, as before.
        ([("1", "50", "50"), ("2", "30", "20")], [LAST.format(20), SUM.format(80)]),
        # Carried above zero: the carried 20 is the principal less the
        # repayments, so the repayments message alone says it (HEAD added a
        # false "last balance is 50").
        ([("1", "50", "50"), ("2", "30", "")], [SUM.format(80)]),
        ([("1", "50", "50"), ("2", "", "")], [SUM.format(50)]),
    ],
)
def test_a_schedule_that_ends_above_zero_is_still_reported(rows, messages: list[str]) -> None:
    source = _schedule_document(rows, {"repayment": 2, "balance": 3, "principal": 100})
    assert _schedule_warnings(source) == messages
    with pytest.raises(StrictValidationError, match="schedule"):
        _strict_render(source)


@pytest.mark.parametrize(
    ("balance", "messages"),
    [
        # A blank balance on the unmarked totals row carries 0 - 100 = -100;
        # that only restates the sum, so one message, as at HEAD.
        ("", [SUM.format(200)]),
        # A dash balance is a stated 0: the row check speaks too, as at HEAD.
        ("-", ["Table 1 schedule: row 3: balance 0 should be 0 − 100 = -100", SUM.format(200)]),
    ],
)
def test_a_totals_row_not_marked_as_one_is_reported_once(balance: str, messages) -> None:
    rows = [("1", "50", "50"), ("2", "50", "0"), ("합계", "100", balance)]
    source = _schedule_document(rows, {"repayment": 2, "balance": 3, "principal": 100})
    assert _schedule_warnings(source) == messages
    with pytest.raises(StrictValidationError, match="repayments add up to 200"):
        _strict_render(source)


# ═══════════════════════════════════════════════════════════════════════════════
# C3: A WORKED-OUT TOTAL BEYOND THE FLOAT RANGE IS SCALED BEFORE IT IS PLOTTED
# ═══════════════════════════════════════════════════════════════════════════════

HUGE_STACK = """chart_type: stacked_bar
title: 가상 누적
labels: [가, 나]
totals: true
series:
  - name: 갑
    values: [1.0e+308, 1]
  - name: 을
    values: [1.0e+308, 2]
"""


def test_a_stack_total_beyond_the_float_range_is_drawn_not_a_code_panel(monkeypatch) -> None:
    # Each value is a finite float; their exact sum 2E+308 is not. (A waterfall
    # whose running total leaves the float range is a spec error already.)
    from matplotlib.figure import Figure

    values: list[str] = []
    saved = Figure.savefig

    def capture(self, *args, **kwargs):
        values.extend(
            text.get_gid()
            for axes in self.axes
            for text in axes.texts
            if (text.get_gid() or "").startswith("value\t") and text.get_visible()
        )
        return saved(self, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    markdown = "---\nprofile: plain\ncharts: true\n---\n\n```chart\n" + HUGE_STACK + "```\n"
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    document = renderer.render(MarkdownParser().parse(markdown))
    assert renderer.errors == []
    assert len(document.inline_shapes) == 1  # the picture, not the code panel
    assert not document.tables
    # The total is the exact sum at its own value, drawn or named as left out.
    total = "200,000" + ",000" * 101
    drawn = [gid.split("\t")[3] for gid in values if gid.split("\t")[1] == "가"]
    named = " ".join(w for w in renderer.warnings if "not written in the picture" in w)
    assert total in drawn or f": {total}" in named
    assert "3" in [gid.split("\t")[3] for gid in values if gid.split("\t")[1] == "나"] or (
        ": 3;" in named + ";"
    )
