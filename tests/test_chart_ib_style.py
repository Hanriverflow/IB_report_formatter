"""IB/consulting chart style (E10): specification keys, drawing and Word paragraphs.

All names and numbers are fictional.
"""

import pytest

from chart_renderer import ChartSpecError, is_classic_spec, parse_chart_spec
from document_model import ChartReferenceLine, ChartSeries

STACKED = """chart_type: stacked_bar
title: 2028년 만기가 가장 크다
subtitle: "연도별 만기 도래액, 단위: 억원"
labels: [2026, 2027, 2028]
series:
  - name: 공모 회사채
    values: [1200, 2500, 3100]
  - name: 은행 차입
    values: [300.4, 500, 900]
totals: true
note: 9월 말 잔액 기준
source: 가상 자료
"""

BAR_LINE = """chart_type: bar_line
title: 배수가 낮아졌다
labels: [2024, 2025]
series:
  - name: 순차입금
    values: [3900, 3500]
  - name: 순차입금/EBITDA
    values: [2.8, 2.4]
    axis: right
right_number_format: multiple
"""

RANGE = """chart_type: range
title: 방법별 가치 범위
labels: [PER, DCF]
low: [10500, 11000]
high: [14000, 14500]
reference:
  - name: 현재 주가
    value: 12300
show_overlap: true
number_format: ',.0f'
"""

HBAR = """chart_type: hbar
title: 잔액 순위
labels: [가상화학, 가상전자, 가상건설]
series:
  - name: 잔액
    values: [18500, 14200, 9800]
highlight: [가상전자]
"""

LINE = """chart_type: line
title: 스프레드 추이
labels: [1월, 2월, 3월]
series:
  - name: 등급 평균
    values: [70, 68, 66]
  - name: 가상전자
    values: [74, 70, 64]
accent_series: 가상전자
reference:
  - name: 희망 스프레드
    value: 60
"""

WATERFALL = """chart_type: waterfall
title: 현금 증감
labels: [기초, 비용, 유입]
series:
  - name: 증감
    values: [100, -30, 20]
start_total: true
"""

CLASSIC_BAR = """chart_type: bar
title: 가상 매출
labels: [상반기, 하반기]
series:
  - name: 금액
    values: [100, -30]
y_label: 백만원
source: 가상 자료
"""


# ═══════════════════════════════════════════════════════════════════════════════
# SPECIFICATION
# ═══════════════════════════════════════════════════════════════════════════════


def test_new_types_parse_in_supplied_units():
    stacked = parse_chart_spec(STACKED)
    assert stacked.chart_type == "stacked_bar" and stacked.totals
    assert stacked.series[1].values == [300.4, 500, 900]
    assert (stacked.subtitle, stacked.note) == (
        "연도별 만기 도래액, 단위: 억원",
        "9월 말 잔액 기준",
    )
    assert stacked.labels == ["2026", "2027", "2028"]

    bar_line = parse_chart_spec(BAR_LINE)
    assert [s.axis for s in bar_line.series] == ["left", "right"]
    assert bar_line.right_number_format == "multiple"

    range_spec = parse_chart_spec(RANGE)
    assert range_spec.series == []
    assert (range_spec.low, range_spec.high) == ([10500, 11000], [14000, 14500])
    assert range_spec.reference_lines == [ChartReferenceLine("현재 주가", 12300)]
    assert range_spec.show_overlap

    hbar = parse_chart_spec(HBAR)
    assert hbar.highlight == ["가상전자"]

    line = parse_chart_spec(LINE)
    assert line.accent_series == "가상전자" and line.reference_lines[0].value == 60

    waterfall = parse_chart_spec(WATERFALL)
    assert waterfall.start_total


def test_classic_specs_are_exactly_the_old_keys_and_types():
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR))
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR.replace("bar", "line")))
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR.replace("bar", "waterfall", 1)))
    for spec in (STACKED, BAR_LINE, RANGE, HBAR, LINE, WATERFALL):
        assert not is_classic_spec(parse_chart_spec(spec))
    for extra in ("subtitle: '단위: 원'\n", "note: 가상\n"):
        assert not is_classic_spec(parse_chart_spec(CLASSIC_BAR + extra))
    single = CLASSIC_BAR.replace("[100, -30]", "[100, 30]")
    assert not is_classic_spec(parse_chart_spec(single + "highlight: [상반기]\n"))
    # A false flag written out is still the default value.
    assert is_classic_spec(parse_chart_spec(WATERFALL.replace("true", "false")))
    assert ChartSeries("a", [1.0]).axis == "left"


@pytest.mark.parametrize(
    "spec, diagnostic",
    [
        (STACKED.replace("stacked_bar", "pie"), "chart_type (or type) must be one of"),
        (STACKED.replace("[1200, 2500, 3100]", "[1200, -1, 3100]"), "must not be negative"),
        (
            STACKED.replace("series:\n", "series:\n" + "  - name: x\n    values: [1, 2, 3]\n" * 3),
            "at most 4 series",
        ),
        (STACKED.replace("totals: true", "totals: yes please"), "totals must be true or false"),
        (CLASSIC_BAR + "totals: true\n", "totals applies only to chart_type stacked_bar"),
        (BAR_LINE.replace("    axis: right\n", ""), "bar_line needs"),
        (BAR_LINE.replace("axis: right", "axis: top"), "axis must be left or right"),
        (
            CLASSIC_BAR.replace("    values: [100, -30]", "    values: [1, 2]\n    axis: right"),
            "unknown",
        ),
        (RANGE.replace("[10500, 11000]", "[15000, 11000]"), "low[0] must not be greater"),
        (RANGE.replace("high: [14000, 14500]\n", ""), "high must be a list"),
        (RANGE + "series: []\n", "series applies only to"),
        (RANGE.replace("value: 12300", "value: .nan"), "finite"),
        (RANGE.replace("    value: 12300\n", ""), "name and value"),
        (HBAR.replace("[가상전자]", "[가상해운]"), "not one of the labels"),
        (HBAR.replace("[가상전자]", "[가상전자, 가상전자]"), "repeats"),
        (HBAR.replace("series:\n", "series:\n  - name: x\n    values: [1, 2, 3]\n"), "exactly one"),
        (LINE.replace("accent_series: 가상전자", "accent_series: 없음"), "names no series"),
        (LINE + "highlight: [1월]\n", "highlight applies only to"),
        (STACKED + "highlight: [2026]\n", "highlight applies only to"),
        (
            CLASSIC_BAR.replace(
                "    values: [100, -30]", "    values: [1, 2]\n  - name: b\n    values: [3, 4]"
            )
            + "highlight: [상반기]\n",
            "single series",
        ),
        (WATERFALL + "reference: []\n", "reference applies only to"),
        (
            LINE.replace(
                "    value: 60\n",
                "    value: 60\n  - name: b\n    value: 1\n  - name: c\n    value: 2\n",
            ),
            "1 to 2 lines",
        ),
        (BAR_LINE + "number_format: wrong\n", "number_format"),
        (BAR_LINE.replace("multiple", "wrong"), "right_number_format"),
        (CLASSIC_BAR + "subtitle: [1]\n", "subtitle must be a string"),
    ],
)
def test_invalid_new_keys_are_chart_spec_errors(spec, diagnostic):
    with pytest.raises(ChartSpecError, match="Chart") as caught:
        parse_chart_spec(spec)
    assert diagnostic in str(caught.value)
