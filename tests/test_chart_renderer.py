"""
Verification tests for standalone YAML chart rendering.

Covered scope:
- YAML chart-spec parsing and field-level validation
- PNG rendering for bar, line, and waterfall chart types
- Korean title and label rendering smoke coverage
- Pure cumulative waterfall position calculations
"""

import pytest

from chart_renderer import (
    ChartSeries,
    ChartSpec,
    ChartSpecError,
    parse_chart_spec,
    render_chart_png,
    waterfall_positions,
)

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def test_parse_valid_bar_chart_spec():
    """Parse all supported fields from a valid bar specification."""
    spec = parse_chart_spec(
        """
chart_type: bar
title: Revenue by year
labels: ["2024", "2025"]
series:
  - name: Revenue
    values: [1234, 1500.5]
y_label: USD millions
source: Company filings
"""
    )

    assert spec == ChartSpec(
        chart_type="bar",
        title="Revenue by year",
        labels=["2024", "2025"],
        series=[ChartSeries(name="Revenue", values=[1234.0, 1500.5])],
        y_label="USD millions",
        source="Company filings",
        total_label=None,
    )


def test_parse_unknown_chart_type_raises_field_error():
    """Reject a chart_type outside the supported set."""
    with pytest.raises(ChartSpecError, match="chart_type"):
        parse_chart_spec(
            """
chart_type: pie
labels: [A]
series:
  - name: Value
    values: [1]
"""
        )


def test_parse_series_length_mismatch_raises_field_error():
    """Reject series values that do not align with labels."""
    with pytest.raises(ChartSpecError, match=r"series\[0\]\.values"):
        parse_chart_spec(
            """
chart_type: bar
labels: [A, B]
series:
  - name: Value
    values: [1]
"""
        )


def test_parse_unknown_top_level_key_raises_field_error():
    """Reject and name unknown top-level fields."""
    with pytest.raises(ChartSpecError, match="unexpected_field"):
        parse_chart_spec(
            """
chart_type: line
labels: [A]
series:
  - name: Value
    values: [1]
unexpected_field: true
"""
        )


def test_parse_waterfall_with_two_series_raises_field_error():
    """Reject multiple series for a waterfall chart."""
    with pytest.raises(ChartSpecError, match="series"):
        parse_chart_spec(
            """
chart_type: waterfall
labels: [A]
series:
  - name: First
    values: [1]
  - name: Second
    values: [2]
"""
        )


def test_parse_bool_series_value_raises_field_error():
    """Reject bool values even though bool subclasses int in Python."""
    with pytest.raises(ChartSpecError, match=r"series\[0\]\.values\[1\]"):
        parse_chart_spec(
            """
chart_type: bar
labels: [A, B]
series:
  - name: Value
    values: [1, true]
"""
        )


@pytest.mark.parametrize(
    "spec",
    [
        ChartSpec(
            chart_type="bar",
            title="Bar chart",
            labels=["A", "B"],
            series=[ChartSeries(name="Value", values=[100.0, -30.0])],
            y_label=None,
            source=None,
        ),
        ChartSpec(
            chart_type="line",
            title="Line chart",
            labels=["A", "B"],
            series=[ChartSeries(name="Value", values=[100.0, 130.0])],
            y_label=None,
            source=None,
        ),
        ChartSpec(
            chart_type="waterfall",
            title="Waterfall chart",
            labels=["Start", "Cost", "Growth"],
            series=[ChartSeries(name="Change", values=[100.0, -30.0, 20.0])],
            y_label=None,
            source="Illustrative data",
            total_label="Ending value",
        ),
    ],
    ids=["bar", "line", "waterfall"],
)
def test_render_chart_type_returns_png_bytes(spec):
    """Render each supported chart type to a non-trivial PNG."""
    result = render_chart_png(spec)

    assert result.startswith(PNG_MAGIC)
    assert len(result) > 1000


def test_render_korean_labels_returns_png_bytes():
    """Render Korean chart text without raising an exception."""
    spec = ChartSpec(
        chart_type="bar",
        title="매출 및 영업이익",
        labels=["매출", "영업이익"],
        series=[ChartSeries(name="금액", values=[1200.0, 180.0])],
        y_label="백만원",
        source=None,
    )

    result = render_chart_png(spec)

    assert result.startswith(PNG_MAGIC)


def test_waterfall_positions_returns_expected_total():
    """Compute cumulative positions and the final waterfall total."""
    bottoms, tops, total = waterfall_positions([100, -30, 20])

    assert bottoms == [0.0, 70.0, 70.0]
    assert tops == [100.0, 100.0, 90.0]
    assert total == 90
