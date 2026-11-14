"""Principle N: every number a chart type writes is drawn whole or named in a warning.

Independent review, 2026-11-14: a small stacked segment's value vanished and
line-end names were pushed out of the picture, silently. For generated charts of
every standard type (series counts, value sizes and name lengths varied widely),
the picture as drawn is measured:

- every text's box lies inside the picture;
- no two value labels overlap;
- the values drawn plus the values a warning names are exactly the values the
  chart type writes.

The test reads the drawn figure only, not the drawer's own bookkeeping, so it
fails on the code before the fix (the two review reproductions are explicit
examples).
"""

import math
from collections import Counter

import pytest
import yaml
from hypothesis import HealthCheck, example, given, seed, settings
from hypothesis import strategies as st

import chart_render_standard as standard
from chart_renderer import format_value, parse_chart_spec, waterfall_positions

LETTERS = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz가나다라마바사아자차카타파하"
names = st.text(alphabet=LETTERS, min_size=1, max_size=14)
FORMATS = [None, "percent", "bps", "multiple", ",.2f"]


@st.composite
def magnitudes(draw, negative: bool = True, zero: bool = False) -> float:
    if zero and draw(st.integers(0, 9)) == 0:
        return 0
    value = draw(st.integers(1, 999)) * 10 ** draw(st.integers(0, 5))
    if draw(st.booleans()):
        value = value / 10
    return -value if negative and draw(st.integers(0, 4)) == 0 else value


@st.composite
def chart_data(draw) -> dict:
    kind = draw(
        st.sampled_from(["bar", "stacked_bar", "waterfall", "line", "bar_line", "range", "hbar"])
    )
    count = draw(st.integers(1, 12 if kind not in {"range", "hbar"} else 16))
    labels = draw(st.lists(names, min_size=count, max_size=count, unique=True))
    data: dict = {"type": kind, "labels": labels}
    fmt = draw(st.sampled_from(FORMATS))
    if fmt:
        data["number_format"] = fmt

    def values(negative=True, zero=False):
        return [draw(magnitudes(negative, zero)) for _ in labels]

    def series(n, negative=True, zero=False):
        series_names = draw(st.lists(names, min_size=n, max_size=n, unique=True))
        return [{"name": name, "values": values(negative, zero)} for name in series_names]

    if kind == "bar":
        data["series"] = series(draw(st.integers(1, 6)))
        if len(data["series"]) == 1 and draw(st.booleans()):
            data["highlight"] = [labels[0]]
    elif kind == "stacked_bar":
        data["series"] = series(draw(st.integers(1, 4)), negative=False, zero=True)
        data["totals"] = draw(st.booleans())
    elif kind in {"waterfall", "hbar"}:
        data["series"] = series(1)
    elif kind == "line":
        data["series"] = series(draw(st.integers(1, 16)))
        if draw(st.booleans()):
            data["reference"] = [
                {"name": draw(names), "value": draw(magnitudes())}
                for _ in range(draw(st.integers(1, 2)))
            ]
    elif kind == "bar_line":
        bars = series(draw(st.integers(1, 2)))
        lines = series(draw(st.integers(1, 2)))
        taken = {s["name"] for s in bars}
        for index, line in enumerate(lines):
            line["axis"] = "right"
            if line["name"] in taken:
                line["name"] += f"L{index}"
        data["series"] = bars + lines
        if draw(st.booleans()):
            data["right_number_format"] = "percent"
    elif kind == "range":
        low = values()
        data["low"] = low
        data["high"] = [value + abs(draw(magnitudes(negative=False, zero=True))) for value in low]
        if draw(st.booleans()):
            data["reference"] = [{"name": draw(names), "value": draw(magnitudes())}]
    return data


def expected_values(spec) -> Counter:
    """The numbers each type writes, as text (the test's own reading of the rules)."""

    def fmt(value, right=False):
        return format_value(value, spec.right_number_format if right else spec.number_format)

    kind = spec.chart_type
    found: list[str] = []
    if kind in {"bar", "hbar"}:
        found = [fmt(v) for s in spec.series for v in s.values]
    elif kind == "stacked_bar":
        found = [fmt(v) for s in spec.series for v in s.values if v != 0]
        if spec.totals:
            found += [
                fmt(math.fsum(s.values[i] for s in spec.series)) for i in range(len(spec.labels))
            ]
    elif kind == "waterfall":
        values = spec.series[0].values
        found = [fmt(v) for v in values] + [fmt(waterfall_positions(values)[2])]
    elif kind == "line":
        names_ = [s.name for s in spec.series]
        accent = (
            names_.index(spec.accent_series) if spec.accent_series in names_ else len(names_) - 1
        )
        points = len(spec.labels)
        at = range(points - 1) if points <= standard.LABELLED_POINTS else range(min(1, points - 1))
        found = [fmt(spec.series[accent].values[i]) for i in at]
        found += [fmt(s.values[-1]) for s in spec.series]
    elif kind == "bar_line":
        found = [fmt(v, s.axis == "right") for s in spec.series for v in s.values]
    elif kind == "range":
        found = [fmt(v) for pair in zip(spec.low, spec.high, strict=True) for v in pair]
    found += [fmt(line.value) for line in spec.reference_lines]
    return Counter(found)


def displayed(spec) -> dict[str, str]:
    """Each value label's text as drawn, mapped to the number it shows."""
    shown: dict[str, str] = {}

    def fmt(value, right=False):
        return format_value(value, spec.right_number_format if right else spec.number_format)

    for s in spec.series:
        for value in s.values:
            text = fmt(value, s.axis == "right")
            shown[text] = text
        shown[f"{fmt(s.values[-1])}  {s.name}"] = fmt(s.values[-1])
    for value in [*spec.low, *spec.high]:
        shown[fmt(value)] = fmt(value)
    for line in spec.reference_lines:
        shown[f"{line.name} {fmt(line.value)}"] = fmt(line.value)
    if spec.chart_type == "stacked_bar":
        for i in range(len(spec.labels)):
            total = fmt(math.fsum(s.values[i] for s in spec.series))
            shown[total] = total
    if spec.chart_type == "waterfall":
        total = fmt(waterfall_positions(spec.series[0].values)[2])
        shown[total] = total
    return shown


def measure(monkeypatch, spec, width: float) -> tuple[list, object, list[str]]:
    """Draw the chart and return every visible text with its box, the canvas and the warnings."""
    from matplotlib.figure import Figure
    from matplotlib.text import Text

    seen: dict = {}
    saved = Figure.savefig
    text_draw = Text.draw
    drawn: list = []

    def record(self, renderer):
        # Only texts the picture actually draws (not ticks outside the view).
        if self.get_visible() and self.get_text().strip():
            drawn.append(self)
        return text_draw(self, renderer)

    def capture(self, *args, **kwargs):
        drawn.clear()
        monkeypatch.setattr(Text, "draw", record)
        self.canvas.draw()
        monkeypatch.setattr(Text, "draw", text_draw)
        renderer = self.canvas.get_renderer()
        # Axis numbers and category names are not value labels.
        ticks = {
            id(tick)
            for axes in self.axes
            for tick in (*axes.get_xticklabels(), *axes.get_yticklabels())
        }
        seen["texts"] = [
            (artist.get_text(), artist.get_window_extent(renderer), id(artist) in ticks)
            for artist in drawn
        ]
        seen["canvas"] = self.bbox.frozen()
        return saved(self, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    warnings: list[str] = []
    standard.render_standard_chart_png(spec, width_inches=width, warnings=warnings)
    monkeypatch.setattr(Figure, "savefig", saved)
    return seen["texts"], seen["canvas"], warnings


def warned_values(warnings: list[str]) -> Counter:
    found: Counter = Counter()
    for warning in warnings:
        if "for lack of room: " not in warning:
            continue
        for item in warning.split("for lack of room: ", 1)[1].split("; "):
            found[item.rsplit(": ", 1)[1]] += 1
    return found


def check(monkeypatch, data: dict, width: float = 6.0) -> list[str]:
    spec = parse_chart_spec(yaml.safe_dump(data, allow_unicode=True))
    texts, canvas, warnings = measure(monkeypatch, spec, width)
    for text, box, _ in texts:
        assert box.x0 >= canvas.x0 - 1 and box.x1 <= canvas.x1 + 1, f"{text!r} leaves the picture"
        assert box.y0 >= canvas.y0 - 1 and box.y1 <= canvas.y1 + 1, f"{text!r} leaves the picture"
    shown = displayed(spec)
    values = [(text, box) for text, box, tick in texts if text in shown and not tick]
    for index, (text, box) in enumerate(values):
        for other, other_box in values[index + 1 :]:
            overlap = (
                box.x0 < other_box.x1 - 0.5
                and other_box.x0 < box.x1 - 0.5
                and box.y0 < other_box.y1 - 0.5
                and other_box.y0 < box.y1 - 0.5
            )
            assert not overlap, f"{text!r} overlaps {other!r}"
    drawn = Counter(shown[text] for text, _ in values)
    assert drawn + warned_values(warnings) == expected_values(spec)
    return warnings


SMALL_STACK = {
    "type": "stacked_bar",
    "labels": ["A", "B"],
    "series": [{"name": "Debt", "values": [1000, 10]}, {"name": "Bank", "values": [1000, 20]}],
}
MANY_LINES = {
    "type": "line",
    "labels": ["A", "B"],
    "series": [{"name": f"S{i:02}", "values": [0, 1]} for i in range(16)],
}


@seed(20261114)
@settings(
    max_examples=150,
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.function_scoped_fixture],
)
@given(chart_data(), st.sampled_from([4.5, 6.0]))
@example(SMALL_STACK, 6.0)
@example(MANY_LINES, 6.0)
def test_every_value_is_drawn_whole_or_named_in_a_warning(monkeypatch, data, width):
    check(monkeypatch, data, width)


def test_small_stacked_segments_are_written_beside_the_bar(monkeypatch):
    assert not [w for w in check(monkeypatch, SMALL_STACK) if "lack of room" in w]


def test_sixteen_line_ends_all_fit_inside_a_taller_picture(monkeypatch):
    assert not [w for w in check(monkeypatch, MANY_LINES) if "lack of room" in w]
    spec = parse_chart_spec(yaml.safe_dump(MANY_LINES))
    assert standard.chart_height(spec) == pytest.approx(16 * standard.LINE_GAP / 72 + 0.75)


def test_a_value_with_no_room_is_named_in_a_warning(monkeypatch):
    # Seven-digit values on 12 narrow bars of six series: some cannot be written.
    data = {
        "type": "bar",
        "labels": list("ABCDEFGHJKLM"),
        "series": [{"name": f"S{i}", "values": [1234567 + i for _ in range(12)]} for i in range(6)],
    }
    warnings = check(monkeypatch, data, 4.5)
    room = [w for w in warnings if "lack of room" in w]
    assert len(room) == 1
    listed = room[0].split("for lack of room: ", 1)[1].split("; ")
    assert room[0].startswith(f"Chart 'bar': {len(listed)} value(s) not written")
    # Drawn left to right, series by series: the first series keeps its values.
    assert "A, S1: 1,234,568" in listed
    assert not any(item.endswith(": 1,234,567") for item in listed)
