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
TWIN_GAP_POINTS = 16.0  # a point's value lies within 16 pt over or under it


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
        # Frozen now: the figure is cleared once the picture is saved.
        seen["axes"] = [axes.transData.frozen() for axes in self.axes]
        seen["dpi"] = self.dpi
        # Whose each value is (category, series, number), with its box; read
        # only to excuse a twin point's value (`twin_excuses`).
        seen["values"] = {
            tuple(artist.get_gid().split("\t")[1:4]): artist.get_window_extent(renderer)
            for artist in drawn
            if (artist.get_gid() or "").startswith("value\t")
        }
        return saved(self, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    warnings: list[str] = []
    standard.render_standard_chart_png(spec, width_inches=width, warnings=warnings)
    monkeypatch.setattr(Figure, "savefig", saved)
    return seen["texts"], seen["canvas"], warnings, seen


def twin_excuses(spec, spot, drawn: dict, warned: set, gap: float) -> Counter:
    """Texts of missing `bar_line` values that a twin point's number speaks for.

    Final rule, exception for twins: points of one category from different
    right-axis series that lie within half a pixel of each other and are
    written with the same text share one number. A value neither drawn nor
    named is excused only when it is itself such a point and its own twin's
    number is drawn at their shared spot (over or under it within `gap`).
    Keyed per point, so a missing equal value elsewhere is never excused.

    Args:
        spec: The chart spec.
        spot: `spot(series, index)` -> the point's pixel, measured from the
            right axis as drawn (not the renderer's record).
        drawn: (category, series, number) -> box of each drawn value.
        warned: (category, series, number) of each value a warning names.
        gap: Largest distance in pixels between a value and its point.

    Returns:
        Each excused value's text, counted once per excused point.
    """
    excused: Counter = Counter()
    if spec.chart_type != "bar_line":
        return excused
    right = [s for s in spec.series if s.axis == "right"]
    for index, category in enumerate(spec.labels):
        for s in right:
            text = format_value(s.values[index], spec.right_number_format)
            key = (category, s.name, text)
            if key in drawn or key in warned:
                continue
            x, y = spot(s, index)
            for other in right:
                ox, oy = spot(other, index)
                box = drawn.get((category, other.name, text))
                if (
                    other is not s
                    and box is not None
                    and format_value(other.values[index], spec.right_number_format) == text
                    and abs(ox - x) <= 0.5
                    and abs(oy - y) <= 0.5
                    and box.x0 - 0.5 <= x <= box.x1 + 0.5
                    and max(box.y0 - y, y - box.y1, 0) <= gap
                ):
                    excused[text] += 1
                    break
    return excused


def warned_values(warnings: list[str]) -> Counter:
    found: Counter = Counter()
    for warning in warnings:
        if "for lack of room: " not in warning:
            continue
        for item in warning.split("for lack of room: ", 1)[1].split("; "):
            found[item.rsplit(": ", 1)[1]] += 1
    return found


def warned_keys(warnings: list[str]) -> set:
    """(category, series, number) of each value a warning names."""
    found = set()
    for warning in warnings:
        if "for lack of room: " not in warning:
            continue
        for item in warning.split("for lack of room: ", 1)[1].split("; "):
            where, number = item.rsplit(": ", 1)
            category, _, series = where.rpartition(", ")
            found.add((category, series, number))
    return found


def check(monkeypatch, data: dict, width: float = 6.0) -> list[str]:
    spec = parse_chart_spec(yaml.safe_dump(data, allow_unicode=True))
    texts, canvas, warnings, seen = measure(monkeypatch, spec, width)
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
    accounted = drawn + warned_values(warnings)
    want = expected_values(spec)
    # Twin points: the number is written once for the pair, so the value not
    # written is neither drawn nor named; judged per point (`twin_excuses`).
    axes = seen["axes"]

    def spot(series, index):
        return axes[1].transform((index, series.values[index]))

    excused = (
        twin_excuses(
            spec, spot, seen["values"], warned_keys(warnings), TWIN_GAP_POINTS * seen["dpi"] / 72
        )
        if len(axes) > 1
        else Counter()
    )
    assert accounted == want - excused
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
# Fifth review (as in tests/test_chart_final_rule.py): two right-axis lines
# that share a spot and a text at A and C.
TWIN_POINTS = {
    "type": "bar_line",
    "title": "Twins",
    "labels": ["A", "B", "C"],
    "series": [
        {"name": "Bars", "values": [10, 12, 14]},
        {"name": "R1", "axis": "right", "values": [5, 6, 9]},
        {"name": "R2", "axis": "right", "values": [5, 7.5, 9]},
    ],
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
@example(TWIN_POINTS, 4.5)
@example(TWIN_POINTS, 6.0)
def test_every_value_is_drawn_whole_or_named_in_a_warning(monkeypatch, data, width):
    check(monkeypatch, data, width)


def test_twin_points_write_one_number_without_a_warning(monkeypatch):
    # Fifth review: at A and C the two right-axis lines share a spot and a
    # text, so 5 and 9 are written once each and nothing is named.
    spec = parse_chart_spec(yaml.safe_dump(TWIN_POINTS, allow_unicode=True))
    for width in (3.0, 3.5, 4.5, 6.0):
        assert check(monkeypatch, TWIN_POINTS, width) == []
        _, _, warnings, seen = measure(monkeypatch, spec, width)
        axes = seen["axes"]
        excused = twin_excuses(
            spec,
            lambda series, index, axes=axes: axes[1].transform((index, series.values[index])),
            seen["values"],
            warned_keys(warnings),
            TWIN_GAP_POINTS * seen["dpi"] / 72,
        )
        assert excused == Counter({"5": 1, "9": 1})


class _Box:
    """A stand-in for a measured text box (pixels)."""

    def __init__(self, x0: float, y0: float, x1: float, y1: float) -> None:
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1


def test_a_missing_equal_value_elsewhere_is_not_excused_by_twins():
    """Review S2, on hand-built measurements (no chart is drawn).

    At A the right-axis points R1 and R2 share a spot and the text 5, and both
    are written, one over and one under the spot. At B, R1 is also 5 but on
    its own spot (R2 is 7.5), and it is neither drawn nor named. A text-keyed
    excuse let A's drawn 5 cover B's miss; the per-point excuse does not.
    At C the twins share 9 and only R2's is drawn: R1's 9 is excused.
    """
    spec = parse_chart_spec(
        yaml.safe_dump(
            {
                "type": "bar_line",
                "labels": ["A", "B", "C"],
                "series": [
                    {"name": "Bars", "values": [10, 12, 14]},
                    {"name": "R1", "axis": "right", "values": [5, 5, 9]},
                    {"name": "R2", "axis": "right", "values": [5, 7.5, 9]},
                ],
            }
        )
    )

    # Pixel of each point: x by category, y by value (a fake right axis).
    def spot(series, index):
        return (100.0 + 100 * index, 10.0 * series.values[index])

    drawn = {
        ("A", "R1", "5"): _Box(95, 55, 105, 65),  # over A's spot (100, 50)
        ("A", "R2", "5"): _Box(95, 35, 105, 45),  # under it
        ("B", "R2", "7.5"): _Box(195, 80, 205, 90),
        ("C", "R2", "9"): _Box(295, 95, 305, 105),  # over C's spot (300, 90)
        **{(c, "Bars", n): _Box(0, 0, 1, 1) for c, n in (("A", "10"), ("B", "12"), ("C", "14"))},
    }
    excused = twin_excuses(spec, spot, drawn, set(), gap=21.0)
    assert excused == Counter({"9": 1})  # B's 5 is not excused
    want = expected_values(spec)
    accounted = Counter(number for _, _, number in drawn)
    # The checker's comparison: rejected for B's 5, and only for it.
    assert want - excused - accounted == Counter({"5": 1})
    assert accounted != want - excused


def test_small_stacked_segments_are_written_on_their_own_segment_or_named(monkeypatch):
    # Third independent review: a value beside its bar must have its centre
    # within its own segment's height. B's segments are 2 and 3 pixels high,
    # so their values (which the first fix wrote beside the bar, level with
    # each other) have no such place and are named; this replaces the pin
    # that nothing is left out. Where a value may be drawn is measured in
    # tests/test_chart_final_rule.py.
    room = [w for w in check(monkeypatch, SMALL_STACK) if "lack of room" in w]
    assert room == [
        "Chart 'stacked_bar': 2 value(s) not written in the picture for lack of room: "
        "B, Debt: 10; B, Bank: 20"
    ]


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
