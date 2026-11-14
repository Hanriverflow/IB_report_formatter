"""Principles N2 and N3: a drawn value sits on its own mark, or it is named in a warning.

Second independent review, 2026-11-14: a stacked segment's value moved beside
the bar landed over another category's bar, and a short bar's value lay over
the taller bar next to it, while the count of values drawn plus values warned
still matched. These tests read the drawn figure and measure every value label
with rules of their own (not the renderer's check):

- each value label names its category, series and number (artist `gid`), and
  the number is the spec's value there;
- it lies inside the picture and overlaps no other drawn text;
- it overlaps no bar but its own (a stack total: its own stack);
- it sits on its own mark: inside its bar, centred over or under its bar or
  point within 16 points, level with its row within 16 points, or, for a
  stacked segment written beside its bar, inside the category's column,
  level with the segment and in the stack's order;
- the values drawn plus the values a warning names are exactly the values the
  chart type writes, by category, series and number.

The two review reproductions are also checked without `gid`, by their unique
numbers, so the tests fail on the code before the fix.
"""

import math
from collections import Counter

import yaml
from hypothesis import HealthCheck, example, given, seed, settings
from hypothesis import strategies as st

import chart_render_standard as standard
from chart_renderer import format_value, parse_chart_spec, waterfall_positions
from render_styles import STYLE
from tests.test_chart_value_labels import MANY_LINES, SMALL_STACK, chart_data

GAP_POINTS = 16.0
REVIEW_STACK = {
    "type": "stacked_bar",
    "labels": ["C0", "C1", "C2", "C3"],
    "series": [
        {"name": "S0", "values": [10000, 100, 100000, 100]},
        {"name": "S1", "values": [2000000, 200000, 2000000, 200000]},
        {"name": "S2", "values": [300000, 30000, 3, 3000000]},
        {"name": "S3", "values": [4000, 4000000, 40, 400000]},
    ],
    "totals": True,
}
REVIEW_STACK_12 = {
    **REVIEW_STACK,
    "labels": [f"C{i}" for i in range(12)],
    "series": [{**s, "values": s["values"] * 3} for s in REVIEW_STACK["series"]],
}
REVIEW_BARS = {
    "type": "bar",
    "labels": list("ABCDEFGH"),
    "series": [
        {"name": "Short", "values": [123456] * 8},
        {"name": "Tall", "values": [1234567] * 8},
    ],
}


def expected(spec) -> Counter:
    """(category, series, number) of every value the chart type writes."""

    def fmt(value, right=False):
        return format_value(value, spec.right_number_format if right else spec.number_format)

    found: list[tuple[str, str, str]] = []
    kind = spec.chart_type
    if kind in {"bar", "hbar", "bar_line"}:
        found = [
            (label, s.name, fmt(v, s.axis == "right"))
            for s in spec.series
            for label, v in zip(spec.labels, s.values, strict=True)
        ]
    elif kind == "stacked_bar":
        found = [
            (label, s.name, fmt(v))
            for s in spec.series
            for label, v in zip(spec.labels, s.values, strict=True)
            if v != 0
        ]
        if spec.totals:
            found += [
                (label, STYLE.CHART_TOTAL_LABEL, fmt(math.fsum(s.values[i] for s in spec.series)))
                for i, label in enumerate(spec.labels)
            ]
    elif kind == "waterfall":
        s = spec.series[0]
        found = [(label, s.name, fmt(v)) for label, v in zip(spec.labels, s.values, strict=True)]
        total = waterfall_positions(s.values)[2]
        found.append((spec.total_label or STYLE.CHART_TOTAL_LABEL, s.name, fmt(total)))
    elif kind == "line":
        names = [s.name for s in spec.series]
        accent = names.index(spec.accent_series) if spec.accent_series in names else len(names) - 1
        points = len(spec.labels)
        at = range(points - 1) if points <= standard.LABELLED_POINTS else range(min(1, points - 1))
        found = [(spec.labels[i], names[accent], fmt(spec.series[accent].values[i])) for i in at]
        found += [(spec.labels[-1], s.name, fmt(s.values[-1])) for s in spec.series]
    elif kind == "range":
        for label, low, high in zip(spec.labels, spec.low, spec.high, strict=True):
            found += [(label, "low", fmt(low)), (label, "high", fmt(high))]
    found += [("", line.name, fmt(line.value)) for line in spec.reference_lines]
    return Counter(found)


def draw(monkeypatch, spec, width: float) -> dict:
    """Draw the chart; return what the picture shows, measured after drawing."""
    from matplotlib.figure import Figure
    from matplotlib.patches import Rectangle
    from matplotlib.text import Text

    seen: dict = {}
    saved, text_draw = Figure.savefig, Text.draw
    texts: list = []

    def record(self, renderer):
        if self.get_visible() and self.get_text().strip():
            texts.append(self)
        return text_draw(self, renderer)

    def capture(self, *args, **kwargs):
        texts.clear()
        monkeypatch.setattr(Text, "draw", record)
        self.canvas.draw()
        monkeypatch.setattr(Text, "draw", text_draw)
        renderer = self.canvas.get_renderer()
        seen["texts"] = [(t.get_text(), t.get_gid(), t.get_window_extent(renderer)) for t in texts]
        seen["bars"] = [
            (patch.get_gid(), patch.get_window_extent(renderer))
            for patch in self.findobj(match=Rectangle)
            if (patch.get_gid() or "").startswith("bar\t")
        ]
        # Frozen now: the figure is cleared once the picture is saved.
        seen["axes"] = [axes.transData.frozen() for axes in self.axes]
        seen["canvas"] = self.bbox.frozen()
        seen["dpi"] = self.dpi
        return saved(self, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    warnings: list[str] = []
    standard.render_standard_chart_png(spec, width_inches=width, warnings=warnings)
    monkeypatch.setattr(Figure, "savefig", saved)
    seen["warnings"] = warnings
    return seen


def _overlap(a, b, tolerance=0.5) -> bool:
    return (
        a.x0 < b.x1 - tolerance
        and b.x0 < a.x1 - tolerance
        and a.y0 < b.y1 - tolerance
        and b.y0 < a.y1 - tolerance
    )


def _inside(outer, inner, tolerance=0.5) -> bool:
    return (
        inner.x0 >= outer.x0 - tolerance
        and inner.x1 <= outer.x1 + tolerance
        and inner.y0 >= outer.y0 - tolerance
        and inner.y1 <= outer.y1 + tolerance
    )


def warned(warnings: list[str]) -> Counter:
    found: Counter = Counter()
    for warning in warnings:
        if "for lack of room: " not in warning:
            continue
        for item in warning.split("for lack of room: ", 1)[1].split("; "):
            where, number = item.rsplit(": ", 1)
            category, _, series = where.rpartition(", ")
            found[(category, series, number)] += 1
    return found


def check(monkeypatch, data: dict, width: float = 6.0) -> dict:
    spec = parse_chart_spec(yaml.safe_dump(data, allow_unicode=True))
    seen = draw(monkeypatch, spec, width)
    canvas, gap = seen["canvas"], GAP_POINTS * seen["dpi"] / 72
    values = [
        (tuple(gid.split("\t")[1:]), text, box)
        for text, gid, box in seen["texts"]
        if (gid or "").startswith("value\t")
    ]
    bars: dict[tuple[str, str], list] = {}
    for gid, box in seen["bars"]:
        _, category, series = gid.split("\t")
        bars.setdefault((category, series), []).append(box)
    want = expected(spec)
    drawn: Counter = Counter()
    kind = spec.chart_type
    ax = seen["axes"][0]
    stacks: dict[str, list] = {}
    for (category, series, number), text, box in values:
        key = (category, series, number)
        assert key in want, f"{key} is not a value of the spec"
        assert number in text, f"{text!r} does not show {number}"
        drawn[key] += 1
        assert _inside(canvas, box), f"{key} leaves the picture"
        for other_text, _, other in seen["texts"]:
            if other is not box and _overlap(box, other):
                assert other_text == text and other == box, f"{key} overlaps {other_text!r}"
        own = bars.get((category, "range" if kind == "range" else series), [])
        if series == STYLE.CHART_TOTAL_LABEL and kind == "stacked_bar":
            own = [b for (c, _), boxes in bars.items() if c == category for b in boxes]
        for (c, s), boxes in bars.items():
            for bar in boxes:
                if not any(bar is o for o in own):
                    assert not _overlap(box, bar), f"{key} lies over the bar of {(c, s)}"
        if category == "" or (kind == "line" and category == spec.labels[-1] and "  " in text):
            continue  # a reference line or a line end carries its own name
        cx, cy = (box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2
        if own and any(_inside(o, box) for o in own):
            if kind == "stacked_bar" and series != STYLE.CHART_TOTAL_LABEL:
                stacks.setdefault(category, []).append((series, cy))
            continue
        if kind in {"line", "bar_line"} and not own:
            axes = seen["axes"][1] if kind == "bar_line" else ax
            index = spec.labels.index(category)
            values_of = next(s for s in spec.series if s.name == series).values
            px, py = axes.transform((index, values_of[index]))
            assert abs(cx - px) <= 1, f"{key} is not over its point"
            assert max(box.y0 - py, py - box.y1, 0) <= gap, f"{key} is far from its point"
            continue
        if kind in {"hbar", "range"}:
            if own:
                row = own[0]
                level = box.y0 <= row.y1 + 0.5 and box.y1 >= row.y0 - 0.5
                beside = max(box.x0 - row.x1, row.x0 - box.x1, 0)
            else:  # a range of one value: its point
                index = spec.labels.index(category)
                bound = spec.low[index]
                px, py = ax.transform((bound, len(spec.labels) - 1 - index))
                level = box.y0 <= py <= box.y1
                beside = max(box.x0 - px, px - box.x1, 0)
            assert level and beside <= gap, f"{key} is not beside its bar"
            continue
        assert own, f"{key} has no bar"
        x0, x1 = min(o.x0 for o in own), max(o.x1 for o in own)
        top, bottom = max(o.y1 for o in own), min(o.y0 for o in own)
        if x0 - 1 <= cx <= x1 + 1:
            assert max(box.y0 - top, bottom - box.y1, 0) <= gap, f"{key} is far from its bar"
            continue
        # A stacked segment's value beside its bar: inside the category's column.
        assert kind == "stacked_bar", f"{key} is not over or under its bar"
        index = spec.labels.index(category)
        left = ax.transform((index - 0.5, 0))[0]
        right = ax.transform((index + 0.5, 0))[0]
        assert left - 0.5 <= box.x0 and box.x1 <= right + 0.5, f"{key} leaves its column"
        segment = own[0]
        assert box.y0 <= segment.y1 + 0.5 and box.y1 >= segment.y0 - 0.5, f"{key} not level"
        stacks.setdefault(category, []).append((series, cy))
    order = [s.name for s in spec.series]
    for category, items in stacks.items():
        by_rank = [cy for _, cy in sorted(items, key=lambda item: order.index(item[0]))]
        assert by_rank == sorted(by_rank), f"{category}: values out of the stack's order"
    assert drawn + warned(seen["warnings"]) == want
    return seen


@seed(20261114)
@settings(
    max_examples=150,
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.function_scoped_fixture],
)
@given(chart_data(), st.sampled_from([3.0, 4.5, 6.0]))
@example(REVIEW_STACK, 3.0)
@example(REVIEW_STACK, 4.5)
@example(REVIEW_STACK, 6.0)
@example(REVIEW_STACK_12, 6.0)
@example(REVIEW_BARS, 3.0)
@example(REVIEW_BARS, 4.5)
@example(REVIEW_BARS, 6.0)
@example(SMALL_STACK, 6.0)
@example(MANY_LINES, 6.0)
def test_every_drawn_value_sits_on_its_own_mark(monkeypatch, data, width):
    check(monkeypatch, data, width)


def _texts(seen, number: str) -> list:
    return [box for text, _, box in seen["texts"] if text == number]


def test_a_moved_stack_value_stays_in_its_own_column(monkeypatch):
    """Review [1], read without `gid`: the two unique numbers that moved."""
    for data, width in ((REVIEW_STACK, 3.0), (REVIEW_STACK_12, 6.0)):
        spec = parse_chart_spec(yaml.safe_dump(data))
        seen = draw(monkeypatch, spec, width)
        ax = seen["axes"][0]
        for number, category in (("4,000,000", "C1"), ("3,000,000", "C3")):
            index = spec.labels.index(category)
            left = ax.transform((index - 0.5, 0))[0]
            right = ax.transform((index + 0.5, 0))[0]
            for box in _texts(seen, number):
                assert left <= box.x0 and box.x1 <= right, f"{number} drawn outside {category}"


def test_a_short_bar_value_never_lies_over_the_tall_bar(monkeypatch):
    """Review [2], read without `gid`: each `123,456` against the taller bars."""
    from matplotlib.figure import Figure
    from matplotlib.patches import Rectangle

    spec = parse_chart_spec(yaml.safe_dump(REVIEW_BARS))
    for width in (3.0, 4.5, 6.0):
        boxes: dict = {}
        saved = Figure.savefig

        def capture(self, *args, saved=saved, boxes=boxes, **kwargs):
            self.canvas.draw()
            renderer = self.canvas.get_renderer()
            boxes["tall"] = [
                p.get_window_extent(renderer)
                for p in self.findobj(match=Rectangle)
                if abs(p.get_height() - 1234567) < 1
            ]
            boxes["short"] = [
                t.get_window_extent(renderer)
                for t in self.texts + [t for a in self.axes for t in a.texts]
                if t.get_text() == "123,456" and t.get_visible()
            ]
            return saved(self, *args, **kwargs)

        monkeypatch.setattr(Figure, "savefig", capture)
        warnings: list[str] = []
        standard.render_standard_chart_png(spec, width_inches=width, warnings=warnings)
        monkeypatch.setattr(Figure, "savefig", saved)
        assert len(boxes["tall"]) == 8
        for box in boxes["short"]:
            assert not any(_overlap(box, tall) for tall in boxes["tall"]), width
        shown = len(boxes["short"])
        named = sum(item.endswith(": 123,456") for w in warnings for item in w.split("; "))
        assert shown + named == 8


def test_long_names_are_broken_inside_the_picture(monkeypatch):
    """Review [6]: a 160-character series name and a 120-character category name."""
    data = {
        "type": "bar",
        "labels": ["X" * 120, "B"],
        "series": [
            {"name": "LONGNAME" * 20, "values": [1, 2]},
            {"name": "Other", "values": [3, 4]},
        ],
    }
    spec = parse_chart_spec(yaml.safe_dump(data))
    seen = draw(monkeypatch, spec, 6.0)
    for text, _, box in seen["texts"]:
        assert _inside(seen["canvas"], box), f"{text[:20]!r} leaves the picture"
    assert not [w for w in seen["warnings"] if "do not fit" in w]
