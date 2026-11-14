"""The final rule for chart values, measured independently of the engine.

Three independent reviews in a row found values in wrong or unreadable places
that the engine's own check (`_Drawer.fits`, used both to place and to verify)
and the earlier placement test let through, because all three read the same
list of rules. This module does not call the engine's check and does not use
its formatter:

- what the picture shows is taken from the drawn figure;
- whether a line or point marker lies over a value is decided on pixels: the
  values are drawn alone, the lines and markers are drawn alone, and no value
  may share ink with them;
- which mark a value belongs to is judged by distance from the text's centre
  to every bar and every plotted point together, not only to its own, and by
  what lies in the band between the text and its own mark;
- the numbers expected are formatted here with `Decimal`, and totals are the
  sums of the numbers as written.

The rule (third independent review and the self-review, 2026-11-14): every
value the picture shows lies whole inside it, overlaps no other text, no
bar but the one it is written in, no line, point marker, reference line or
zero line; no other bar or point is nearer to its centre than its own by more
than 3 pt; a stacked value beside its bar has its centre within its own
segment's height and is not level with another value beside the same bar.
A value with no such place is named in the warning, never drawn elsewhere.

Fourth independent review and the coordinator's reading of the Word pages
(2026-11-14): a value written outside its mark has nothing in the band between
the text and the mark (no other bar, line, point or reference line: a bar's
value above another series' line read as that line's value); a bar's value
has no plotted point nearer than its bar, and a point's value no bar nearer
than its point; a line-end name is nearest to its own line's end; and a
value-axis tick is compared with its position as a decimal, not within a
float tolerance. Self-review after that: a point's value under its point and
just over a bar's end stood as near to the bar as to the point, so the own
mark has to be nearer than every mark of the other kind by 3 pt; and a value
beside a stack that is as high as a whole other thin segment of the bar names
neither segment, so it is not drawn. Reading the Word pages: a line's value
stood halfway between its own point and another series' point of the same
category, so there too the own point has to be nearer by 3 pt. In a
horizontal chart (`hbar`, `range`) a value runs sideways
from its bar's end, so a longer neighbouring bar is always nearer to the
text's centre: there the value is judged by its centre lying within its own
row's height and at most 16 pt beside the bar (an approved exception), with
the band rule on top.

Fifth independent review (2026-11-15): the engine's check and this measure
both took a mark for a value's own because it lay where the value's point
lay. Another series' line through that spot was left out of the band, and
another series' point on that spot was no rival: `A, Own: 100.00` stood over
a spot it shared with the other series' 100.01. Here a series' line is now
found by its data (the solid line that plots exactly that series' values) and
a point by its series and category; where a mark lies decides nothing. Only
the two pieces of the value's own series' line that meet at its own point,
and its own marker, are left out of the band. Another series' point of the
same category is a rival also on the very same spot, with one exception: two
points on one spot (within half a pixel) that are written with the same text,
as where two index lines both start at 100. That number is true of both, so
it is written once, and the value not written is not named as missing.
"""

import math
import random
import warnings as python_warnings
from collections import Counter
from decimal import ROUND_HALF_UP, Context, Decimal

import numpy as np
import pytest
import yaml

import chart_render_standard as standard
from chart_renderer import parse_chart_spec
from render_styles import STYLE

MARGIN_POINTS = 3.0
GAP_POINTS = 16.0
# Pixels left out at the mark's own edge when the band between a text and its
# mark is searched: a line along that edge (the baseline, a waterfall
# connector leaving the bar's end) is not between them.
EDGE_PIXELS = 1.5

# ═══════════════════════════════════════════════════════════════════════════════
# WHAT THE PICTURE SHOWS
# ═══════════════════════════════════════════════════════════════════════════════


def _ink(image) -> np.ndarray:
    rgb = image[:, :, :3].astype(int)
    return (255 - rgb).max(axis=2) > 48


def _only(fig, keep: set[int]) -> np.ndarray:
    """Ink of the figure drawn with only the artists in `keep`."""
    from matplotlib.collections import Collection
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    from matplotlib.text import Text

    backgrounds = {id(fig.patch)} | {id(axes.patch) for axes in fig.axes}
    hidden = []
    for artist in fig.findobj():
        if isinstance(artist, Text | Line2D | Patch | Collection) and id(artist) not in backgrounds:
            hidden.append((artist, artist.get_visible()))
            artist.set_visible(artist.get_visible() and id(artist) in keep)
    try:
        fig.canvas.draw()
        return _ink(np.asarray(fig.canvas.buffer_rgba()).copy())
    finally:
        for artist, visible in hidden:
            artist.set_visible(visible)


def draw(monkeypatch, spec, width: float) -> dict:
    """Draw the chart and return what the saved picture holds."""
    from matplotlib.figure import Figure
    from matplotlib.text import Text

    seen: dict = {}
    saved, text_draw = Figure.savefig, Text.draw

    def capture(fig, *args, **kwargs):
        drawn: list = []

        def record(text, renderer, *a, **k):
            if text.get_visible() and text.get_text().strip():
                drawn.append(text)
            return text_draw(text, renderer, *a, **k)

        monkeypatch.setattr(Text, "draw", record)
        fig.canvas.draw()
        monkeypatch.setattr(Text, "draw", text_draw)
        renderer = fig.canvas.get_renderer()
        ticks = {
            id(tick.label1): (index, float(tick.get_loc()))
            for index, axes in enumerate(fig.axes)
            for tick in axes.yaxis.get_major_ticks()
        }
        texts, done = [], set()
        for text in drawn:
            if id(text) in done:
                continue
            done.add(id(text))
            box = text.get_window_extent(renderer)
            texts.append(
                {
                    "text": text.get_text().replace("\\$", "$"),
                    "raw": text.get_text(),
                    "gid": text.get_gid() or "",
                    "box": (box.x0, box.y0, box.x1, box.y1),
                    "tick": ticks.get(id(text)),
                    "artist": text,
                }
            )
        seen["texts"] = texts
        seen["bars"] = [
            (patch.get_gid(), tuple(patch.get_window_extent(renderer).extents))
            for axes in fig.axes
            for patch in axes.patches
            if (patch.get_gid() or "").startswith("bar\t")
        ]
        lines = [line for axes in fig.axes for line in axes.lines]
        seen["lines"] = [
            {
                "axes": fig.axes.index(line.axes),
                "gid": line.get_gid() or "",
                "ydata": [float(value) for value in line.get_ydata()],
                "xdata": [float(value) for value in line.get_xdata()],
                "dashed": line.is_dashed(),
                "pixels": [
                    (float(x), float(y))
                    for x, y in line.get_transform().transform(line.get_xydata())
                ],
                "drawn": line.get_linestyle() not in ("None", "", " "),
                "radius": (
                    0.0
                    if line.get_marker() in (None, "None", "", " ", "|")
                    else float(line.get_markersize()) * fig.dpi / 72 / 2
                ),
            }
            for line in lines
        ]
        seen["transforms"] = [axes.transData.frozen() for axes in fig.axes]
        seen["size"] = (fig.bbox.width, fig.bbox.height)
        seen["dpi"] = fig.dpi
        values = [item["artist"] for item in texts if item["gid"].startswith("value\t")]
        colours = [(artist, artist.get_color()) for artist in values]
        for artist in values:  # white values inside bars are ink too
            artist.set_color("#000000")
        seen["value_ink"] = _only(fig, {id(artist) for artist in values})
        for artist, colour in colours:
            artist.set_color(colour)
        seen["line_ink"] = _only(fig, {id(line) for line in lines})
        fig.canvas.draw()
        return saved(fig, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    notes: list[str] = []
    with python_warnings.catch_warnings(record=True) as caught:
        python_warnings.simplefilter("always")
        seen["png"] = standard.render_standard_chart_png(spec, width_inches=width, warnings=notes)
    monkeypatch.setattr(Figure, "savefig", saved)
    seen["warnings"] = notes
    seen["python_warnings"] = [str(item.message) for item in caught]
    return seen


# ═══════════════════════════════════════════════════════════════════════════════
# WHAT IS EXPECTED
# ═══════════════════════════════════════════════════════════════════════════════


WIDE = Context(prec=2000)


def written_number(value: float | Decimal, number_format: str | None) -> str:
    """The number as the README says it is written, worked out here.

    The decimal the YAML held (the float's shortest text) is rounded as a
    spreadsheet rounds, a 5 away from zero (owner decision 2026-11-14).
    """
    exact = value if isinstance(value, Decimal) else Decimal(repr(float(value)))
    magnitude = exact.copy_abs()
    suffix = {"percent": "%", "bps": " bps", "multiple": "x"}.get(number_format or "", "")
    if number_format and number_format not in ("number", "percent", "bps", "multiple"):
        digits = int(number_format.split(".")[1].rstrip("f")) if "." in number_format else 6
        step = Decimal(1).scaleb(-digits)
        rounded = magnitude.quantize(step, rounding=ROUND_HALF_UP, context=WIDE)
        text = f"{rounded:,f}" if number_format.startswith(",") else f"{rounded:f}"
    elif magnitude == magnitude.to_integral_value():
        text = f"{int(magnitude):,}"
    else:
        rounded = magnitude.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP, context=WIDE)
        text = f"{rounded:,f}"
    text += suffix
    # A negative amount whose written digits are all zero is not marked negative.
    nonzero = any(char in "123456789" for char in text)
    return f"({text})" if exact < 0 and nonzero else text


def written_sum(values) -> Decimal:
    """The exact sum of the numbers as written in the YAML, rounded by no one."""
    total = Decimal(0)
    for value in values:
        total = WIDE.add(total, Decimal(repr(float(value))))
    return total


def expected(spec) -> list[tuple[str, str, str, str]]:
    """Every number the chart type writes: category, series, text, role."""
    kind = spec.chart_type

    def fmt(value, right=False):
        return written_number(value, spec.right_number_format if right else spec.number_format)

    found = []
    if kind in ("bar", "hbar"):
        for series in spec.series:
            found += [
                (label, series.name, fmt(value), "bar")
                for label, value in zip(spec.labels, series.values, strict=True)
            ]
    elif kind == "stacked_bar":
        for series in spec.series:
            found += [
                (label, series.name, fmt(value), "segment")
                for label, value in zip(spec.labels, series.values, strict=True)
                if value != 0
            ]
        if spec.totals:
            for index, label in enumerate(spec.labels):
                total = written_sum(series.values[index] for series in spec.series)
                found.append((label, STYLE.CHART_TOTAL_LABEL, fmt(total), "total"))
    elif kind == "waterfall":
        series = spec.series[0]
        found += [
            (label, series.name, fmt(value), "bar")
            for label, value in zip(spec.labels, series.values, strict=True)
        ]
        total_label = spec.total_label or STYLE.CHART_TOTAL_LABEL
        found.append((total_label, series.name, fmt(written_sum(series.values)), "bar"))
    elif kind == "line":
        names = [series.name for series in spec.series]
        accent = names.index(spec.accent_series) if spec.accent_series in names else len(names) - 1
        count = len(spec.labels)
        for index in range(count - 1) if count <= 8 else range(min(1, count - 1)):
            value = spec.series[accent].values[index]
            found.append((spec.labels[index], names[accent], fmt(value), "point"))
        found += [
            (spec.labels[-1], series.name, fmt(series.values[-1]), "line") for series in spec.series
        ]
    elif kind == "bar_line":
        for series in spec.series:
            right = series.axis == "right"
            found += [
                (label, series.name, fmt(value, right), "point" if right else "bar")
                for label, value in zip(spec.labels, series.values, strict=True)
            ]
    elif kind == "range":
        for label, low, high in zip(spec.labels, spec.low, spec.high, strict=True):
            found += [(label, "low", fmt(low), "row"), (label, "high", fmt(high), "row")]
    found += [("", line.name, fmt(line.value), "reference") for line in spec.reference_lines]
    return found


# ═══════════════════════════════════════════════════════════════════════════════
# THE MEASURE
# ═══════════════════════════════════════════════════════════════════════════════


def _overlap(a, b, tolerance=0.5) -> bool:
    return (
        a[0] < b[2] - tolerance
        and b[0] < a[2] - tolerance
        and a[1] < b[3] - tolerance
        and b[1] < a[3] - tolerance
    )


def _inside(outer, inner, tolerance=0.5) -> bool:
    return (
        inner[0] >= outer[0] - tolerance
        and inner[2] <= outer[2] + tolerance
        and inner[1] >= outer[1] - tolerance
        and inner[3] <= outer[3] + tolerance
    )


def _from_box(box, x, y) -> float:
    return math.hypot(max(box[0] - x, x - box[2], 0.0), max(box[1] - y, y - box[3], 0.0))


def _band(box, target):
    """The rectangle between a text and its mark, as wide (or as high) as the text."""
    if box[1] >= target[3]:  # the text is above its mark
        band = (box[0], target[3] + EDGE_PIXELS, box[2], box[1])
    elif box[3] <= target[1]:  # below
        band = (box[0], box[3], box[2], target[1] - EDGE_PIXELS)
    elif box[0] >= target[2]:  # to the right
        band = (target[2] + EDGE_PIXELS, box[1], box[0], box[3])
    elif box[2] <= target[0]:  # to the left
        band = (box[2], box[1], target[0] - EDGE_PIXELS, box[3])
    else:
        return None
    return band if band[2] - band[0] > 0.5 and band[3] - band[1] > 0.5 else None


def _within(band, x, y, grow=0.0) -> bool:
    return band[0] - grow < x < band[2] + grow and band[1] - grow < y < band[3] + grow


def in_the_band(band, seen, own_bars, own=None, twins=()) -> list[str]:
    """What lies between a text and its own mark: bars, lines and point markers.

    Lines are walked half a pixel at a time. `own` is a point's value's own
    point, as its series' line and the category's index: only the two pieces
    of that line that meet at that point, and that point's marker, are its own
    mark and left out. `twins` are points (line, index) the number is also
    true of; their markers are left out, their lines are not. Nothing is left
    out for lying where the own point lies.
    """
    found = []
    for _, bar in seen["bars"]:
        if not any(bar is mine for mine in own_bars) and _overlap(band, bar):
            found.append("another bar")
    for line in seen["lines"]:
        pixels = line["pixels"]
        mine = own is not None and line is own[0]
        if line["drawn"]:
            pieces = enumerate(zip(pixels, pixels[1:], strict=False))
            for start, ((x0, y0), (x1, y1)) in pieces:
                if mine and own[1] in (start, start + 1):
                    continue
                steps = max(1, int(math.hypot(x1 - x0, y1 - y0) * 2))
                if any(
                    _within(band, x0 + (x1 - x0) * step / steps, y0 + (y1 - y0) * step / steps)
                    for step in range(steps + 1)
                ):
                    found.append("a line")
                    break
        if not line["radius"]:
            continue
        for index, (x, y) in enumerate(pixels):
            if mine and index == own[1]:
                continue
            if any(line is other and index == at for other, at in twins):
                continue
            if _within(band, x, y, line["radius"] * 0.7):
                found.append("a point marker")
                break
    return found


def _same_data(drawn, values) -> bool:
    """Whether a drawn line plots exactly these values.

    Values beyond 1e290 are drawn at one power-of-two scale, which is exact.
    """
    values = [float(value) for value in values]
    if len(drawn) != len(values):
        return False
    if list(drawn) == values:
        return True
    largest = max(values, key=abs, default=0.0)
    if abs(largest) <= 1e290:
        return False
    scale = drawn[values.index(largest)] / largest
    return math.frexp(scale)[0] == 0.5 and all(
        y == value * scale for y, value in zip(drawn, values, strict=True)
    )


def series_lines(spec, seen) -> dict[int, dict]:
    """The drawn line of each plotted series, by the series' place in the spec.

    A series' line is the solid line over all categories, on the axes its
    series is plotted on, whose y data are exactly the series' values. Where
    the line lies in the picture plays no part. Two series with the same
    values have two such lines, taken in drawing order: they are one line
    drawn twice.
    """
    kind = spec.chart_type
    if kind not in ("line", "bar_line"):
        return {}
    xs = [float(index) for index in range(len(spec.labels))]
    axes = 1 if kind == "bar_line" else 0
    free = [
        line
        for line in seen["lines"]
        if line["xdata"] == xs and not line["dashed"] and line["axes"] == axes
    ]
    found = {}
    for index, series in enumerate(spec.series):
        if kind == "bar_line" and series.axis != "right":
            continue
        match = next((line for line in free if _same_data(line["ydata"], series.values)), None)
        if match is not None:
            free = [line for line in free if line is not match]
            found[index] = match
    return found


def problems(spec, seen) -> list[str]:
    """Every breach of the rule in the drawn picture, in words."""
    found: list[str] = []
    width, height = seen["size"]
    point = seen["dpi"] / 72
    canvas = (0.0, 0.0, width, height)
    texts = seen["texts"]
    values = [item for item in texts if item["gid"].startswith("value\t")]
    bars: dict[tuple[str, str], list] = {}
    for gid, box in seen["bars"]:
        _, category, series = gid.split("\t")
        bars.setdefault((category, series), []).append(box)
    every_bar = [box for boxes in bars.values() for box in boxes]
    kind = spec.chart_type
    name = spec.title or spec.chart_type

    # 0. Every plotted series has its line, and the engine's own record of
    # whose line it is (the line's gid) says the same as the line's data.
    lines_of = series_lines(spec, seen)
    for index, series in enumerate(spec.series):
        plotted = kind == "line" or (kind == "bar_line" and series.axis == "right")
        if plotted and index not in lines_of:
            found.append(f"no line plots the values of {series.name!r}")
    for index, line in lines_of.items():
        if not line["gid"]:
            continue
        marked = line["gid"].split("\t")
        said = int(marked[1]) if marked[0] == "line" and marked[1].isdigit() else -1
        if said != index and not (
            0 <= said < len(spec.series)
            and list(spec.series[said].values) == list(spec.series[index].values)
        ):
            found.append(f"the line of series {index} is marked {line['gid']!r}")
    # Every plotted point: its series, its category's index, its pixel.
    points = [
        (series_index, index, x, y)
        for series_index, line in lines_of.items()
        for index, (x, y) in enumerate(line["pixels"])
    ]

    def text_of(series_index: int, index: int) -> str:
        number_format = spec.right_number_format if kind == "bar_line" else spec.number_format
        return written_number(spec.series[series_index].values[index], number_format)

    def candidates(category: str, series: str, number: str) -> list[tuple[int, int]]:
        """The plotted points a value's description (its gid) can stand for."""
        return [
            (series_index, index)
            for series_index in lines_of
            for index, label in enumerate(spec.labels)
            if spec.series[series_index].name == series
            and label == category
            and text_of(series_index, index) == number
        ]

    def twins_of(series_index: int, index: int) -> list[tuple[int, int]]:
        """Other series' points of the category on the same spot with the same text."""
        x, y = lines_of[series_index]["pixels"][index]
        return [
            (other, index)
            for other, line in lines_of.items()
            if other != series_index
            and abs(line["pixels"][index][0] - x) <= 0.5
            and abs(line["pixels"][index][1] - y) <= 0.5
            and text_of(other, index) == text_of(series_index, index)
        ]

    # 1. Every number is drawn once or named in the one warning.
    want = expected(spec)
    roles = {}
    for category, series, number, role in want:
        roles.setdefault((category, series, number), role)
    drawn = Counter(tuple(item["gid"].split("\t")[1:4]) for item in values)
    missing = Counter((c, s, n) for c, s, n, _ in want) - drawn
    # A point's value is not missing when the same text is written at the
    # same spot for its twin: that one number is true of both points.
    for (category, series, number), times in list(missing.items()):
        if roles[(category, series, number)] != "point":
            continue
        spoken = sum(
            1
            for series_index, index in candidates(category, series, number)
            if any(
                drawn[(category, spec.series[twin].name, number)] > 0
                for twin, _ in twins_of(series_index, index)
            )
        )
        missing[(category, series, number)] = times - min(times, spoken)
    missing = +missing
    if drawn - Counter((c, s, n) for c, s, n, _ in want):
        found.append(f"drawn but not a value of the spec: {sorted(drawn - Counter(roles))}")
    left_out = [note for note in seen["warnings"] if "not written in the picture" in note]
    listed = Counter()
    if left_out:
        if len(left_out) != 1:
            found.append("more than one warning names values left out")
        body = left_out[0].removeprefix(f"Chart {name!r}: ")
        count, rest = body.split(" value(s) not written in the picture for lack of room: ", 1)
        listed = Counter(rest.split("; ")) if rest else Counter()
        if int(count) != sum(missing.values()):
            found.append(f"the warning counts {count}, {sum(missing.values())} are not drawn")
    named = Counter()
    for (category, series, number), times in missing.items():
        named[(f"{category}, {series}" if category else series) + f": {number}"] += times
    if "; ".join(sorted(named.elements())) != "; ".join(sorted(listed.elements())):
        if sorted("; ".join(sorted(named.elements()))) != sorted("; ".join(listed.elements())):
            found.append(f"the warning names {sorted(listed)}, not drawn are {sorted(named)}")

    # 2. A line or reference line left without its text is said to have no name.
    unnamed = [note for note in seen["warnings"] if "drawn without a name" in note]
    for role in ("line", "reference"):
        lost = sorted(s for (c, s, n) in missing if roles[(c, s, n)] == role)
        said = [
            note for note in unnamed if (" reference line(s) " in note) == (role == "reference")
        ]
        if bool(lost) != bool(said):
            found.append(f"{role}s without a name: {lost}; warnings: {said}")
        elif lost and not all(series in said[0] for series in lost):
            found.append(f"the warning does not name every {role}: {lost}")

    def point_problems(place: tuple[int, int], box, key: str) -> list[str]:
        """What is wrong with the text at `box` as the value of the point `place`."""
        wrong: list[str] = []
        own_series, own_index = place
        centre = ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2)
        gap = GAP_POINTS * point
        margin = MARGIN_POINTS * point + 0.5
        mine = lines_of[own_series]["pixels"][own_index]
        near = math.hypot(centre[0] - mine[0], centre[1] - mine[1])
        twins = twins_of(own_series, own_index)
        # on its own point: centred over or under it, at most 16 pt away
        if abs(centre[0] - mine[0]) > 1 or max(box[1] - mine[1], mine[1] - box[3], 0) > gap:
            wrong.append(f"{key}: is not over or under its own point")
        for other_series, index, x, y in points:
            if (other_series, index) == (own_series, own_index):
                continue  # its own point, by series and category
            away = math.hypot(centre[0] - x, centre[1] - y)
            if _from_box(box, x, y) == 0:
                wrong.append(f"{key}: lies over another plotted point")
            elif index != own_index:
                if away < near - margin:
                    wrong.append(f"{key}: another plotted point is nearer than its own")
            elif (other_series, index) not in twins and away < near + margin - 1:
                # another series' point of the same category is on the
                # text's own centre line, or on the very same spot: the
                # own point is nearer by 3 pt
                wrong.append(f"{key}: another series' point is as near as its own")
        # ... its own point is nearer than every bar by the margin (a point
        # and its category's bar are on one centre line), and nothing lies
        # between.
        clearly = MARGIN_POINTS * point - 0.5
        if any(_from_box(bar, *centre) < near + clearly for bar in every_bar):
            wrong.append(f"{key}: a bar is as near as its own point")
        band = _band(box, (mine[0], mine[1], mine[0], mine[1]))
        own_mark = (lines_of[own_series], own_index)
        twin_marks = [(lines_of[twin], index) for twin, index in twins]
        for thing in in_the_band(band, seen, [], own_mark, twin_marks) if band else []:
            wrong.append(f"{key}: {thing} lies between the value and its own point")
        return wrong

    written_points: list[tuple[int, int, str]] = []
    for item in values:
        category, series, number = item["gid"].split("\t")[1:4]
        role = roles.get((category, series, number), "?")
        box = item["box"]
        key = f"{category} / {series} / {number}"
        centre = ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2)
        # 3. The text is the number (with the name, for a line end or reference).
        want_text = {"line": f"{number}  {series}", "reference": f"{series} {number}"}.get(
            role, number
        )
        if "".join(item["text"].split()) != "".join(want_text.split()):
            found.append(f"{key}: the text is {item['text']!r}")
        # (a) inside the picture; (b) clear of every other text.
        if not _inside(canvas, box):
            found.append(f"{key}: leaves the picture")
        for other in texts:
            if other is not item and _overlap(box, other["box"]):
                found.append(f"{key}: overlaps the text {other['text']!r}")
        # (c) clear of every bar but the one it is written in ...
        own = bars.get((category, "range" if kind == "range" else series), [])
        if role == "total":
            own = [b for (c, _), boxes in bars.items() if c == category for b in boxes]
        written_in = [bar for bar in own if _inside(bar, box)]
        for bar in every_bar:
            if not any(bar is mine for mine in written_in) and _overlap(box, bar):
                found.append(f"{key}: lies over a bar it is not written in")
        # ... and of every line and point marker: no shared ink.
        x0, x1 = max(0, int(box[0]) - 1), int(math.ceil(box[2])) + 1
        y0, y1 = max(0, int(height - box[3]) - 1), int(math.ceil(height - box[1])) + 1
        shared = int((seen["value_ink"][y0:y1, x0:x1] & seen["line_ink"][y0:y1, x0:x1]).sum())
        if shared >= 3:
            found.append(f"{key}: a line or point marker covers {shared} pixels of the text")
        gap = GAP_POINTS * point
        margin = MARGIN_POINTS * point + 0.5
        if role == "line":
            # A line-end name is nearest to the end of its own line.
            # Whose end is whose is known by the series, never by where an
            # end lies; a name can stand for each series its text describes.
            last = len(spec.labels) - 1
            ends = {index: line["pixels"][last][1] for index, line in lines_of.items()}
            own_ends = [
                index
                for index in ends
                if spec.series[index].name == series and text_of(index, last) == number
            ]

            nearest_to_its_own = any(
                all(
                    abs(centre[1] - end) >= abs(centre[1] - ends[own]) - margin
                    for other, end in ends.items()
                    if other != own
                )
                for own in own_ends
            )
            if not nearest_to_its_own:
                found.append(f"{key}: another line ends nearer to this name than its own")
        if role in ("line", "reference", "?"):
            continue
        every_point = [(x, y) for _, _, x, y in points]
        if role == "point":
            # Its own point: a point its description stands for. Two points can
            # have the same category, series name and text (two series of one
            # name that meet); the value is judged as the value of each, and
            # stands if it stands as the value of one of them.
            described = candidates(category, series, number)
            if not described:
                found.append(f"{key}: has no plotted point")
                continue
            judged = [(point_problems(place, box, key), place) for place in described]
            wrong, place = min(judged, key=lambda item: len(item[0]))
            found += wrong
            written_points.append((*place, key))
            continue
        if not own:
            if kind == "range":  # a range of one value: its point
                index = spec.labels.index(category)
                px, py = seen["transforms"][0].transform(
                    (spec.low[index], len(spec.labels) - 1 - index)
                )
                if abs(centre[1] - py) > 1 or max(box[0] - px, px - box[2], 0) > gap:
                    found.append(f"{key}: is not beside its own point")
                band = _band(box, (px, py, px, py))
                for thing in in_the_band(band, seen, []) if band else []:
                    found.append(f"{key}: {thing} lies between the value and its own point")
            else:
                found.append(f"{key}: has no bar")
            continue
        if written_in:
            continue
        near = min(_from_box(bar, *centre) for bar in own)
        # (f) Nothing lies between a value written outside its mark and that mark.
        whole = (
            min(bar[0] for bar in own),
            min(bar[1] for bar in own),
            max(bar[2] for bar in own),
            max(bar[3] for bar in own),
        )
        band = _band(box, whole)
        for thing in in_the_band(band, seen, own, None) if band else []:
            found.append(f"{key}: {thing} lies between the value and its own bar")
        if kind in ("hbar", "range"):
            # Rows do not overlap: the centre within the row's height settles
            # whose value it is (see the module's note on horizontal charts).
            row = own[0]
            beside = max(box[0] - row[2], row[0] - box[2], 0)
            if not (row[1] - 1 <= centre[1] <= row[3] + 1) or beside > gap:
                found.append(f"{key}: is not beside its own row")
            continue
        if role == "segment":
            segment = own[0]
            index = spec.labels.index(category)
            left = seen["transforms"][0].transform((index - 0.5, 0))[0]
            right = seen["transforms"][0].transform((index + 0.5, 0))[0]
            if box[0] < left - 0.5 or box[2] > right + 0.5:
                found.append(f"{key}: leaves its category's column")
            if not (segment[1] - 1 <= centre[1] <= segment[3] + 1):
                found.append(f"{key}: its centre is not within its own segment's height")
            # Not as high as a whole other segment of its bar: two thin
            # segments level with one number cannot be told apart.
            for (other_category, other_series), boxes in bars.items():
                if other_category != category or other_series == series:
                    continue
                for other in boxes:
                    drawn_high = other[3] - other[1] >= 1.0
                    if drawn_high and other[1] >= box[1] and other[3] <= box[3]:
                        found.append(f"{key}: is level with the whole segment {other_series!r}")
            continue
        low, high = min(bar[0] for bar in own), max(bar[2] for bar in own)
        top, bottom = max(bar[3] for bar in own), min(bar[1] for bar in own)
        if not (low - 1 <= centre[0] <= high + 1) or max(box[1] - top, bottom - box[3], 0) > gap:
            found.append(f"{key}: is not over or under its own bar")
        rivals = [bar for bar in every_bar if not any(bar is mine for mine in own)]
        if any(_from_box(bar, *centre) < near - margin for bar in rivals):
            found.append(f"{key}: another bar is nearer than its own")
        clearly = MARGIN_POINTS * point - 0.5
        if any(math.hypot(centre[0] - x, centre[1] - y) < near + clearly for x, y in every_point):
            found.append(f"{key}: a plotted point is as near as its own bar")

    # One number for two points on one spot is written once.
    for first, (series_index, index, key) in enumerate(written_points):
        for other_series, other_index, _ in written_points[first + 1 :]:
            if (other_series, other_index) in twins_of(series_index, index):
                found.append(f"{key}: the same number is written twice at one spot")

    # (e) Values beside one stacked bar are not level with each other.
    beside: dict[str, list] = {}
    for item in values:
        category, series, number = item["gid"].split("\t")[1:4]
        if roles.get((category, series, number)) == "segment":
            own = bars.get((category, series), [])
            if not any(_inside(bar, item["box"]) for bar in own):
                beside.setdefault(category, []).append(item["box"])
    for category, boxes in beside.items():
        for index, first in enumerate(boxes):
            for second in boxes[index + 1 :]:
                if min(first[3], second[3]) - max(first[1], second[1]) > 0.5:
                    found.append(f"{category}: two values beside the bar are level with each other")

    # The value axis writes a number only where it is that number.
    labels = [item for item in texts if item["tick"] is not None and item["tick"][0] == 0]
    if kind == "line":
        shown = [item["text"] for item in labels]
        if len(set(shown)) != len(shown):
            found.append(f"the value axis repeats a number: {shown}")
        for item in labels:
            if item["text"] != written_number(item["tick"][1], spec.number_format):
                found.append(f"the value axis writes {item['text']!r} at {item['tick'][1]}")
            # The text read as a decimal is the tick's position (the shortest
            # decimal of the float it is drawn at), with no tolerance.
            digits = "".join(char for char in item["text"] if char in "0123456789.")
            position = item["tick"][1]
            read = Decimal(digits) * (-1 if item["text"].startswith("(") else 1)
            if read != Decimal(repr(float(position))):
                found.append(f"the value axis writes {item['text']!r} at {position!r}")

    # Texts no check removes: inside the picture and apart, or named in a warning.
    fixed = [item for item in texts if not item["gid"].startswith("value\t")]
    untidy = any(not _inside(canvas, item["box"]) for item in fixed) or any(
        _overlap(a["box"], b["box"]) for i, a in enumerate(fixed) for b in fixed[i + 1 :]
    )
    if untidy and not any("do not fit in the picture or overlap" in w for w in seen["warnings"]):
        found.append(
            "an axis name or legend entry leaves the picture or overlaps, without a warning"
        )
    # matplotlib's own warnings never reach the user outside the log format.
    if seen["python_warnings"]:
        found.append(f"Python warnings: {seen['python_warnings'][:2]}")
    return found


def check(monkeypatch, data: dict, width: float = 6.0) -> dict:
    spec = parse_chart_spec(yaml.safe_dump(data, allow_unicode=True, sort_keys=False))
    seen = draw(monkeypatch, spec, width)
    found = problems(spec, seen)
    assert not found, f"{width} in, {data}: " + "; ".join(found[:6])
    return seen


def shown(seen: dict) -> list[str]:
    return [item["text"] for item in seen["texts"] if item["gid"].startswith("value\t")]


# ═══════════════════════════════════════════════════════════════════════════════
# THE REVIEWS' DOCUMENTS
# ═══════════════════════════════════════════════════════════════════════════════

# Third independent review, [1]: each line's value sat nearer the other
# line's point, and the other line ran through it.
TWO_LINES = {
    "type": "bar_line",
    "labels": ["A", "B"],
    "right_number_format": ".2f",
    "series": [
        {"name": "Bar", "values": [10, 20]},
        {"name": "L1", "axis": "right", "values": [1, 2]},
        {"name": "L2", "axis": "right", "values": [1.15, 2.15]},
    ],
}
LINE_THROUGH_VALUE = {
    "type": "line",
    "labels": ["A", "B"],
    "accent_series": "Mid",
    "series": [
        {"name": "Low", "values": [70, 70]},
        {"name": "High", "values": [1000, 1000]},
        {"name": "Mid", "values": [100, 100]},
    ],
}
REFERENCE_THROUGH_VALUE = {
    "type": "range",
    "labels": ["A"],
    "low": [0],
    "high": [100],
    "reference": [{"name": "Ref", "value": 105}],
}
# Third independent review, [2]: the thin segment's value was centred on the
# next segment's height.
THIN_SEGMENT = {
    "type": "stacked_bar",
    "labels": ["A"],
    "totals": True,
    "series": [
        {"name": "Small", "values": [1000]},
        {"name": "Medium", "values": [100000]},
        {"name": "Large", "values": [1000000]},
    ],
}
# Self-review [1]: a line of another series ran through a bar's value, and
# two lines' values lay over each other's points.
LINE_OVER_BAR_VALUE = {
    "type": "bar_line",
    "title": "차입금과 레버리지",
    "labels": ["2025", "2026E", "2027E", "2028E"],
    "number_format": ",.1f",
    "right_number_format": ".1f",
    "series": [
        {"name": "단기", "values": [1740.1, 2477.4, 1389.8, 1969.3]},
        {"name": "장기", "values": [12047.0, 16454.0, 12357.0, 14117.0]},
        {"name": "레버리지", "values": [5.8, 4.6, 5.1, 2.5], "axis": "right"},
        {"name": "이자보상", "values": [5.4, 2.8, 3.6, 5.3], "axis": "right"},
    ],
}
CROSSING_LINES = {
    "type": "bar_line",
    "title": "매출과 배수",
    "labels": ["2023", "2024", "2025", "2026E", "2027E", "2028E"],
    "number_format": "percent",
    "right_number_format": "multiple",
    "series": [
        {"name": "매출", "values": [1632, 1565, 2045, 1702, 2324, 1374]},
        {"name": "배수 A", "values": [4.2, 1.7, 4.1, 3.4, 3.3, 5.4], "axis": "right"},
        {"name": "배수 B", "values": [2.4, 2.4, 2.1, 4.0, 6.3, 5.2], "axis": "right"},
    ],
}
OWN_LINE_OVER_VALUE = {
    "type": "line",
    "title": "스프레드",
    "labels": ["가", "나", "다", "라", "마"],
    "series": [
        {"name": "비교", "values": [389, 395, 398, 290, 356]},
        {"name": "당사", "values": [343.4, 341.3, 280.3, 294.5, 314.5]},
    ],
}
# Self-review [2]: 119 on the left and 5 on the right of the bar, at one height.
LEVEL_VALUES = {
    "type": "stacked_bar",
    "title": "조달 구성",
    "labels": ["2026E", "2027E", "2028E"],
    "totals": True,
    "series": [
        {"name": "기타", "values": [5, 5, 5]},
        {"name": "회사채", "values": [119, 121, 119]},
        {"name": "은행 차입", "values": [3386, 3406, 3369]},
    ],
}
REVIEW_CASES = {
    "two-lines": TWO_LINES,
    "line-through-value": LINE_THROUGH_VALUE,
    "reference-through-value": REFERENCE_THROUGH_VALUE,
    "thin-segment": THIN_SEGMENT,
    "line-over-bar-value": LINE_OVER_BAR_VALUE,
    "crossing-lines": CROSSING_LINES,
    "own-line-over-value": OWN_LINE_OVER_VALUE,
    "level-values": LEVEL_VALUES,
}


@pytest.mark.parametrize("name", REVIEW_CASES)
@pytest.mark.parametrize("width", [3.0, 3.5, 4.5, 6.0])
def test_the_reviews_charts_keep_the_rule(monkeypatch, name, width):
    check(monkeypatch, REVIEW_CASES[name], width)


def test_the_thin_segments_value_is_named_not_drawn_on_the_next_segment(monkeypatch):
    seen = check(monkeypatch, THIN_SEGMENT)
    assert "1,000" not in shown(seen)
    assert seen["warnings"] == [
        "Chart 'stacked_bar': 1 value(s) not written in the picture for lack of room: "
        "A, Small: 1,000"
    ]


def test_values_level_beside_one_bar_are_not_both_drawn(monkeypatch):
    seen = check(monkeypatch, LEVEL_VALUES)
    assert "5" not in shown(seen) and shown(seen).count("119") == 2
    assert "2026E, 기타: 5; 2027E, 기타: 5; 2028E, 기타: 5" in seen["warnings"][0]


# ═══════════════════════════════════════════════════════════════════════════════
# THE FOURTH REVIEW
# ═══════════════════════════════════════════════════════════════════════════════

# The coordinator's reading of the Word page (O1): at 6 inches 16,454.0 stood
# over its bar with another series' line and point between the two.
# `LINE_OVER_BAR_VALUE` above is that chart.


def test_a_bars_value_is_not_written_beyond_another_series_line(monkeypatch):
    seen = check(monkeypatch, LINE_OVER_BAR_VALUE)
    assert "16,454.0" not in shown(seen)
    left_out = [note for note in seen["warnings"] if "not written in the picture" in note]
    assert len(left_out) == 1 and "2026E, 장기: 16,454.0" in left_out[0]


def other_line_from_the_same_spot(band, own_line, point, twin=False) -> list[str]:
    other = {"pixels": [(60.0, 20.0), point, (130.0, 75.0)], "drawn": True, "radius": 4.0}
    seen = {"bars": [], "lines": [own_line, other]}
    return in_the_band(band, seen, [], (own_line, 1), [(other, 1)] if twin else [])


def test_the_band_measure_sees_what_lies_between_a_text_and_its_mark():
    """The measure itself, on hand-made geometry (it does not ask the engine)."""
    bar = (100.0, 0.0, 140.0, 50.0)
    text = (105.0, 70.0, 135.0, 82.0)
    band = _band(text, bar)
    assert band == (105.0, 50.0 + EDGE_PIXELS, 135.0, 70.0)

    def line(pixels, drawn=True, radius=0.0):
        return {"pixels": pixels, "drawn": drawn, "radius": radius}

    def between(lines, bars=()):
        seen = {"bars": [("bar\tA\tS", bar), *bars], "lines": lines}
        return in_the_band(band, seen, [bar])

    assert between([]) == []
    # a line across the band, a line only beside it, a line along the bar's own edge
    assert between([line([(60.0, 55.0), (180.0, 65.0)])]) == ["a line"]
    assert between([line([(60.0, 55.0), (104.0, 65.0)])]) == []
    assert between([line([(60.0, 50.5), (180.0, 50.5)])]) == []
    # a point marker in the band, with and without its line drawn
    assert between([line([(120.0, 60.0)], drawn=False, radius=4.0)]) == ["a point marker"]
    assert between([line([(60.0, 20.0), (120.0, 60.0)], radius=4.0)]) == [
        "a line",
        "a point marker",
    ]
    # another bar in the band; the text's own bar is not
    assert between([], bars=[("bar\tA\tT", (90.0, 55.0, 150.0, 60.0))]) == ["another bar"]
    # Of a point's own line only the two pieces that meet at the point, and
    # the point's own marker, are its own mark.
    point = (120.0, 40.0)
    above = _band(text, (*point, *point))
    own_line = line([(60.0, 60.0), point, (180.0, 60.0)], radius=4.0)
    seen = {"bars": [], "lines": [own_line, line([(60.0, 60.0), (180.0, 60.0)])]}
    assert in_the_band(above, seen, [], (own_line, 1)) == ["a line"]
    assert in_the_band(above, {"bars": [], "lines": [own_line]}, [], (own_line, 1)) == []
    # Without the record of whose point it is, nothing is the value's own.
    assert in_the_band(above, {"bars": [], "lines": [own_line]}, []) == [
        "a line",
        "a point marker",
    ]
    # A later piece of the own line that comes back through the band counts.
    back = line([(60.0, 60.0), point, (180.0, 45.0), (100.0, 62.0)], radius=4.0)
    assert in_the_band(above, {"bars": [], "lines": [back]}, [], (back, 1)) == ["a line"]
    # Fifth review: another series' line that leaves from the very same spot
    # is between the value and its point, and so is that series' marker there,
    assert other_line_from_the_same_spot(above, own_line, point) == ["a line", "a point marker"]
    # unless the number is true of that point too (a twin): then only its line.
    assert other_line_from_the_same_spot(above, own_line, point, twin=True) == ["a line"]
    # text and mark touching: there is no band
    assert _band((105.0, 50.5, 135.0, 62.0), bar) is None


# Self-review: the value 3.6% stood under its point and as near to the bar
# below it (4.5 pt) as to the point (4 pt); a point and its category's bar are
# on one centre line.
POINT_VALUE_OVER_A_BAR = {
    "type": "bar_line",
    "labels": ["2026E", "2027E", "2028E"],
    "right_number_format": "percent",
    "series": [
        {"name": "회사채", "values": [5.0, 4.0, 3.0]},
        {"name": "Alpha", "values": [3.6, 6.3, 2.4], "axis": "right"},
    ],
}
# Self-review: the 2 beside bar E was as high as both thin segments (2 and 3).
# That holds for a text of about 20.2 px or more (Malgun Gothic: 21 px). With
# a lower text (18 px in Gulim and Batang) the segment of 3 reaches 1.6 to
# 1.7 px beyond the text, so the 2 may be written there; CI on Ubuntu, which
# draws with NanumGothic, writes it too. The chart keeps the rule either way
# (`FOURTH_CASES`); which values it writes depends on the font.
THIN_SEGMENTS_LEVEL = {
    "type": "stacked_bar",
    "totals": True,
    "labels": ["A", "B", "C", "D", "E", "F"],
    "series": [
        {"name": "영업", "values": [120, 90, 60, 4, 75, 130]},
        {"name": "투자", "values": [8, 30, 70, 3, 2, 9]},
        {"name": "재무", "values": [10, 4, 25, 5, 3, 1]},
    ],
}
# The same stacks with two segments of 1 in bar E (under 3 px each): a value
# beside the bar has its centre within its own segment, so its text, of any
# font's height, reaches over the whole other thin segment.
THINNER_SEGMENTS = {
    "type": "stacked_bar",
    "totals": True,
    "labels": ["A", "B", "C", "D", "E", "F"],
    "series": [
        {"name": "영업", "values": [120, 90, 60, 4, 75, 130]},
        {"name": "투자", "values": [8, 30, 70, 3, 1, 9]},
        {"name": "재무", "values": [10, 4, 25, 5, 1, 1]},
    ],
}
# Three lines end within a few points of each other, under a reference line.
LINE_ENDS_CLOSE = {
    "type": "line",
    "labels": ["1월", "2월", "3월", "4월", "5월"],
    "accent_series": "AA-",
    "series": [
        {"name": "국고 3년", "values": [3.1, 3.2, 3.15, 3.3, 3.42]},
        {"name": "AA-", "values": [3.9, 3.95, 3.8, 3.7, 3.46]},
        {"name": "A+", "values": [4.4, 4.5, 4.2, 3.9, 3.5]},
    ],
    "reference": [{"name": "기준금리", "value": 3.5}],
}
# A reference line between a range's end and the value written beside it.
REFERENCE_BESIDE_A_RANGE = {
    "type": "range",
    "labels": ["가상전자", "한빛물산", "새솔", "동녘"],
    "low": [3.2, 3.6, 4.1, 3.9],
    "high": [3.8, 4.4, 4.2, 5.0],
    "reference": [{"name": "제시 금리", "value": 3.85}],
}
# Fourth independent review, [3]: the value beside row A is nearer to the
# longer bar of row B than to its own. Approved exception: in a horizontal
# chart the row decides (see the module's note).
ROW_VALUE_NEAR_A_LONGER_ROW = {
    "type": "hbar",
    "labels": ["A", "B", "C"],
    "number_format": ",.2f",
    "series": [{"name": "S", "values": [100000000, 1000000000, 1000000000]}],
}
# Reading the Word pages: at 2027E the grey 4.2x stood halfway between its own
# point (under it) and the other line's point (over it).
TWO_LINES_ONE_CATEGORY = {
    "type": "bar_line",
    "labels": ["2021", "2022", "2023", "2024", "2025", "2026E", "2027E", "2028E"],
    "right_number_format": "multiple",
    "series": [
        {
            "name": "가상전자",
            "values": [210.136, 271.186, 343.628, 349.963, 395.407, 331.555, 262.73, 239.962],
        },
        {"name": "Alpha", "values": [1.8, 3.5, 1.5, 4.8, 5.8, 6.0, 6.2, 4.6], "axis": "right"},
        {"name": "새솔", "values": [5.2, 4.2, 4.7, 4.2, 3.0, 2.1, 4.2, 3.0], "axis": "right"},
    ],
}
FOURTH_CASES = {
    "point-value-over-a-bar": POINT_VALUE_OVER_A_BAR,
    "two-lines-one-category": TWO_LINES_ONE_CATEGORY,
    "thin-segments-level": THIN_SEGMENTS_LEVEL,
    "line-ends-close": LINE_ENDS_CLOSE,
    "reference-beside-a-range": REFERENCE_BESIDE_A_RANGE,
    "row-value-near-a-longer-row": ROW_VALUE_NEAR_A_LONGER_ROW,
}


@pytest.mark.parametrize("name", FOURTH_CASES)
@pytest.mark.parametrize("width", [3.0, 3.5, 4.5, 6.0])
def test_the_fourth_reviews_charts_keep_the_rule(monkeypatch, name, width):
    check(monkeypatch, FOURTH_CASES[name], width)


def test_a_value_beside_a_row_is_judged_by_its_row(monkeypatch):
    """The approved exception is measured: the row's height and 16 pt beside it."""
    for width in (3.0, 3.5, 4.5, 6.0):
        seen = check(monkeypatch, ROW_VALUE_NEAR_A_LONGER_ROW, width)
        assert sorted(shown(seen)) == ["1,000,000,000.00", "1,000,000,000.00", "100,000,000.00"]
        assert seen["warnings"] == []
        # ... and the measure is the row, not the nearest bar: the value of A
        # is nearer to B's bar than to its own, as the review found.
        rows = {gid.split("\t")[1]: box for gid, box in seen["bars"]}
        value = next(item["box"] for item in seen["texts"] if item["text"] == "100,000,000.00")
        centre = ((value[0] + value[2]) / 2, (value[1] + value[3]) / 2)
        assert _from_box(rows["B"], *centre) < _from_box(rows["A"], *centre)
        assert rows["A"][1] <= centre[1] <= rows["A"][3]
        assert value[0] - rows["A"][2] <= GAP_POINTS * seen["dpi"] / 72


def test_a_value_is_not_written_beyond_a_reference_line_beside_its_range(monkeypatch):
    seen = check(monkeypatch, REFERENCE_BESIDE_A_RANGE)
    assert "3.8" not in shown(seen) and "3.9" not in shown(seen)
    assert seen["warnings"] == [
        "Chart 'range': 2 value(s) not written in the picture for lack of room: "
        "가상전자, high: 3.8; 동녘, low: 3.9"
    ]


def test_a_points_value_is_clearly_nearer_to_its_point_than_to_a_bar(monkeypatch):
    seen = check(monkeypatch, POINT_VALUE_OVER_A_BAR)
    assert seen["warnings"] == []
    value = next(item["box"] for item in seen["texts"] if item["text"] == "3.6%")
    bar = next(box for gid, box in seen["bars"] if gid == "bar\t2026E\t회사채")
    point = seen["transforms"][1].transform((0, 3.6))
    # It was under its point and over the bar; now it is over its point.
    assert value[1] > point[1] > bar[3]


def test_a_points_value_is_clearly_nearer_to_its_point_than_to_another_series(monkeypatch):
    seen = check(monkeypatch, TWO_LINES_ONE_CATEGORY)
    lines = {"Alpha": [1.8, 3.5, 1.5, 4.8, 5.8, 6.0, 6.2, 4.6]}
    lines["새솔"] = [5.2, 4.2, 4.7, 4.2, 3.0, 2.1, 4.2, 3.0]
    labels = TWO_LINES_ONE_CATEGORY["labels"]
    clearly = MARGIN_POINTS * seen["dpi"] / 72 - 1
    measured = 0
    for item in seen["texts"]:
        parts = item["gid"].split("\t")
        if parts[0] != "value" or parts[2] not in lines:
            continue
        index, box = labels.index(parts[1]), item["box"]
        centre = ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2)
        away = {
            name: math.hypot(*(centre - seen["transforms"][1].transform((index, values[index]))))
            for name, values in lines.items()
        }
        other = next(name for name in lines if name != parts[2])
        assert away[other] >= away[parts[2]] + clearly, item["gid"]
        measured += 1
    assert measured >= 8  # most of the sixteen values still have a place


def test_a_value_as_high_as_a_whole_other_segment_is_not_drawn(monkeypatch):
    """Measured in the picture, so that it holds whatever font draws the values.

    The premise is checked first: wherever a value's centre lies within its
    own thin segment, the other thin segment of bar E (1 px or more high) is
    wholly within the text's height, with 2 px to spare. Then neither value
    is drawn, and both are named.
    """
    seen = check(monkeypatch, THINNER_SEGMENTS)
    bars = {tuple(gid.split("\t")[1:3]): box for gid, box in seen["bars"]}
    drawn = {item["gid"]: item["box"] for item in seen["texts"]}
    # A value's text height in this picture, from one that is always drawn.
    inside = drawn["value\tE\t영업\t75"]
    height = inside[3] - inside[1]
    for own, other in (("투자", "재무"), ("재무", "투자")):
        segment, whole = bars[("E", own)], bars[("E", other)]
        assert whole[3] - whole[1] >= 1.0
        for centre in (segment[1], segment[3]):  # the lowest and the highest level place
            assert centre - height / 2 <= whole[1] - 2
            assert whole[3] + 2 <= centre + height / 2
        assert f"value\tE\t{own}\t1" not in drawn
    left_out = [note for note in seen["warnings"] if "not written in the picture" in note]
    assert len(left_out) == 1 and "E, 투자: 1" in left_out[0] and "E, 재무: 1" in left_out[0]


# ═══════════════════════════════════════════════════════════════════════════════
# THE FIFTH REVIEW
# ═══════════════════════════════════════════════════════════════════════════════

# Fifth independent review, [4]: at A the two lines start 0.01 apart, on one
# spot of the picture (0.2 px). `A, Own: 100.00` was written at that spot, as
# near to the other series' point as to its own, with the other series' line
# between; the engine's check and this module's measure both let it pass,
# because each took whatever lay on the value's own point for the value's own.
SHARED_SPOT = {
    "type": "line",
    "title": "Shared point",
    "number_format": ",.2f",
    "labels": ["A", "B", "C"],
    "series": [
        {"name": "Other", "values": [100.01, 98, 120]},
        {"name": "Own", "values": [100, 110, 120]},
    ],
}
# Two index lines that both start at 100: the one number is true of both points.
SHARED_START = {
    "type": "line",
    "title": "Index",
    "number_format": ",.2f",
    "labels": ["A", "B", "C"],
    "series": [
        {"name": "Other", "values": [100, 98, 120]},
        {"name": "Own", "values": [100, 110, 125]},
    ],
}
# Two lines that both write their values: at A and C they share a spot and a text.
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
# The same spots with different texts (5.00 and 5.01): neither can be told apart.
NEAR_POINTS = {
    "type": "bar_line",
    "title": "Near",
    "right_number_format": ",.2f",
    "labels": ["A", "B", "C"],
    "series": [
        {"name": "Bars", "values": [10, 12, 14]},
        {"name": "R1", "axis": "right", "values": [5, 6, 9]},
        {"name": "R2", "axis": "right", "values": [5.01, 7.5, 9.01]},
    ],
}
FIFTH_CASES = {
    "shared-spot": SHARED_SPOT,
    "shared-start": SHARED_START,
    "twin-points": TWIN_POINTS,
    "near-points": NEAR_POINTS,
}
WIDTHS = [3.0, 3.5, 4.5, 6.0]


@pytest.mark.parametrize("name", FIFTH_CASES)
@pytest.mark.parametrize("width", WIDTHS)
def test_the_fifth_reviews_charts_keep_the_rule(monkeypatch, name, width):
    check(monkeypatch, FIFTH_CASES[name], width)


def _spec(data: dict):
    return parse_chart_spec(yaml.safe_dump(data, allow_unicode=True, sort_keys=False))


@pytest.mark.parametrize("width", WIDTHS)
def test_a_value_on_a_spot_it_shares_with_another_series_is_named_not_drawn(monkeypatch, width):
    seen = check(monkeypatch, SHARED_SPOT, width)
    lines = series_lines(_spec(SHARED_SPOT), seen)
    other, own = lines[0]["pixels"][0], lines[1]["pixels"][0]
    # The review's geometry: the two points of A are one spot of the picture.
    assert math.hypot(other[0] - own[0], other[1] - own[1]) < 1.0
    assert "100.00" not in shown(seen)
    left_out = [note for note in seen["warnings"] if "not written in the picture" in note]
    assert len(left_out) == 1 and "A, Own: 100.00" in left_out[0]


def test_the_measure_knows_a_values_own_point_by_its_series_not_by_its_place(monkeypatch):
    """The measure itself: the review's value, put back over the shared spot, is reported.

    The engine is not asked: a value is added by hand to what the picture
    holds, over the spot the two points of A share, as the reviewed engine
    drew it. The old measure left out every point and line on that spot.
    """
    spec = _spec(SHARED_SPOT)
    seen = draw(monkeypatch, spec, 6.0)
    assert problems(spec, seen) == []
    own = series_lines(spec, seen)[1]["pixels"][0]
    point = seen["dpi"] / 72
    for low, high in ((4 * point, 13 * point), (-13 * point, -4 * point)):  # over, under
        put_back = dict(seen)
        put_back["texts"] = [
            *seen["texts"],
            {
                "text": "100.00",
                "raw": "100.00",
                "gid": "value\tA\tOwn\t100.00",
                "box": (own[0] - 15, own[1] + low, own[0] + 15, own[1] + high),
                "tick": None,
                "artist": None,
            },
        ]
        found = problems(spec, put_back)
        assert "A / Own / 100.00: another series' point is as near as its own" in found
        assert "A / Own / 100.00: a point marker lies between the value and its own point" in found


def test_one_number_is_written_once_for_two_points_on_one_spot(monkeypatch):
    """Same spot, same text: written once, and the value not written is not named."""
    for width in WIDTHS:
        seen = check(monkeypatch, TWIN_POINTS, width)
        assert shown(seen).count("5") == 1 and shown(seen).count("9") == 1
        assert sorted(shown(seen)) == sorted(["10", "12", "14", "5", "6", "7.5", "9"])
        assert seen["warnings"] == []
        # The one number stands over the spot both points share.
        lines = series_lines(_spec(TWIN_POINTS), seen)
        assert lines[1]["pixels"][0] == lines[2]["pixels"][0]
    seen = check(monkeypatch, SHARED_START)
    assert shown(seen).count("100.00") == 1 and seen["warnings"] == []


def test_two_values_that_differ_on_one_spot_are_named_not_drawn(monkeypatch):
    for width in WIDTHS:
        seen = check(monkeypatch, NEAR_POINTS, width)
        assert sorted(shown(seen)) == sorted(["10", "12", "14", "6.00", "7.50"])
        assert seen["warnings"] == [
            "Chart 'Near': 4 value(s) not written in the picture for lack of room: "
            "A, R1: 5.00; C, R1: 9.00; A, R2: 5.01; C, R2: 9.01"
        ]


def test_a_series_line_is_found_by_its_data_and_the_engines_record_agrees(monkeypatch):
    """Whose line is whose: by the values it plots, checked against the line's gid."""
    for data in (SHARED_SPOT, TWIN_POINTS, CROSSING_LINES, LINE_ENDS_CLOSE):
        spec = _spec(data)
        seen = draw(monkeypatch, spec, 6.0)
        lines = series_lines(spec, seen)
        plotted = [
            index
            for index, series in enumerate(spec.series)
            if spec.chart_type == "line" or series.axis == "right"
        ]
        assert sorted(lines) == plotted
        for index, line in lines.items():
            assert line["ydata"] == [float(value) for value in spec.series[index].values]
            assert line["gid"] == f"line\t{index}"
    # A line marked as another series' is reported.
    spec = _spec(SHARED_SPOT)
    seen = draw(monkeypatch, spec, 6.0)
    series_lines(spec, seen)[0]["gid"] = "line\t1"
    assert "the line of series 0 is marked 'line\\t1'" in problems(spec, seen)


# Two series of one name that start on one spot: the value's description
# (category, series name, text) stands for both points.
SAME_NAME_TWINS = {
    "type": "line",
    "labels": ["A", "B", "C"],
    "series": [
        {"name": "S", "values": [10, 12, 14]},
        {"name": "S", "values": [10, 9, 20]},
    ],
}


@pytest.mark.parametrize("width", WIDTHS)
def test_a_value_is_judged_as_each_point_its_description_stands_for(monkeypatch, width):
    seen = check(monkeypatch, SAME_NAME_TWINS, width)
    assert shown(seen).count("10") <= 1


def _line(values, number_format=None):
    data = {
        "type": "line",
        "labels": ["A", "B", "C"][: len(values)],
        "series": [{"name": "L", "values": values}],
    }
    return {**data, "number_format": number_format} if number_format else data


# Fourth independent review, [1]: a tick was kept when its text was within
# 1e-9 of its position as a share of the value, so 1,000,000,000.5 was written
# 1,000,000,001 and six ticks from 1e-10 to 2e-10 were all written 0.0.
TICKS = [
    (
        _line([1000000000, 1000000002], ",.0f"),
        ["1,000,000,000", "1,000,000,001", "1,000,000,002"],
    ),
    (_line([1.0e-10, 2.0e-10]), []),
    # ... and sizes between the two.
    (_line([1.0e-5, 3.0e-5]), []),
    (_line([0.001, 0.003], ".2f"), []),
    (_line([0.1, 0.2, 0.3]), ["0.1", "0.2", "0.3"]),
    (
        _line([1000.5, 1002.5], ",.1f"),
        ["1,000.5", "1,001.0", "1,001.5", "1,002.0", "1,002.5"],
    ),
    (
        _line([100000, 100004], ",.0f"),
        ["100,000", "100,001", "100,002", "100,003", "100,004"],
    ),
    (_line([10000000, 10000001], ",.0f"), ["10,000,000", "10,000,001"]),
    (
        _line([1000000000.5, 1000000002.5], ",.1f"),
        [
            "1,000,000,000.5",
            "1,000,000,001.0",
            "1,000,000,001.5",
            "1,000,000,002.0",
            "1,000,000,002.5",
        ],
    ),
    (
        _line([123456789012, 123456789016], ",.0f"),
        [
            "123,456,789,012",
            "123,456,789,013",
            "123,456,789,014",
            "123,456,789,015",
            "123,456,789,016",
        ],
    ),
]


@pytest.mark.parametrize("data, ticks", TICKS)
def test_a_tick_is_written_only_where_its_text_is_its_position(monkeypatch, data, ticks):
    seen = check(monkeypatch, data)
    labels = [item for item in seen["texts"] if item["tick"] is not None and item["tick"][0] == 0]
    assert [item["text"] for item in labels] == ticks
    for item in labels:  # as decimals, with no tolerance (also in `problems`)
        digits = "".join(char for char in item["text"] if char in "0123456789.")
        assert Decimal(digits) == Decimal(repr(float(item["tick"][1])))
    # The points keep their values whatever the axis writes.
    assert len(shown(seen)) == 2 if len(data["labels"]) == 2 else len(shown(seen)) == 3


# ═══════════════════════════════════════════════════════════════════════════════
# NUMBERS AND NAMES
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_value_axis_tick_is_written_only_where_it_is_that_number(monkeypatch):
    # Self-review [3]: ticks at 1, 1.5, 2, 2.5, 3 were written 1, 2, 2, 2, 3.
    whole = {
        "type": "line",
        "labels": ["2023", "2024", "2025"],
        "number_format": ",.0f",
        "series": [{"name": "건수", "values": [1, 2, 3]}],
    }
    seen = check(monkeypatch, whole)
    ticks = [item["text"] for item in seen["texts"] if item["tick"] is not None]
    assert ticks == ["1", "2", "3"]
    # No tick is its own number: the axis writes none, the points keep theirs.
    tiny = {
        **whole,
        "number_format": "percent",
        "series": [{"name": "금리", "values": [0.012, 0.018, 0.02]}],
    }
    seen = check(monkeypatch, tiny)
    assert [item["text"] for item in seen["texts"] if item["tick"] is not None] in ([], ["0%"])
    flat = {**whole, "number_format": None, "series": [{"name": "S", "values": [0, 0, 0]}]}
    flat.pop("number_format")
    check(monkeypatch, flat)


@pytest.mark.parametrize(
    "values, number_format, total",
    [
        ([0.3, -0.1, -0.2], None, "0"),
        ([0.1, 0.2, -0.3], None, "0"),
        ([0.3, -0.1, -0.2], ",.2f", "0.00"),
        ([100.1, -50.05, -50.05], ",.1f", "0.0"),
        ([0.1, 0.2, 0.7], None, "1"),
        ([77.1, -103.3, 26.2], None, "0"),
        ([1.5, -3.0], None, "(1.5)"),
    ],
)
def test_a_waterfall_total_is_the_sum_of_the_written_numbers(
    monkeypatch, values, number_format, total
):
    # Self-review [6]: `[0.3, -0.1, -0.2]` totalled `(0.0)`.
    data = {
        "type": "waterfall",
        "labels": [f"s{index}" for index in range(len(values))],
        "series": [{"name": "Flow", "values": values}],
    }
    if number_format:
        data["number_format"] = number_format
    seen = check(monkeypatch, data)
    totals = [
        item["text"]
        for item in seen["texts"]
        if item["gid"].startswith(f"value\t{STYLE.CHART_TOTAL_LABEL}\t")
    ]
    assert totals == [total]


def test_a_value_that_rounds_to_zero_is_not_written_as_negative(monkeypatch):
    data = {
        "type": "bar",
        "labels": ["A", "B", "C"],
        "series": [{"name": "S", "values": [-0.04, -0.4, 0.04]}],
    }
    assert shown(check(monkeypatch, data)) == ["0.0", "(0.4)", "0.0"]
    data["number_format"] = ",.0f"
    assert shown(check(monkeypatch, data)) == ["0", "0", "0"]


def test_a_dollar_sign_in_a_name_is_drawn_as_written(monkeypatch):
    # Self-review [4]: two `$` in one name were read as mathematical notation.
    from matplotlib import cbook

    names = ["매출($) / 이익($)", "A$/US$", "Cost $_$ net"]
    charts = [
        {
            "type": "bar",
            "labels": names,
            "series": [
                {"name": "USD ($m) vs KRW ($m)", "values": [1, 2, 3]},
                {"name": "S", "values": [2, 3, 1]},
            ],
        },
        {
            "type": "line",
            "labels": ["a", "b"],
            "series": [{"name": "Price ($) and cost ($)", "values": [1, 2]}],
            "reference": [{"name": "Cap $5-$10", "value": 1.5}],
        },
        {"type": "hbar", "labels": names, "series": [{"name": "s", "values": [1, 2, 3]}]},
        {
            "type": "waterfall",
            "labels": ["a", "b"],
            "total_label": "Net $_$ total",
            "series": [{"name": "s", "values": [1, 2]}],
        },
        {
            "type": "stacked_bar",
            "labels": ["a"],
            "series": [{"name": "A$ and US$", "values": [1]}, {"name": "T", "values": [2]}],
        },
    ]
    for data in charts:
        seen = check(monkeypatch, data)
        assert not [item["raw"] for item in seen["texts"] if cbook.is_math_text(item["raw"])]
        drawn = " ".join(" ".join(item["text"].split()) for item in seen["texts"])
        wanted = [*data["labels"], *(s["name"] for s in data["series"])]
        wanted += [data["total_label"]] if "total_label" in data else []
        if data["type"] in ("bar", "stacked_bar"):  # names of series are in the legend
            wanted = [*data["labels"], *(s["name"] for s in data["series"])]
        if data["type"] in ("hbar", "waterfall"):  # a single series has no legend
            wanted = [*data["labels"], *([data["total_label"]] if "total_label" in data else [])]
        for name in wanted:
            assert "".join(name.split()) in "".join(drawn.split()), (name, drawn)


def test_a_horizontal_bar_chart_of_zeros_writes_its_zeros(monkeypatch):
    # Self-review [9]: the place beside a bar was a share of an empty span.
    data = {"type": "hbar", "labels": ["a", "b"], "series": [{"name": "S", "values": [0, 0]}]}
    seen = check(monkeypatch, data)
    assert shown(seen) == ["0", "0"] and seen["warnings"] == []


@pytest.mark.parametrize(
    "data",
    [
        {"type": "bar", "labels": ["A", "B"], "series": [{"name": "S", "values": [1e308, -1e308]}]},
        {
            "type": "waterfall",
            "labels": ["A", "B"],
            "series": [{"name": "S", "values": [1.7e308, 1.0]}],
        },
        {"type": "line", "labels": ["A", "B"], "series": [{"name": "S", "values": [1e300, 2e300]}]},
        {"type": "bar", "labels": ["A", "B"], "series": [{"name": "S", "values": [1e300, 5.0]}]},
    ],
    ids=["bar-both-ends", "waterfall-largest", "line-1e300", "bar-1e300"],
)
def test_the_largest_finite_values_are_drawn_not_rejected(monkeypatch, data):
    # Self-review [7]: the plot's limits overflowed where the classic style drew.
    seen = check(monkeypatch, data)
    assert seen["png"].startswith(b"\x89PNG")


def test_a_long_line_name_is_broken_and_keeps_the_plot(monkeypatch):
    # Self-review [5]: one long name squeezed the plot to a seventh of the picture.
    long_name = (
        "가상전자 무보증 공모사채 제12-2회 선순위 고정금리부 채권 (만기 2031년 3월, 분기 이자 지급)"
    )
    data = {
        "type": "line",
        "labels": ["2024", "2025", "2026E"],
        "series": [
            {"name": long_name, "values": [3.1, 3.4, 3.2]},
            {"name": "국고채", "values": [2.8, 2.9, 3.0]},
            {"name": "회사채 AA-", "values": [3.6, 3.8, 3.7]},
        ],
    }
    for width in (3.0, 6.0):
        seen = check(monkeypatch, data, width)
        end = next(item for item in seen["texts"] if long_name[:6] in item["text"])
        assert "\n" in end["text"]
        assert end["box"][2] - end["box"][0] <= 0.45 * seen["size"][0]
        assert not [w for w in seen["warnings"] if "without a name" in w]


def test_a_line_left_without_its_name_is_said_so(monkeypatch):
    data = {
        "type": "line",
        "labels": ["A", "B"],
        "series": [{"name": f"S{index:02}", "values": [0, 1]} for index in range(70)],
    }
    seen = check(monkeypatch, data)
    said = [w for w in seen["warnings"] if "line(s) drawn without a name" in w]
    assert len(said) == 1


def test_no_python_warning_leaves_the_engine_at_a_narrow_width(monkeypatch):
    # Self-review [8]: matplotlib's "Tight layout not applied" reached the CLI.
    data = {
        "type": "bar",
        "labels": [f"가상회사 {index}" for index in range(13)],
        "series": [{"name": f"계열 {index}", "values": [150] * 13} for index in range(4)],
    }
    for width in (2.27, 3.0):
        assert check(monkeypatch, data, width)["python_warnings"] == []


# ═══════════════════════════════════════════════════════════════════════════════
# ROUNDING
# ═══════════════════════════════════════════════════════════════════════════════

# Owner decision 2026-11-14: a standard chart rounds as a spreadsheet does, a
# 5 away from zero, on the decimal the spec wrote. The texts are written out
# here by hand; `old` is what the same value gave before (ties to even on the
# binary float), which the classic style keeps.
ROUNDED = [
    (4.125, ",.2f", "4.13", "4.12"),
    (1500.5, ",.0f", "1,501", "1,500"),
    (-4.125, ",.2f", "(4.13)", "(4.12)"),
    (2.5, ",.0f", "3", "2"),
    (0.125, ".2f", "0.13", "0.12"),
    (2.675, ".2f", "2.68", "2.67"),
    (0.25, None, "0.3", "0.2"),
    (-0.25, "percent", "(0.3%)", "(0.2%)"),
    (1.25, "multiple", "1.3x", "1.2x"),
    (12.45, "bps", "12.5 bps", "12.4 bps"),
    (0.5, ",.0f", "1", "0"),
    (-0.5, ",.0f", "(1)", "(0)"),
    # A negative amount that rounds to zero is not written as negative.
    (-0.04, None, "0.0", "(0.0)"),
    (-0.004, ",.2f", "0.00", "(0.00)"),
    (-1.0e-7, None, "0.0", "(0.0)"),
    (-0.0, None, "0", "0"),
    # Numbers the YAML reads in exponent form are written out in digits.
    (1.0e21, None, "1,000,000,000,000,000,000,000", "1,000,000,000,000,000,000,000"),
    (1.0e-7, "f", "0.000000", "0.000000"),
    (9007199254740993, None, "9,007,199,254,740,992", "9,007,199,254,740,992"),
]


@pytest.mark.parametrize("value, number_format, text, old", ROUNDED)
def test_a_standard_chart_rounds_a_five_away_from_zero(
    monkeypatch, value, number_format, text, old
):
    from chart_renderer import format_value

    data = {"type": "bar", "labels": ["A", "B"], "series": [{"name": "S", "values": [value, 1]}]}
    if number_format:
        data["number_format"] = number_format
    seen = check(monkeypatch, data)
    assert shown(seen)[0] == text
    # The classic style's formatter is not touched: its pictures and XML are
    # pinned whole by the classic corpus snapshots (`tests/test_chart_ib_style.py`).
    assert format_value(float(value), number_format) == old


@pytest.mark.parametrize(
    "kind, values, number_format, written, total",
    [
        # 1.005 + 1.01 is 2.015 as written; the binary sum is 2.0149999999999997.
        ("waterfall", [1.005, 1.01], ".2f", ["1.01", "1.01"], "2.02"),
        ("waterfall", [2.5, -5.0], ",.0f", ["3", "(5)"], "(3)"),
        ("waterfall", [1000.25, 500.2], ",.1f", ["1,000.3", "500.2"], "1,500.5"),
        ("waterfall", [0.05, -0.1], None, ["0.1", "(0.1)"], "(0.1)"),
        ("stacked_bar", [1.25, 1.25], ",.0f", ["1", "1"], "3"),
        ("stacked_bar", [0.15, 0.2], None, ["0.2", "0.2"], "0.4"),
        ("stacked_bar", [1.005, 1.01], ".2f", ["1.01", "1.01"], "2.02"),
    ],
)
def test_a_sum_is_added_exactly_and_rounded_once(
    monkeypatch, kind, values, number_format, written, total
):
    data: dict = {"type": kind, "number_format": number_format}
    if kind == "waterfall":
        data |= {"labels": ["a", "b"], "series": [{"name": "Flow", "values": values}]}
    else:
        data |= {
            "labels": ["a"],
            "totals": True,
            "series": [
                {"name": f"S{index}", "values": [value]} for index, value in enumerate(values)
            ],
        }
    if number_format is None:
        data.pop("number_format")
    seen = check(monkeypatch, data)
    assert shown(seen) == [*written, total]


def test_line_ends_reference_lines_and_ticks_round_the_same_way(monkeypatch):
    data = {
        "type": "line",
        "labels": ["2024", "2025", "2026E"],
        "number_format": ",.2f",
        "series": [{"name": "Spread", "values": [4.125, 2.675, 0.125]}],
        "reference": [{"name": "Cap", "value": 1.005}],
    }
    seen = check(monkeypatch, data)
    texts = [" ".join(item["text"].split()) for item in seen["texts"]]
    assert {"4.13", "2.68", "0.13 Spread", "Cap 1.01"} <= set(texts)
    whole = {**data, "number_format": ",.0f", "series": [{"name": "Count", "values": [1, 2, 3]}]}
    whole.pop("reference")
    seen = check(monkeypatch, whole)
    # Ticks at 1.5 and 2.5 would now be written 2 and 3: they are not kept.
    assert [item["text"] for item in seen["texts"] if item["tick"] is not None] == ["1", "2", "3"]


# ═══════════════════════════════════════════════════════════════════════════════
# MANY CHARTS
# ═══════════════════════════════════════════════════════════════════════════════

SHORT = [
    "가상전자",
    "한빛물산",
    "새솔",
    "동녘",
    "Alpha",
    "Beta",
    "Gamma",
    "Delta",
    "회사채",
    "은행 차입",
]
LONG = [
    "Northfield Infrastructure Senior Secured Notes 2031",
    "가상전자 무보증 공모사채 제12-2회 (선순위)",
]
YEARS = ["2021", "2022", "2023", "2024", "2025", "2026E", "2027E", "2028E"]
FORMATS = [None, None, "number", "percent", "bps", "multiple", ",.0f", ",.2f", ".1f"]
PROFILES = ["smooth", "trend", "close", "spread", "equal", "zeros", "signs"]
TYPES = ["bar", "line", "waterfall", "stacked_bar", "bar_line", "range", "hbar"]


def _names(rng: random.Random, count: int) -> list[str]:
    names: list[str] = []
    while len(names) < count:
        base = rng.choice(LONG) if rng.random() < 0.06 else rng.choice(SHORT)
        name = base if rng.random() < 0.5 else f"{base} {len(names) + 1}"
        if name not in names:
            names.append(name)
    return names


def _numbers(rng: random.Random, count: int, profile: str, positive: bool = False) -> list[float]:
    if profile == "smooth":
        base = rng.choice([3, 40, 250, 1800, 12500])
        numbers = [base * (1 + rng.uniform(-0.3, 0.6)) for _ in range(count)]
    elif profile == "trend":
        base, step = rng.uniform(1, 300), rng.uniform(-40, 60)
        numbers = [base + step * index + rng.uniform(-5, 5) for index in range(count)]
    elif profile == "close":
        base = rng.choice([5.0, 120.0, 3400.0])
        numbers = [base * (1 + rng.uniform(-0.01, 0.01)) for _ in range(count)]
    elif profile == "spread":
        numbers = [10 ** rng.uniform(-2, 6) for _ in range(count)]
    elif profile == "equal":
        numbers = [rng.choice([0.0, 7.0, 150.0])] * count
    elif profile == "zeros":
        numbers = [0.0 if rng.random() < 0.4 else rng.uniform(0, 500) for _ in range(count)]
    else:
        numbers = [rng.uniform(-400, 600) for _ in range(count)]
    if positive:
        numbers = [abs(number) for number in numbers]
    elif profile not in ("equal", "zeros") and rng.random() < 0.25:
        numbers = [number if rng.random() < 0.6 else -number for number in numbers]
    digits = rng.choice([0, 0, 1, 2, 3])
    return [round(number, digits) if digits else float(round(number)) for number in numbers]


def random_chart(rng: random.Random, kind: str) -> dict:
    """A chart as reports hold them, or (one in three) a crowded or awkward one."""
    wide = rng.random() < 0.34
    count = rng.choice([1, 2, 5, 9, 12, 13, 16]) if wide else rng.randint(3, 8)
    labels = _names(rng, count) if wide or rng.random() < 0.4 else YEARS[-count:]
    data: dict = {"type": kind, "labels": labels}
    if rng.random() < 0.6:
        data["title"] = "Probe"
    number_format = rng.choice(FORMATS)
    if number_format:
        data["number_format"] = number_format
    profile = rng.choice(PROFILES if wide else PROFILES[:3])
    if kind in ("bar", "line"):
        series = rng.choice([1, 2, 3, 5, 7]) if wide else rng.choice([1, 2, 2, 3])
        names = _names(rng, series)
        positive = kind == "line" and not wide
        data["series"] = [
            {"name": name, "values": _numbers(rng, count, profile, positive)} for name in names
        ]
        if series > 1 and rng.random() < 0.4:
            data["accent_series"] = rng.choice(names)
        if kind == "bar" and series == 1 and rng.random() < 0.4:
            data["highlight"] = rng.sample(labels, k=1)
        if kind == "line" and rng.random() < 0.45:
            flat = [value for item in data["series"] for value in item["values"]]
            data["reference"] = [
                {"name": rng.choice(["한도", "Covenant"]), "value": round(rng.choice(flat), 2)}
                for _ in range(rng.choice([1, 1, 2]))
            ]
    elif kind in ("hbar", "waterfall"):
        data["series"] = [{"name": _names(rng, 1)[0], "values": _numbers(rng, count, profile)}]
        if kind == "waterfall" and rng.random() < 0.4:
            data["start_total"] = True
        if kind == "waterfall" and rng.random() < 0.3:
            data["total_label"] = rng.choice(["기말", "Closing balance"])
        if kind == "hbar" and rng.random() < 0.4:
            data["highlight"] = rng.sample(labels, k=1)
    elif kind == "stacked_bar":
        names = _names(rng, rng.choice([1, 2, 3, 3, 4, 4]))
        data["series"] = [
            {"name": name, "values": _numbers(rng, count, profile, positive=True)} for name in names
        ]
        if rng.random() < 0.6:
            data["totals"] = True
    elif kind == "bar_line":
        bars, lines = rng.choice([1, 1, 2]), rng.choice([1, 1, 2])
        series = []
        for index, name in enumerate(_names(rng, bars + lines)):
            if index >= bars:
                values = [round(rng.uniform(1.5, 6.5), 1) for _ in range(count)]
                if wide:
                    values = _numbers(rng, count, rng.choice(PROFILES))
                series.append({"name": name, "values": values, "axis": "right"})
            else:
                series.append({"name": name, "values": _numbers(rng, count, profile, not wide)})
        data["series"] = series
        right_format = rng.choice(FORMATS)
        if right_format:
            data["right_number_format"] = right_format
    else:  # range
        lows = _numbers(rng, count, profile)
        highs = [
            low if rng.random() < 0.12 else round(low + abs(low) * rng.uniform(0.01, 0.8) + 1, 2)
            for low in lows
        ]
        data["low"], data["high"] = lows, highs
        if rng.random() < 0.5:
            data["show_overlap"] = True
        if rng.random() < 0.6:
            data["reference"] = [
                {
                    "name": rng.choice(["현재가", "Median"]),
                    "value": round(rng.choice([*lows, *highs]), 2),
                }
                for _ in range(rng.choice([1, 1, 2]))
            ]
    return data


@pytest.mark.parametrize("kind", TYPES)
def test_many_charts_of_each_type_keep_the_rule(monkeypatch, kind):
    """Thirty charts per type from a fixed seed, at 3, 3.5, 4.5 and 6 inches."""
    rng = random.Random(f"final-rule-{kind}")
    for index in range(30):
        check(monkeypatch, random_chart(rng, kind), (6.0, 4.5, 3.0, 6.0, 3.5)[index % 5])


def meeting_lines_chart(rng: random.Random, kind: str) -> dict:
    """Two or three plotted series that meet: at some categories their values
    are equal (one spot, one text) or 0.01 or 0.02 apart (one spot, two texts)."""
    count = rng.randint(3, 7)
    base = [round(rng.uniform(80, 120), 1) for _ in range(count)]
    series = []
    for name in _names(rng, 2 if kind == "bar_line" else rng.choice([2, 2, 3])):
        values = []
        for value in base:
            pick = rng.random()
            if pick < 0.4:
                values.append(value)
            elif pick < 0.7:
                values.append(round(value + rng.choice([0.01, -0.01, 0.02]), 2))
            else:
                values.append(round(value + rng.uniform(-15, 15), 1))
        series.append({"name": name, "values": values})
    number_format = rng.choice([",.2f", ",.2f", ".1f", None])
    data: dict = {"type": kind, "labels": YEARS[-count:], "series": series}
    if kind == "line":
        if number_format:
            data["number_format"] = number_format
        return data
    for item in series:
        item["axis"] = "right"
    bars = {"name": "Bars 0", "values": [float(round(rng.uniform(50, 400))) for _ in base]}
    data["series"] = [bars, *series]
    if number_format:
        data["right_number_format"] = number_format
    return data


@pytest.mark.parametrize("kind", ["line", "bar_line"])
def test_many_charts_whose_lines_meet_keep_the_rule(monkeypatch, kind):
    """Thirty charts per type from a fixed seed whose series share spots (fifth review)."""
    rng = random.Random(f"meeting-lines-{kind}")
    met = 0
    for index in range(30):
        data = meeting_lines_chart(rng, kind)
        seen = check(monkeypatch, data, WIDTHS[index % 4])
        lines = list(series_lines(_spec(data), seen).values())
        met += sum(
            1
            for first, line in enumerate(lines)
            for other in lines[first + 1 :]
            for mine, theirs in zip(line["pixels"], other["pixels"], strict=True)
            if math.hypot(mine[0] - theirs[0], mine[1] - theirs[1]) <= 0.5
        )
    assert met >= 30  # the sample does hold points on one spot
