"""IB/consulting style chart drawing (`chart_style: standard`, the default).

The picture holds only the plot: the title, subtitle, note and source are Word
paragraphs written by `ib_render_media.ChartRenderer`. Rules (E10, owner
decisions 2026-11-14): one accent colour (the theme's primary colour) with
greys, values written on the marks instead of a value axis, a zero baseline
and no frame, the legend in one row above the plot, line names at the line
ends, decreases hatched, and tones chosen so that series stay apart in black
and white. Values are never rescaled and are written with the decimal places
of `chart_renderer.format_value`, but rounded as a spreadsheet rounds: on the
decimal the spec wrote, a 5 away from zero (`round_half_up_text`; owner
decision 2026-11-14; the classic style keeps its rounding). The only numbers
the engine works out are the waterfall's total and, when the spec asks with
`totals: true`, a stack's total: exact sums of the written decimals, rounded
once. A line
chart keeps a value axis; only ticks whose written number equals their
position are kept. A legend too wide for one row takes more rows. Names are
drawn as written: a `$` never starts mathematical notation.

Changelog (E10):
    - NEW: stacked_bar, bar_line, range and hbar; bar, line and waterfall
      redrawn in the same style; tones chosen by grayscale distance.
    - FIXED (third review, 2026-11-14): value labels keep clear of every line,
      point marker, reference line and zero line, sit nearest to their own
      mark, and a stacked value beside its bar is centred on its own segment
      (the final rule, `_Drawer.fits`).
    - FIXED (fifth review, 2026-11-15): a value's own point, line pieces and
      marker are known by series and category (`_Label.point`,
      `_Drawer.points`), never by lying where its point lies; two points on
      one spot written with the same text share one number (`_Drawer.twins`).
"""

import logging
import math
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Context, Decimal
from io import BytesIO
from itertools import combinations
from typing import Any

from chart_renderer import ChartSpecError, waterfall_positions
from document_model import ChartSpec
from render_styles import STYLE, RasterFontPolicy

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════════
# TONES
# ═══════════════════════════════════════════════════════════════════════════════

# Neutral greys from near black to near white (share of white), the
# candidates for non-accent series; fine steps let the choice below keep wide
# grayscale gaps whatever the accent's own brightness.
GREY_SHARES = tuple(round(0.08 + 0.05 * step, 2) for step in range(17))
# The grey beside the accent in a two-tone chart when it stays far enough.
PAIR_GREY_SHARE = 0.75
# Tints of the accent colour (share of white) offered to stacked series.
TINT_SHARES = (0.35, 0.5, 0.6, 0.7)
# Smallest grayscale (0-255 luma) distance between the tones of one chart, by
# the number of tones; tests/test_chart_ib_style.py checks it for any accent.
MIN_TONE_GAP = {2: 60, 3: 50, 4: 40}
# A waterfall's start and total bars keep at least this distance from the
# accent, so balances and increases differ when printed in black and white.
ANCHOR_GAP = 45
ANCHOR_GREY_SHARE = 0.35
# More series than this cannot keep MIN_TONE_GAP; the renderer warns.
MAX_DISTINCT_SERIES = 4


def _hex(color: object) -> str:
    return f"#{str(color).upper()}"


def _channels(color: str) -> tuple[int, int, int]:
    value = color.lstrip("#")
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def mix(color: str, other: str, share: float) -> str:
    """Blend `share` of `other` into `color` (both `#RRGGBB`)."""
    channels = [
        round(a + (b - a) * share) for a, b in zip(_channels(color), _channels(other), strict=True)
    ]
    return "#{:02X}{:02X}{:02X}".format(*channels)


def luma(color: str) -> float:
    """Grayscale brightness 0-255 (ITU-R BT.601), as a black-and-white print shows it."""
    red, green, blue = _channels(color)
    return 0.299 * red + 0.587 * green + 0.114 * blue


def _min_gap(colors: Sequence[str]) -> float:
    return min(
        (abs(luma(a) - luma(b)) for a, b in combinations(colors, 2)),
        default=255.0,
    )


@dataclass(frozen=True)
class ChartTones:
    """The colours of one chart, worked out from the active style."""

    accent: str
    negative: str
    text: str
    muted: str

    @classmethod
    def from_style(cls) -> "ChartTones":
        return cls(
            accent=_hex(STYLE.NAVY),
            negative=_hex(STYLE.CHART_NEGATIVE_COLOR),
            text=_hex(STYLE.DARK_GRAY),
            muted=_hex(STYLE.MEDIUM_GRAY),
        )

    def greys(self) -> list[str]:
        return [mix("#000000", "#FFFFFF", share) for share in GREY_SHARES]

    def tints(self) -> list[str]:
        return [mix(self.accent, "#FFFFFF", share) for share in TINT_SHARES]

    @property
    def tint(self) -> str:
        """The accent's light tint (60 % white), for range bars."""
        return mix(self.accent, "#FFFFFF", 0.6)

    def _pick(self, count: int, candidates: list[str], prefer_tint: bool) -> list[str]:
        """Choose `count` tones besides the accent with the widest grayscale gaps.

        Every choice whose smallest gap reaches the best gap less 4, or
        MIN_TONE_GAP plus 10 if that is lower, is acceptable; among those a
        stack prefers one with a single tint of the accent, then the set whose
        tones are spread most evenly, then the earlier candidate order.
        """
        if count <= 0:
            return []
        tints = set(self.tints())
        scored = []
        for order, combo in enumerate(combinations(candidates, count)):
            tones = [self.accent, *combo]
            gap = _min_gap(tones)
            has_tint = sum(color in tints for color in combo) == 1
            lumas = sorted(luma(tone) for tone in tones)
            spread = sum(b - a for a, b in zip(lumas, lumas[1:], strict=False) if b - a < gap + 20)
            scored.append((gap, has_tint, spread, -order, list(combo)))
        best = max(item[0] for item in scored)
        floor = min(best - 4, MIN_TONE_GAP.get(count + 1, 0) + 10)
        near = [item for item in scored if item[0] >= floor]
        chosen = max(near, key=lambda item: (item[1] if prefer_tint else 0, item[2], item[3]))
        return chosen[4]

    def stack(self, count: int) -> list[str]:
        """Bottom-to-top tones of a stack: accent, a tint if it fits, then greys."""
        others = self._pick(count - 1, [*self.tints(), *self.greys()], prefer_tint=True)
        tints = set(self.tints())
        others.sort(key=lambda c: (c not in tints, luma(c)))
        return [self.accent, *others]

    def compare(self, count: int, accent_index: int) -> list[str]:
        """Tones of series compared side by side; `accent_index` gets the accent.

        The other series get greys, lighter for earlier series ("later is darker").
        """
        if count == 2:
            greys = [self.pair_grey()]
        elif count > MAX_DISTINCT_SERIES:
            greys = [
                mix("#000000", "#FFFFFF", 0.3 + 0.5 * index / max(1, count - 2))
                for index in range(count - 1)
            ]
        else:
            greys = self._pick(count - 1, self.greys(), prefer_tint=False)
        greys.sort(key=luma, reverse=True)
        tones = []
        for index in range(count):
            tones.append(self.accent if index == accent_index else greys.pop(0))
        return tones

    def pair_grey(self) -> str:
        """The grey beside the accent in a two-tone chart (highlight, bar_line).

        A light grey (75 % white) when it keeps MIN_TONE_GAP[2] from the
        accent, otherwise the grey nearest to it that does.
        """
        preferred = mix("#000000", "#FFFFFF", PAIR_GREY_SHARE)
        far = [g for g in self.greys() if abs(luma(g) - luma(self.accent)) >= MIN_TONE_GAP[2]]
        return min(far, key=lambda g: abs(luma(g) - luma(preferred)))

    def anchor(self) -> str:
        """The grey for waterfall balances: nearest to a dark grey (35 % white)
        among those that keep ANCHOR_GAP from the accent."""
        preferred = mix("#000000", "#FFFFFF", ANCHOR_GREY_SHARE)
        far = [g for g in self.greys() if abs(luma(g) - luma(self.accent)) >= ANCHOR_GAP]
        return min(far, key=lambda g: abs(luma(g) - luma(preferred)))


def on_fill(fill: str) -> str:
    """Label colour inside a filled mark."""
    return "#FFFFFF" if luma(fill) < 140 else "#202020"


# ═══════════════════════════════════════════════════════════════════════════════
# LAYOUT HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_HEIGHT = 2.9  # inches
ROW_HEIGHT = {"hbar": 0.42, "range": 0.5}
MAX_HEIGHT = 9.0
VALUE_SIZE = 8.0
CATEGORY_SIZE = 8.5
LEGEND_SIZE = 8.0
LEGEND_BAND = 0.32  # inches above the plot for a one-row legend
MANY_POINTS = 12  # more points: no markers and thinned category names
LABELLED_POINTS = 8  # up to this many points, every accent point gets its value
LINE_GAP = 11.5  # points between two line-end names, and one nudge of a value
END_ROOM = 0.75  # inches of a line chart besides its line-end names


def chart_height(spec: ChartSpec) -> float:
    """Picture height in inches: fixed, by rows for horizontal charts, and for
    a line chart at least enough to stack every line-end name."""
    if spec.chart_type in ROW_HEIGHT:
        rows = len(spec.labels)
        extra = 0.25 if spec.reference_lines else 0.0
        return min(MAX_HEIGHT, max(1.3, ROW_HEIGHT[spec.chart_type] * rows + 0.45 + extra))
    if spec.chart_type == "line":
        ends = len(spec.series) * LINE_GAP / 72 + END_ROOM
        return min(MAX_HEIGHT, max(DEFAULT_HEIGHT, ends))
    return DEFAULT_HEIGHT


def _text_width(text: str, size: float) -> float:
    """Approximate width in inches; wide (CJK) characters count as one em."""
    ems = sum(1.0 if ord(char) >= 0x1100 else 0.56 for char in text)
    return ems * size / 72.0


def wrap_label(label: str, slot_inches: float, size: float = CATEGORY_SIZE) -> str:
    """Break a category name that does not fit its slot at the middle space.

    Korean is never rotated; a name without spaces is left as written.
    """
    if _text_width(label, size) <= slot_inches * 0.92 or " " not in label:
        return label
    middle = len(label) / 2
    spaces = [index for index, char in enumerate(label) if char == " "]
    split = min(spaces, key=lambda index: abs(index - middle))
    return label[:split] + "\n" + label[split + 1 :]


def _value_size(bar_width_inches: float) -> float:
    if bar_width_inches >= 0.3:
        return VALUE_SIZE
    return 7.0 if bar_width_inches >= 0.18 else 6.5


def spread(desired: Sequence[float], gap: float, low: float, high: float) -> list[float] | None:
    """Move sorted positions apart by `gap` within [low, high], each as little as possible.

    Returns None when they cannot all fit.
    """
    if not desired:
        return []
    if (len(desired) - 1) * gap > high - low:
        return None
    placed = [min(max(value, low), high) for value in desired]
    for index in range(1, len(placed)):
        placed[index] = max(placed[index], placed[index - 1] + gap)
    if placed[-1] > high:
        placed[-1] = high
        for index in range(len(placed) - 2, -1, -1):
            placed[index] = min(placed[index], placed[index + 1] - gap)
    return placed


# ═══════════════════════════════════════════════════════════════════════════════
# VALUE LABELS
# ═══════════════════════════════════════════════════════════════════════════════

MAX_GAP = 16.0  # points: a value farther than this from its bar end or point is not on it
# Points: a value's own bar or point may be at most this much farther from the
# text's centre than the nearest other bar or point (the final rule, d).
NEAREST_MARGIN = 3.0
# Pixels kept between a text's box and a line or point marker: the soft
# edges of both (anti-aliasing) reach about a pixel beyond their outlines.
INK_MARGIN = 1.5
# Values this large are drawn at an exact power-of-two scale, so that the
# plot's limits do not overflow; the numbers written are the original ones.
SAFE_MAGNITUDE = 1e290


@dataclass
class _Spot:
    """One place a value label may take, tried in order.

    `rule` says how the label must sit on its mark (`anchor`, a bar or
    segment patch; None for a point at `xy`):

    - `inside`: inside the anchor patch;
    - `end`: centred over the bar or point, at most `MAX_GAP` above or below it;
    - `side`: level with the bar or point, at most `MAX_GAP` beside it;
    - `beside`: a stacked segment's value beside its bar, inside the
      category's own column (`slot`, data x), its centre within the
      segment's height, above the baseline and not level with another value
      written beside the same bar;
    - `free`: a line-end name or a reference line, which carries its name.
    """

    xy: tuple[float, float]
    text: str
    offset: tuple[float, float] = (0.0, 0.0)  # points
    ha: str = "center"
    va: str = "bottom"
    size: float = VALUE_SIZE
    color: str = "#202020"
    bold: bool = False
    axes: Any = None  # None: the main axes
    rule: str = "end"
    anchor: Any = None
    slot: tuple[float, float] | None = None


@dataclass
class _Label:
    """A number the chart type writes: drawn at the first spot that fits, else reported.

    `owners` are its own bars (its bar or its stack). `group` is the stack a
    segment value belongs to. `point` is the plotted point a point's value or
    a line-end name belongs to: its series' line and the category's index.
    Marks are a label's own by this record alone, never because they lie
    where its mark lies. `role` is `line` for a line-end name and
    `reference` for a reference line's text: a line left without one has no
    name in the picture, which the warning says.
    """

    number: str
    category: str
    series: str
    spots: list[_Spot]
    owners: list[Any] = field(default_factory=list)
    group: int | None = None
    role: str = "value"
    # A line-end name: where its own line ends (data x and value).
    end: tuple[float, float] | None = None
    point: tuple[Any, int] | None = None

    @property
    def where(self) -> str:
        return f"{self.category}, {self.series}" if self.category else self.series


def _overlaps(first: Any, second: Any, tolerance: float = 0.5) -> bool:
    return bool(
        first.x0 < second.x1 - tolerance
        and second.x0 < first.x1 - tolerance
        and first.y0 < second.y1 - tolerance
        and second.y0 < first.y1 - tolerance
    )


def _contains(outer: Any, inner: Any, tolerance: float = 0.5) -> bool:
    return bool(
        inner.x0 >= outer.x0 - tolerance
        and inner.x1 <= outer.x1 + tolerance
        and inner.y0 >= outer.y0 - tolerance
        and inner.y1 <= outer.y1 + tolerance
    )


def _crosses(box: Any, segment: tuple[float, float, float, float, float]) -> bool:
    """Whether a drawn line segment (its width included) passes over a box.

    Args:
        box: The text's box in pixels.
        segment: Start and end in pixels and half the line's width.
    """
    x0, y0, x1, y1, half = segment
    grow = half + INK_MARGIN
    left, right = box.x0 - grow, box.x1 + grow
    bottom, top = box.y0 - grow, box.y1 + grow
    if left >= right or bottom >= top:
        return False
    enter, leave = 0.0, 1.0
    dx, dy = x1 - x0, y1 - y0
    for step, room in ((-dx, x0 - left), (dx, right - x0), (-dy, y0 - bottom), (dy, top - y0)):
        if step == 0:
            if room <= 0:
                return False
            continue
        at = room / step
        if step < 0:
            enter = max(enter, at)
        else:
            leave = min(leave, at)
        if enter >= leave:
            return False
    return True


def _away(box: Any, x: float, y: float) -> float:
    """Distance in pixels from a point to a box (0 inside it)."""
    return math.hypot(max(box.x0 - x, x - box.x1, 0.0), max(box.y0 - y, y - box.y1, 0.0))


def plain(text: str) -> str:
    """`text` for matplotlib with every `$` literal (no mathematical notation)."""
    return text.replace("$", "\\$")


def written(text: str) -> str:
    """The text an artist shows, as the spec wrote it (undoes `plain`)."""
    return text.replace("\\$", "$")


def _safe_scale(spec: ChartSpec) -> tuple[ChartSpec, float]:
    """The spec at a scale whose plot limits cannot overflow, and that scale.

    Only values beyond `SAFE_MAGNITUDE` are scaled, by a power of two (exact);
    `_Drawer.fmt` divides by the scale again, so the numbers written are the
    spec's own.
    """
    numbers = [value for series in spec.series for value in series.values]
    numbers += [*spec.low, *spec.high, *(line.value for line in spec.reference_lines)]
    largest = max((abs(value) for value in numbers), default=0.0)
    if largest <= SAFE_MAGNITUDE:
        return spec, 1.0
    scale = 2.0 ** -(math.frexp(largest)[1] - math.frexp(SAFE_MAGNITUDE)[1])
    return (
        replace(
            spec,
            series=[
                replace(series, values=[value * scale for value in series.values])
                for series in spec.series
            ],
            low=[value * scale for value in spec.low],
            high=[value * scale for value in spec.high],
            reference_lines=[
                replace(line, value=line.value * scale) for line in spec.reference_lines
            ],
        ),
        scale,
    )


# Sums and rounding of written decimals are exact: the largest finite float
# with six decimals has 315 digits.
EXACT = Context(prec=1000)
SUFFIXES = {"percent": "%", "bps": " bps", "multiple": "x"}


def written_decimal(value: float) -> Decimal:
    """The number as the spec wrote it: the float's shortest decimal text.

    `2.675` is the decimal 2.675, not the binary float just below it, so that
    rounding is done on what the author typed.
    """
    return Decimal(repr(float(value)))


def exact_sum(values: Sequence[float], scale: float = 1.0) -> Decimal:
    """The exact sum of the numbers as the spec wrote them.

    The decimals are added without rounding, so `[0.3, -0.1, -0.2]` sums to 0
    and not to the binary rounding error -2.8e-17, and a sum is rounded once,
    when it is written. `values` are at the plot's scale (`_safe_scale`); the
    sum is at the spec's.
    """
    total = Decimal(0)
    for value in values:
        total = EXACT.add(total, written_decimal(value / scale))
    return total


def round_half_up_text(number: Decimal, number_format: str | None) -> str:
    """A number as a standard chart writes it: a 5 rounds away from zero.

    The decimal places are those of `chart_renderer.format_value` (the named
    formats and no format write one decimal unless the number is whole; a
    fixed-point format writes its own), with thousands separators, the same
    suffixes and parentheses for a negative amount. Only the rounding
    differs: `4.125` is `4.13` and `2.5` with `,.0f` is `3`, as a spreadsheet
    rounds, where the classic style keeps Python's rounding of the binary
    float (ties to even). A negative amount that rounds to zero is written
    without parentheses.

    Args:
        number: The exact decimal to write.
        number_format: Named format or a fixed-point format such as ',.2f'.

    Returns:
        The text written in the picture.
    """
    magnitude = number.copy_abs()
    if number_format and number_format not in ("number", *SUFFIXES):
        places = number_format.lstrip(",").rstrip("f")
        digits = int(places[1:]) if places else 6
        rounded = magnitude.quantize(Decimal(1).scaleb(-digits), ROUND_HALF_UP, EXACT)
        text = format(rounded, ",f" if number_format.startswith(",") else "f")
    elif magnitude == magnitude.to_integral_value():
        text = f"{int(magnitude):,}"
    else:
        text = format(magnitude.quantize(Decimal("0.1"), ROUND_HALF_UP, EXACT), ",f")
    text += SUFFIXES.get(number_format or "", "")
    negative = number.is_signed() and any(char in "123456789" for char in text)
    return f"({text})" if negative else text


def fit_text(text: str, width_inches: float, size: float) -> str:
    """Break each line of `text` that is wider than `width_inches`.

    A line breaks at its middle space when both halves fit, else wherever
    the width runs out (also inside a word: a name without spaces is broken
    rather than cut off at the picture's edge).
    """
    lines: list[str] = []
    for line in text.split("\n"):
        if _text_width(line, size) <= width_inches:
            lines.append(line)
            continue
        if " " in line:
            middle = len(line) / 2
            spaces = [index for index, char in enumerate(line) if char == " "]
            split = min(spaces, key=lambda index: abs(index - middle))
            halves = [line[:split], line[split + 1 :]]
            if all(_text_width(half, size) <= width_inches for half in halves):
                lines.extend(halves)
                continue
        piece = ""
        for char in line:
            if piece and _text_width(piece + char, size) > width_inches:
                lines.append(piece.rstrip())
                piece = char.lstrip()
            else:
                piece += char
        if piece:
            lines.append(piece)
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
# DRAWING
# ═══════════════════════════════════════════════════════════════════════════════


def chart_texts(spec: ChartSpec) -> list[str]:
    """Every string the picture draws (for the raster font choice)."""
    texts = [*spec.labels, *(series.name for series in spec.series)]
    texts += [line.name for line in spec.reference_lines]
    if spec.chart_type == "waterfall":
        texts.append(spec.total_label or STYLE.CHART_TOTAL_LABEL)
    return texts


def render_standard_chart_png(
    spec: ChartSpec,
    width_inches: float = 6.0,
    dpi: int = 200,
    warnings: list[str] | None = None,
) -> bytes:
    """Draw `spec` in the IB/consulting style into PNG bytes.

    Every number the chart type writes is drawn on its own mark, whole inside
    the picture and clear of the other texts and of every other mark (bars,
    lines, point markers, reference and zero lines), or it is left out and
    named in a warning (the final rule of `_Drawer.fits`, independent reviews
    2026-11-14). After the numbers are placed the drawn picture is checked
    again; a number that fails is removed and named. Warnings are logged once
    here and added to `warnings`.

    Args:
        spec: Validated chart data.
        width_inches: Picture width, bounded by the Word content area by its caller.
        dpi: PNG resolution.
        warnings: Receives non-rejecting notes (too many series to tell apart,
            values left out, texts that do not fit).

    Returns:
        PNG bytes ready for insertion from memory.
    """
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.text import Text

    family = RasterFontPolicy.resolve(STYLE.KOREAN_FONT, "\n".join(chart_texts(spec)))
    spec, scale = _safe_scale(spec)
    fig = Figure(figsize=(width_inches, chart_height(spec)), dpi=dpi, facecolor="#FFFFFF")
    FigureCanvasAgg(fig)
    try:
        ax = fig.add_subplot(111)
        ax.set_facecolor("#FFFFFF")
        drawer = _Drawer(fig, ax, spec, ChartTones.from_style(), width_inches, warnings, family)
        drawer.scale = scale
        draw: Callable[[], None] | None = getattr(drawer, "draw_" + spec.chart_type, None)
        if draw is None:
            raise ChartSpecError(f"chart_type is invalid: {spec.chart_type}")
        draw()
        drawer.reserve()
        for artist in fig.findobj(match=Text):
            artist.set_fontfamily(family)
        drawer.finish()
        with BytesIO() as buffer:
            fig.savefig(
                buffer, format="png", dpi=dpi, facecolor=fig.get_facecolor(), edgecolor="none"
            )
            return buffer.getvalue()
    finally:
        fig.clear()


class _Drawer:
    """Draws one chart on one axes; every method reads the spec as validated."""

    def __init__(
        self,
        fig: Any,
        ax: Any,
        spec: ChartSpec,
        tones: ChartTones,
        width: float,
        warnings: list[str] | None,
        family: str,
    ) -> None:
        self.fig, self.ax, self.spec, self.tones = fig, ax, spec, tones
        self.width = width
        self.warnings = warnings
        self.family = family
        self.legend: tuple[list[Any], list[str]] | None = None
        self.legend_artist: Any = None
        self.legend_columns = 1
        self.legend_room = max(0.5, width - 0.6)  # inches a legend name may take
        # Line ends: value, label text, colour, accent, series name, the line.
        self.ends: list[tuple[float, str, str, bool, str, Any]] = []
        self.labels: list[_Label] = []
        self.bars: list[Any] = []
        self.provisional: list[Any] = []
        self.placed: list[tuple[Any, _Label, _Spot]] = []
        self.axes: list[Any] = [ax]
        self.label_name = spec.title or spec.chart_type
        # Plotted data points: axes, x, y, the series' line, the category's
        # index and the value as the picture writes it. Whose a point is, is
        # read from here, never from where it lies.
        self.points: list[tuple[Any, float, float, Any, int, str]] = []
        # The stack (category index) of each stacked segment.
        self.stack_of: dict[int, int] = {}
        # Values not written because the same text at the same spot already
        # speaks for them: (the value, the value that is written).
        self.shared: list[tuple[_Label, _Label]] = []
        self.scale = 1.0  # see `_safe_scale`
        self.value_axis = False

    # ── shared pieces ────────────────────────────────────────────────────────

    def note(self, message: str) -> None:
        """A warning that never rejects: logged once here and handed to the renderer."""
        logger.warning("%s", message)
        if self.warnings is not None:
            self.warnings.append(message)

    def fmt(self, value: float | Decimal, right: bool = False) -> str:
        """The number as written in the picture (`round_half_up_text`).

        A float is a value at the plot's scale and is written as the decimal
        the spec held; a `Decimal` is a sum already at the spec's scale
        (`exact_sum`), rounded here once.
        """
        number_format = self.spec.right_number_format if right else self.spec.number_format
        if not isinstance(value, Decimal):
            value = written_decimal(value / self.scale)
        return round_half_up_text(value, number_format)

    def warn_many(self, count: int) -> None:
        if count > MAX_DISTINCT_SERIES:
            self.note(
                f"Chart {self.label_name!r} has {count} series; more than "
                f"{MAX_DISTINCT_SERIES} cannot be told apart when printed in black and white"
            )

    def accent_index(self) -> int:
        names = [series.name for series in self.spec.series]
        if self.spec.accent_series in names:
            return names.index(self.spec.accent_series)
        return len(names) - 1

    def plain_axes(self, value_axis: bool = False, grid: bool = False) -> None:
        ax = self.ax
        for side in ("top", "right", "left", "bottom"):
            ax.spines[side].set_visible(False)
        if value_axis:
            self.value_axis = True
            ax.tick_params(axis="y", length=0, labelsize=VALUE_SIZE, colors=self.tones.text, pad=3)
            ax.yaxis.set_major_formatter(_formatter(self.fmt))
        else:
            ax.set_yticks([])
        if grid:
            ax.set_axisbelow(True)
            ax.yaxis.grid(True, color="#E3E3E3", linewidth=0.6)
        ax.tick_params(axis="x", length=0, labelsize=CATEGORY_SIZE, colors=self.tones.text, pad=5)

    def baseline(self) -> None:
        self.ax.axhline(0, color=self.tones.text, linewidth=0.8, zorder=3)

    def categories(self, positions: Sequence[float], labels: Sequence[str], step: int = 1) -> None:
        slot = self.width * 0.85 / max(1, len(labels))
        shown = [
            fit_text(wrap_label(label, slot * step), slot * step * 0.92, CATEGORY_SIZE)
            if index % step == 0
            else ""
            for index, label in enumerate(labels)
        ]
        self.ax.set_xticks(list(positions))
        self.ax.set_xticklabels(
            [plain(name) for name in shown], fontsize=CATEGORY_SIZE, color=self.tones.text
        )

    def row_names(self, labels: Sequence[str]) -> list[str]:
        """Category names of a horizontal chart, broken at a third of the picture."""
        return [plain(fit_text(label, self.width / 3, CATEGORY_SIZE)) for label in labels]

    def value_limits(self, low: float, high: float, head: float = 0.14) -> None:
        low, high = min(0.0, low), max(0.0, high)
        span = (high - low) or 1.0
        self.ax.set_ylim(low - (span * head if low < 0 else 0.0), high + span * head)

    def patch_legend(self, colors: Sequence[str], names: Sequence[str]) -> None:
        from matplotlib.patches import Patch

        self.legend = ([Patch(color=color) for color in colors], list(names))

    @staticmethod
    def mark(patches: Any, category: str, series: str) -> None:
        """Name a bar for the drawn picture (`bar<TAB>category<TAB>series`)."""
        for patch in patches:
            patch.set_gid(f"bar\t{category}\t{series}")

    # ── value labels ─────────────────────────────────────────────────────────

    def end_spot(
        self,
        x: float,
        y: float,
        text: str,
        below: bool,
        size: float = VALUE_SIZE,
        color: str | None = None,
        bold: bool = False,
        lift: float = 2.5,
        axes: Any = None,
        anchor: Any = None,
    ) -> _Spot:
        """Above a mark (below it for a negative value), `lift` points away."""
        return _Spot(
            (x, y),
            text,
            (0.0, -lift if below else lift),
            "center",
            "top" if below else "bottom",
            size,
            color or "#202020",
            bold,
            axes,
            "end",
            anchor,
        )

    def outward(self, spot: _Spot, below: bool) -> _Spot:
        """The same spot moved one nudge further from its mark."""
        dx, dy = spot.offset
        step = LINE_GAP * (-1 if below else 1)
        return _Spot(**{**spot.__dict__, "offset": (dx, dy + step)})

    def add_label(
        self,
        number: str,
        category: str,
        series: str,
        spots: list[_Spot],
        owners: Sequence[Any] = (),
        group: int | None = None,
        role: str = "value",
        point: tuple[Any, int] | None = None,
    ) -> None:
        self.labels.append(
            _Label(number, category, series, spots, list(owners), group, role, point=point)
        )

    def _make(self, spot: _Spot) -> Any:
        axes = spot.axes or self.ax
        return axes.annotate(
            plain(spot.text),
            spot.xy,
            xytext=spot.offset,
            textcoords="offset points",
            ha=spot.ha,
            va=spot.va,
            fontsize=spot.size,
            color=spot.color,
            fontweight="bold" if spot.bold else "normal",
            annotation_clip=False,
            fontfamily=self.family,
        )

    def reserve(self) -> None:
        """Draw each label at its first spot so that the layout leaves room for it."""
        for label in self.labels:
            self.provisional.append(self._make(label.spots[0]))

    def finish(self) -> None:
        """Fit the plot and legend inside the picture, settle the numbers, check."""
        band = self.legend_band()
        top = 1.0 - band / self.fig.get_figheight() if self.legend is not None else 1.0
        self.fit_inside(top)
        self.keep_true_ticks()
        if self.legend is not None:
            self.place_legend()
        for artist in self.provisional:
            artist.remove()
        self.provisional = []
        missing = self.settle()
        missing += self.verify()
        if missing:
            listed = "; ".join(f"{label.where}: {label.number}" for label in missing)
            self.note(
                f"Chart {self.label_name!r}: {len(missing)} value(s) not written in the "
                f"picture for lack of room: {listed}"
            )
        for role, what in (("line", "line(s)"), ("reference", "reference line(s)")):
            unnamed = [label.series for label in missing if label.role == role]
            if unnamed:
                self.note(
                    f"Chart {self.label_name!r}: {len(unnamed)} {what} drawn without a name "
                    f"(its text had no room): " + "; ".join(unnamed)
                )
        self.check_fixed_texts()

    def fit_inside(self, top: float) -> None:
        """Place the plot so that its texts lie inside the picture.

        The plot starts as large as the picture below the legend band and each
        pass moves its edges in by what sticks out (axis names, values at
        their first spots, line-end names), at most a few times. matplotlib's
        `tight_layout` is not used: it gives up with a Python warning when the
        texts need more room than the picture has, and misses some overhangs.
        When the texts cannot fit, the plot keeps the last usable place and
        the texts that stick out are named by `check_fixed_texts` or left out.
        """
        from matplotlib.transforms import Bbox

        renderer = self.fig.canvas.get_renderer()
        width, height = self.fig.bbox.width, self.fig.bbox.height
        edge = 3.0 * self.fig.dpi / 72  # pixels kept free around the texts
        start = Bbox.from_extents(
            edge / width, edge / height, 1.0 - edge / width, top - edge / height
        )
        for axes in self.axes:
            axes.set_position(start)
        for _ in range(6):
            box = Bbox.union([axes.get_tightbbox(renderer) for axes in self.axes])
            left = max(0.0, edge - box.x0)
            right = max(0.0, box.x1 - (width - edge))
            bottom = max(0.0, edge - box.y0)
            over = max(0.0, box.y1 - (height * top - edge))
            if not (left or right or bottom or over):
                return
            place = self.ax.get_position()
            moved = Bbox.from_extents(
                place.x0 + left / width,
                place.y0 + bottom / height,
                place.x1 - right / width,
                place.y1 - over / height,
            )
            if moved.width <= 0.1 or moved.height <= 0.1:
                return
            for axes in self.axes:
                axes.set_position(moved)

    def keep_true_ticks(self) -> None:
        """Keep the value-axis ticks whose written number equals their position.

        A tick's text is rounded like every value (`fmt`), so a tick at 1.5
        written `2` would repeat the next tick's number. A tick is kept only
        when its text, read as a decimal, is its position (the shortest
        decimal of the float it is drawn at), and when no other kept tick
        has the same text; with fewer than two left the axis writes no
        numbers (the values are written at the points).
        """
        if not self.value_axis:
            return
        ax = self.ax
        if self.scale != 1.0:
            # Ticks are round numbers at the plot's scale, not at the spec's.
            ax.tick_params(axis="y", labelleft=False)
            return
        low, high = sorted(ax.get_ylim())
        # The locator's own arithmetic noise (3 * 0.1 is 0.30000000000000004):
        # a few units in the last place of the axis's largest number, and far
        # below a pixel. Only that much may a tick be moved.
        noise = min(4 * math.ulp(max(abs(low), abs(high))), 1e-9 * (high - low))
        true: list[tuple[float, str]] = []
        for tick in ax.get_yticks():
            if not low <= tick <= high:
                continue
            text = self.fmt(float(tick))
            digits = "".join(char for char in text if char in "0123456789.")
            shown = Decimal(digits or "0")
            if text.startswith("("):
                shown = -shown
            # Compared as decimals, not within a float tolerance: a tick at
            # 1,000,000,000.5 is not 1,000,000,001. The tick is drawn at the
            # number its text says (the float whose shortest decimal is that
            # text), which is where the locator put it but for its noise.
            at = float(shown)
            if Decimal(repr(at)) == shown and abs(at - float(tick)) <= noise:
                true.append((at, text))
        # A number written at two ticks says nothing: neither is kept.
        texts = [text for _, text in true]
        kept = [tick for tick, text in true if texts.count(text) == 1]
        if len(kept) >= 2:
            ax.set_yticks(kept)
        else:
            ax.tick_params(axis="y", labelleft=False)

    def _legend(self, ncol: int, left: float) -> Any:
        from matplotlib.font_manager import FontProperties

        handles, names = self.legend or ([], [])
        # A name wider than the picture breaks into lines instead of leaving it.
        return self.fig.legend(
            handles,
            [plain(fit_text(name, self.legend_room, LEGEND_SIZE)) for name in names],
            loc="upper left",
            bbox_to_anchor=(left, 1.0),
            ncol=ncol,
            frameon=False,
            handlelength=0.9,
            handleheight=0.9,
            handletextpad=0.4,
            columnspacing=1.4,
            borderaxespad=0.2,
            borderpad=0.2,
            prop=FontProperties(family=self.family, size=LEGEND_SIZE),
            labelcolor=self.tones.text,
        )

    def legend_band(self) -> float:
        """Inches above the plot for the legend: one row when it fits the width.

        A legend too wide for one row takes as many rows as it needs, and
        the picture grows by the extra rows so that the plot keeps its height.
        """
        if self.legend is None:
            return 0.0
        renderer = self.fig.canvas.get_renderer()
        count = len(self.legend[0])
        self.legend_columns = count
        # Names are broken by an estimate of their width; when even one column
        # is wider than the picture, break them shorter by the measured excess.
        for _ in range(4):
            trial = self._legend(1, 0.0)
            wide = float(trial.get_window_extent(renderer).width)
            trial.remove()
            if wide <= self.fig.bbox.width - 2:
                break
            self.legend_room *= (self.fig.bbox.width - 2) / wide * 0.95
        for ncol in range(count, 0, -1):
            trial = self._legend(ncol, 0.0)
            box = trial.get_window_extent(renderer)
            trial.remove()
            self.legend_columns = ncol
            if box.width <= self.fig.bbox.width - 2:
                break
        band = max(LEGEND_BAND, float(box.height) / self.fig.dpi + 0.06)
        if band > LEGEND_BAND:
            self.fig.set_figheight(self.fig.get_figheight() + band - LEGEND_BAND)
        return float(band)

    def place_legend(self) -> None:
        renderer = self.fig.canvas.get_renderer()
        left = self.ax.get_position().x0
        legend = self._legend(self.legend_columns, left)
        if legend.get_window_extent(renderer).x1 > self.fig.bbox.width - 1:
            legend.remove()
            legend = self._legend(self.legend_columns, 0.0)
        self.legend_artist = legend

    def place_ends(self, renderer: Any) -> None:
        """Line-end names after the last point, spread at least `LINE_GAP` apart
        between the plot's bottom and the picture's top."""
        if not self.ends:
            return
        transform = self.ax.transData
        last_x = len(self.spec.labels) - 1
        items = sorted(
            (
                (transform.transform((last_x, value))[1], value, text, color, main, name, line)
                for value, text, color, main, name, line in self.ends
            ),
            key=lambda item: item[:6],
        )
        half = 0.0
        for item in items:  # a long name is broken into lines: the tallest sets the gap
            probe = self._make(_Spot((last_x, item[1]), item[2], (6.0, 0.0), "left", "center"))
            half = max(half, probe.get_window_extent(renderer).height / 2)
            probe.remove()
        gap = max(LINE_GAP * self.fig.dpi / 72, 2 * half + 1)
        low = self.ax.get_window_extent(renderer).y0 + half
        high = self.fig.bbox.y1 - half - 1
        placed = spread([item[0] for item in items], gap, low, high)
        positions = placed if placed is not None else [item[0] for item in items]
        inverse = transform.inverted()
        category = self.spec.labels[-1]
        for pixel_y, (_, value, text, color, main, name, line) in zip(
            positions, items, strict=True
        ):
            y = inverse.transform((0, pixel_y))[1]
            spot = _Spot(
                (last_x, y),
                text,
                (6.0, 0.0),
                "left",
                "center",
                VALUE_SIZE,
                color if main else self.tones.text,
                main,
                rule="free",
            )
            self.labels.append(
                _Label(
                    self.fmt(value),
                    category,
                    name,
                    [spot],
                    role="line",
                    end=(last_x, value),
                    point=(line, last_x),
                )
            )

    def _context(self, renderer: Any) -> dict[str, Any]:
        """What every check reads: the picture, the fixed texts and every drawn mark."""
        blocked = [box for _, box in self.fixed_texts(renderer)]
        segments, discs, segment_owners, disc_owners = self.marks(renderer)
        return {
            "canvas": self.fig.bbox,
            "blocked": blocked,
            "bars": [(bar, bar.get_window_extent(renderer)) for bar in self.bars],
            "segments": segments,
            "discs": discs,
            # Whose each line piece and point marker is: (line, index, index)
            # and (line, index) for a series' line, None for every other line.
            "segment_owners": segment_owners,
            "disc_owners": disc_owners,
            # Every plotted point: pixel x and y, its line, its category's
            # index and its written value.
            "points": [
                (*map(float, axes.transData.transform((x, y))), line, index, text)
                for axes, x, y, line, index, text in self.points
            ],
            "ends": [
                (float(self.ax.transData.transform((len(self.spec.labels) - 1, value))[1]), line)
                for value, *_, line in self.ends
            ],
            "floor": self.ax.get_window_extent(renderer).y0,
            "renderer": renderer,
        }

    def marks(
        self, renderer: Any
    ) -> tuple[
        list[tuple[float, float, float, float, float]],
        list[tuple[float, float, float]],
        list[tuple[Any, int, int] | None],
        list[tuple[Any, int] | None],
    ]:
        """Every drawn line and point marker, in pixels, and whose each is.

        Lines are the series' lines, reference lines, the zero line, waterfall
        connectors and the visible axis lines (spines), each as segments with
        half the line's width; a `|` marker is a short segment and every other
        marker a disc. Grid lines and the tint of a range's overlap lie
        behind the plot and are not marks.

        A series' line is known by the record of its plotted points
        (`self.points`): its piece between the points of categories `i` and
        `i + 1` is owned by `(line, i, i + 1)` and its marker at category `i`
        by `(line, i)`. Every other line (reference line, zero line,
        connector, a range's single-value mark, axis line) has no owner.

        Returns:
            Segments `(x0, y0, x1, y1, half_width)`, discs `(x, y, radius)`,
            and the owner of each segment and of each disc, in step.
        """
        scale = self.fig.dpi / 72
        segments: list[tuple[float, float, float, float, float]] = []
        discs: list[tuple[float, float, float]] = []
        segment_owners: list[tuple[Any, int, int] | None] = []
        disc_owners: list[tuple[Any, int] | None] = []
        series_lines = {id(line) for _, _, _, line, _, _ in self.points}
        for axes in self.axes:
            for line in axes.lines:
                points = [
                    (float(x), float(y))
                    for x, y in line.get_transform().transform(line.get_xydata())
                ]
                owned = id(line) in series_lines
                half = float(line.get_linewidth()) * scale / 2
                if line.get_linestyle() not in ("None", "", " "):
                    for index, ((x0, y0), (x1, y1)) in enumerate(
                        zip(points, points[1:], strict=False)
                    ):
                        segments.append((x0, y0, x1, y1, half))
                        segment_owners.append((line, index, index + 1) if owned else None)
                marker = line.get_marker()
                if marker in (None, "None", "", " "):
                    continue
                size = float(line.get_markersize()) * scale
                edge = float(line.get_markeredgewidth()) * scale
                for index, (x, y) in enumerate(points):
                    if marker == "|":
                        segments.append((x, y - size / 2, x, y + size / 2, edge / 2))
                        segment_owners.append(None)
                    else:
                        discs.append((x, y, size / 2 + edge / 2))
                        disc_owners.append((line, index) if owned else None)
            place = axes.get_window_extent(renderer)
            ends = {
                "left": (place.x0, place.y0, place.x0, place.y1),
                "right": (place.x1, place.y0, place.x1, place.y1),
                "bottom": (place.x0, place.y0, place.x1, place.y0),
                "top": (place.x0, place.y1, place.x1, place.y1),
            }
            for side, spine in axes.spines.items():
                if spine.get_visible() and side in ends:
                    segments.append((*ends[side], float(spine.get_linewidth()) * scale / 2))
                    segment_owners.append(None)
        return segments, discs, segment_owners, disc_owners

    def twins(self, label: _Label, context: dict[str, Any]) -> set[tuple[int, int]]:
        """Other series' points that a point's value is also true of.

        Two series' points of one category that lie on the same spot (within
        half a pixel) and are written with the same text, as where two index
        lines both start at 100: the one number names both, so the other
        point is no rival of the value and its marker does not come between.
        With different texts (100.00 and 100.01) the other point is a rival
        like any, however near it lies.

        Returns:
            `(id(line), category index)` of each such point.
        """
        if label.point is None or label.role != "value":
            return set()
        line, index = label.point
        own = [
            (x, y)
            for x, y, other_line, other_index, _ in context["points"]
            if other_line is line and other_index == index
        ]
        if not own:
            return set()
        own_x, own_y = own[0]
        return {
            (id(other_line), other_index)
            for x, y, other_line, other_index, text in context["points"]
            if other_line is not line
            and other_index == index
            and text == label.number
            and abs(x - own_x) <= 0.5
            and abs(y - own_y) <= 0.5
        }

    def fixed_texts(self, renderer: Any) -> list[tuple[Any, Any]]:
        """Axis names and legend entries, with their boxes: texts no check removes."""
        found = []
        for axes in self.axes:
            for axis, (low, high) in ((axes.xaxis, axes.get_xlim()), (axes.yaxis, axes.get_ylim())):
                if not axis.get_visible():
                    continue  # a twin's shared axis is not drawn twice
                low, high = min(low, high), max(low, high)
                margin = (high - low) * 1e-9
                for tick in axis.get_major_ticks():
                    label = tick.label1
                    # Only ticks inside the view are drawn.
                    inside = low - margin <= tick.get_loc() <= high + margin
                    if inside and label.get_visible() and label.get_text().strip():
                        found.append((label, label.get_window_extent(renderer)))
        if self.legend_artist is not None:
            for text in self.legend_artist.get_texts():
                if text.get_text().strip():
                    found.append((text, text.get_window_extent(renderer)))
            for handle in self.legend_artist.legend_handles:
                found.append((handle, handle.get_window_extent(renderer)))
        return found

    def _anchor_box(self, spot: _Spot, renderer: Any) -> Any:
        from matplotlib.transforms import Bbox

        if spot.anchor is not None:
            return spot.anchor.get_window_extent(renderer)
        x, y = (spot.axes or self.ax).transData.transform(spot.xy)
        return Bbox.from_extents(x, y, x, y)

    def fits(
        self,
        label: _Label,
        spot: _Spot,
        box: Any,
        others: Sequence[Any],
        stacked: Sequence[Any],
        context: dict[str, Any],
    ) -> bool:
        """The final rule: whether a text drawn at `box` may stay in the picture.

        Every text (values, line-end names, reference texts):

        - (a) lies whole inside the picture;
        - (b) overlaps no other text;
        - (c) overlaps no mark: no bar or segment (its own only when it is
          written inside it), no line with its width, no point marker (its
          own series' included), no reference line, zero line, axis line or
          waterfall connector (`marks`).

        A value (not a line-end name or reference text, which carry names):

        - sits on its own mark: inside its bar; centred over or under its
          bar or point within `MAX_GAP`; or beside its row with its centre
          at the row's height, within `MAX_GAP`;
        - (d) has no other mark nearer to its centre than its own: no other
          bar for a bar's value and no other point for a point's value by
          more than `NEAREST_MARGIN`, and its own mark nearer than every
          mark of the other kind (a point of any series for a bar's value, a
          bar for a point's value) and than every other series' point of the
          same category by at least `NEAREST_MARGIN`: those are on the
          text's own centre line, so only the distance tells them apart.
          Whose a point is, is read from the record of the plotted points
          (series and category), never from where it lies: another series'
          point of the same category is a rival also when it lies on the
          value's own point. The one exception is `twins`: on the same spot
          and written with the same text, the number is true of both.
          Rows of a horizontal chart (`hbar`, `range`) do not overlap
          and their values run sideways from the bar's end, where a longer
          neighbouring bar is always nearer to the text's centre: there the
          centre at the row's height settles it (an exception approved by
          the owner's coordinator, fourth review);
        - (e) beside a stack: inside the category's column, its centre within
          its own segment's height, above the baseline, not level with
          another value written beside the same bar, on either side, and not
          as high as a whole other segment of that bar (a drawn segment of
          1 px or more that lies wholly within the text's height);
        - (f) written outside its mark (over, under or beside it), has
          nothing between: the band from the text to its own mark, as wide
          as the text, holds no other bar or segment, no line, no point and
          no reference line (`_band_is_clear`). Otherwise a bar's value over
          another series' line reads as that line's value. Only the two
          pieces of the value's own series' line that meet at its own point
          are not counted; another series' line is, also when it leaves from
          the same spot.

        A line-end name sits nearest to its own line's end: no other series'
        line ends nearer to its centre's height by more than `NEAREST_MARGIN`.

        Args:
            label: The value.
            spot: The place tried.
            box: The label's box there, in pixels.
            others: Boxes of the other value labels drawn.
            stacked: Boxes of the other values written beside the same bar.
            context: From `_context`.
        """
        renderer = context["renderer"]
        if not _contains(context["canvas"], box):
            return False
        if any(_overlaps(box, other) for other in (*others, *context["blocked"])):
            return False
        own = {id(owner) for owner in label.owners} if spot.rule == "inside" else set()
        if any(_overlaps(box, bar_box) for bar, bar_box in context["bars"] if id(bar) not in own):
            return False
        if any(_crosses(box, segment) for segment in context["segments"]):
            return False
        if any(_away(box, x, y) < radius + INK_MARGIN for x, y, radius in context["discs"]):
            return False
        tolerance = 0.5
        gap = MAX_GAP * self.fig.dpi / 72
        margin = NEAREST_MARGIN * self.fig.dpi / 72
        centre_x, centre_y = (box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2
        if spot.rule == "free":
            if label.end is None:
                return True
            own_end = float(self.ax.transData.transform(label.end)[1])
            away = abs(centre_y - own_end)
            own_line = label.point[0] if label.point is not None else None
            return all(
                abs(centre_y - other) >= away - margin
                for other, line in context["ends"]
                if line is not own_line
            )
        anchor = self._anchor_box(spot, renderer)
        if spot.rule == "inside":
            return _contains(anchor, box)
        if not self._band_is_clear(label, spot, box, anchor, context):
            return False
        if spot.rule == "end":
            beyond = max(box.y0 - anchor.y1, anchor.y0 - box.y1, 0.0)
            if not (anchor.x0 - 1 <= centre_x <= anchor.x1 + 1 and beyond <= gap):
                return False
            if spot.anchor is None:  # a point: no other plotted point is nearer
                if label.point is None:
                    return False  # a point's value is always recorded with its point
                own_line, own_index = label.point
                twins = self.twins(label, context)
                near = math.hypot(centre_x - anchor.x0, centre_y - anchor.y0)
                rivals = []
                # Another series' point of the same category, and every bar,
                # is on the text's own centre line.
                others = []
                for x, y, line, index, _ in context["points"]:
                    if line is own_line and index == own_index:
                        continue  # the value's own point
                    away = math.hypot(centre_x - x, centre_y - y)
                    if index != own_index:
                        rivals.append(away)
                    elif (id(line), index) not in twins:
                        others.append(away)
                others += [_away(bar_box, centre_x, centre_y) for _, bar_box in context["bars"]]
            else:  # a bar: no other bar is nearer than the label's own
                mine = {id(owner) for owner in label.owners}
                near = min(
                    (
                        _away(bar_box, centre_x, centre_y)
                        for bar, bar_box in context["bars"]
                        if id(bar) in mine
                    ),
                    default=0.0,
                )
                rivals = [
                    _away(bar_box, centre_x, centre_y)
                    for bar, bar_box in context["bars"]
                    if id(bar) not in mine
                ]
                # ... and no plotted point of any series is nearer at all.
                others = [math.hypot(centre_x - x, centre_y - y) for x, y, *_ in context["points"]]
            # A mark of the same kind beside the text's own may be nearer by
            # the margin: the text is centred on its own. A mark on the same
            # centre line (a point over its category's bar, two series'
            # points of one category) is told apart by distance alone, so
            # there the text's own mark has to be the nearer one by that
            # margin.
            return all(rival >= near - margin for rival in rivals) and all(
                other >= near + margin for other in others
            )
        level = anchor.y0 - tolerance <= centre_y <= anchor.y1 + tolerance
        if spot.rule == "side":
            beside = max(box.x0 - anchor.x1, anchor.x0 - box.x1, 0.0)
            return bool(level and beside <= gap)
        # beside: a stacked segment's value next to its own bar
        transform = (spot.axes or self.ax).transData
        assert spot.slot is not None
        left = transform.transform((spot.slot[0], 0))[0]
        right = transform.transform((spot.slot[1], 0))[0]
        if box.x0 < left - tolerance or box.x1 > right + tolerance:
            return False
        if not level or box.y0 < context["floor"] - tolerance:
            return False
        # A text as high as a whole other segment of its bar names neither:
        # two thin segments level with one number cannot be told apart.
        mine = {id(owner) for owner in label.owners}
        for bar, bar_box in context["bars"]:
            if id(bar) in mine or bar_box.y1 - bar_box.y0 < 1.0:
                continue
            # The same bar: a segment of the label's own stack, by the record.
            same_bar = label.group is not None and self.stack_of.get(id(bar)) == label.group
            if same_bar and bar_box.y0 >= box.y0 - tolerance and bar_box.y1 <= box.y1 + tolerance:
                return False
        return not any(
            box.y0 < other.y1 - tolerance and other.y0 < box.y1 - tolerance for other in stacked
        )

    def _band_is_clear(
        self, label: _Label, spot: _Spot, box: Any, anchor: Any, context: dict[str, Any]
    ) -> bool:
        """Rule (f): nothing lies between a text and the mark it is written beside.

        The band is the rectangle from the text to its own mark, as wide (or,
        beside a row, as high) as the text. It starts one pixel off the mark,
        so a line that runs along the mark's own edge (the baseline under a
        bar of zero, a waterfall connector leaving the bar's end) is not in
        it. For a point's value, the two pieces of its own series' line that
        meet at its own point are not counted, nor its own marker, nor the
        marker of a twin (`twins`): these are known by series and category,
        never by lying where the point lies. Another series' line counts,
        also when it leaves from the very same spot. Every other bar, line,
        point marker or reference line in the band hides which mark the
        number belongs to.
        """
        from matplotlib.transforms import Bbox

        edge = 1.0
        if spot.rule == "end":
            if box.y0 >= anchor.y1:
                band = Bbox.from_extents(box.x0, anchor.y1 + edge, box.x1, box.y0)
            else:
                band = Bbox.from_extents(box.x0, box.y1, box.x1, anchor.y0 - edge)
        elif box.x0 >= anchor.x1:
            band = Bbox.from_extents(anchor.x1 + edge, box.y0, box.x0, box.y1)
        else:
            band = Bbox.from_extents(box.x1, box.y0, anchor.x0 - edge, box.y1)
        if band.x1 <= band.x0 or band.y1 <= band.y0:
            return True  # the text touches its mark
        mine = {id(owner) for owner in label.owners}
        if any(_overlaps(band, bar_box) for bar, bar_box in context["bars"] if id(bar) not in mine):
            return False
        point = label.point if spot.anchor is None else None
        twins = self.twins(label, context) if point is not None else set()

        def own_piece(owner: tuple[Any, int, int] | None) -> bool:
            return (
                point is not None
                and owner is not None
                and owner[0] is point[0]
                and point[1] in owner[1:]
            )

        def own_marker(owner: tuple[Any, int] | None) -> bool:
            if point is None or owner is None:
                return False
            if owner[0] is point[0] and owner[1] == point[1]:
                return True
            return (id(owner[0]), owner[1]) in twins

        for (x0, y0, x1, y1, _), owner in zip(
            context["segments"], context["segment_owners"], strict=True
        ):
            if own_piece(owner):
                continue
            if _crosses(band, (x0, y0, x1, y1, -INK_MARGIN)):  # the line's centre
                return False
        return not any(
            _away(band, x, y) < radius
            for (x, y, radius), owner in zip(context["discs"], context["disc_owners"], strict=True)
            if not own_marker(owner)
        )

    def settle(self) -> list[_Label]:
        """Draw every number at its first spot that `fits`; return those with none."""
        self.fig.canvas.draw()
        renderer = self.fig.canvas.get_renderer()
        self.place_ends(renderer)
        context = self._context(renderer)
        drawn: list[Any] = []
        missing: list[_Label] = []
        for label in self.labels:
            # The same text at the same spot is written once: a value whose
            # twin is already written is neither drawn again nor missing.
            speaker = self._speaker(label, context)
            if speaker is not None:
                self.shared.append((label, speaker))
                continue
            for spot in label.spots:
                artist = self._make(spot)
                box = artist.get_window_extent(renderer)
                stacked = [
                    other_artist.get_window_extent(renderer)
                    for other_artist, other, other_spot in self.placed
                    if label.group is not None
                    and other.group == label.group
                    and other_spot.rule == "beside"
                ]
                if self.fits(label, spot, box, drawn, stacked, context):
                    artist.set_gid(f"value\t{label.category}\t{label.series}\t{label.number}")
                    drawn.append(box)
                    self.placed.append((artist, label, spot))
                    break
                artist.remove()
            else:
                missing.append(label)
        return self._spoken_for(missing, context)

    def _speaker(self, label: _Label, context: dict[str, Any]) -> _Label | None:
        """The written value of a twin of `label` (`twins`), if there is one."""
        twins = self.twins(label, context)
        if not twins:
            return None
        return next(
            (
                other
                for _, other, _ in self.placed
                if other.point is not None
                and other.role == "value"
                and (id(other.point[0]), other.point[1]) in twins
            ),
            None,
        )

    def _spoken_for(self, missing: list[_Label], context: dict[str, Any]) -> list[_Label]:
        """The values of `missing` that no written twin speaks for.

        A value with no place of its own is not missing when its twin's
        value is written: the same text at the same spot is true of both
        points. It is kept in `shared` with the value that speaks for it.
        """
        still: list[_Label] = []
        for label in missing:
            speaker = self._speaker(label, context)
            if speaker is None:
                still.append(label)
            else:
                self.shared.append((label, speaker))
        return still

    def verify(self) -> list[_Label]:
        """Check the drawn picture again (principle N3) and remove a value that fails.

        The figure is drawn once more and every value label is measured where
        it ended up, against all the others still drawn; one that is no longer
        whole, clear and on its own mark is removed and reported.
        """
        self.fig.canvas.draw()
        renderer = self.fig.canvas.get_renderer()
        context = self._context(renderer)
        boxes = [artist.get_window_extent(renderer) for artist, _, _ in self.placed]
        kept = list(range(len(self.placed)))
        missing: list[_Label] = []
        for index, (artist, label, spot) in enumerate(self.placed):
            others = [boxes[other] for other in kept if other != index]
            stacked = [
                boxes[other]
                for other in kept
                if other != index and label.group is not None
                if self.placed[other][1].group == label.group
                and self.placed[other][2].rule == "beside"
            ]
            if not self.fits(label, spot, boxes[index], others, stacked, context):
                artist.remove()
                kept.remove(index)
                missing.append(label)
        self.placed = [self.placed[index] for index in kept]
        # A value that stood for its twin is gone: the twin is not written
        # either, unless another twin still is.
        gone = {id(label) for label in missing}
        missing += [label for label, speaker in self.shared if id(speaker) in gone]
        self.shared = [pair for pair in self.shared if id(pair[1]) not in gone]
        return self._spoken_for(missing, context)

    def check_fixed_texts(self) -> None:
        """Warn when an axis name or legend entry leaves the picture or overlaps another."""
        renderer = self.fig.canvas.get_renderer()
        texts = [
            (artist, box)
            for artist, box in self.fixed_texts(renderer)
            if hasattr(artist, "get_text")
        ]
        canvas = self.fig.bbox
        wrong = []
        for index, (artist, box) in enumerate(texts):
            clash = any(_overlaps(box, other) for _, other in texts[index + 1 :])
            if not _contains(canvas, box) or clash:
                wrong.append(written(artist.get_text()).replace("\n", " "))
        if wrong:
            shown = ", ".join(repr(text if len(text) <= 40 else text[:39] + "…") for text in wrong)
            self.note(
                f"Chart {self.label_name!r}: {len(wrong)} text(s) do not fit in the "
                f"picture or overlap: {shown}"
            )

    # ── bar ──────────────────────────────────────────────────────────────────

    def draw_bar(self) -> None:
        spec, ax = self.spec, self.ax
        count = len(spec.series)
        self.warn_many(count)
        if count == 1 and spec.highlight:
            grey = self.tones.pair_grey()
            colors = [
                [self.tones.accent if label in spec.highlight else grey for label in spec.labels]
            ]
        elif count == 1:
            colors = [[self.tones.accent] * len(spec.labels)]
        else:
            tones = self.tones.compare(count, self.accent_index())
            colors = [[tone] * len(spec.labels) for tone in tones]
        group = 0.72
        width = group / count
        bar_inches = width * 0.94 * self.width * 0.85 / len(spec.labels)
        size = _value_size(bar_inches)
        values = [value for series in spec.series for value in series.values]
        for index, series in enumerate(spec.series):
            offset = (index - (count - 1) / 2.0) * width
            xs = [x + offset for x in range(len(spec.labels))]
            patches = ax.bar(
                xs, series.values, width=width * 0.94, color=colors[index], linewidth=0
            )
            self.bars.extend(patches)
            for x, value, color, patch, label in zip(
                xs, series.values, colors[index], patches, spec.labels, strict=True
            ):
                self.mark([patch], label, series.name)
                bold = bool(count == 1 and spec.highlight and color == self.tones.accent)
                text = self.fmt(value)
                below = value < 0
                first = self.end_spot(
                    x,
                    value,
                    text,
                    below,
                    size,
                    self.tones.accent if bold else None,
                    bold,
                    anchor=patch,
                )
                inside = _Spot(
                    (x, value),
                    text,
                    (0.0, 3.0 if below else -3.0),
                    "center",
                    "bottom" if below else "top",
                    size,
                    on_fill(color),
                    bold,
                    rule="inside",
                    anchor=patch,
                )
                spots = [first, self.outward(first, below), inside]
                self.add_label(text, label, series.name, spots, [patch])
        self.plain_axes()
        self.baseline()
        self.categories(range(len(spec.labels)), spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        self.value_limits(min(values), max(values))
        if count > 1:
            self.patch_legend([c[0] for c in colors], [s.name for s in spec.series])

    # ── stacked bar ──────────────────────────────────────────────────────────

    def draw_stacked_bar(self) -> None:
        """Stacks; each segment's value inside it when it fits, else beside the bar.

        A value written beside the bar stays inside its category's column,
        with its centre within its own segment's height, above the baseline,
        not level with another value beside the same bar and not as high as
        a whole other segment of the bar (the final rule, e); with no such
        place it is left out and named. A segment of 0 has
        no height and no value. The stack total is written only with
        `totals: true`.
        """
        spec, ax = self.spec, self.ax
        tones = self.tones.stack(len(spec.series))
        positions = range(len(spec.labels))
        sums = [exact_sum([s.values[i] for s in spec.series], self.scale) for i in positions]
        totals = [float(total) * self.scale for total in sums]
        top = max(totals) or 1.0
        width = 0.56  # leaves each column room for a value beside its bar
        bar_inches = width * self.width * 0.85 / len(spec.labels)
        size = _value_size(bar_inches)
        bottoms = [0.0] * len(spec.labels)
        stacks: list[list[Any]] = [[] for _ in positions]
        for series, tone in zip(spec.series, tones, strict=True):
            patches = ax.bar(
                positions, series.values, bottom=bottoms, width=width, color=tone, linewidth=0
            )
            self.bars.extend(patches)
            for i, (value, patch) in enumerate(zip(series.values, patches, strict=True)):
                stacks[i].append(patch)
                self.stack_of[id(patch)] = i
                self.mark([patch], spec.labels[i], series.name)
                if value == 0:
                    continue
                text = self.fmt(value)
                spots = [
                    _Spot(
                        (i, bottoms[i] + value / 2),
                        text,
                        (0.0, 0.0),
                        "center",
                        "center",
                        size - 0.5,
                        on_fill(tone),
                        rule="inside",
                        anchor=patch,
                    )
                ]
                for side, ha in ((1, "left"), (-1, "right")):
                    # A point inside the segment's edge keeps the text clear of
                    # the zero line and of the next segment's edge.
                    for va, y, lift in (
                        ("center", bottoms[i] + value / 2, 0.0),
                        ("bottom", bottoms[i], 1.0),
                        ("top", bottoms[i] + value, -1.0),
                    ):
                        spots.append(
                            _Spot(
                                (i + side * width / 2, y),
                                text,
                                (side * 2.0, lift),
                                ha,
                                va,
                                size - 0.5,
                                self.tones.text,
                                rule="beside",
                                anchor=patch,
                                slot=(i - 0.5, i + 0.5),
                            )
                        )
                self.add_label(text, spec.labels[i], series.name, spots, [patch], i)
            bottoms = [b + v for b, v in zip(bottoms, series.values, strict=True)]
        if spec.totals:
            for i, total in enumerate(totals):
                first = self.end_spot(
                    i, total, self.fmt(sums[i]), False, size + 0.5, bold=True, anchor=stacks[i][-1]
                )
                spots = [first, self.outward(first, False)]
                self.add_label(
                    self.fmt(sums[i]), spec.labels[i], STYLE.CHART_TOTAL_LABEL, spots, stacks[i]
                )
        self.plain_axes()
        self.baseline()
        self.categories(positions, spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        ax.set_ylim(0, top * 1.12)
        self.patch_legend(tones, [s.name for s in spec.series])

    # ── waterfall ────────────────────────────────────────────────────────────

    def draw_waterfall(self) -> None:
        spec, ax, tones = self.spec, self.ax, self.tones
        series = spec.series[0]
        values = series.values
        bottoms, tops, _ = waterfall_positions(values)
        # The total written is the exact sum of the values as written, not
        # the running sum's rounding error: `[0.3, -0.1, -0.2]` totals 0.
        total_written = exact_sum(values, self.scale)
        total = float(total_written) * self.scale
        anchor = tones.anchor()
        width = 0.6
        cumulative = 0.0
        for index, value in enumerate(values):
            height = tops[index] - bottoms[index]
            if index == 0 and spec.start_total:
                kwargs: dict[str, Any] = {"color": anchor}
            elif value >= 0:
                kwargs = {"color": tones.accent}
            else:
                kwargs = {"color": tones.negative, "hatch": "////", "edgecolor": "#FFFFFF"}
            patches = ax.bar(
                index, height, bottom=bottoms[index], width=width, linewidth=0, **kwargs
            )
            self.bars.extend(patches)
            self.mark(patches, spec.labels[index], series.name)
            cumulative += value
            end = tops[index] if value >= 0 else bottoms[index]
            first = self.end_spot(index, end, self.fmt(value), value < 0, anchor=patches[0])
            spots = [first, self.outward(first, value < 0)]
            self.add_label(self.fmt(value), spec.labels[index], series.name, spots, patches)
            ax.plot(
                [index + width / 2, index + 1 - width / 2],
                [cumulative, cumulative],
                color=tones.muted,
                linewidth=0.6,
            )
        total_x = len(values)
        total_label = spec.total_label or STYLE.CHART_TOTAL_LABEL
        patches = ax.bar(
            total_x, abs(total), bottom=min(0.0, total), width=width, color=anchor, linewidth=0
        )
        self.bars.extend(patches)
        self.mark(patches, total_label, series.name)
        first = self.end_spot(
            total_x, total, self.fmt(total_written), total < 0, bold=True, anchor=patches[0]
        )
        spots = [first, self.outward(first, total < 0)]
        self.add_label(self.fmt(total_written), total_label, series.name, spots, patches)
        labels = [*spec.labels, total_label]
        self.plain_axes()
        self.baseline()
        self.categories(range(len(labels)), labels)
        ax.set_xlim(-0.55, len(labels) - 0.45)
        self.value_limits(min([*bottoms, total]), max([*tops, total]))

    # ── line ─────────────────────────────────────────────────────────────────

    def draw_line(self) -> None:
        spec, ax = self.spec, self.ax
        count, points = len(spec.series), len(spec.labels)
        self.warn_many(count)
        accent = self.accent_index()
        tones = self.tones.compare(count, accent) if count > 1 else [self.tones.accent]
        xs = list(range(points))
        markers = points <= MANY_POINTS
        order = [i for i in range(count) if i != accent] + [accent]
        lines: dict[int, Any] = {}
        for index in order:
            series = spec.series[index]
            main = index == accent
            (lines[index],) = ax.plot(
                xs,
                series.values,
                color=tones[index],
                linewidth=2.0 if main else 1.4,
                marker="o" if markers else None,
                markersize=4 if main else 3,
                zorder=4 if main else 3,
            )
            lines[index].set_gid(f"line\t{index}")
            self.points.extend(
                (ax, x, value, lines[index], x, self.fmt(value))
                for x, value in zip(xs, series.values, strict=True)
            )
        main_values = spec.series[accent].values
        others = [s.values for i, s in enumerate(spec.series) if i != accent]
        # The last value is written with the line's name at its end.
        every = points <= LABELLED_POINTS
        label_at = range(points - 1) if every else range(min(1, points - 1))
        for i in label_at:
            value = main_values[i]
            below = any(other[i] > value for other in others)
            text = self.fmt(value)
            # 4 pt clears the point's own marker (2.5 pt with its edge).
            first = self.end_spot(i, value, text, below, VALUE_SIZE, self.tones.accent, True, 4.0)
            flipped = self.end_spot(
                i, value, text, not below, VALUE_SIZE, self.tones.accent, True, 4.0
            )
            spots = [first, flipped, self.outward(first, below)]
            self.add_label(
                text, spec.labels[i], spec.series[accent].name, spots, point=(lines[accent], i)
            )
        for line in spec.reference_lines:
            # Over the points only, so that it does not run through the
            # line-end names beside the last point.
            ax.plot(
                [0.0, points - 1.0] if points > 1 else [-0.4, 0.4],
                [line.value, line.value],
                color=self.tones.muted,
                linewidth=1.0,
                linestyle=(0, (4, 2)),
            )
            self.add_reference(line.name, line.value, (0, line.value), "left")
        self.plain_axes(value_axis=True, grid=True)
        step = math.ceil(points / 8) if points > MANY_POINTS else 1
        self.categories(xs, spec.labels, step)
        if not markers:
            ax.tick_params(axis="x", length=3, color=self.tones.text)
        ax.spines["bottom"].set_visible(True)
        ax.spines["bottom"].set_color(self.tones.text)
        ax.spines["bottom"].set_linewidth(0.8)
        all_values = [v for s in spec.series for v in s.values]
        all_values += [line.value for line in spec.reference_lines]
        low, high = min(all_values), max(all_values)
        span = (high - low) or abs(high) or 1.0
        ax.set_ylim(low - span * 0.12, high + span * 0.16)
        ax.set_xlim(-0.4, points - 0.6)
        self.ends = [
            (
                spec.series[i].values[-1],
                self.end_text(spec.series[i].values[-1], spec.series[i].name),
                tones[i],
                i == accent,
                spec.series[i].name,
                lines[i],
            )
            for i in range(count)
        ]
        # Reserve the line ends' width beside the plot until they are settled.
        last_x = points - 1
        for value, text, _, main, _, _ in self.ends:
            spot = _Spot((last_x, value), text, (6.0, 0.0), "left", "center", bold=main)
            self.provisional.append(self._make(spot))

    def end_text(self, value: float, name: str) -> str:
        """A line end's value and name; a long name is broken into lines.

        The text is at most 40 % of the picture wide (1 inch at least), so
        that one long name does not squeeze the plot; the value and the
        start of the name stay on the first line.
        """
        head = f"{self.fmt(value)}  "
        room = max(1.0, self.width * 0.4) - _text_width(head, VALUE_SIZE)
        return head + fit_text(name, max(0.5, room), VALUE_SIZE)

    def add_reference(self, name: str, value: float, xy: tuple[float, float], ha: str) -> None:
        """A reference line's name and value, at its start, else just past it."""
        text = f"{name} {self.fmt(value)}"
        color, size = self.tones.muted, VALUE_SIZE - 0.5
        spots = [
            _Spot(xy, text, (0.0, 2.5), ha, "bottom", size, color, rule="free"),
            _Spot(xy, text, (0.0, -2.5), ha, "top", size, color, rule="free"),
            _Spot(xy, text, (0.0, 2.5 + LINE_GAP), ha, "bottom", size, color, rule="free"),
        ]
        self.add_label(self.fmt(value), "", name, spots, role="reference")

    # ── bar + line ───────────────────────────────────────────────────────────

    def draw_bar_line(self) -> None:
        from matplotlib.lines import Line2D
        from matplotlib.patches import Patch

        spec, ax = self.spec, self.ax
        bars = [s for s in spec.series if s.axis == "left"]
        lines = [s for s in spec.series if s.axis == "right"]
        tones = self.tones.compare(len(bars) + len(lines), len(bars))
        bar_tones, line_tones = tones[: len(bars)], [tones[len(bars)], *tones[len(bars) + 1 :]]
        positions = list(range(len(spec.labels)))
        width = 0.58 / len(bars)
        bar_values = [v for s in bars for v in s.values]
        low, high = min(0.0, *bar_values), max(0.0, *bar_values)
        span = (high - low) or 1.0
        ax.set_ylim(low - (span * 0.1 if low < 0 else 0.0), high + span * 0.72)
        plot_points = (chart_height(spec) - 0.9) * 72
        size = _value_size(width * self.width * 0.85 / len(spec.labels))
        for index, (series, tone) in enumerate(zip(bars, bar_tones, strict=True)):
            offset = (index - (len(bars) - 1) / 2.0) * width
            xs = [x + offset for x in positions]
            patches = ax.bar(xs, series.values, width=width * 0.94, color=tone, linewidth=0)
            self.bars.extend(patches)
            for x, value, patch, label in zip(xs, series.values, patches, spec.labels, strict=True):
                self.mark([patch], label, series.name)
                text = self.fmt(value)
                inside = _Spot(
                    (x, value),
                    text,
                    (0.0, -3.0),
                    "center",
                    "top",
                    size,
                    on_fill(tone),
                    rule="inside",
                    anchor=patch,
                )
                outside = self.end_spot(x, value, text, value < 0, size, anchor=patch)
                tall = abs(value) / (span * 1.72) * plot_points >= 16 and value > 0
                spots = [inside, outside] if tall else [outside, inside]
                spots.append(self.outward(outside, value < 0))
                self.add_label(text, label, series.name, spots, [patch])
        self.plain_axes()
        self.baseline()
        self.categories(positions, spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        right = ax.twinx()
        self.axes.append(right)
        line_values = [v for s in lines for v in s.values]
        lo, hi = min(line_values), max(line_values)
        band = (hi - lo) or abs(hi) or 1.0
        right.set_ylim(lo - band * 2.2, hi + band * 0.45)
        for series, tone in zip(lines, line_tones, strict=True):
            (drawn,) = right.plot(
                positions, series.values, color=tone, linewidth=2.0, marker="o", markersize=4
            )
            # The series' place in the specification, as for a line chart.
            place = next(i for i, other in enumerate(spec.series) if other is series)
            drawn.set_gid(f"line\t{place}")
            self.points.extend(
                (right, x, value, drawn, x, self.fmt(value, right=True))
                for x, value in zip(positions, series.values, strict=True)
            )
            for x, value, label in zip(positions, series.values, spec.labels, strict=True):
                text = self.fmt(value, right=True)
                above = self.end_spot(
                    x, value, text, False, VALUE_SIZE + 0.5, tone, True, 4.0, right
                )
                below = self.end_spot(
                    x, value, text, True, VALUE_SIZE + 0.5, tone, True, 4.0, right
                )
                spots = [above, below, self.outward(above, False)]
                self.add_label(text, label, series.name, spots, point=(drawn, x))
        right.set_yticks([])
        for side in ("top", "right", "left", "bottom"):
            right.spines[side].set_visible(False)
        handles: list[Any] = [Patch(color=tone) for tone in bar_tones]
        handles += [
            Line2D([], [], color=tone, linewidth=2.0, marker="o", markersize=4)
            for tone in line_tones
        ]
        self.legend = (handles, [s.name for s in (*bars, *lines)])

    # ── range ────────────────────────────────────────────────────────────────

    def draw_range(self) -> None:
        spec, ax, tones = self.spec, self.ax, self.tones
        rows = len(spec.labels)
        ys = list(range(rows))[::-1]
        bounds = [*spec.low, *spec.high, *(line.value for line in spec.reference_lines)]
        low, high = min(bounds), max(bounds)
        span = (high - low) or abs(high) or 1.0
        pad = span * 0.02
        if spec.show_overlap and max(spec.low) < min(spec.high):
            ax.axvspan(max(spec.low), min(spec.high), color=tones.accent, alpha=0.08, linewidth=0)
        for y, bottom, top, label in zip(ys, spec.low, spec.high, spec.labels, strict=True):
            patch = None
            if top > bottom:
                patches = ax.barh(
                    y, top - bottom, left=bottom, height=0.5, color=tones.tint, linewidth=0
                )
                self.bars.extend(patches)
                self.mark(patches, label, "range")
                patch = patches[0]
            else:
                ax.plot([bottom], [y], marker="|", markersize=12, color=tones.accent)
            for value, x, ha, name in (
                (bottom, bottom - pad, "right", "low"),
                (top, top + pad, "left", "high"),
            ):
                text = self.fmt(value)
                spots = [
                    _Spot(
                        (x, y),
                        text,
                        (0.0, nudge),
                        ha,
                        "center",
                        VALUE_SIZE,
                        tones.text,
                        rule="side",
                        anchor=patch,
                    )
                    for nudge in (0.0, LINE_GAP * 0.6, -LINE_GAP * 0.6)
                ]
                if patch is None:
                    # A single value: the anchor is the point itself, and only
                    # the place level with it is on it.
                    spots = spots[:1]
                    spots[0].xy = (value, y)
                    spots[0].offset = (-3.0 if ha == "right" else 3.0, 0.0)
                self.add_label(text, label, name, spots, [patch] if patch is not None else [])
        top_y = rows - 0.5
        for line in spec.reference_lines:
            ax.plot(
                [line.value, line.value],
                [-0.5, top_y],
                color=tones.accent,
                linewidth=1.2,
                linestyle=(0, (4, 2)),
            )
            text = f"{line.name} {self.fmt(line.value)}"
            spots = [
                _Spot(
                    (line.value, top_y + 0.08),
                    text,
                    (0.0, lift),
                    "center",
                    "bottom",
                    VALUE_SIZE,
                    tones.accent,
                    True,
                    rule="free",
                )
                for lift in (0.0, LINE_GAP, 2 * LINE_GAP)
            ]
            self.add_label(self.fmt(line.value), "", line.name, spots, role="reference")
        ax.set_yticks(ys)
        ax.set_yticklabels(self.row_names(spec.labels), fontsize=CATEGORY_SIZE, color=tones.text)
        ax.set_ylim(-0.6, top_y + (0.5 if spec.reference_lines else 0.1))
        ax.set_xlim(low - span * 0.2, high + span * 0.2)
        self._row_axes()

    def _row_axes(self) -> None:
        ax = self.ax
        for side in ("top", "right", "bottom"):
            ax.spines[side].set_visible(False)
        ax.spines["left"].set_color(self.tones.text)
        ax.spines["left"].set_linewidth(0.8)
        ax.set_xticks([])
        ax.tick_params(axis="y", length=0, labelsize=CATEGORY_SIZE, colors=self.tones.text, pad=6)

    # ── horizontal bar ───────────────────────────────────────────────────────

    def draw_hbar(self) -> None:
        spec, ax, tones = self.spec, self.ax, self.tones
        series = spec.series[0]
        values = series.values
        rows = len(spec.labels)
        ys = list(range(rows))[::-1]
        grey = tones.pair_grey()
        low, high = min(0.0, *values), max(0.0, *values)
        span = (high - low) or 1.0
        for y, label, value in zip(ys, spec.labels, values, strict=True):
            marked = not spec.highlight or label in spec.highlight
            color = tones.accent if marked else grey
            patches = ax.barh(y, value, height=0.6, color=color, linewidth=0)
            self.bars.extend(patches)
            self.mark(patches, label, series.name)
            strong = bool(spec.highlight) and marked
            text = self.fmt(value)
            # A gap in points: a share of the span is far from the bar when
            # every value is 0.
            outside = _Spot(
                (value, y),
                text,
                (4.0 if value >= 0 else -4.0, 0.0),
                "left" if value >= 0 else "right",
                "center",
                VALUE_SIZE,
                tones.accent if strong else tones.text,
                strong,
                rule="side",
                anchor=patches[0],
            )
            inside = _Spot(
                (value, y),
                text,
                (-3.0 if value >= 0 else 3.0, 0.0),
                "right" if value >= 0 else "left",
                "center",
                VALUE_SIZE,
                on_fill(color),
                strong,
                rule="inside",
                anchor=patches[0],
            )
            self.add_label(text, label, series.name, [outside, inside], patches)
        ax.set_yticks(ys)
        ax.set_yticklabels(self.row_names(spec.labels), fontsize=CATEGORY_SIZE, color=tones.text)
        for tick, label in zip(ax.get_yticklabels(), spec.labels, strict=True):
            if spec.highlight and label in spec.highlight:
                tick.set_fontweight("bold")
                tick.set_color(tones.accent)
        ax.set_xlim(low - (span * 0.16 if low < 0 else 0.0), high + span * 0.16)
        ax.set_ylim(-0.6, rows - 0.4)
        self._row_axes()
        if low < 0:
            ax.axvline(0, color=tones.text, linewidth=0.8)


def _formatter(fmt: Callable[[float], str]) -> Any:
    from matplotlib.ticker import FuncFormatter

    return FuncFormatter(lambda value, _position: fmt(value))
