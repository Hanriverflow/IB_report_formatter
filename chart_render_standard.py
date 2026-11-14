"""IB/consulting style chart drawing (`chart_style: standard`, the default).

The picture holds only the plot: the title, subtitle, note and source are Word
paragraphs written by `ib_render_media.ChartRenderer`. Rules (E10, owner
decisions 2026-11-14): one accent colour (the theme's primary colour) with
greys, values written on the marks instead of a value axis, a zero baseline
and no frame, the legend in one row above the plot, line names at the line
ends, decreases hatched, and tones chosen so that series stay apart in black
and white. Values are formatted with `chart_renderer.format_value` and never
rescaled; the only numbers the engine works out are the waterfall's running
total and, when the spec asks with `totals: true`, a stack's total.

Changelog (E10):
    - NEW: stacked_bar, bar_line, range and hbar; bar, line and waterfall
      redrawn in the same style; tones chosen by grayscale distance.
"""

import math
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from io import BytesIO
from itertools import combinations
from typing import Any

from chart_renderer import ChartSpecError, format_value, waterfall_positions
from document_model import ChartSpec
from render_styles import STYLE, RasterFontPolicy

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


def chart_height(spec: ChartSpec) -> float:
    """Picture height in inches: fixed, or by rows for horizontal charts."""
    if spec.chart_type in ROW_HEIGHT:
        rows = len(spec.labels)
        extra = 0.25 if spec.reference_lines else 0.0
        return min(MAX_HEIGHT, max(1.3, ROW_HEIGHT[spec.chart_type] * rows + 0.45 + extra))
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

    Args:
        spec: Validated chart data.
        width_inches: Picture width, bounded by the Word content area by its caller.
        dpi: PNG resolution.
        warnings: Receives non-rejecting notes (too many series to tell apart).

    Returns:
        PNG bytes ready for insertion from memory.
    """
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.text import Text

    family = RasterFontPolicy.resolve(STYLE.KOREAN_FONT, "\n".join(chart_texts(spec)))
    fig = Figure(figsize=(width_inches, chart_height(spec)), dpi=dpi, facecolor="#FFFFFF")
    FigureCanvasAgg(fig)
    try:
        ax = fig.add_subplot(111)
        ax.set_facecolor("#FFFFFF")
        drawer = _Drawer(fig, ax, spec, ChartTones.from_style(), width_inches, warnings, family)
        draw: Callable[[], None] | None = getattr(drawer, "draw_" + spec.chart_type, None)
        if draw is None:
            raise ChartSpecError(f"chart_type is invalid: {spec.chart_type}")
        draw()
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
        # Line-end names, placed again once the layout has settled.
        self.ends: list[tuple[float, str, str, bool]] = []
        self.end_artists: list[Any] = []
        self.label_name = spec.title or spec.chart_type

    # ── shared pieces ────────────────────────────────────────────────────────

    def fmt(self, value: float, right: bool = False) -> str:
        number_format = self.spec.right_number_format if right else self.spec.number_format
        return format_value(value, number_format)

    def warn_many(self, count: int) -> None:
        if count > MAX_DISTINCT_SERIES and self.warnings is not None:
            self.warnings.append(
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
            wrap_label(label, slot * step) if index % step == 0 else ""
            for index, label in enumerate(labels)
        ]
        self.ax.set_xticks(list(positions))
        self.ax.set_xticklabels(shown, fontsize=CATEGORY_SIZE, color=self.tones.text)

    def value_label(
        self,
        x: float,
        y: float,
        text: str,
        below: bool,
        size: float = VALUE_SIZE,
        color: str | None = None,
        bold: bool = False,
    ) -> None:
        self.ax.annotate(
            text,
            (x, y),
            xytext=(0, -2.5 if below else 2.5),
            textcoords="offset points",
            ha="center",
            va="top" if below else "bottom",
            fontsize=size,
            color=color or "#202020",
            fontweight="bold" if bold else "normal",
        )

    def value_limits(self, low: float, high: float, head: float = 0.14) -> None:
        low, high = min(0.0, low), max(0.0, high)
        span = (high - low) or 1.0
        self.ax.set_ylim(low - (span * head if low < 0 else 0.0), high + span * head)

    def finish(self) -> None:
        """Fit the plot inside the picture and put the legend row above it."""
        top = 1.0
        if self.legend is not None:
            top = 1.0 - LEGEND_BAND / self.fig.get_figheight()
        self.fig.tight_layout(pad=0.3, rect=(0.0, 0.0, 1.0, top))
        if self.ends:
            # The first placement reserved room; place again on the final axes.
            for artist in self.end_artists:
                artist.remove()
            self.place_end_labels()
        if self.legend is not None:
            from matplotlib.font_manager import FontProperties

            handles, names = self.legend
            left = self.ax.get_position().x0
            self.fig.legend(
                handles,
                names,
                loc="upper left",
                bbox_to_anchor=(left, 1.0),
                ncol=len(handles),
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

    def patch_legend(self, colors: Sequence[str], names: Sequence[str]) -> None:
        from matplotlib.patches import Patch

        self.legend = ([Patch(color=color) for color in colors], list(names))

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
            ax.bar(xs, series.values, width=width * 0.94, color=colors[index], linewidth=0)
            for x, value, color in zip(xs, series.values, colors[index], strict=True):
                bold = count == 1 and spec.highlight and color == self.tones.accent
                self.value_label(
                    x,
                    value,
                    self.fmt(value),
                    value < 0,
                    size,
                    self.tones.accent if bold else None,
                    bool(bold),
                )
        self.plain_axes()
        self.baseline()
        self.categories(range(len(spec.labels)), spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        self.value_limits(min(values), max(values))
        if count > 1:
            self.patch_legend([c[0] for c in colors], [s.name for s in spec.series])

    # ── stacked bar ──────────────────────────────────────────────────────────

    def draw_stacked_bar(self) -> None:
        spec, ax = self.spec, self.ax
        tones = self.tones.stack(len(spec.series))
        positions = range(len(spec.labels))
        totals = [math.fsum(s.values[i] for s in spec.series) for i in positions]
        top = max(totals) or 1.0
        width = 0.62
        bar_inches = width * self.width * 0.85 / len(spec.labels)
        size = _value_size(bar_inches)
        # Points per data unit, to keep a label only in a segment tall enough.
        plot_points = (chart_height(spec) - 0.9) * 72
        bottoms = [0.0] * len(spec.labels)
        for series, tone in zip(spec.series, tones, strict=True):
            ax.bar(positions, series.values, bottom=bottoms, width=width, color=tone, linewidth=0)
            for i, value in enumerate(series.values):
                tall = value / (top * 1.12) * plot_points >= size * 1.6
                if value > 0 and value >= top * 0.09 and tall:
                    ax.text(
                        i,
                        bottoms[i] + value / 2,
                        self.fmt(value),
                        ha="center",
                        va="center",
                        fontsize=size - 0.5,
                        color=on_fill(tone),
                    )
            bottoms = [b + v for b, v in zip(bottoms, series.values, strict=True)]
        if spec.totals:
            for i, total in enumerate(totals):
                self.value_label(i, total, self.fmt(total), False, size + 0.5, bold=True)
        self.plain_axes()
        self.baseline()
        self.categories(positions, spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        ax.set_ylim(0, top * 1.12)
        self.patch_legend(tones, [s.name for s in spec.series])

    # ── waterfall ────────────────────────────────────────────────────────────

    def draw_waterfall(self) -> None:
        spec, ax, tones = self.spec, self.ax, self.tones
        values = spec.series[0].values
        bottoms, tops, total = waterfall_positions(values)
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
            ax.bar(index, height, bottom=bottoms[index], width=width, linewidth=0, **kwargs)
            cumulative += value
            end = tops[index] if value >= 0 else bottoms[index]
            self.value_label(index, end, self.fmt(value), value < 0)
            ax.plot(
                [index + width / 2, index + 1 - width / 2],
                [cumulative, cumulative],
                color=tones.muted,
                linewidth=0.6,
            )
        total_x = len(values)
        ax.bar(total_x, abs(total), bottom=min(0.0, total), width=width, color=anchor, linewidth=0)
        self.value_label(total_x, total, self.fmt(total), total < 0, bold=True)
        labels = [*spec.labels, spec.total_label or STYLE.CHART_TOTAL_LABEL]
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
        for index in order:
            series = spec.series[index]
            main = index == accent
            ax.plot(
                xs,
                series.values,
                color=tones[index],
                linewidth=2.0 if main else 1.4,
                marker="o" if markers else None,
                markersize=4 if main else 3,
                zorder=4 if main else 3,
            )
        main_values = spec.series[accent].values
        others = [s.values for i, s in enumerate(spec.series) if i != accent]
        # The last value is written with the line's name at its end.
        every = points <= LABELLED_POINTS
        label_at = range(points - 1) if every else range(min(1, points - 1))
        for i in label_at:
            value = main_values[i]
            below = any(other[i] > value for other in others)
            self.value_label(i, value, self.fmt(value), below, color=self.tones.accent, bold=True)
        for line in spec.reference_lines:
            ax.axhline(line.value, color=self.tones.muted, linewidth=1.0, linestyle=(0, (4, 2)))
            ax.annotate(
                f"{line.name} {self.fmt(line.value)}",
                (0, line.value),
                xytext=(0, 2.5),
                textcoords="offset points",
                ha="left",
                va="bottom",
                fontsize=VALUE_SIZE - 0.5,
                color=self.tones.muted,
            )
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
                f"{self.fmt(spec.series[i].values[-1])}  {spec.series[i].name}",
                tones[i],
                i == accent,
            )
            for i in range(count)
        ]
        self.place_end_labels()

    def place_end_labels(self) -> None:
        """Write each line's name after its last point, kept 11 points apart."""
        ax = self.ax
        items = self.ends
        transform = ax.transData
        inverse = transform.inverted()
        last_x = len(self.spec.labels) - 1
        placed = sorted(
            (transform.transform((last_x, value))[1], text, color, main)
            for value, text, color, main in items
        )
        gap = 11 * self.fig.dpi / 72
        for index in range(1, len(placed)):
            if placed[index][0] - placed[index - 1][0] < gap:
                placed[index] = (placed[index - 1][0] + gap, *placed[index][1:])
        self.end_artists = []
        for pixel_y, text, color, main in placed:
            y = inverse.transform((0, pixel_y))[1]
            artist = ax.annotate(
                text,
                (last_x, y),
                xytext=(6, 0),
                textcoords="offset points",
                ha="left",
                va="center",
                fontsize=VALUE_SIZE,
                color=color if main else self.tones.text,
                fontweight="bold" if main else "normal",
                annotation_clip=False,
                fontfamily=self.family,
            )
            self.end_artists.append(artist)

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
            ax.bar(xs, series.values, width=width * 0.94, color=tone, linewidth=0)
            for x, value in zip(xs, series.values, strict=True):
                inside = abs(value) / (span * 1.72) * plot_points >= 16 and value > 0
                if inside:
                    ax.annotate(
                        self.fmt(value),
                        (x, value),
                        xytext=(0, -3),
                        textcoords="offset points",
                        ha="center",
                        va="top",
                        fontsize=size,
                        color=on_fill(tone),
                    )
                else:
                    self.value_label(x, value, self.fmt(value), value < 0, size)
        self.plain_axes()
        self.baseline()
        self.categories(positions, spec.labels)
        ax.set_xlim(-0.55, len(spec.labels) - 0.45)
        right = ax.twinx()
        line_values = [v for s in lines for v in s.values]
        lo, hi = min(line_values), max(line_values)
        band = (hi - lo) or abs(hi) or 1.0
        right.set_ylim(lo - band * 2.2, hi + band * 0.45)
        for series, tone in zip(lines, line_tones, strict=True):
            right.plot(
                positions, series.values, color=tone, linewidth=2.0, marker="o", markersize=4
            )
            for x, value in zip(positions, series.values, strict=True):
                right.annotate(
                    self.fmt(value, right=True),
                    (x, value),
                    xytext=(0, 4),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    fontsize=VALUE_SIZE + 0.5,
                    fontweight="bold",
                    color=tone,
                )
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
        for y, bottom, top in zip(ys, spec.low, spec.high, strict=True):
            if top > bottom:
                ax.barh(y, top - bottom, left=bottom, height=0.5, color=tones.tint, linewidth=0)
            else:
                ax.plot([bottom], [y], marker="|", markersize=12, color=tones.accent)
            ax.text(
                bottom - pad,
                y,
                self.fmt(bottom),
                ha="right",
                va="center",
                fontsize=VALUE_SIZE,
                color=tones.text,
            )
            ax.text(
                top + pad,
                y,
                self.fmt(top),
                ha="left",
                va="center",
                fontsize=VALUE_SIZE,
                color=tones.text,
            )
        top_y = rows - 0.5
        for line in spec.reference_lines:
            ax.plot(
                [line.value, line.value],
                [-0.5, top_y],
                color=tones.accent,
                linewidth=1.2,
                linestyle=(0, (4, 2)),
            )
            ax.text(
                line.value,
                top_y + 0.08,
                f"{line.name} {self.fmt(line.value)}",
                ha="center",
                va="bottom",
                fontsize=VALUE_SIZE,
                fontweight="bold",
                color=tones.accent,
            )
        ax.set_yticks(ys)
        ax.set_yticklabels(spec.labels, fontsize=CATEGORY_SIZE, color=tones.text)
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
        values = spec.series[0].values
        rows = len(spec.labels)
        ys = list(range(rows))[::-1]
        grey = tones.pair_grey()
        low, high = min(0.0, *values), max(0.0, *values)
        span = (high - low) or 1.0
        for y, label, value in zip(ys, spec.labels, values, strict=True):
            marked = not spec.highlight or label in spec.highlight
            color = tones.accent if marked else grey
            ax.barh(y, value, height=0.6, color=color, linewidth=0)
            strong = bool(spec.highlight) and marked
            ax.text(
                value + (span * 0.015 if value >= 0 else -span * 0.015),
                y,
                self.fmt(value),
                ha="left" if value >= 0 else "right",
                va="center",
                fontsize=VALUE_SIZE,
                fontweight="bold" if strong else "normal",
                color=tones.accent if strong else tones.text,
            )
        ax.set_yticks(ys)
        ax.set_yticklabels(spec.labels, fontsize=CATEGORY_SIZE, color=tones.text)
        for tick in ax.get_yticklabels():
            if spec.highlight and tick.get_text() in spec.highlight:
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
