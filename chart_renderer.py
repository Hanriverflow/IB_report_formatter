"""Validated YAML charts using isolated Agg figures and request-scoped styles.

Changelog (feature port):
    - Preserve PR #5 chart_type/y_label and cumulative waterfall semantics.
    - Add type alias, unit and presentation-only number formats.
    - Avoid pyplot, rcParams mutation, temporary files and cached theme values.
"""

import math
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import yaml

from document_model import ChartSeries, ChartSpec
from render_styles import STYLE


class ChartSpecError(ValueError):
    """Invalid chart data with a field-level diagnostic."""


class _ChartSchema:
    """Reusable syntax patterns; no style values are captured here."""

    NUMBER_FORMAT = re.compile(r"^,?(?:\.[0-9]{1,2})?f$")
    TYPES = {"bar", "line", "waterfall"}
    FIELDS = {"type", "chart_type", "title", "labels", "series", "y_label",
              "unit", "source", "total_label", "number_format"}
    FORMATS = {"number", "percent", "bps", "multiple"}


def _optional_string(data: Dict[str, Any], key: str) -> Optional[str]:
    value = data.get(key)
    if value is not None and not isinstance(value, str):
        raise ChartSpecError(f"{key} must be a string or null")
    return value


def parse_chart_spec(text: str) -> ChartSpec:
    """Validate the legacy schema and its documented aliases without rescaling.

    Args:
        text: YAML from a chart fence.

    Returns:
        Chart data in the original units.

    Raises:
        ChartSpecError: Invalid YAML, fields or values, including the title if known.
    """
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ChartSpecError(f"YAML is invalid: {exc}") from exc
    if not isinstance(data, dict):
        raise ChartSpecError("spec must be a YAML mapping")
    try:
        return _validate_spec(data)
    except (ValueError, OverflowError) as exc:
        title = data.get("title")
        name = f" {title!r}:" if isinstance(title, str) and title else ""
        raise ChartSpecError(f"Chart{name} {exc}") from exc


def _validate_spec(data: Dict[str, Any]) -> ChartSpec:
    unknown = set(data) - _ChartSchema.FIELDS
    if unknown:
        raise ChartSpecError("unknown field(s): " + ", ".join(sorted(map(str, unknown))))
    if "type" in data and "chart_type" in data and data["type"] != data["chart_type"]:
        raise ChartSpecError("type and chart_type conflict")
    kind = data.get("chart_type", data.get("type"))
    if not isinstance(kind, str) or kind not in _ChartSchema.TYPES:
        raise ChartSpecError("chart_type (or type) must be bar, line or waterfall")
    labels_data = data.get("labels")
    if not isinstance(labels_data, list) or not labels_data:
        raise ChartSpecError("labels must be a non-empty list")
    labels = []
    for label in labels_data:
        if isinstance(label, bool) or not isinstance(label, (str, int, float)):
            raise ChartSpecError("labels must contain only strings or numbers")
        if isinstance(label, float) and not math.isfinite(label):
            raise ChartSpecError("numeric labels must be finite")
        labels.append(str(label))
    raw_series = data.get("series")
    if not isinstance(raw_series, list) or not raw_series:
        raise ChartSpecError("series must contain at least one series")
    if kind == "waterfall" and len(raw_series) != 1:
        raise ChartSpecError("series must contain exactly one series for waterfall")
    series = []
    for index, item in enumerate(raw_series):
        field = f"series[{index}]"
        if not isinstance(item, dict):
            raise ChartSpecError(f"{field} must be a mapping")
        unknown = set(item) - {"name", "values"}
        if unknown:
            raise ChartSpecError(f"{field} has unknown fields: {list(map(str, unknown))}")
        if not isinstance(item.get("name"), str):
            raise ChartSpecError(f"{field}.name must be a string")
        raw_values = item.get("values")
        if not isinstance(raw_values, list) or len(raw_values) != len(labels):
            raise ChartSpecError(f"{field}.values must be a list with labels length")
        values = []
        for value_index, value in enumerate(raw_values):
            value_field = f"{field}.values[{value_index}]"
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ChartSpecError(f"{value_field} must be numeric")
            try:
                number = float(value)
            except OverflowError as exc:
                raise ChartSpecError(f"{value_field} must be finite") from exc
            if not math.isfinite(number):
                raise ChartSpecError(f"{value_field} must be finite")
            values.append(number)
        series.append(ChartSeries(item["name"], values))
    strings = {key: _optional_string(data, key) for key in
               ("title", "y_label", "unit", "source", "total_label", "number_format")}
    number_format = strings["number_format"]
    if number_format is not None and number_format not in _ChartSchema.FORMATS:
        if not _ChartSchema.NUMBER_FORMAT.fullmatch(number_format):
            raise ChartSpecError("number_format must be number, percent, bps, multiple or e.g. ',.2f'")
    if kind == "waterfall":
        waterfall_positions(series[0].values)
    return ChartSpec(chart_type=kind, labels=labels, series=series, **strings)


def waterfall_positions(values: List[float]) -> Tuple[List[float], List[float], float]:
    """Compute PR #5's cumulative deltas and final total from an implicit zero.

    Args:
        values: Sequential changes in the supplied units.

    Returns:
        Bar bottoms, tops, and the cumulative total (including negative totals).
    """
    bottoms, tops = [], []
    cumulative = 0.0
    for value in values:
        next_total = cumulative + value
        if not math.isfinite(next_total):
            raise ChartSpecError("waterfall cumulative total must be finite")
        bottoms.append(min(cumulative, next_total))
        tops.append(max(cumulative, next_total))
        cumulative = next_total
    return bottoms, tops, cumulative


def format_value(value: float, number_format: Optional[str] = None) -> str:
    """Format a value for display; percent/bps/multiple never change its scale.

    Args:
        value: Unscaled numeric value.
        number_format: Named format or a fixed-point format such as ',.2f'.

    Returns:
        Label with parentheses for negative values, matching PR #5's default.
    """
    magnitude = abs(float(value))
    suffix = {"percent": "%", "bps": " bps", "multiple": "x"}.get(number_format or "", "")
    if number_format and _ChartSchema.NUMBER_FORMAT.fullmatch(number_format):
        text = format(magnitude, number_format)
    else:
        text = f"{int(magnitude):,}" if magnitude.is_integer() else f"{magnitude:,.1f}"
    text += suffix
    return f"({text})" if value < 0 else text


def render_chart_png(
    spec: ChartSpec, width_inches: float = 6.0, height_inches: float = 3.2, dpi: int = 200,
) -> bytes:
    """Render a chart with per-artist font/color settings and an unregistered figure.

    Args:
        spec: Validated chart specification.
        width_inches: Figure width, bounded by the Word content area by its caller.
        height_inches: Figure height.
        dpi: PNG resolution.

    Returns:
        PNG bytes ready for insertion from memory.
    """
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.font_manager import FontProperties
    from matplotlib.text import Text
    from matplotlib.ticker import FuncFormatter

    from ib_renderer import FontPolicy

    font = FontProperties(family=FontPolicy.resolve_korean_font())
    navy, dark, medium = (f"#{value}" for value in (STYLE.NAVY, STYLE.DARK_GRAY, STYLE.MEDIUM_GRAY))
    palette = [navy, dark, medium]
    fig = Figure(figsize=(width_inches, height_inches), dpi=dpi, facecolor=f"#{STYLE.WHITE}")
    FigureCanvasAgg(fig)
    try:
        ax = fig.add_subplot(111)
        ax.set_facecolor(f"#{STYLE.WHITE}")
        positions = list(range(len(spec.labels)))

        def annotate(x: float, end: float, value: float) -> None:
            offset, align = (3, "bottom") if value >= 0 else (-3, "top")
            ax.annotate(format_value(value, spec.number_format), (x, end), xytext=(0, offset),
                        textcoords="offset points", ha="center", va=align, fontsize=8, color=dark)

        if spec.chart_type == "bar":
            width = 0.8 / len(spec.series)
            for index, series in enumerate(spec.series):
                offset = (index - (len(spec.series) - 1) / 2.0) * width
                xs = [x + offset for x in positions]
                ax.bar(xs, series.values, width=width, color=palette[index % 3], label=series.name)
                for x, value in zip(xs, series.values):
                    annotate(x, value, value)
        elif spec.chart_type == "line":
            for index, series in enumerate(spec.series):
                ax.plot(positions, series.values, color=palette[index % 3], label=series.name,
                        marker="o", linewidth=1.8, markersize=4)
        elif spec.chart_type == "waterfall":
            values = spec.series[0].values
            bottoms, tops, total = waterfall_positions(values)
            heights = [top - bottom for bottom, top in zip(bottoms, tops)]
            colors = [navy if value >= 0 else f"#{STYLE.CHART_NEGATIVE_COLOR}" for value in values]
            ax.bar(positions, heights, bottom=bottoms, width=0.65, color=colors)
            cumulative = 0.0
            for index, value in enumerate(values):
                cumulative += value
                annotate(index, cumulative, value)
                ax.plot([index + 0.325, index + 0.675], [cumulative, cumulative],
                        color=medium, linestyle=":", linewidth=1.0)
            total_x = len(positions)
            ax.bar([total_x], [abs(total)], bottom=[min(0.0, total)], width=0.65, color=medium)
            annotate(total_x, total, total)
            positions.append(total_x)
        else:
            raise ChartSpecError(f"chart_type is invalid: {spec.chart_type}")

        labels = list(spec.labels)
        if spec.chart_type == "waterfall":
            labels.append(spec.total_label or "Total")
        ax.set_xticks(positions)
        ax.set_xticklabels(labels, fontsize=8)
        ax.set_title(spec.title or "", fontsize=10, fontweight="bold", color=navy, pad=10)
        ax.set_ylabel(spec.y_label or spec.unit or "", fontsize=9, color=dark)
        if len(spec.series) > 1:
            ax.legend(frameon=False, prop=FontProperties(family=font.get_family(), size=8))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda value, pos: format_value(value, spec.number_format)))
        ax.set_axisbelow(True)
        ax.yaxis.grid(True, color=medium, alpha=0.25, linewidth=0.7)
        ax.xaxis.grid(False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        for side in ("left", "bottom"):
            ax.spines[side].set_color(medium)
        ax.tick_params(axis="both", labelsize=8, colors=dark)
        ax.margins(y=0.18)
        if spec.source:
            fig.text(0.01, 0.01, f"Source: {spec.source}", fontsize=7, color=medium)
        for artist in fig.findobj(match=Text):
            artist.set_fontfamily(font.get_family())
        fig.tight_layout(rect=(0.0, 0.09 if spec.source else 0.02, 1.0, 1.0))
        with BytesIO() as buffer:
            fig.savefig(buffer, format="png", dpi=dpi, facecolor=fig.get_facecolor(), edgecolor="none")
            return buffer.getvalue()
    finally:
        fig.clear()
