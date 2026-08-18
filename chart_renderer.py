"""
Chart Renderer for IB Report Formatter
Parses YAML chart specifications and renders IB-styled PNG images.

Supports grouped bar charts, line charts, and cumulative waterfall charts.
"""

import logging
from dataclasses import dataclass
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import yaml

logger = logging.getLogger(__name__)

_CHART_TYPES = {"bar", "line", "waterfall"}
_TOP_LEVEL_FIELDS = {
    "chart_type", "title", "labels", "series", "y_label", "source", "total_label",
}


class ChartSpecError(ValueError):
    """Raised when a chart specification contains an invalid field."""


@dataclass
class ChartSeries:
    """One named sequence of numeric chart values."""

    name: str
    values: List[float]


@dataclass
class ChartSpec:
    """Validated input for an IB-styled chart."""

    chart_type: str
    title: Optional[str]
    labels: List[str]
    series: List[ChartSeries]
    y_label: Optional[str]
    source: Optional[str]
    total_label: Optional[str] = None


def _optional_string(data: Dict[str, Any], field_name: str) -> Optional[str]:
    value = data.get(field_name)
    if value is not None and not isinstance(value, str):
        raise ChartSpecError(f"{field_name} must be a string or null")
    return value


def parse_chart_spec(text: str) -> ChartSpec:
    """
    Parse and validate a YAML chart specification.

    Args:
        text: YAML text describing the chart

    Returns:
        A validated ChartSpec

    Raises:
        ChartSpecError: If any field is missing, unknown, or invalid
    """
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ChartSpecError(f"yaml is invalid: {exc}") from exc
    if not isinstance(data, dict):
        raise ChartSpecError("spec must be a YAML mapping")

    unknown_fields = [key for key in data if key not in _TOP_LEVEL_FIELDS]
    if unknown_fields:
        names = ", ".join(sorted(str(key) for key in unknown_fields))
        raise ChartSpecError(f"unknown top-level field(s): {names}")

    chart_type = data.get("chart_type")
    if not isinstance(chart_type, str) or chart_type not in _CHART_TYPES:
        raise ChartSpecError("chart_type must be one of: bar, line, waterfall")

    labels_data = data.get("labels")
    if not isinstance(labels_data, list) or not labels_data:
        raise ChartSpecError("labels must be a non-empty list")
    if any(not isinstance(label, str) for label in labels_data):
        raise ChartSpecError("labels must contain only strings")
    labels = list(labels_data)

    series_data = data.get("series")
    if not isinstance(series_data, list) or not series_data:
        raise ChartSpecError("series must contain at least one series")
    if chart_type == "waterfall" and len(series_data) != 1:
        raise ChartSpecError("series must contain exactly one series for waterfall")

    series = []
    for index, item in enumerate(series_data):
        field = f"series[{index}]"
        if not isinstance(item, dict):
            raise ChartSpecError(f"{field} must be a mapping")
        unexpected = [key for key in item if key not in {"name", "values"}]
        if unexpected:
            raise ChartSpecError(f"{field}.{unexpected[0]} is an unknown field")
        name = item.get("name")
        if not isinstance(name, str):
            raise ChartSpecError(f"{field}.name must be a string")
        values_data = item.get("values")
        if not isinstance(values_data, list):
            raise ChartSpecError(f"{field}.values must be a list")
        if len(values_data) != len(labels):
            raise ChartSpecError(f"{field}.values length must equal labels length")
        values = []
        for value_index, value in enumerate(values_data):
            value_field = f"{field}.values[{value_index}]"
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ChartSpecError(f"{value_field} must be numeric")
            values.append(float(value))
        series.append(ChartSeries(name=name, values=values))

    return ChartSpec(
        chart_type=chart_type,
        title=_optional_string(data, "title"),
        labels=labels,
        series=series,
        y_label=_optional_string(data, "y_label"),
        source=_optional_string(data, "source"),
        total_label=_optional_string(data, "total_label"),
    )


def waterfall_positions(values: List[float]) -> Tuple[List[float], List[float], float]:
    """
    Compute bottom/top coordinates and the final total for waterfall deltas.

    Args:
        values: Sequential deltas from an implicit zero baseline

    Returns:
        Lists of bar bottoms and tops, followed by the cumulative total
    """
    bottoms, tops = [], []
    cumulative = 0.0
    for value in values:
        next_total = cumulative + value
        bottoms.append(min(cumulative, next_total))
        tops.append(max(cumulative, next_total))
        cumulative = next_total
    return bottoms, tops, cumulative


def _matplotlib_available() -> bool:
    try:
        import matplotlib  # noqa: F401
    except ImportError:
        logger.warning("matplotlib is unavailable; chart rendering is disabled")
        return False
    return True


def _hex_color(rgb_color: Any) -> str:
    return f"#{rgb_color}"


def _format_value(value: float) -> str:
    magnitude = abs(value)
    text = f"{int(magnitude):,}" if magnitude.is_integer() else f"{magnitude:,.1f}"
    return f"({text})" if value < 0 else text


def _add_value_label(ax: Any, x_value: float, end_value: float, value: float) -> None:
    offset, alignment = (3, "bottom") if value >= 0 else (-3, "top")
    ax.annotate(
        _format_value(value), xy=(x_value, end_value), xytext=(0, offset),
        textcoords="offset points", ha="center", va=alignment, fontsize=8,
    )


def render_chart_png(
    spec: ChartSpec,
    width_inches: float = 6.0,
    height_inches: float = 3.2,
    dpi: int = 200,
) -> bytes:
    """
    Render a ChartSpec to PNG bytes using the current IB theme.

    Args:
        spec: Chart specification to render
        width_inches: Figure width in inches
        height_inches: Figure height in inches
        dpi: PNG resolution

    Returns:
        Complete PNG file contents
    """
    if not _matplotlib_available():
        raise RuntimeError("matplotlib is required to render charts")

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    import ib_renderer

    style = ib_renderer.STYLE
    navy, dark_gray = _hex_color(style.NAVY), _hex_color(style.DARK_GRAY)
    medium_gray, red = _hex_color(style.MEDIUM_GRAY), _hex_color(style.RED)
    additional = [dark_gray, medium_gray, _hex_color(style.ACCENT_BLUE)]
    rc_settings = {
        "font.family": [style.KOREAN_FONT, "sans-serif"],
        "axes.unicode_minus": False,
    }

    fig = None
    try:
        with matplotlib.rc_context(rc_settings):
            fig, ax = plt.subplots(figsize=(width_inches, height_inches))
            x_positions = list(range(len(spec.labels)))

            if spec.chart_type == "bar":
                bar_width = 0.8 / len(spec.series)
                for index, chart_series in enumerate(spec.series):
                    color = navy if index == 0 else additional[(index - 1) % 3]
                    offset = (index - (len(spec.series) - 1) / 2.0) * bar_width
                    bar_x = [x_value + offset for x_value in x_positions]
                    ax.bar(bar_x, chart_series.values, width=bar_width,
                           label=chart_series.name, color=color)
                    for x_value, value in zip(bar_x, chart_series.values):
                        _add_value_label(ax, x_value, value, value)
            elif spec.chart_type == "line":
                for index, chart_series in enumerate(spec.series):
                    color = navy if index == 0 else additional[(index - 1) % 3]
                    ax.plot(x_positions, chart_series.values, marker="o", linewidth=1.8,
                            markersize=4, label=chart_series.name, color=color)
            elif spec.chart_type == "waterfall":
                values = spec.series[0].values
                bottoms, tops, total = waterfall_positions(values)
                heights = [top - bottom for bottom, top in zip(bottoms, tops)]
                colors = [navy if value >= 0 else red for value in values]
                ax.bar(x_positions, heights, bottom=bottoms, width=0.65, color=colors)
                cumulative = 0.0
                for index, value in enumerate(values):
                    cumulative += value
                    _add_value_label(ax, index, cumulative, value)
                    ax.plot([index + 0.325, index + 1 - 0.325],
                            [cumulative, cumulative], color=medium_gray,
                            linestyle=":", linewidth=1.0)
                total_x = len(x_positions)
                ax.bar([total_x], [abs(total)], bottom=[min(0.0, total)],
                       width=0.65, color=medium_gray)
                _add_value_label(ax, total_x, total, total)
                x_positions.append(total_x)
            else:
                raise ChartSpecError(f"chart_type is invalid: {spec.chart_type}")

            tick_labels = list(spec.labels)
            if spec.chart_type == "waterfall":
                tick_labels.append(spec.total_label or "Total")
            ax.set_xticks(x_positions)
            ax.set_xticklabels(tick_labels, fontsize=8)
            if spec.title:
                ax.set_title(spec.title, fontsize=10, fontweight="bold", pad=10)
            if spec.y_label:
                ax.set_ylabel(spec.y_label, fontsize=9)
            if len(spec.series) > 1:
                ax.legend(frameon=False, fontsize=8)

            ax.set_axisbelow(True)
            ax.yaxis.grid(True, color=medium_gray, alpha=0.25, linewidth=0.7)
            ax.xaxis.grid(False)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_color(medium_gray)
            ax.spines["bottom"].set_color(medium_gray)
            ax.tick_params(axis="y", labelsize=8, colors=dark_gray)
            ax.tick_params(axis="x", colors=dark_gray)
            if spec.source:
                fig.text(0.01, 0.01, f"Source: {spec.source}", ha="left", va="bottom",
                         fontsize=7, color=medium_gray)

            fig.tight_layout(rect=(0.0, 0.09 if spec.source else 0.02, 1.0, 1.0))
            with BytesIO() as buffer:
                fig.savefig(buffer, format="png", dpi=dpi, facecolor="white", edgecolor="none")
                return buffer.getvalue()
    finally:
        if fig is not None:
            plt.close(fig)
