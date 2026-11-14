"""Validated YAML charts using isolated Agg figures and request-scoped styles.

Changelog (feature port):
    - Preserve PR #5 chart_type/y_label and cumulative waterfall semantics.
    - Add type alias, unit and presentation-only number formats.
    - Avoid pyplot, rcParams mutation, temporary files and cached theme values.
    - Resolve installed raster fonts and diagnose missing CJK glyph coverage.

Changelog (E10, IB/consulting charts):
    - Validate stacked_bar, bar_line, range and hbar and the style keys; the
      classic drawing below is unchanged and draws only `is_classic_spec` charts.
"""

import math
import re
from io import BytesIO
from typing import Any

import yaml

from document_model import ChartReferenceLine, ChartSeries, ChartSpec
from input_safety import load_yaml
from render_styles import STYLE, RasterFontPolicy


class ChartSpecError(ValueError):
    """Invalid chart data with a field-level diagnostic."""


class _ChartSchema:
    """Reusable syntax patterns; no style values are captured here."""

    NUMBER_FORMAT = re.compile(r"^,?(?:\.[0-9]{1,2})?f$")
    # The types the engine drew before the IB/consulting style; only these
    # can be drawn in the classic style.
    CLASSIC_TYPES = ("bar", "line", "waterfall")
    TYPES = (*CLASSIC_TYPES, "stacked_bar", "bar_line", "range", "hbar")
    CLASSIC_FIELDS = {
        "type",
        "chart_type",
        "title",
        "labels",
        "series",
        "y_label",
        "unit",
        "source",
        "total_label",
        "number_format",
    }
    # Keys that only some types take, and those types.
    APPLIES = {
        "series": set(TYPES) - {"range"},
        "highlight": {"bar", "hbar"},
        "accent_series": {"bar", "line"},
        "reference": {"line", "range"},
        "totals": {"stacked_bar"},
        "start_total": {"waterfall"},
        "low": {"range"},
        "high": {"range"},
        "show_overlap": {"range"},
        "right_number_format": {"bar_line"},
    }
    FIELDS = CLASSIC_FIELDS | {"subtitle", "note"} | set(APPLIES)
    FORMATS = {"number", "percent", "bps", "multiple"}
    # More stacked series cannot be told apart in black and white.
    MAX_STACKED_SERIES = 4
    MAX_REFERENCE_LINES = 2
    MAX_BAR_LINE_SERIES_PER_AXIS = 2


def _optional_string(data: dict[str, Any], key: str) -> str | None:
    value = data.get(key)
    if value is not None and not isinstance(value, str):
        raise ChartSpecError(f"{key} must be a string or null")
    return value


def _optional_bool(data: dict[str, Any], key: str) -> bool:
    value = data.get(key, False)
    if not isinstance(value, bool):
        raise ChartSpecError(f"{key} must be true or false")
    return value


def _finite_number(value: Any, field: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ChartSpecError(f"{field} must be numeric")
    try:
        number = float(value)
    except OverflowError as exc:
        raise ChartSpecError(f"{field} must be finite") from exc
    if not math.isfinite(number):
        raise ChartSpecError(f"{field} must be finite")
    return number


def _number_list(value: Any, field: str, length: int) -> list[float]:
    if not isinstance(value, list) or len(value) != length:
        raise ChartSpecError(f"{field} must be a list with labels length")
    return [_finite_number(item, f"{field}[{index}]") for index, item in enumerate(value)]


def _check_number_format(value: str | None, key: str) -> None:
    if value is not None and value not in _ChartSchema.FORMATS:
        if not _ChartSchema.NUMBER_FORMAT.fullmatch(value):
            raise ChartSpecError(f"{key} must be number, percent, bps, multiple or e.g. ',.2f'")


def parse_chart_spec(text: str) -> ChartSpec:
    """Validate the schema and its documented aliases without rescaling.

    Args:
        text: YAML from a chart fence.

    Returns:
        Chart data in the original units.

    Raises:
        ChartSpecError: Invalid YAML, fields or values, including the title if known.
    """
    try:
        data = load_yaml(text)
    except (yaml.YAMLError, ValueError) as exc:
        raise ChartSpecError(f"YAML is invalid: {exc}") from exc
    if not isinstance(data, dict):
        raise ChartSpecError("spec must be a YAML mapping")
    try:
        return _validate_spec(data)
    except (ValueError, OverflowError) as exc:
        title = data.get("title")
        name = f" {title!r}:" if isinstance(title, str) and title else ""
        raise ChartSpecError(f"Chart{name} {exc}") from exc


def is_classic_spec(spec: ChartSpec) -> bool:
    """Whether the classic style can draw `spec`: a first-three type, no newer key.

    Args:
        spec: Validated chart data.

    Returns:
        True when every key the chart uses existed before the IB/consulting style.
    """
    return spec.chart_type in _ChartSchema.CLASSIC_TYPES and spec == ChartSpec(
        chart_type=spec.chart_type,
        title=spec.title,
        labels=spec.labels,
        series=[ChartSeries(series.name, series.values) for series in spec.series],
        y_label=spec.y_label,
        source=spec.source,
        total_label=spec.total_label,
        unit=spec.unit,
        number_format=spec.number_format,
    )


def _validate_spec(data: dict[str, Any]) -> ChartSpec:
    unknown = set(data) - _ChartSchema.FIELDS
    if unknown:
        raise ChartSpecError("unknown field(s): " + ", ".join(sorted(map(str, unknown))))
    if "type" in data and "chart_type" in data and data["type"] != data["chart_type"]:
        raise ChartSpecError("type and chart_type conflict")
    kind = data.get("chart_type", data.get("type"))
    if not isinstance(kind, str) or kind not in _ChartSchema.TYPES:
        raise ChartSpecError("chart_type (or type) must be one of " + ", ".join(_ChartSchema.TYPES))
    for key, kinds in _ChartSchema.APPLIES.items():
        if key in data and kind not in kinds:
            raise ChartSpecError(
                f"{key} applies only to chart_type "
                + ", ".join(t for t in _ChartSchema.TYPES if t in kinds)
            )
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
    series = [] if kind == "range" else _validate_series(data.get("series"), kind, len(labels))
    low: list[float] = []
    high: list[float] = []
    if kind == "range":
        low = _number_list(data.get("low"), "low", len(labels))
        high = _number_list(data.get("high"), "high", len(labels))
        for index, (bottom, top) in enumerate(zip(low, high, strict=True)):
            if bottom > top:
                raise ChartSpecError(f"low[{index}] must not be greater than high[{index}]")
    strings = {
        key: _optional_string(data, key)
        for key in (
            "title",
            "y_label",
            "unit",
            "source",
            "total_label",
            "number_format",
            "subtitle",
            "note",
            "accent_series",
            "right_number_format",
        )
    }
    _check_number_format(strings["number_format"], "number_format")
    _check_number_format(strings["right_number_format"], "right_number_format")
    highlight = _validate_highlight(data, labels, series)
    accent = strings.pop("accent_series")
    if accent is not None and accent not in {item.name for item in series}:
        raise ChartSpecError(f"accent_series {accent!r} names no series")
    references = _validate_reference_lines(data.get("reference"))
    if kind == "waterfall":
        waterfall_positions(series[0].values)
    return ChartSpec(
        chart_type=kind,
        labels=labels,
        series=series,
        highlight=highlight,
        accent_series=accent,
        reference_lines=references,
        totals=_optional_bool(data, "totals"),
        start_total=_optional_bool(data, "start_total"),
        low=low,
        high=high,
        show_overlap=_optional_bool(data, "show_overlap"),
        **strings,
    )


def _validate_series(raw_series: Any, kind: str, length: int) -> list[ChartSeries]:
    if not isinstance(raw_series, list) or not raw_series:
        raise ChartSpecError("series must contain at least one series")
    if kind in {"waterfall", "hbar"} and len(raw_series) != 1:
        raise ChartSpecError(f"series must contain exactly one series for {kind}")
    if kind == "stacked_bar" and len(raw_series) > _ChartSchema.MAX_STACKED_SERIES:
        raise ChartSpecError(
            f"stacked_bar takes at most {_ChartSchema.MAX_STACKED_SERIES} series; "
            "group the smaller ones into one"
        )
    keys = {"name", "values", "axis"} if kind == "bar_line" else {"name", "values"}
    series = []
    for index, item in enumerate(raw_series):
        field = f"series[{index}]"
        if not isinstance(item, dict):
            raise ChartSpecError(f"{field} must be a mapping")
        unknown = set(item) - keys
        if unknown:
            raise ChartSpecError(f"{field} has unknown fields: {list(map(str, unknown))}")
        if not isinstance(item.get("name"), str):
            raise ChartSpecError(f"{field}.name must be a string")
        values = _number_list(item.get("values"), f"{field}.values", length)
        if kind == "stacked_bar" and any(value < 0 for value in values):
            raise ChartSpecError(f"{field}.values must not be negative in stacked_bar")
        axis = item.get("axis", "left")
        if axis not in {"left", "right"}:
            raise ChartSpecError(f"{field}.axis must be left or right")
        series.append(ChartSeries(item["name"], values, axis))
    if kind == "bar_line":
        counts = [sum(item.axis == side for item in series) for side in ("left", "right")]
        limit = _ChartSchema.MAX_BAR_LINE_SERIES_PER_AXIS
        if not all(1 <= count <= limit for count in counts):
            raise ChartSpecError(
                f"bar_line needs 1 to {limit} bar series and 1 to {limit} series with axis: right"
            )
    return series


def _validate_highlight(
    data: dict[str, Any], labels: list[str], series: list[ChartSeries]
) -> list[str]:
    if "highlight" not in data:
        return []
    value = data["highlight"]
    if not isinstance(value, list) or not value or not all(isinstance(v, str) for v in value):
        raise ChartSpecError("highlight must be a non-empty list of labels")
    if len(series) != 1:
        raise ChartSpecError("highlight needs a single series; use accent_series to pick a series")
    for label in value:
        if label not in labels:
            raise ChartSpecError(f"highlight {label!r} is not one of the labels")
    if len(set(value)) != len(value):
        raise ChartSpecError("highlight repeats a label")
    return list(value)


def _validate_reference_lines(value: Any) -> list[ChartReferenceLine]:
    if value is None:
        return []
    limit = _ChartSchema.MAX_REFERENCE_LINES
    if not isinstance(value, list) or not 1 <= len(value) <= limit:
        raise ChartSpecError(f"reference must be a list of 1 to {limit} lines")
    lines = []
    for index, item in enumerate(value):
        field = f"reference[{index}]"
        if not isinstance(item, dict) or set(item) != {"name", "value"}:
            raise ChartSpecError(f"{field} must be a mapping with name and value")
        if not isinstance(item["name"], str):
            raise ChartSpecError(f"{field}.name must be a string")
        lines.append(
            ChartReferenceLine(item["name"], _finite_number(item["value"], f"{field}.value"))
        )
    return lines


def waterfall_positions(values: list[float]) -> tuple[list[float], list[float], float]:
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


def format_value(value: float, number_format: str | None = None) -> str:
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
    spec: ChartSpec,
    width_inches: float = 6.0,
    height_inches: float = 3.2,
    dpi: int = 200,
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

    texts = [spec.title or "", *spec.labels, spec.y_label or spec.unit or "", spec.source or ""]
    if len(spec.series) > 1:
        texts.extend(series.name for series in spec.series)
    if spec.chart_type == "waterfall":
        texts.append(spec.total_label or "Total")
    font = FontProperties(family=RasterFontPolicy.resolve(STYLE.KOREAN_FONT, "\n".join(texts)))
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
            ax.annotate(
                format_value(value, spec.number_format),
                (x, end),
                xytext=(0, offset),
                textcoords="offset points",
                ha="center",
                va=align,
                fontsize=8,
                color=dark,
            )

        if spec.chart_type == "bar":
            width = 0.8 / len(spec.series)
            for index, series in enumerate(spec.series):
                offset = (index - (len(spec.series) - 1) / 2.0) * width
                xs = [x + offset for x in positions]
                ax.bar(xs, series.values, width=width, color=palette[index % 3], label=series.name)
                for x, value in zip(xs, series.values, strict=False):
                    annotate(x, value, value)
        elif spec.chart_type == "line":
            for index, series in enumerate(spec.series):
                ax.plot(
                    positions,
                    series.values,
                    color=palette[index % 3],
                    label=series.name,
                    marker="o",
                    linewidth=1.8,
                    markersize=4,
                )
        elif spec.chart_type == "waterfall":
            values = spec.series[0].values
            bottoms, tops, total = waterfall_positions(values)
            heights = [top - bottom for bottom, top in zip(bottoms, tops, strict=True)]
            colors = [navy if value >= 0 else f"#{STYLE.CHART_NEGATIVE_COLOR}" for value in values]
            ax.bar(positions, heights, bottom=bottoms, width=0.65, color=colors)
            cumulative = 0.0
            for index, value in enumerate(values):
                cumulative += value
                annotate(index, cumulative, value)
                ax.plot(
                    [index + 0.325, index + 0.675],
                    [cumulative, cumulative],
                    color=medium,
                    linestyle=":",
                    linewidth=1.0,
                )
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
        ax.yaxis.set_major_formatter(
            FuncFormatter(lambda value, pos: format_value(value, spec.number_format))
        )
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
            fig.savefig(
                buffer, format="png", dpi=dpi, facecolor=fig.get_facecolor(), edgecolor="none"
            )
            return buffer.getvalue()
    finally:
        fig.clear()
