"""Principle V: a chart block validates to a spec or to `ChartSpecError`, nothing else.

Any YAML value in any key (the classic keys, the E10 keys, series and reference
items) either validates or raises `ChartSpecError`; no other exception escapes
the parser. A document whose charts are off converts every chart block, valid
or not, to a code panel, in strict mode too.
"""

import math
from io import BytesIO

import yaml
from docx import Document
from hypothesis import HealthCheck, example, given, seed, settings
from hypothesis import strategies as st

from chart_renderer import ChartSpecError, _ChartSchema, parse_chart
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

BASES = {
    "bar": {"type": "bar", "labels": ["A", "B"], "series": [{"name": "S", "values": [1, 2]}]},
    "line": {"type": "line", "labels": ["A", "B"], "series": [{"name": "S", "values": [1, 2]}]},
    "waterfall": {
        "type": "waterfall",
        "labels": ["A", "B"],
        "series": [{"name": "S", "values": [1, -2]}],
    },
    "stacked_bar": {
        "type": "stacked_bar",
        "labels": ["A", "B"],
        "series": [{"name": "S", "values": [1, 2]}, {"name": "T", "values": [3, 4]}],
    },
    "bar_line": {
        "type": "bar_line",
        "labels": ["A", "B"],
        "series": [
            {"name": "S", "values": [1, 2]},
            {"name": "T", "values": [3, 4], "axis": "right"},
        ],
    },
    "range": {"type": "range", "labels": ["A", "B"], "low": [1, 2], "high": [3, 4]},
    "hbar": {"type": "hbar", "labels": ["A", "B"], "series": [{"name": "S", "values": [1, 2]}]},
}
TOP_KEYS = sorted(_ChartSchema.FIELDS)
SERIES_KEYS = ["name", "values", "axis"]
REFERENCE_KEYS = ["name", "value"]

scalars = st.one_of(
    st.none(),
    st.booleans(),
    st.integers(min_value=-(10**30), max_value=10**30),
    st.sampled_from([400, -400]).map(lambda e: 10 ** abs(e) * (1 if e > 0 else -1)),
    st.floats(allow_nan=True, allow_infinity=True),
    st.text(max_size=12),
    st.sampled_from(
        ["x" * 20000, "1,000", "1e3", "nan", "inf", "-0", "0x10", " 12 ", "left", "right"]
    ),
    st.from_regex(r"-?[0-9]{1,7}(\.[0-9]{1,3})?", fullmatch=True),
)
values = st.recursive(
    scalars,
    lambda children: st.one_of(
        st.lists(children, max_size=4),
        st.dictionaries(st.sampled_from([*TOP_KEYS, *SERIES_KEYS, "x"]), children, max_size=3),
    ),
    max_leaves=8,
)


@st.composite
def chart_specs(draw):
    """A valid base spec with one to three keys replaced by arbitrary values."""
    data = yaml.safe_load(yaml.safe_dump(BASES[draw(st.sampled_from(sorted(BASES)))]))
    for _ in range(draw(st.integers(min_value=1, max_value=3))):
        where = draw(st.sampled_from(["top", "series", "reference", "item"]))
        if where == "top":
            data[draw(st.sampled_from(TOP_KEYS))] = draw(values)
        elif where == "series" and isinstance(data.get("series"), list) and data["series"]:
            item = data["series"][draw(st.integers(0, len(data["series"]) - 1))]
            if isinstance(item, dict):
                item[draw(st.sampled_from(SERIES_KEYS))] = draw(values)
        elif where == "reference":
            data["reference"] = [{key: draw(values) for key in REFERENCE_KEYS}]
        else:
            key = draw(st.sampled_from(["labels", "values", "low", "high", "highlight"]))
            target = data["series"][0] if key == "values" and data.get("series") else data
            if isinstance(target, dict):
                target[key] = draw(st.lists(values, max_size=3))
    return data


def _dump(data) -> str:
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False)


AXIS_LIST = _dump(
    {
        "type": "bar_line",
        "labels": ["A"],
        "series": [{"name": "B", "values": [1]}, {"name": "L", "values": [2], "axis": []}],
    }
)
AXIS_MAP = AXIS_LIST.replace("axis: []", "axis: {}")


@seed(20261114)
@settings(max_examples=600, deadline=None, suppress_health_check=[HealthCheck.too_slow])
@given(chart_specs())
@example(yaml.safe_load(AXIS_LIST))
@example(yaml.safe_load(AXIS_MAP))
def test_any_value_in_any_key_is_a_spec_or_a_spec_error(data):
    text = _dump(data)
    try:
        spec, _ = parse_chart(text)
    except ChartSpecError:
        return
    # A spec that validates holds only finite numbers.
    numbers = [*spec.low, *spec.high, *(v for s in spec.series for v in s.values)]
    numbers += [line.value for line in spec.reference_lines]
    assert all(math.isfinite(number) for number in numbers)


@seed(20261114)
@settings(max_examples=60, deadline=None, suppress_health_check=[HealthCheck.too_slow])
@given(chart_specs())
@example(yaml.safe_load(AXIS_LIST))
@example(yaml.safe_load(AXIS_MAP))
def test_charts_off_converts_any_chart_block_to_a_code_panel(data):
    text = _dump(data)
    markdown = "# 차트\n\n```chart\n" + text + "```\n"
    model = MarkdownParser(profile="plain").parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=False, strict=True))
    doc = renderer.render(model)
    buffer = BytesIO()
    doc.save(buffer)
    saved = Document(BytesIO(buffer.getvalue()))
    assert not saved.inline_shapes
    written = "\n".join(paragraph.text for paragraph in saved.paragraphs)
    for table in saved.tables:
        for row in table.rows:
            written += "\n" + "\n".join(cell.text for cell in row.cells)
    first = next(line for line in text.splitlines() if line.strip())
    assert first.strip()[:20] in written


def test_axis_list_and_mapping_are_spec_errors():
    for text in (AXIS_LIST, AXIS_MAP):
        try:
            parse_chart(text)
        except ChartSpecError as error:
            assert "axis must be left or right" in str(error)
        else:
            raise AssertionError("axis: [] validated")


def test_numbers_too_long_to_print_are_spec_errors():
    """A 5,000-digit integer cannot be turned into text (Python's digit limit)."""
    digits = "9" * 5000
    series = "series:\n  - name: S\n    values: [{v}]\n"
    for text in (
        f"type: bar\nlabels: [{digits}]\n" + series.format(v=1),
        "type: bar\nlabels: [A]\n" + series.format(v=digits),
        "type: line\nlabels: [A]\n"
        + series.format(v=1)
        + f"reference:\n  - name: R\n    value: {digits}\n",
    ):
        try:
            parse_chart(text)
        except ChartSpecError:
            continue
        raise AssertionError("a 5,000-digit number validated")
