"""IB/consulting chart style (E10): specification keys, drawing and Word paragraphs.

All names and numbers are fictional.
"""

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from hypothesis import given, settings
from hypothesis import strategies as st

from chart_renderer import ChartSpecError, is_classic_spec, parse_chart_spec
from document_model import ChartReferenceLine, ChartSeries, ElementType
from document_profiles import RenderOptions, get_profile, load_style
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser
from render_styles import use_style

STACKED = """chart_type: stacked_bar
title: 2028년 만기가 가장 크다
subtitle: "연도별 만기 도래액, 단위: 억원"
labels: [2026, 2027, 2028]
series:
  - name: 공모 회사채
    values: [1200, 2500, 3100]
  - name: 은행 차입
    values: [300.4, 500, 900]
totals: true
note: 9월 말 잔액 기준
source: 가상 자료
"""

BAR_LINE = """chart_type: bar_line
title: 배수가 낮아졌다
labels: [2024, 2025]
series:
  - name: 순차입금
    values: [3900, 3500]
  - name: 순차입금/EBITDA
    values: [2.8, 2.4]
    axis: right
right_number_format: multiple
"""

RANGE = """chart_type: range
title: 방법별 가치 범위
labels: [PER, DCF]
low: [10500, 11000]
high: [14000, 14500]
reference:
  - name: 현재 주가
    value: 12300
show_overlap: true
number_format: ',.0f'
"""

HBAR = """chart_type: hbar
title: 잔액 순위
labels: [가상화학, 가상전자, 가상건설]
series:
  - name: 잔액
    values: [18500, 14200, 9800]
highlight: [가상전자]
"""

LINE = """chart_type: line
title: 스프레드 추이
labels: [1월, 2월, 3월]
series:
  - name: 등급 평균
    values: [70, 68, 66]
  - name: 가상전자
    values: [74, 70, 64]
accent_series: 가상전자
reference:
  - name: 희망 스프레드
    value: 60
"""

WATERFALL = """chart_type: waterfall
title: 현금 증감
labels: [기초, 비용, 유입]
series:
  - name: 증감
    values: [100, -30, 20]
start_total: true
"""

CLASSIC_BAR = """chart_type: bar
title: 가상 매출
labels: [상반기, 하반기]
series:
  - name: 금액
    values: [100, -30]
y_label: 백만원
source: 가상 자료
"""


# ═══════════════════════════════════════════════════════════════════════════════
# SPECIFICATION
# ═══════════════════════════════════════════════════════════════════════════════


def test_new_types_parse_in_supplied_units():
    stacked = parse_chart_spec(STACKED)
    assert stacked.chart_type == "stacked_bar" and stacked.totals
    assert stacked.series[1].values == [300.4, 500, 900]
    assert (stacked.subtitle, stacked.note) == (
        "연도별 만기 도래액, 단위: 억원",
        "9월 말 잔액 기준",
    )
    assert stacked.labels == ["2026", "2027", "2028"]

    bar_line = parse_chart_spec(BAR_LINE)
    assert [s.axis for s in bar_line.series] == ["left", "right"]
    assert bar_line.right_number_format == "multiple"

    range_spec = parse_chart_spec(RANGE)
    assert range_spec.series == []
    assert (range_spec.low, range_spec.high) == ([10500, 11000], [14000, 14500])
    assert range_spec.reference_lines == [ChartReferenceLine("현재 주가", 12300)]
    assert range_spec.show_overlap

    hbar = parse_chart_spec(HBAR)
    assert hbar.highlight == ["가상전자"]

    line = parse_chart_spec(LINE)
    assert line.accent_series == "가상전자" and line.reference_lines[0].value == 60

    waterfall = parse_chart_spec(WATERFALL)
    assert waterfall.start_total


def test_classic_specs_are_exactly_the_old_keys_and_types():
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR))
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR.replace("bar", "line")))
    assert is_classic_spec(parse_chart_spec(CLASSIC_BAR.replace("bar", "waterfall", 1)))
    for spec in (STACKED, BAR_LINE, RANGE, HBAR, LINE, WATERFALL):
        assert not is_classic_spec(parse_chart_spec(spec))
    for extra in ("subtitle: '단위: 원'\n", "note: 가상\n"):
        assert not is_classic_spec(parse_chart_spec(CLASSIC_BAR + extra))
    single = CLASSIC_BAR.replace("[100, -30]", "[100, 30]")
    assert not is_classic_spec(parse_chart_spec(single + "highlight: [상반기]\n"))
    # A false flag written out is still the default value.
    assert is_classic_spec(parse_chart_spec(WATERFALL.replace("true", "false")))
    assert ChartSeries("a", [1.0]).axis == "left"


@pytest.mark.parametrize(
    "spec, diagnostic",
    [
        (STACKED.replace("stacked_bar", "pie"), "chart_type (or type) must be one of"),
        (STACKED.replace("[1200, 2500, 3100]", "[1200, -1, 3100]"), "must not be negative"),
        (
            STACKED.replace("series:\n", "series:\n" + "  - name: x\n    values: [1, 2, 3]\n" * 3),
            "at most 4 series",
        ),
        (STACKED.replace("totals: true", "totals: yes please"), "totals must be true or false"),
        (CLASSIC_BAR + "totals: true\n", "totals applies only to chart_type stacked_bar"),
        (BAR_LINE.replace("    axis: right\n", ""), "bar_line needs"),
        (BAR_LINE.replace("axis: right", "axis: top"), "axis must be left or right"),
        (
            CLASSIC_BAR.replace("    values: [100, -30]", "    values: [1, 2]\n    axis: right"),
            "unknown",
        ),
        (RANGE.replace("[10500, 11000]", "[15000, 11000]"), "low[0] must not be greater"),
        (RANGE.replace("high: [14000, 14500]\n", ""), "high must be a list"),
        (RANGE + "series: []\n", "series applies only to"),
        (RANGE.replace("value: 12300", "value: .nan"), "finite"),
        (RANGE.replace("    value: 12300\n", ""), "name and value"),
        (HBAR.replace("[가상전자]", "[가상해운]"), "not one of the labels"),
        (HBAR.replace("[가상전자]", "[가상전자, 가상전자]"), "repeats"),
        (HBAR.replace("series:\n", "series:\n  - name: x\n    values: [1, 2, 3]\n"), "exactly one"),
        (LINE.replace("accent_series: 가상전자", "accent_series: 없음"), "names no series"),
        (LINE + "highlight: [1월]\n", "highlight applies only to"),
        (STACKED + "highlight: [2026]\n", "highlight applies only to"),
        (
            CLASSIC_BAR.replace(
                "    values: [100, -30]", "    values: [1, 2]\n  - name: b\n    values: [3, 4]"
            )
            + "highlight: [상반기]\n",
            "single series",
        ),
        (WATERFALL + "reference: []\n", "reference applies only to"),
        (
            LINE.replace(
                "    value: 60\n",
                "    value: 60\n  - name: b\n    value: 1\n  - name: c\n    value: 2\n",
            ),
            "1 to 2 lines",
        ),
        (BAR_LINE + "number_format: wrong\n", "number_format"),
        (BAR_LINE.replace("multiple", "wrong"), "right_number_format"),
        (CLASSIC_BAR + "subtitle: [1]\n", "subtitle must be a string"),
    ],
)
def test_invalid_new_keys_are_chart_spec_errors(spec, diagnostic):
    with pytest.raises(ChartSpecError, match="Chart") as caught:
        parse_chart_spec(spec)
    assert diagnostic in str(caught.value)


# ═══════════════════════════════════════════════════════════════════════════════
# DRAWING
# ═══════════════════════════════════════════════════════════════════════════════

CHARTS_SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "qa" / "charts.md"


def _round(value):
    if isinstance(value, float):
        return round(value, 6)
    if isinstance(value, (list, tuple)):
        return [_round(item) for item in value]
    return value


def _figure_elements(fig) -> dict:
    """Font-independent description of a drawn figure: marks, colours and texts."""
    from matplotlib.colors import to_hex

    fig.canvas.draw()
    axes = []
    for ax in fig.axes:
        legend = ax.get_legend()
        axes.append(
            {
                "patches": [
                    [
                        _round(float(p.get_x())),
                        _round(float(p.get_y())),
                        _round(float(p.get_width())),
                        _round(float(p.get_height())),
                        to_hex(p.get_facecolor()),
                        p.get_hatch(),
                    ]
                    for p in ax.patches
                ],
                "lines": [
                    [
                        _round([float(x) for x in line.get_xdata()]),
                        _round([float(y) for y in line.get_ydata()]),
                        to_hex(line.get_color()),
                        str(line.get_linestyle()),
                        _round(float(line.get_linewidth())),
                        str(line.get_marker()),
                    ]
                    for line in ax.lines
                ],
                "texts": [
                    [t.get_text(), to_hex(t.get_color()), _round(float(t.get_fontsize()))]
                    for t in ax.texts
                ],
                "title": ax.get_title(),
                "ylabel": ax.get_ylabel(),
                "xticks": [t.get_text() for t in ax.get_xticklabels()],
                "xlim": _round([float(v) for v in ax.get_xlim()]),
                "ylim": _round([float(v) for v in ax.get_ylim()]),
                "legend": [t.get_text() for t in legend.get_texts()] if legend else [],
            }
        )
    return {
        "size": _round([float(v) for v in fig.get_size_inches()]),
        "dpi": _round(float(fig.dpi)),
        "axes": axes,
        "figure_texts": [t.get_text() for t in fig.texts],
        "figure_legends": [[t.get_text() for t in legend.get_texts()] for legend in fig.legends],
    }


@pytest.fixture
def drawn(monkeypatch):
    """Collect the figure elements of every chart picture saved during a test."""
    from matplotlib.figure import Figure

    saved = Figure.savefig
    found: list[dict] = []

    def capture(self, *args, **kwargs):
        found.append(_figure_elements(self))
        return saved(self, *args, **kwargs)

    monkeypatch.setattr(Figure, "savefig", capture)
    return found


def render_doc(markdown, profile="plain", **options):
    """Actual parsing, rendering and DOCX round trip."""
    model = MarkdownParser(profile=profile).parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile, charts=True, **options))
    doc = renderer.render(model)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return renderer, Document(buffer), buffer.getvalue()


def fence(spec: str) -> str:
    return "```chart\n" + spec + "```\n"


def _axes(drawn, index=0):
    return drawn[index]["axes"]


def _tones(profile):
    from chart_render_standard import ChartTones

    with use_style(load_style(get_profile(profile))):
        return ChartTones.from_style()


@pytest.mark.parametrize("profile", ["business-report", "ib-report"])
def test_classic_draws_the_sample_charts_as_before(profile, drawn, snapshot):
    """`chart_style: classic` keeps the pre-E10 figures (snapshot made before E10's drawing)."""
    markdown = CHARTS_SAMPLE.read_text(encoding="utf-8").replace(
        "profile: business-report", f"profile: {profile}"
    )
    renderer, doc, _ = render_doc(markdown, profile=profile, chart_style="classic", strict=True)
    assert len(doc.inline_shapes) == 3 and not renderer.errors
    assert all(shape.height == 2926080 for shape in doc.inline_shapes)  # 3.2 inches
    assert drawn == snapshot


def test_classic_pictures_are_the_unchanged_classic_function():
    """The classic path inserts exactly what `render_chart_png` draws for the same style."""
    from chart_renderer import render_chart_png

    source = CHARTS_SAMPLE.read_text(encoding="utf-8")
    _, _, data = render_doc(source, "business-report", chart_style="classic")
    with ZipFile(BytesIO(data)) as package:
        names = sorted(name for name in package.namelist() if name.startswith("word/media/"))
        pictures = [package.read(name) for name in names]
    model = MarkdownParser(profile="business-report").parse(source)
    specs = [e.content.spec for e in model.elements if e.element_type == ElementType.CHART]
    with use_style(load_style(get_profile("business-report"))):
        expected = [render_chart_png(spec, width_inches=6.0) for spec in specs]
    assert pictures == expected


def test_classic_style_draws_newer_charts_in_the_standard_style(drawn):
    markdown = "---\nchart_style: classic\n---\n" + fence(CLASSIC_BAR) + fence(STACKED)
    renderer, doc, _ = render_doc(markdown, strict=True)
    assert not renderer.errors
    assert [shape.height for shape in doc.inline_shapes] == [2926080, 2651760]
    assert _axes(drawn, 0)[0]["title"] == "가상 매출"  # classic draws its title inside
    assert _axes(drawn, 1)[0]["title"] == ""


def test_standard_is_the_default_and_options_override_frontmatter(drawn):
    _, doc, _ = render_doc(fence(CLASSIC_BAR))
    assert doc.inline_shapes[0].height == 2651760 and _axes(drawn)[0]["title"] == ""
    classic = "---\nchart_style: classic\n---\n" + fence(CLASSIC_BAR)
    _, doc, _ = render_doc(classic, chart_style="standard")
    assert doc.inline_shapes[0].height == 2651760
    _, doc, _ = render_doc(classic)
    assert doc.inline_shapes[0].height == 2926080


@pytest.mark.parametrize("value", ["fancy", "1", "[classic]"])
def test_invalid_chart_style_is_rejected_in_both_modes(value):
    from input_safety import StrictValidationError

    for strict in (False, True):
        with pytest.raises(StrictValidationError, match="chart_style must be standard or classic"):
            render_doc(f"---\nchart_style: {value}\n---\n" + fence(CLASSIC_BAR), strict=strict)
    with pytest.raises(ValueError, match="chart_style"):
        render_doc(fence(CLASSIC_BAR), chart_style="fancy")


def test_standard_stacked_bar_tones_totals_and_unscaled_labels(drawn):
    tones = _tones("ib-report").stack(2)
    render_doc(fence(STACKED), "ib-report", strict=True)
    ax = _axes(drawn)[0]
    assert [p[4] for p in ax["patches"]] == [tones[0].lower()] * 3 + [tones[1].lower()] * 3
    assert [p[3] for p in ax["patches"]] == [1200, 2500, 3100, 300.4, 500, 900]
    assert [p[1] for p in ax["patches"][3:]] == [1200, 2500, 3100]
    # Totals are added before any rounding and shown by the chart's own format.
    assert [t[0] for t in ax["texts"]][-3:] == ["1,500.4", "3,000", "4,000"]
    assert drawn[0]["figure_legends"] == [["공모 회사채", "은행 차입"]]


def test_standard_stacked_bar_writes_no_totals_unless_asked(drawn):
    render_doc(fence(STACKED.replace("totals: true\n", "")), strict=True)
    texts = [t[0] for t in _axes(drawn)[0]["texts"]]
    assert "4,000" not in texts and "3,000" not in texts and "1,500.4" not in texts


def test_standard_waterfall_start_total_hatch_and_profile_total_word(drawn):
    for profile, word in (("ib-report", "Total"), ("business-report", "합계")):
        drawn.clear()
        render_doc(fence(WATERFALL), profile, strict=True)
        ax = _axes(drawn)[0]
        tones = _tones(profile)
        assert [p[4] for p in ax["patches"]] == [
            tones.anchor().lower(),
            tones.negative.lower(),
            tones.accent.lower(),
            tones.anchor().lower(),
        ]
        assert [p[5] for p in ax["patches"]] == [None, "////", None, None]
        assert ax["xticks"][-1] == word
        assert [t[0] for t in ax["texts"]] == ["100", "(30)", "20", "90"]
    drawn.clear()
    render_doc(fence(WATERFALL.replace("start_total: true\n", "total_label: 기말\n")), strict=True)
    ax = _axes(drawn)[0]
    assert ax["patches"][0][4] == ax["patches"][2][4]  # without the flag the start is an increase
    assert ax["xticks"][-1] == "기말"


def test_standard_hbar_order_highlight_and_range_geometry(drawn):
    tones = _tones("ib-report")
    render_doc(fence(HBAR), "ib-report", strict=True)
    ax = _axes(drawn)[0]
    assert [p[1] + p[3] / 2 for p in ax["patches"]] == [2, 1, 0]  # written order, top down
    grey = tones.pair_grey().lower()
    assert [p[4] for p in ax["patches"]] == [grey, tones.accent.lower(), grey]
    assert [t[0] for t in ax["texts"]] == ["18,500", "14,200", "9,800"]
    drawn.clear()
    render_doc(fence(RANGE), "ib-report", strict=True)
    ax = _axes(drawn)[0]
    bars = [p for p in ax["patches"] if p[3] == 0.5]
    assert [(p[0], p[2]) for p in bars] == [(10500, 3500), (11000, 3500)]
    assert "현재 주가 12,300" in [t[0] for t in ax["texts"]]
    assert [12300, 12300] in [line[0] for line in ax["lines"]]


def test_standard_bar_highlight_marks_only_the_named_label(drawn):
    spec = CLASSIC_BAR.replace("[100, -30]", "[100, 30]") + "highlight: [하반기]\n"
    render_doc(fence(spec), "ib-report", strict=True)
    tones = _tones("ib-report")
    ax = _axes(drawn)[0]
    assert [p[4] for p in ax["patches"]] == [tones.pair_grey().lower(), tones.accent.lower()]


def test_standard_bar_line_uses_two_axes_and_the_right_format(drawn):
    render_doc(fence(BAR_LINE), strict=True)
    left, right = _axes(drawn)
    assert [p[3] for p in left["patches"]] == [3900, 3500]
    assert [t[0] for t in left["texts"]] == ["3,900", "3,500"]
    assert right["lines"][0][1] == [2.8, 2.4]
    assert [t[0] for t in right["texts"]] == ["2.8x", "2.4x"]
    assert drawn[0]["figure_legends"] == [["순차입금", "순차입금/EBITDA"]]


def test_standard_line_end_names_accent_and_unscaled_percent(drawn):
    spec = LINE.replace("title: 스프레드 추이", "title: 비율\nnumber_format: percent")
    render_doc(fence(spec), strict=True)
    ax = _axes(drawn)[0]
    texts = [t[0] for t in ax["texts"]]
    # Accent points before the last, the reference line, then each line's end.
    assert texts[:2] == ["74%", "70%"]
    assert "희망 스프레드 60%" in texts
    assert sorted(texts[-2:]) == ["64%  가상전자", "66%  등급 평균"]
    assert {line[4] for line in ax["lines"] if len(line[0]) == 3} == {1.4, 2.0}


def test_more_than_four_compared_series_warn_without_rejecting(drawn):
    many = "chart_type: bar\ntitle: 많은 계열\nlabels: [a]\nseries:\n" + "".join(
        f"  - name: s{index}\n    values: [{index}]\n" for index in range(5)
    )
    renderer, doc, _ = render_doc(fence(many), strict=True)
    assert len(doc.inline_shapes) == 1 and not renderer.errors
    assert any("5 series" in warning for warning in renderer.warnings)


@given(st.from_regex(r"#[0-9A-F]{6}", fullmatch=True))
@settings(max_examples=300, deadline=None)
def test_chart_tones_stay_apart_in_black_and_white(accent):
    """Any accent colour: the tones of one chart keep MIN_TONE_GAP in grayscale."""
    from chart_render_standard import ANCHOR_GAP, MIN_TONE_GAP, ChartTones, _min_gap, luma

    tones = ChartTones(accent, "#C00000", "#404040", "#808080")
    for count in (2, 3, 4):
        assert _min_gap(tones.stack(count)) >= MIN_TONE_GAP[count]
        for accent_index in range(count):
            assert _min_gap(tones.compare(count, accent_index)) >= MIN_TONE_GAP[count]
    assert abs(luma(tones.anchor()) - luma(accent)) >= ANCHOR_GAP
    assert abs(luma(tones.pair_grey()) - luma(accent)) >= MIN_TONE_GAP[2]


# ═══════════════════════════════════════════════════════════════════════════════
# WORD PARAGRAPHS
# ═══════════════════════════════════════════════════════════════════════════════

TABLE_WITH_ID = """```yaml table
id: tbl-plan
caption: 가상 조달 계획
```

| 구분 | 금액 |
|---|---:|
| 회사채 | 1,000 |

"""


def _chart_block(doc):
    """Paragraphs from the first non-empty one before the picture to the picture's last line."""
    paragraphs = doc.paragraphs
    picture = next(i for i, p in enumerate(paragraphs) if p._p.xpath(".//w:drawing"))
    return paragraphs, picture


def _keep_next(paragraph) -> bool:
    return bool(paragraph.paragraph_format.keep_with_next)


@pytest.mark.parametrize(
    "profile, note_word, source_word",
    [("business-report", "주", "자료"), ("plain", "주", "자료"), ("ib-report", "Note", "Source")],
)
def test_title_subtitle_note_and_source_are_word_paragraphs(profile, note_word, source_word):
    spec = STACKED.replace("note: 9월 말 잔액 기준", "note: 9월 말 잔액 기준")
    renderer, doc, _ = render_doc(fence(spec), profile, strict=True)
    assert not renderer.errors
    paragraphs, picture = _chart_block(doc)
    texts = [p.text for p in paragraphs[picture - 2 : picture + 3]]
    assert texts == [
        "2028년 만기가 가장 크다",
        "연도별 만기 도래액, 단위: 억원",
        "",
        f"{note_word}: 9월 말 잔액 기준",
        f"{source_word}: 가상 자료",
    ]
    title = paragraphs[picture - 2]
    assert all(run.bold for run in title.runs)
    assert [_keep_next(p) for p in paragraphs[picture - 2 : picture + 3]] == [
        True,
        True,
        True,
        True,
        False,
    ]
    descr = doc.inline_shapes[0]._inline.docPr.get("descr")
    assert descr == "2028년 만기가 가장 크다"


def test_subtitle_falls_back_to_the_axis_label_as_written_and_adds_no_words():
    _, doc, _ = render_doc(fence(CLASSIC_BAR), "business-report", strict=True)
    paragraphs, picture = _chart_block(doc)
    assert [p.text for p in paragraphs[picture - 2 : picture + 2]] == [
        "가상 매출",
        "백만원",
        "",
        "자료: 가상 자료",
    ]
    unit_only = CLASSIC_BAR.replace("y_label: 백만원", "unit: 이익률")
    _, doc, _ = render_doc(fence(unit_only), "business-report", strict=True)
    paragraphs, picture = _chart_block(doc)
    assert paragraphs[picture - 1].text == "이익률"
    _, doc, _ = render_doc(fence(unit_only + "subtitle: ''\n"), "business-report", strict=True)
    paragraphs, picture = _chart_block(doc)
    assert paragraphs[picture - 1].text == "가상 매출"  # an empty subtitle hides the line


def test_a_chart_without_texts_is_only_its_picture():
    bare = "chart_type: bar\nlabels: [a, b]\nseries:\n  - name: x\n    values: [1, 2]\n"
    plain_doc = render_doc("문단\n", "plain")[1]
    renderer, doc, _ = render_doc("문단\n\n" + fence(bare), "plain", strict=True)
    assert not renderer.errors
    assert len(doc.paragraphs) == len(plain_doc.paragraphs) + 1
    assert not _keep_next(doc.paragraphs[-1])
    assert doc.inline_shapes[0]._inline.docPr.get("descr") == "Chart 1"


def test_a_theme_file_can_set_the_chart_words(tmp_path):
    theme = tmp_path / "theme.yaml"
    theme.write_text('CHART_SOURCE_LABEL: "출처"\nCHART_NOTE_LABEL: "참고"\n', encoding="utf-8")
    _, doc, _ = render_doc(fence(STACKED), "ib-report", theme=str(theme), strict=True)
    paragraphs, picture = _chart_block(doc)
    assert [p.text for p in paragraphs[picture + 1 : picture + 3]] == [
        "참고: 9월 말 잔액 기준",
        "출처: 가상 자료",
    ]


@pytest.mark.parametrize("place", ["title", "subtitle", "note", "source"])
def test_a_reference_in_chart_text_stays_text_and_is_reported(place):
    texts = {"title": "가상 매출", "subtitle": "'단위: 원'", "note": "가상", "source": "가상 자료"}
    texts[place] = "'@tbl-plan 참고'"
    spec = "chart_type: bar\nlabels: [a]\nseries:\n  - name: x\n    values: [1]\n" + "".join(
        f"{key}: {value}\n" for key, value in texts.items()
    )
    markdown = TABLE_WITH_ID + fence(spec)
    renderer, doc, _ = render_doc(markdown, "business-report")
    message = ";".join(renderer.errors)
    assert f"chart {place}" in message and "@tbl-plan" in message
    assert any("@tbl-plan 참고" in p.text for p in doc.paragraphs)
    assert not any(
        "REF" in (node.text or "")
        for node in doc.element.body.iter()
        if node.tag.endswith("instrText")
    )
    with pytest.raises(ValueError, match="chart"):
        render_doc(markdown, "business-report", strict=True)


def test_classic_chart_text_is_drawn_and_not_read_for_references():
    spec = CLASSIC_BAR.replace("source: 가상 자료", "source: '@tbl-plan 참고'")
    renderer, doc, _ = render_doc(
        TABLE_WITH_ID + fence(spec), "business-report", chart_style="classic", strict=True
    )
    assert not renderer.errors
    assert not any("@tbl-plan" in p.text for p in doc.paragraphs)
