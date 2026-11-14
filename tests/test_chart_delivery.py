"""Chart warnings reach every output path, and chart pictures fit the page body.

Second independent review, 2026-11-14:

- [3] the warning that names values left out of a chart reached only
  `IBDocumentRenderer.warnings`; the CLI printed nothing. It is now logged once
  where it is made, so the CLI, batch conversion, the registry's
  `render_warnings` and `scripts/agent_render.py` all show it, once each.
- [4] a picture made taller for many line ends could exceed the page body.
  A standard chart's picture is now at most the section's body height less
  `ChartRenderer.CHART_TEXT_ROOM` (1.1 inches for its title, subtitle, note
  and source), narrowed in proportion when it would be taller.
"""

import logging
import struct
import subprocess
import sys
from io import BytesIO
from pathlib import Path

import pytest
import yaml
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches

from converters import get_default_registry
from document_profiles import RenderOptions
from ib_render_media import ChartRenderer
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser
from scripts.agent_render import render_document

ROOT = Path(__file__).resolve().parents[1]
LEFT_OUT = "value(s) not written in the picture for lack of room"
DENSE = {
    "type": "bar",
    "labels": [f"C{i}" for i in range(12)],
    "series": [{"name": f"S{i}", "values": [1234567 + i] * 12} for i in range(6)],
}
TALL = {
    "type": "line",
    "title": "Sixty lines",
    "labels": ["A", "B"],
    "series": [{"name": f"L{i:02}", "values": [i, i + 1]} for i in range(60)],
}


def fence(data: dict) -> str:
    return "```chart\n" + yaml.safe_dump(data, allow_unicode=True) + "```\n"


def write(tmp_path: Path, data: dict, front: str = "profile: plain\n") -> Path:
    path = tmp_path / "chart.md"
    text = f"---\n{front}title: Chart test\ncharts: true\n---\n\n# Charts\n\n" + fence(data)
    path.write_text(text, encoding="utf-8")
    return path


# ── [3] warnings reach the user ─────────────────────────────────────────────


def test_the_cli_prints_a_left_out_value_warning_once(tmp_path):
    source = write(tmp_path, DENSE)
    target = tmp_path / "out" / "chart.docx"
    target.parent.mkdir()
    done = subprocess.run(
        [sys.executable, str(ROOT / "md_to_word.py"), str(source), str(target), "--strict"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
    )
    assert done.returncode == 0, done.stderr
    assert target.exists()
    shown = done.stdout + done.stderr
    assert shown.count(LEFT_OUT) == 1
    assert "[WARNING] Chart 'bar':" in shown


def test_the_registry_and_the_log_each_get_it_once(tmp_path, caplog):
    source = write(tmp_path, DENSE)
    registry = get_default_registry()
    model = registry.convert(str(source), profile="plain")
    sink: list[str] = []
    caplog.set_level(logging.WARNING)
    registry.convert(
        model,
        output_format="docx",
        output_path=str(tmp_path / "chart.docx"),
        strict=True,
        render_warnings=sink,
    )
    assert [w for w in sink if LEFT_OUT in w] and len([w for w in sink if LEFT_OUT in w]) == 1
    logged = [r for r in caplog.records if LEFT_OUT in r.getMessage()]
    assert len(logged) == 1 and logged[0].levelno == logging.WARNING


def test_the_agent_runner_lists_it_and_stays_ok(tmp_path):
    source = write(tmp_path, DENSE)
    result = render_document(str(source), str(tmp_path / "result"))
    assert result["ok"]
    assert len([w for w in result["render_warnings"] if LEFT_OUT in w]) == 1


# ── [4] the picture fits the page body ──────────────────────────────────────


def _picture_heights(data: bytes) -> list[tuple[float, float, float]]:
    """Each picture's height and width with its section's body height, in inches."""
    document = Document(BytesIO(data))
    section = document.sections[-1]
    body = (section.page_height - section.top_margin - section.bottom_margin) / 914400
    return [(shape.height / 914400, shape.width / 914400, body) for shape in document.inline_shapes]


def _render(markdown: str, profile: str, **options) -> bytes:
    model = MarkdownParser(profile=profile).parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile, charts=True, **options))
    buffer = BytesIO()
    renderer.render(model).save(buffer)
    return buffer.getvalue()


@pytest.mark.parametrize(
    ("profile", "theme"),
    [
        ("plain", None),
        ("plain", "TOP_MARGIN: 3\nBOTTOM_MARGIN: 3\n"),
        ("business-report", None),
        ("office-letter", None),
        ("term-sheet", None),
        ("ib-report", None),
    ],
)
def test_a_tall_chart_fits_the_body_less_the_text_room(tmp_path, profile, theme):
    options = {}
    if theme is not None:
        path = tmp_path / "theme.yaml"
        path.write_text(theme, encoding="utf-8")
        options["theme"] = str(path)
    front = {
        "office-letter": (
            "---\ntitle: 차트 시험\nrecipients: [가상기관]\n"
            "sender:\n  organization: 가상회사\n  signatory: 가상회사 대표이사\n---\n\n"
        ),
        "term-sheet": (
            "---\ntitle: 가상 조건표\nprepared_by: 가상증권\ndisclaimer: 가상 면책 문구\n---\n\n"
        ),
    }.get(profile, "")
    data = _render(front + "# Charts\n\n" + fence(TALL), profile, **options)
    pictures = _picture_heights(data)
    assert pictures
    for height, _, body in pictures:
        assert height <= body - ChartRenderer.CHART_TEXT_ROOM + 1e-3


def test_a_landscape_section_fits_and_keeps_the_proportions():
    from docx.enum.section import WD_SECTION

    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=True))
    document = renderer.render(MarkdownParser(profile="plain").parse("# Charts\n"))
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    model = MarkdownParser(profile="plain").parse(fence(TALL))
    chart = next(element.content for element in model.elements if hasattr(element.content, "spec"))
    ChartRenderer(document).render(chart, "standard")
    shape = document.inline_shapes[0]
    body = (section.page_height - section.top_margin - section.bottom_margin) / 914400
    assert shape.height / 914400 <= body - ChartRenderer.CHART_TEXT_ROOM + 1e-3
    blob = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
    png = document.part.related_parts[blob].blob
    wide, high = struct.unpack(">II", png[16:24])
    assert shape.height / shape.width == pytest.approx(high / wide, rel=1e-3)
    assert shape.width < Inches(6)


def test_an_ordinary_chart_keeps_its_size():
    data = _render("# Charts\n\n" + fence({**TALL, "series": TALL["series"][:3]}), "plain")
    ((height, width, _),) = _picture_heights(data)
    assert width == pytest.approx(6.0, abs=1e-3)
    assert height == pytest.approx(2.9, abs=0.01)


@pytest.mark.parametrize("body_inches", [0.2, 0.3, 0.5, 1.0, 1.4, 2.0])
def test_a_chart_is_never_taller_than_the_page_body(body_inches):
    """Third independent review, C4: the 0.5 in minimum height exceeded a shorter body."""
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=True))
    document = renderer.render(MarkdownParser(profile="plain").parse("# Charts\n"))
    section = document.sections[-1]
    spare = section.page_height - Inches(body_inches)
    section.top_margin = spare // 2
    section.bottom_margin = spare - spare // 2
    model = MarkdownParser(profile="plain").parse(fence(TALL))
    chart = next(element.content for element in model.elements if hasattr(element.content, "spec"))
    ChartRenderer(document).render(chart, "standard")
    shape = document.inline_shapes[0]
    assert shape.height / 914400 <= body_inches + 1e-3
    # Where the body leaves room for the text, the earlier rule still holds.
    if body_inches - ChartRenderer.CHART_TEXT_ROOM >= ChartRenderer.MIN_CHART_HEIGHT:
        assert shape.height / 914400 <= body_inches - ChartRenderer.CHART_TEXT_ROOM + 1e-3
