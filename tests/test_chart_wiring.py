"""
Verification tests for opt-in chart rendering in the Markdown-to-Word pipeline.

Covered scope:
- Default chart fences remain code panels
- Valid chart specs render as inline images when enabled
- Invalid chart specs fall back to code panels without aborting rendering
- Korean chart labels complete the full conversion pipeline
"""

from pathlib import Path

from ib_renderer import IBDocumentRenderer
from md_parser import parse_markdown_file
from md_to_word import IBReportConverter, RenderOptions

VALID_CHART = """```chart
chart_type: bar
labels: ["2025", "2026"]
series:
  - name: Revenue
    values: [100, 120]
```
"""

INVALID_CHART = """```chart
chart_type: bar
labels: ["2025", "2026"]
series:
  - name: Revenue
    values: [100]
```
"""

KOREAN_CHART = """# 실적 보고서

```chart
chart_type: bar
title: 매출 및 영업이익
labels: [매출, 영업이익]
series:
  - name: 금액
    values: [1200, 180]
```
"""


def _write_markdown(tmp_path: Path, text: str, name: str = "chart.md") -> Path:
    markdown_path = tmp_path / name
    markdown_path.write_text(text, encoding="utf-8")
    return markdown_path


def _render_elements(markdown_path: Path, enable_charts: bool = False):
    model = parse_markdown_file(markdown_path)
    renderer = IBDocumentRenderer(enable_charts=enable_charts)
    for element in model.elements:
        renderer._render_element(element)
    return renderer.doc


def test_chart_flag_off_renders_code_panel(tmp_path):
    """Keep chart fences as code panels when chart rendering is disabled."""
    doc = _render_elements(_write_markdown(tmp_path, VALID_CHART))

    assert len(doc.inline_shapes) == 0
    assert len(doc.tables) == 1
    assert "chart_type: bar" in doc.tables[0].cell(0, 0).text


def test_chart_flag_on_renders_inline_image(tmp_path):
    """Render a valid chart specification as an inline image when enabled."""
    doc = _render_elements(
        _write_markdown(tmp_path, VALID_CHART),
        enable_charts=True,
    )

    assert len(doc.inline_shapes) == 1
    assert len(doc.tables) == 0


def test_invalid_chart_falls_back_to_code_panel(tmp_path):
    """Fall back to the code panel when chart validation fails."""
    output_path = tmp_path / "invalid-chart.docx"
    doc = _render_elements(
        _write_markdown(tmp_path, INVALID_CHART),
        enable_charts=True,
    )
    doc.save(str(output_path))

    assert output_path.is_file()
    assert len(doc.inline_shapes) == 0
    assert len(doc.tables) == 1
    assert "values: [100]" in doc.tables[0].cell(0, 0).text


def test_korean_chart_completes_full_pipeline(tmp_path):
    """Convert a Korean chart through parsing, rendering, and DOCX saving."""
    markdown_path = _write_markdown(tmp_path, KOREAN_CHART, "korean-chart.md")
    output_path = tmp_path / "korean-chart.docx"

    rendered_path = IBReportConverter(
        md_file_path=str(markdown_path),
        output_path=str(output_path),
        render_options=RenderOptions(enable_charts=True),
    ).convert()

    assert Path(rendered_path) == output_path.resolve()
    assert output_path.is_file()
    assert output_path.stat().st_size > 0
