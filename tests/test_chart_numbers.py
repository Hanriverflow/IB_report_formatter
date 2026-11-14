"""Chart numbers (E10 with E5): `id: fig-…` numbers a chart as a figure in its title line.

Every test goes through the production parser and renderer. Names and numbers
are fictional.
"""

import shutil
from pathlib import Path

import pytest
from docx.document import Document as DocxDocument

from chart_renderer import parse_chart
from document_profiles import RenderOptions
from docx_audit import audit_file, audit_to_dict
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import parse_markdown_file
from tests.test_captions_crossrefs import SAMPLE_IMAGE, bookmarks, fields

CHART = """```chart
chart_type: bar
id: {ref_id}
title: {title}
labels: [상반기, 하반기]
series:
  - name: 금액
    values: [100, 120]
source: 가상 자료
```
"""


def chart(ref_id: str = "fig-sales", title: str = "가상 매출이 늘었다") -> str:
    return CHART.format(ref_id=ref_id, title=title)


def write(tmp_path: Path, body: str, charts: str = "true") -> Path:
    shutil.copy(SAMPLE_IMAGE, tmp_path / "s.png")
    path = tmp_path / "doc.md"
    front = f"---\nprofile: business-report\ntitle: 시험 문서\ncharts: {charts}\n---\n\n"
    path.write_text(front + "# 시험 문서\n\n" + body, encoding="utf-8")
    return path


def render(path: Path, strict: bool = True, **options) -> tuple[DocxDocument, list[str]]:
    renderer = IBDocumentRenderer(options=RenderOptions(strict=strict, **options))
    document = renderer.render(parse_markdown_file(str(path)))
    return document, renderer.errors


def titles(document: DocxDocument) -> list[str]:
    """The paragraph just above each picture's subtitle, or above the picture."""
    paragraphs = document.paragraphs
    found = []
    for index, paragraph in enumerate(paragraphs):
        if paragraph._p.xpath(".//w:drawing") and index >= 1:
            found.append(paragraphs[index - 1].text)
    return found


def test_a_chart_with_an_id_is_a_numbered_figure_and_can_be_referenced(tmp_path: Path) -> None:
    body = "매출은 @fig-sales와 같다.\n\n" + chart()
    document, errors = render(write(tmp_path, body))
    assert errors == []
    assert fields(document) == [
        ("REF _Ref_fig_sales \\h \\* CHARFORMAT", "그림 1"),
        ("SEQ 그림 \\* ARABIC", "1"),
    ]
    assert "그림 1. 가상 매출이 늘었다" in titles(document)
    assert bookmarks(document)["_Ref_fig_sales"][1] == "그림 1"
    alt = document.inline_shapes[0]._inline.docPr.get("descr")
    assert alt == "가상 매출이 늘었다"  # the description stays the title


def test_charts_and_image_figures_share_one_figure_count(tmp_path: Path) -> None:
    body = (
        "![가상 구조도](s.png){#fig-structure}\n\n"
        + chart()
        + "\n![번호 없는 그림](s.png)\n\n"
        + chart("fig-second", "두 번째 차트")
        + "\n@fig-second, @fig-structure\n"
    )
    document, errors = render(write(tmp_path, body))
    assert errors == []
    assert [code for code, _ in fields(document) if code.startswith("SEQ")] == [
        "SEQ 그림 \\* ARABIC"
    ] * 3
    assert [result for code, result in fields(document)] == ["1", "2", "3", "그림 3", "그림 1"]
    assert {"그림 2. 가상 매출이 늘었다", "그림 3. 두 번째 차트"} <= set(titles(document))


def test_a_chart_without_an_id_or_without_charts_on_keeps_its_output(tmp_path: Path) -> None:
    plain = chart().replace("id: fig-sales\n", "")
    document, errors = render(write(tmp_path, plain))
    assert errors == [] and fields(document) == []
    assert "가상 매출이 늘었다" in titles(document)


def test_a_chart_that_is_not_drawn_gets_no_number_and_its_references_are_reported(
    tmp_path: Path,
) -> None:
    path = write(tmp_path, "@fig-sales 참고\n\n" + chart(), charts="false")
    with pytest.raises(StrictValidationError, match="charts are off"):
        render(path)
    document, errors = render(path, strict=False)
    message = ";".join(errors)
    assert "'fig-sales' is on a chart, and charts are off" in message
    assert "Cross-reference @fig-sales" in message
    assert fields(document) == [] and not document.inline_shapes
    assert any("@fig-sales 참고" in p.text for p in document.paragraphs)
    # The CLI/API switch decides, not only the frontmatter. Without a reference
    # the chart is a code panel and nothing is reported (principle V2, second
    # independent review): this replaces the earlier pin that the ID alone was
    # reported, which rejected such documents in strict mode.
    path = write(tmp_path, chart())
    document, errors = render(path, strict=False, charts=False)
    assert errors == [] and not document.inline_shapes
    document, errors = render(path, charts=False)
    assert errors == []


@pytest.mark.parametrize(
    "ref_id, title, problem",
    [
        ("fig_sales", "가상 매출", "is not a figure id"),
        ("tbl-sales", "가상 매출", "is not a figure id"),
        ("fig-sales", "''", "has no caption to number"),
    ],
)
def test_an_unusable_chart_id_leaves_the_chart_drawn_without_a_number(
    tmp_path: Path, ref_id: str, title: str, problem: str
) -> None:
    path = write(tmp_path, chart(ref_id, title))
    with pytest.raises(StrictValidationError, match=problem):
        render(path)
    document, errors = render(path, strict=False)
    assert any(problem in error for error in errors)
    assert len(document.inline_shapes) == 1 and fields(document) == []


def test_a_repeated_chart_id_keeps_the_count_and_reports(tmp_path: Path) -> None:
    path = write(tmp_path, chart() + "\n" + chart(title="또 다른 차트"))
    with pytest.raises(StrictValidationError, match="more than one table or figure"):
        render(path)
    document, _ = render(path, strict=False)
    assert [result for _, result in fields(document)] == ["1", "2"]


def test_an_id_must_be_text_and_a_classic_style_still_numbers(tmp_path: Path) -> None:
    from chart_renderer import ChartSpecError

    with pytest.raises(ChartSpecError, match="id must be a string"):
        parse_chart(
            chart().replace("```chart\n", "").replace("```\n", "").replace("fig-sales", "[1]")
        )
    document, errors = render(write(tmp_path, chart()), chart_style="classic")
    assert errors == []
    assert "그림 1. 가상 매출이 늘었다" in titles(document)


def test_numbered_chart_output_passes_the_structural_audit(tmp_path: Path) -> None:
    body = "매출은 @fig-sales와 같다.\n\n" + chart()
    document, errors = render(write(tmp_path, body))
    assert errors == []
    output = tmp_path / "out.docx"
    document.save(str(output))
    report = audit_to_dict(audit_file(str(output)))
    assert report["issues"] == [] and report["warnings"] == []
