"""With charts off a chart block is a code block in every respect (principle V3).

Third independent review, 2026-11-14: a chart's `id` was still collected as a
reference target when charts were off. `See:@fig-test` was then reported as a
reference written right after other text (strict mode rejected a document that
converted before chart IDs existed), and `See @fig-test` split one Word run
into three. The parser now reads a document with chart IDs both with and
without its charts; the parts the two readings read differently carry the
other reading, and the renderer takes it for its own setting (on its copy,
for parts still as parsed; `tests/test_chart_setting_model.py`). So every way
of turning charts off (no frontmatter key, `charts: false`, the API or CLI
override, with or without telling the parser) gives the document it gave
before: the chart's `id` changes nothing but its own code text.
"""

import logging
from io import BytesIO
from zipfile import ZipFile

import pytest

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file
from tests.snapshot_helpers import XMLNormalizer

CHART = (
    "```chart\ntype: bar\nid: {id}\ntitle: 가상 매출\nlabels: [A, B]\n"
    "series: [{{name: S, values: [1, 2]}}]\n```\n"
)
TABLES = "tables:\n  - id: tbl-results\n    caption: 실적\n    columns: [text, number]\n"
TABLE = "| 구분 | 금액 |\n|---|---|\n| 가 | 1 |\n"
REFERENCES = {
    "none": "본문이다.",
    "spaced": "See @fig-test here.",
    "attached": "See:@fig-test here.",
    "attached-korean": "다음@fig-test 참조.",
    "escaped": "See \\@fig-test here.",
    "bold": "See **@fig-test** here.",
    "split": "See @fig-**test** here.",
    "heading": "## 제목 @fig-test\n\n본문.",
    "cell": "| 구분 | 내용 |\n|---|---|\n| 가 | @fig-test 참조 |",
    "callout": "> **Note:** @fig-test 참조.",
    "quote": "> 인용 @fig-test 참조.",
    "footnote": "본문[^1].\n\n[^1]: @fig-test 참조.",
    "list": "- 항목 @fig-test\n- 둘째",
    "code": "`@fig-test` 코드.",
    "link-text": "[@fig-test](https://example.com) 링크.",
    "with-table": "실적은 @tbl-results 참조, 그림은 @fig-test 참조.",
    "attached-table": "실적:@tbl-results 참조.",
    "twice": "@fig-test 그리고 @fig-test.",
}
# How charts end up off: frontmatter line, renderer override, parser override.
PATHS = {
    "no-key": ("", None, None),
    "frontmatter-false": ("charts: false\n", None, None),
    "renderer-override": ("charts: true\n", False, None),
    "both-overrides": ("charts: true\n", False, False),
}


def convert(
    markdown: str,
    renderer_charts: bool | None,
    parser_charts: bool | None,
    strict: bool,
    profile: str = "plain",
) -> tuple[str, dict[str, str], list[str], list[str]]:
    """Returns the rejection (or ""), the XML parts, the diagnostics and the render warnings."""
    model = MarkdownParser(profile=profile, charts=parser_charts).parse(markdown)
    renderer = IBDocumentRenderer(
        options=RenderOptions(profile=profile, charts=renderer_charts, strict=strict)
    )
    try:
        document = renderer.render(model)
    except ValueError as error:
        return (str(error), {}, [], [])
    buffer = BytesIO()
    document.save(buffer)
    with ZipFile(BytesIO(buffer.getvalue())) as package:
        parts = {
            name: XMLNormalizer.normalize(package.read(name), name)
            for name in package.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        }
    return ("", parts, list(renderer.errors), list(renderer.warnings))


def document(front: str, reference: str, chart_id: str) -> str:
    body = "# 본문\n\n" + TABLE + "\n" + reference + "\n\n" + CHART.format(id=chart_id)
    return f"---\ntitle: 가상 시험\n{front}{TABLES}---\n\n" + body


@pytest.mark.parametrize("path", PATHS)
@pytest.mark.parametrize("reference", REFERENCES)
@pytest.mark.parametrize("strict", [False, True])
def test_a_chart_id_changes_nothing_but_its_own_code_text(path, reference, strict):
    front, renderer_charts, parser_charts = PATHS[path]
    with_id = convert(
        document(front, REFERENCES[reference], "fig-test"), renderer_charts, parser_charts, strict
    )
    # The same document whose chart has an ID nothing refers to.
    other = convert(
        document(front, REFERENCES[reference], "fig-none"), renderer_charts, parser_charts, strict
    )
    rejection, parts, errors, warnings = with_id
    other_rejection, other_parts, other_errors, other_warnings = other
    assert rejection == other_rejection
    assert errors == other_errors and warnings == other_warnings
    assert parts.keys() == other_parts.keys()
    for name, xml in parts.items():
        assert xml == other_parts[name].replace("fig-none", "fig-test"), name


@pytest.mark.parametrize("path", PATHS)
def test_the_reviews_document_converts_in_strict_mode_again(path):
    front, renderer_charts, parser_charts = PATHS[path]
    markdown = f"---\nprofile: plain\n{front}---\n\nSee:@fig-test here.\n\n" + CHART.format(
        id="fig-test"
    )
    result = convert(markdown, renderer_charts, parser_charts, strict=True)
    assert result[0] == "", result[0]
    assert result[2] == [] and result[3] == []


@pytest.mark.parametrize("path", PATHS)
def test_a_reference_written_for_an_undrawn_chart_stays_one_run(path):
    from docx import Document

    front, renderer_charts, parser_charts = PATHS[path]
    markdown = f"---\nprofile: plain\n{front}---\n\nSee @fig-test here.\n\n" + CHART.format(
        id="fig-test"
    )
    model = MarkdownParser(profile="plain", charts=parser_charts).parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=renderer_charts))
    buffer = BytesIO()
    renderer.render(model).save(buffer)
    paragraph = next(
        p for p in Document(BytesIO(buffer.getvalue())).paragraphs if "@fig-test" in p.text
    )
    assert [run.text for run in paragraph.runs] == ["See @fig-test here."]
    assert renderer.errors == ["Cross-reference @fig-test has no table or figure with that id"]


def test_nothing_is_logged_for_the_reading_that_is_not_used(caplog):
    markdown = "---\nprofile: plain\ncharts: true\n---\n\nSee:@fig-test here.\n\n" + CHART.format(
        id="fig-test"
    )
    caplog.set_level(logging.WARNING)
    model = MarkdownParser(profile="plain").parse(markdown)
    # The frontmatter draws charts: that reading reports the attached reference,
    assert any("fig-test" in message for message in model.warnings)
    # but it is logged only by a render that uses that reading.
    assert not [record for record in caplog.records if "fig-test" in record.getMessage()]
    off = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=False, strict=True))
    off.render(model)
    assert not [record for record in caplog.records if "fig-test" in record.getMessage()]
    on = IBDocumentRenderer(options=RenderOptions(profile="plain"))
    on.render(model)
    assert len([record for record in caplog.records if "fig-test" in record.getMessage()]) == 1


def test_the_other_reading_is_carried_only_when_a_chart_id_is_a_target():
    plain = MarkdownParser(profile="plain").parse("# 본문\n\n@fig-test 참조.\n")
    assert "_chart_setting" not in vars(plain)
    with_id = MarkdownParser(profile="plain").parse(
        "# 본문\n\n@fig-test 참조.\n\n" + CHART.format(id="fig-test")
    )
    # No frontmatter key: the model is the charts-off reading.
    assert with_id.warnings == ["Cross-reference @fig-test has no table or figure with that id"]
    carrying = [element for element in with_id.elements if "_chart_alternative" in vars(element)]
    assert [element.content.text for element in carrying] == ["@fig-test 참조."]
    on = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=True, strict=True))
    assert len(on.render(with_id).inline_shapes) == 1 and on.errors == []
    told = MarkdownParser(profile="plain", charts=True).parse(
        "# 본문\n\n@fig-test 참조.\n\n" + CHART.format(id="fig-test")
    )
    assert told.warnings == []


def test_a_file_gives_both_readings_its_folder(tmp_path):
    from tests.test_captions_crossrefs import SAMPLE_IMAGE

    (tmp_path / "s.png").write_bytes(SAMPLE_IMAGE.read_bytes())
    source = tmp_path / "doc.md"
    source.write_text(
        "---\nprofile: plain\n---\n\n# 본문\n\n@fig-test 참조.\n\n![가상 그림](s.png)\n\n"
        + CHART.format(id="fig-test"),
        encoding="utf-8",
    )
    model = parse_markdown_file(str(source))
    assert model.source_dir == tmp_path.resolve()
    for charts, pictures in ((True, 2), (False, 1)):
        rendered = IBDocumentRenderer(options=RenderOptions(charts=charts)).render(model)
        assert len(rendered.inline_shapes) == pictures
