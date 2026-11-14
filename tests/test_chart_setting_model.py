"""The renderer draws the model it is given, whatever the chart setting.

Fourth independent review, 2026-11-14: a document with chart IDs was kept as
two whole models, and a render whose chart setting differed from the parser's
swapped the model it was given for the other one. Whatever the caller had
changed after parsing (a title, an added paragraph) was silently lost.

Now there is one model. Only the parts the other chart setting reads
differently carry that other reading, and the renderer takes it, on its own
copy, only for a part that is still as parsed. Three invariants:

1. what the caller changes after parsing is in the output, whether the
   parser's and the render's chart settings agree or not;
2. with charts off the output is the document's without chart IDs
   (principle V3; `tests/test_chart_off_code_block.py`);
3. a model parsed with charts off and rendered with charts on draws its
   charts, numbers them and resolves references to them. An element that was
   changed after parsing and holds a reference the two settings read
   differently is rendered as given and reported.

The oracle is the same document parsed with the render's own setting and
changed the same way: both must give the same output and diagnostics.
"""

from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile

import pytest
from docx import Document

from document_model import DocumentModel
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, Paragraph
from tests.snapshot_helpers import XMLNormalizer

CHART = (
    "```chart\ntype: bar\nid: fig-test\ntitle: 가상 매출\nlabels: [A, B]\n"
    "series: [{name: S, values: [1, 2]}]\n```\n"
)
FRONT = (
    "---\ntitle: 원래 제목\n{charts}tables:\n  - id: tbl-results\n    caption: 실적\n"
    "    columns: [text, number]\n---\n\n"
)
# Reference shapes the two settings read differently, each in its own document
# so that strict mode has documents that convert under either setting.
BODIES = {
    "spaced": "그림은 @fig-test 참조.",
    "attached": "그림:@fig-test 참조.",
    "heading": "## 소제목 @fig-test",
    "cell": "| 구분 | 내용 |\n|---|---|\n| 나 | @fig-test 참조 |",
    "none": "참조가 없는 문단.",
}


def document(body: str, charts: str) -> str:
    return (
        FRONT.format(charts=charts)
        + "# 본문\n\n| 구분 | 금액 |\n|---|---|\n| 가 | 1 |\n\n첫 문단[^1].\n\n"
        + body
        + "\n\n둘째 문단. 실적은 @tbl-results 참조.\n\n지울 문단.\n\n"
        + CHART
        + "\n[^1]: 원래 각주.\n"
    )


def _paragraph(model: DocumentModel, start: str) -> int:
    return next(
        index
        for index, element in enumerate(model.elements)
        if isinstance(element.content, Paragraph) and element.content.text.startswith(start)
    )


def change_title(model: DocumentModel) -> None:
    model.metadata.title = "고친 제목"


def add_elements(model: DocumentModel) -> None:
    model.elements.extend(MarkdownParser(profile="plain").parse("덧붙인 본문.").elements)


def remove_element(model: DocumentModel) -> None:
    del model.elements[_paragraph(model, "지울 문단")]


def reorder_elements(model: DocumentModel) -> None:
    first, second = _paragraph(model, "첫 문단"), _paragraph(model, "지울 문단")
    model.elements[first], model.elements[second] = model.elements[second], model.elements[first]


def edit_paragraph(model: DocumentModel) -> None:
    content = model.elements[_paragraph(model, "지울 문단")].content
    content.text = "고친 문단."
    content.runs[0].text = "고친 문단."


def edit_footnote(model: DocumentModel) -> None:
    model.footnotes[1] = "고친 각주."


def edit_table_spec(model: DocumentModel) -> None:
    table = next(e.content for e in model.elements if hasattr(e.content, "caption"))
    table.caption = "고친 표 제목"


def add_warning(model: DocumentModel) -> None:
    model.warnings.append("호출자가 더한 진단")


CHANGES = {
    "title": (change_title, "고친 제목"),
    "added": (add_elements, "덧붙인 본문."),
    "removed": (remove_element, None),
    "reordered": (reorder_elements, None),
    "paragraph": (edit_paragraph, "고친 문단."),
    "footnote": (edit_footnote, "고친 각주."),
    "table-spec": (edit_table_spec, "고친 표 제목"),
    "nothing": (lambda model: None, None),
}


def render(model: DocumentModel, charts: bool, strict: bool):
    """Rejection text (or ""), XML parts, diagnostics, warnings and picture count."""
    renderer = IBDocumentRenderer(
        options=RenderOptions(profile="plain", charts=charts, strict=strict)
    )
    try:
        rendered = renderer.render(model)
    except ValueError as error:
        return (str(error), {}, [], [], 0)
    buffer = BytesIO()
    rendered.save(buffer)
    with ZipFile(BytesIO(buffer.getvalue())) as package:
        parts = {
            name: XMLNormalizer.normalize(package.read(name), name)
            for name in package.namelist()
            if name.endswith(".xml")
        }
    pictures = len(Document(BytesIO(buffer.getvalue())).inline_shapes)
    return ("", parts, list(renderer.errors), list(renderer.warnings), pictures)


def shown(parts: dict[str, str]) -> str:
    return "\n".join(parts.values())


@pytest.mark.parametrize("change", CHANGES)
@pytest.mark.parametrize("body", BODIES)
@pytest.mark.parametrize("render_charts", [False, True], ids=["render-off", "render-on"])
@pytest.mark.parametrize("strict", [False, True], ids=["default", "strict"])
def test_a_change_after_parsing_is_rendered_whatever_the_settings(
    change, body, render_charts, strict
):
    """Invariants 1 to 3, both directions: the parser was told the opposite setting."""
    apply, text = CHANGES[change]
    markdown = document(BODIES[body], "charts: true\n")
    # The oracle: parsed with the render's own setting, changed the same way.
    told = MarkdownParser(profile="plain", charts=render_charts).parse(markdown)
    apply(told)
    expected = render(told, render_charts, strict)
    # The case: parsed with the opposite setting, changed, then rendered.
    model = MarkdownParser(profile="plain", charts=not render_charts).parse(markdown)
    apply(model)
    before = deepcopy(model)
    result = render(model, render_charts, strict)
    assert result == expected
    assert model == before and vars(model).keys() == vars(before).keys()  # the caller's model
    if text is not None and result[0] == "":
        assert text in shown(result[1])
    if change == "removed" and result[0] == "":
        assert "지울 문단" not in shown(result[1])
    if result[0] == "":
        assert result[4] == (1 if render_charts else 0)


@pytest.mark.parametrize("frontmatter", ["charts: true\n", "charts: false\n", ""])
@pytest.mark.parametrize("render_charts", [False, True], ids=["render-off", "render-on"])
def test_the_frontmatter_setting_is_no_different_from_telling_the_parser(
    frontmatter, render_charts
):
    """The parser reads the document for its frontmatter; the API option decides."""
    markdown = document(BODIES["spaced"], frontmatter)
    told = MarkdownParser(profile="plain", charts=render_charts).parse(markdown)
    model = MarkdownParser(profile="plain").parse(markdown)
    for target in (told, model):
        change_title(target)
        add_elements(target)
    assert render(model, render_charts, False) == render(told, render_charts, False)


def test_the_reviews_reproduction_keeps_the_title_and_the_added_body():
    markdown = (
        "---\nprofile: business-report\ntitle: Original\ncharts: true\n---\n\n"
        "See @fig-test here.\n\n" + CHART
    )
    for parser_charts, render_charts in ((None, False), (False, True)):
        model = MarkdownParser(profile="business-report", charts=parser_charts).parse(markdown)
        model.metadata.title = "Revised title"
        model.elements.extend(
            MarkdownParser(profile="business-report").parse("API ADDED BODY.").elements
        )
        renderer = IBDocumentRenderer(options=RenderOptions(charts=render_charts))
        buffer = BytesIO()
        renderer.render(model).save(buffer)
        saved = Document(BytesIO(buffer.getvalue()))
        text = "\n".join(paragraph.text for paragraph in saved.paragraphs)
        assert "Revised title" in text and "Original" not in text
        assert "API ADDED BODY." in text
        assert saved.core_properties.title == "Revised title"
        assert len(saved.inline_shapes) == (1 if render_charts else 0)


@pytest.mark.parametrize("parser_charts", [False, True], ids=["parsed-off", "parsed-on"])
def test_renders_with_changing_settings_are_independent_and_leave_the_model(parser_charts):
    markdown = document(BODIES["spaced"] + "\n\n" + BODIES["attached"], "")
    model = MarkdownParser(profile="plain", charts=parser_charts).parse(markdown)
    change_title(model)
    add_elements(model)
    before = deepcopy(model)
    fresh = {flag: render(_changed(markdown, flag), flag, False) for flag in (False, True)}
    for flag in (False, True, True, False, True, False):
        assert render(model, flag, False) == fresh[flag]
        assert model == before
        for mine, kept in zip(model.elements, before.elements, strict=True):
            assert vars(mine).keys() == vars(kept).keys()


def _changed(markdown: str, charts: bool) -> DocumentModel:
    model = MarkdownParser(profile="plain", charts=charts).parse(markdown)
    change_title(model)
    add_elements(model)
    return model


@pytest.mark.parametrize("render_charts", [False, True], ids=["render-off", "render-on"])
def test_a_changed_element_with_a_chart_reference_is_rendered_as_given_and_reported(
    render_charts,
):
    """Its references were read for the other setting: nothing is replaced, and it is said."""
    markdown = document(BODIES["spaced"], "")
    model = MarkdownParser(profile="plain", charts=not render_charts).parse(markdown)
    index = _paragraph(model, "그림은")
    content = model.elements[index].content
    content.runs[0].text = "고쳐 쓴 " + content.runs[0].text
    rejection, parts, errors, _, pictures = render(model, render_charts, False)
    assert rejection == "" and "고쳐 쓴 그림은" in shown(parts)
    assert pictures == (1 if render_charts else 0)
    read, now = ("off", "on") if render_charts else ("on", "off")
    message = (
        f"Element {index + 1} was changed after parsing: its references were read with "
        f"charts {read}, and this render has charts {now}. Parse the document again "
        f"with charts={render_charts} to read them for this render"
    )
    assert message in errors
    assert message in render(model, render_charts, True)[0]
    # With the parser's own setting the same change needs no report.
    assert not [e for e in render(model, not render_charts, False)[2] if "changed after" in e]


def test_a_caller_diagnostic_is_kept_when_the_settings_differ():
    markdown = document(BODIES["spaced"], "")
    model = MarkdownParser(profile="plain", charts=True).parse(markdown)
    add_warning(model)
    errors = render(model, False, False)[2]
    assert "호출자가 더한 진단" in errors


def test_a_model_without_chart_ids_carries_nothing():
    model = MarkdownParser(profile="plain").parse(
        FRONT.format(charts="")
        + "# 본문\n\n| 구분 | 금액 |\n|---|---|\n| 가 | 1 |\n\n@tbl-results 참조.\n"
    )
    assert set(vars(model)) == set(DocumentModel.__dataclass_fields__)
    for element in model.elements:
        assert "_chart_alternative" not in vars(element)
