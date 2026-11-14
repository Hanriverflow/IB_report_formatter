"""Which model a render draws when its chart setting is not the parser's.

Fifth independent review, 2026-11-15. A document with chart IDs is read both
with and without its charts, because a reference to a chart resolves only
when charts are drawn. The fourth review's answer exchanged single elements
between the two readings at render time; that exchange lost references in an
element replaced by an equal one, reported an edit that had removed the
reference, and kept a stale diagnostic. It is gone. The contract:

- G1. The record (the parser's setting X, the other reading, a copy of the
  model as parsed) is on the model object alone. Elements and metadata carry
  nothing; two equal readings leave no record. `copy.deepcopy` and pickle
  keep it; `dataclasses.replace` makes a model without it.
- G2. A render with setting Y == X, or of a model without a record, draws the
  model it is given.
- G3. Y != X and the model equal in value to its copy as parsed: the render
  draws the other reading, whole.
- G4. Y != X and the model changed: the render draws the model it is given
  and reports it once for the document; strict mode rejects.

Oracles: for G2 and G4 the same model without its record (`dataclasses.replace`),
which the renderer can only draw as given; for G3 the document parsed with the
render's own setting.
"""

import dataclasses
import pickle
from collections import Counter
from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile

import pytest
from docx import Document

from document_model import DocumentModel, TextRun
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, Paragraph
from tests.snapshot_helpers import XMLNormalizer

RECORD = "_chart_setting"
FIELDS = set(DocumentModel.__dataclass_fields__)

CHART = (
    "```chart\ntype: bar\nid: fig-test\ntitle: 가상 매출\nlabels: [A, B]\n"
    "series: [{name: S, values: [1, 2]}]\n```\n"
)
FRONT = (
    "---\ntitle: 원래 제목\n{charts}tables:\n  - id: tbl-results\n    caption: 실적\n"
    "    columns: [text, number]\n---\n\n"
)
# Reference shapes the two settings read differently, each in its own document
# so that strict mode has documents that convert under either setting. Two of
# them read the same under both settings and so leave no record: `none` holds
# no reference to the chart, and a heading is kept as written (its references
# are read when it is drawn).
BODIES = {
    "spaced": "그림은 @fig-test 참조.",
    "attached": "그림:@fig-test 참조.",
    "heading": "## 소제목 @fig-test",
    "cell": "| 구분 | 내용 |\n|---|---|\n| 나 | @fig-test 참조 |",
    "none": "참조가 없는 문단.",
}
EQUAL_READINGS = ("heading", "none")
DIRECTIONS = pytest.mark.parametrize(
    "parsed_with", [False, True], ids=["parsed-off-drawn-on", "parsed-on-drawn-off"]
)


def document(body: str, charts: str = "") -> str:
    return (
        FRONT.format(charts=charts)
        + "# 본문\n\n| 구분 | 금액 |\n|---|---|\n| 가 | 1 |\n\n첫 문단[^1].\n\n"
        + body
        + "\n\n둘째 문단. 실적은 @tbl-results 참조.\n\n지울 문단.\n\n"
        + CHART
        + "\n[^1]: 원래 각주.\n"
    )


def parse(markdown: str, charts: bool | None) -> DocumentModel:
    return MarkdownParser(profile="plain", charts=charts).parse(markdown)


def changed_message(parsed_with: bool, drawn_with: bool) -> str:
    """The report of G4, written out here and not taken from the engine."""
    read, now = ("on", "off") if parsed_with else ("off", "on")
    return (
        f"The document was changed after it was parsed with charts {read}; this render has "
        f"charts {now}, so references to chart IDs are drawn as parsed. Parse it again with "
        f"charts={drawn_with}"
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
    "diagnostic": (add_warning, None),
}


def _replace_every_element(model: DocumentModel) -> DocumentModel:
    model.elements[:] = [dataclasses.replace(element) for element in model.elements]
    return model


def _new_lists(model: DocumentModel) -> DocumentModel:
    model.elements = list(model.elements)
    model.warnings = list(model.warnings)
    model.metadata = dataclasses.replace(model.metadata)
    return model


def _changed_and_back(model: DocumentModel) -> DocumentModel:
    title = model.metadata.title
    model.metadata.title = "잠시 바꾼 제목"
    model.metadata.title = title
    model.warnings.append("잠시 더한 진단")
    model.warnings.pop()
    return model


# Models equal in value to the parsed one, whatever their objects are.
SAME_VALUE = {
    "untouched": lambda model: model,
    "elements-replaced": _replace_every_element,
    "new-lists": _new_lists,
    "changed-and-back": _changed_and_back,
    "deepcopy": deepcopy,
    "pickle": lambda model: pickle.loads(pickle.dumps(model)),
}


def render_with(renderer: IBDocumentRenderer, model: DocumentModel):
    """Rejection text (or ""), XML parts, diagnostics, warnings and picture count."""
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


def render(model: DocumentModel, charts: bool | None, strict: bool):
    options = RenderOptions(profile="plain", charts=charts, strict=strict)
    return render_with(IBDocumentRenderer(options=options), model)


def shown(parts: dict[str, str]) -> str:
    return "\n".join(parts.values())


def without_record(model: DocumentModel) -> DocumentModel:
    """The same elements and metadata in a model that carries no record."""
    bare = dataclasses.replace(model)
    assert RECORD not in vars(bare)
    return bare


def snapshot(model: DocumentModel):
    """What a render must leave alone: the model's value and its record."""
    record = vars(model).get(RECORD)
    kept = None if record is None else (deepcopy(record["other"]), deepcopy(record["parsed"]))
    return deepcopy(model), set(vars(model)), kept


def assert_untouched(model: DocumentModel, before) -> None:
    value, names, kept = before
    assert model == value and set(vars(model)) == names
    record = vars(model).get(RECORD)
    assert (record is None) == (kept is None)
    if record is not None and kept is not None:
        assert (record["other"], record["parsed"]) == kept
    assert_nothing_on_the_parts(model)


def assert_nothing_on_the_parts(model: DocumentModel) -> None:
    """G1: only the model object may carry anything."""
    assert set(vars(model)) - FIELDS <= {RECORD}
    for part in (model.metadata, *model.elements):
        assert not [name for name in vars(part) if "chart_" in name and name.startswith("_")]


# ═══════════════════════════════════════════════════════════════════════════════
# G1: THE RECORD
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("body", BODIES)
@pytest.mark.parametrize("parsed_with", [False, True, None], ids=["off", "on", "frontmatter"])
def test_the_record_is_on_the_model_alone_and_only_when_the_readings_differ(body, parsed_with):
    model = parse(document(BODIES[body]), parsed_with)
    assert (RECORD in vars(model)) == (body not in EQUAL_READINGS)
    assert_nothing_on_the_parts(model)
    if body not in EQUAL_READINGS:
        record = vars(model)[RECORD]
        assert record["parsed"] == model and record["parsed"] is not model
        assert RECORD not in vars(record["parsed"]) and RECORD not in vars(record["other"])
        assert record["other"] != model or record["own_only"] or record["other_only"]


def test_a_model_without_chart_ids_carries_nothing():
    model = MarkdownParser(profile="plain").parse(
        FRONT.format(charts="")
        + "# 본문\n\n| 구분 | 금액 |\n|---|---|\n| 가 | 1 |\n\n@tbl-results 참조.\n"
    )
    assert set(vars(model)) == FIELDS
    assert_nothing_on_the_parts(model)


@pytest.mark.parametrize("parsed_with", [False, True], ids=["parsed-off", "parsed-on"])
def test_deepcopy_and_pickle_keep_the_record_and_replace_does_not(parsed_with):
    model = parse(document(BODIES["spaced"]), parsed_with)
    for copy in (deepcopy(model), pickle.loads(pickle.dumps(model))):
        assert copy == model and RECORD in vars(copy)
        record = vars(copy)[RECORD]
        assert record["drawn"] is parsed_with and record["parsed"] == copy
        assert record["other"] == vars(model)[RECORD]["other"]
        assert record["other"] is not vars(model)[RECORD]["other"]
    assert set(vars(dataclasses.replace(model))) == FIELDS


# ═══════════════════════════════════════════════════════════════════════════════
# G2: THE SAME SETTING, OR NO RECORD
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("change", [*CHANGES, "nothing"])
@pytest.mark.parametrize("body", BODIES)
@pytest.mark.parametrize("setting", [False, True], ids=["off", "on"])
@pytest.mark.parametrize("strict", [False, True], ids=["default", "strict"])
def test_g2_the_parsers_own_setting_draws_the_given_model(change, body, setting, strict):
    apply, text = CHANGES.get(change, (lambda model: None, None))
    model = parse(document(BODIES[body]), setting)
    apply(model)
    before = snapshot(model)
    result = render(model, setting, strict)
    assert result == render(without_record(model), setting, strict)
    assert_untouched(model, before)
    assert not [message for message in result[2] if "changed after it was parsed" in message]
    assert "changed after it was parsed" not in result[0]
    if result[0] == "":
        assert text is None or text in shown(result[1])
        assert result[4] == (1 if setting else 0)
        if change == "removed":
            assert "지울 문단" not in shown(result[1])


# ═══════════════════════════════════════════════════════════════════════════════
# G3: THE OTHER SETTING, MODEL AS PARSED
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("variant", SAME_VALUE)
@pytest.mark.parametrize("body", BODIES)
@DIRECTIONS
@pytest.mark.parametrize("strict", [False, True], ids=["default", "strict"])
def test_g3_a_model_equal_to_the_parsed_one_is_drawn_as_the_other_reading(
    variant, body, parsed_with, strict
):
    markdown = document(BODIES[body])
    drawn_with = not parsed_with
    expected = render(parse(markdown, drawn_with), drawn_with, strict)
    model = SAME_VALUE[variant](parse(markdown, parsed_with))
    before = snapshot(model)
    result = render(model, drawn_with, strict)
    assert result == expected
    assert_untouched(model, before)
    assert "changed after it was parsed" not in result[0] + "".join(result[2])
    if result[0] == "":
        assert result[4] == (1 if drawn_with else 0)
    if body == "spaced" and drawn_with and not strict:
        # The reference resolves: a field, no text left, nothing reported.
        assert " REF _Ref_fig_test " in shown(result[1]) and result[2] == []
        assert "@fig-test" not in shown(result[1])


# ═══════════════════════════════════════════════════════════════════════════════
# G4: THE OTHER SETTING, MODEL CHANGED
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("change", CHANGES)
@pytest.mark.parametrize("body", BODIES)
@DIRECTIONS
def test_g4_a_changed_model_is_drawn_as_given_and_reported_once(change, body, parsed_with):
    apply, text = CHANGES[change]
    drawn_with = not parsed_with
    model = parse(document(BODIES[body]), parsed_with)
    parser_diagnostics = list(model.warnings)
    apply(model)
    before = snapshot(model)
    given = render(without_record(model), drawn_with, False)
    result = render(model, drawn_with, False)
    strict = render(model, drawn_with, True)
    assert_untouched(model, before)
    message = changed_message(parsed_with, drawn_with)
    assert "Element" not in message  # for the document, not for one element
    if body in EQUAL_READINGS:
        # Its two readings are equal: no record, so nothing to report (G2).
        assert RECORD not in vars(model)
        assert result == given and message not in result[2] and message not in strict[0]
        told = parse(document(BODIES[body]), drawn_with)
        apply(told)
        assert result == render(told, drawn_with, False)
        return
    assert result[0] == "" and result[1] == given[1]  # the same document
    assert result[3:] == given[3:]
    assert Counter(result[2]) == Counter([*given[2], message])  # one report more
    assert all(diagnostic in result[2] for diagnostic in parser_diagnostics)
    assert message in strict[0] and strict[1] == {}  # strict rejects before saving
    assert text is None or text in shown(result[1])
    assert result[4] == (1 if drawn_with else 0)
    # With the parser's own setting the same change needs no report.
    assert message not in render(model, parsed_with, False)[2]


# ═══════════════════════════════════════════════════════════════════════════════
# THE FIFTH REVIEW'S REPRODUCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

REVIEW_CHART = (
    "```chart\ntype: bar\nid: fig-test\ntitle: Test\nlabels: [A, B]\n"
    "series:\n  - name: S\n    values: [1, 2]\n```\n"
)
REVIEW_MD = "See @fig-test here.\n\n" + REVIEW_CHART
UNRESOLVED = "Cross-reference @fig-test has no table or figure with that id"
ATTACHED = (
    "Cross-reference @fig-test is written right after other text and stays text: "
    "put a space before it, or write \\@fig-test to keep it as text"
)


@DIRECTIONS
def test_review_1a_a_replaced_model_has_no_record_and_is_drawn_as_given(parsed_with):
    """The approved limit: `dataclasses.replace(model)` is a model built by hand."""
    drawn_with = not parsed_with
    model = dataclasses.replace(parse(REVIEW_MD, parsed_with))
    assert set(vars(model)) == FIELDS
    rejection, parts, errors, _, pictures = render(model, drawn_with, False)
    assert rejection == "" and pictures == (1 if drawn_with else 0)
    # Not silent in either direction: the reference stays text and is reported.
    assert "@fig-test" in shown(parts) and " REF _Ref_fig_test " not in shown(parts)
    assert errors == [UNRESOLVED]
    assert UNRESOLVED in render(model, drawn_with, True)[0]


@DIRECTIONS
def test_review_1b_an_element_replaced_by_an_equal_one_is_still_the_parsed_model(parsed_with):
    drawn_with = not parsed_with
    model = parse(REVIEW_MD, parsed_with)
    model.elements[0] = dataclasses.replace(model.elements[0])
    for strict in (False, True):
        result = render(model, drawn_with, strict)
        assert result == render(parse(REVIEW_MD, drawn_with), drawn_with, strict)
    result = render(model, drawn_with, False)
    if drawn_with:
        assert " REF _Ref_fig_test " in shown(result[1]) and "@fig-test" not in shown(result[1])
        assert result[2] == [] and result[4] == 1
        assert render(model, drawn_with, True)[0] == ""
    else:
        assert result[2] == [UNRESOLVED] and result[4] == 0


@DIRECTIONS
def test_review_2_an_edit_that_removes_the_reference_is_reported_for_the_document(parsed_with):
    drawn_with = not parsed_with
    model = parse(REVIEW_MD, parsed_with)
    model.elements[0].content.text = "Revised plain text."
    model.elements[0].content.runs = [TextRun(text="Revised plain text.")]
    rejection, parts, errors, _, pictures = render(model, drawn_with, False)
    message = changed_message(parsed_with, drawn_with)
    assert rejection == "" and "Revised plain text." in shown(parts)
    assert "@fig-test" not in shown(parts) and pictures == (1 if drawn_with else 0)
    # The report names no element, and the parser's own diagnostics are kept.
    assert errors == [*model.warnings, message]
    assert not [error for error in errors if error.startswith("Element ")]
    assert message in render(model, drawn_with, True)[0]


@DIRECTIONS
def test_review_3_a_diagnostic_added_by_the_caller_is_kept_with_the_parsers(parsed_with):
    drawn_with = not parsed_with
    model = parse(REVIEW_MD, parsed_with)
    parser_diagnostics = list(model.warnings)
    model.warnings.append("CALLER DIAGNOSTIC")
    rejection, parts, errors, _, pictures = render(model, drawn_with, False)
    message = changed_message(parsed_with, drawn_with)
    assert rejection == "" and pictures == (1 if drawn_with else 0)
    assert errors[: len(parser_diagnostics) + 2] == [
        *parser_diagnostics,
        "CALLER DIAGNOSTIC",
        message,
    ]
    assert errors.count(message) == 1 and errors.count("CALLER DIAGNOSTIC") == 1
    # The model is drawn as given: its reference was read for the parser's setting.
    assert "@fig-test" in shown(parts) and " REF _Ref_fig_test " not in shown(parts)
    strict = render(model, drawn_with, True)[0]
    assert message in strict and "CALLER DIAGNOSTIC" in strict


# ═══════════════════════════════════════════════════════════════════════════════
# A MODEL WITHOUT A RECORD, DRAWN WITH THE OTHER SETTING
# ═══════════════════════════════════════════════════════════════════════════════
#
# What the approved limit gives, pinned as it is on 2026-11-15. Three of the
# four cases are reported. The fourth is not, and is left for the owner's
# decision (stage 8 report): do not read its assertions as a requirement.


def test_no_record_parsed_off_drawn_on_a_spaced_reference_is_reported():
    model = dataclasses.replace(parse(REVIEW_MD, False))
    assert model.warnings == [UNRESOLVED]  # the parser's diagnostic travels with the model
    rejection, parts, errors, _, pictures = render(model, True, False)
    assert rejection == "" and pictures == 1 and errors == [UNRESOLVED]
    assert "@fig-test" in shown(parts)
    assert UNRESOLVED in render(model, True, True)[0]  # strict rejects


def test_no_record_parsed_off_drawn_on_an_attached_reference_is_not_reported():
    """Open question for the owner: the off reading has no diagnostic to carry."""
    markdown = "See:@fig-test here.\n\n" + REVIEW_CHART
    model = dataclasses.replace(parse(markdown, False))
    assert model.warnings == []
    result = render(model, True, False)
    told = render(parse(markdown, True), True, False)
    # The document is the one a parser told `charts=True` gives, text included;
    assert result[1] == told[1] and result[4] == 1 and "See:@fig-test here." in shown(result[1])
    # only the report that the reference stays text is missing, and strict passes.
    assert told[2] == [ATTACHED] and result[2] == []
    assert render(model, True, True)[0] == ""
    assert ATTACHED in render(parse(markdown, True), True, True)[0]


def test_no_record_parsed_on_drawn_off_a_reference_to_the_undrawn_chart_is_reported():
    model = dataclasses.replace(parse(REVIEW_MD, True))
    assert model.warnings == []
    rejection, parts, errors, _, pictures = render(model, False, False)
    assert rejection == "" and pictures == 0 and errors == [UNRESOLVED]
    assert "@fig-test" in shown(parts) and " REF _Ref_fig_test " not in shown(parts)
    assert UNRESOLVED in render(model, False, True)[0]  # strict rejects


def test_no_record_parsed_on_drawn_off_an_attached_reference_keeps_the_parsers_report():
    """Reported although a charts-off parse reports nothing: more, never less."""
    markdown = "See:@fig-test here.\n\n" + REVIEW_CHART
    model = dataclasses.replace(parse(markdown, True))
    assert model.warnings == [ATTACHED]
    result = render(model, False, False)
    told = render(parse(markdown, False), False, False)
    assert result[1] == told[1] and result[4] == 0
    assert result[2] == [ATTACHED] and told[2] == []
    assert ATTACHED in render(model, False, True)[0]


# ═══════════════════════════════════════════════════════════════════════════════
# REPEATED RENDERS
# ═══════════════════════════════════════════════════════════════════════════════

SEQUENCE = (False, True, True, False, True, False)


@pytest.mark.parametrize("parsed_with", [False, True], ids=["parsed-off", "parsed-on"])
def test_renders_of_a_parsed_model_with_changing_settings_are_independent(parsed_with):
    markdown = document(BODIES["spaced"] + "\n\n" + BODIES["attached"])
    model = parse(markdown, parsed_with)
    before = snapshot(model)
    fresh = {flag: render(parse(markdown, flag), flag, False) for flag in (False, True)}
    for flag in SEQUENCE:
        assert render(model, flag, False) == fresh[flag]
        assert_untouched(model, before)


@pytest.mark.parametrize("parsed_with", [False, True], ids=["parsed-off", "parsed-on"])
def test_renders_of_a_changed_model_with_changing_settings_are_independent(parsed_with):
    model = parse(document(BODIES["spaced"] + "\n\n" + BODIES["attached"]), parsed_with)
    change_title(model)
    add_elements(model)
    before = snapshot(model)
    given = {flag: render(without_record(model), flag, False) for flag in (False, True)}
    message = changed_message(parsed_with, not parsed_with)
    for flag in SEQUENCE:
        result = render(model, flag, False)
        assert result[1] == given[flag][1] and "고친 제목" in shown(result[1])
        assert "덧붙인 본문." in shown(result[1])
        extra = [] if flag == parsed_with else [message]
        assert Counter(result[2]) == Counter([*given[flag][2], *extra])
        assert_untouched(model, before)


@DIRECTIONS
def test_one_renderer_draws_parsed_and_changed_models_in_turn(parsed_with):
    drawn_with = not parsed_with
    markdown = document(BODIES["spaced"])
    parsed = parse(markdown, parsed_with)
    changed = parse(markdown, parsed_with)
    change_title(changed)
    expected = render(parse(markdown, drawn_with), drawn_with, False)
    message = changed_message(parsed_with, drawn_with)
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", charts=drawn_with))
    for _ in range(2):
        assert render_with(renderer, parsed) == expected  # G3, no report left over
        result = render_with(renderer, changed)  # G4
        assert result[2].count(message) == 1 and "고친 제목" in shown(result[1])
    assert render_with(renderer, parsed) == expected


@DIRECTIONS
def test_a_changed_copy_is_reported_and_the_original_is_not(parsed_with):
    drawn_with = not parsed_with
    markdown = document(BODIES["spaced"])
    model = parse(markdown, parsed_with)
    message = changed_message(parsed_with, drawn_with)
    for copy in (deepcopy(model), pickle.loads(pickle.dumps(model))):
        change_title(copy)
        assert message in render(copy, drawn_with, False)[2]
        assert message in render(copy, drawn_with, True)[0]
    assert render(model, drawn_with, False) == render(
        parse(markdown, drawn_with), drawn_with, False
    )


# ═══════════════════════════════════════════════════════════════════════════════
# THE SETTING'S SOURCES
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("frontmatter", ["charts: true\n", "charts: false\n", ""])
@pytest.mark.parametrize("drawn_with", [False, True], ids=["render-off", "render-on"])
@pytest.mark.parametrize("body", ["spaced", "attached"])
def test_the_frontmatter_setting_is_no_different_from_telling_the_parser(
    frontmatter, drawn_with, body
):
    """The parser reads the document for its frontmatter; the API option decides."""
    markdown = document(BODIES[body], frontmatter)
    told = parse(markdown, drawn_with)
    model = parse(markdown, None)
    for strict in (False, True):
        assert render(model, drawn_with, strict) == render(told, drawn_with, strict)
    # Without an API option the frontmatter decides at render time too.
    assert render(model, None, False) == render(without_record(model), None, False)


def test_a_change_of_the_frontmatter_setting_after_parsing_is_a_change():
    model = parse(document(BODIES["spaced"], "charts: false\n"), None)
    model.metadata.extra["charts"] = True
    result = render(model, None, False)
    assert changed_message(False, True) in result[2] and result[4] == 1


def test_the_fourth_reviews_reproduction_keeps_the_title_and_the_added_body():
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
        assert changed_message(not render_charts, render_charts) in renderer.errors


# ═══════════════════════════════════════════════════════════════════════════════
# NOT A NUMBER
# ═══════════════════════════════════════════════════════════════════════════════


@DIRECTIONS
@pytest.mark.parametrize("copied", ["as-parsed", "pickled"])
def test_a_model_that_holds_a_nan_is_drawn_as_the_other_reading_or_reported(parsed_with, copied):
    """A NaN is not equal to itself, so such a model may count as changed (G4).

    Frontmatter scalars are kept as text; a NaN stays a number inside a list
    (here under a `letter` key the plain profile does not read).
    As parsed, the model and its copy share that one number, and the list
    comparison takes the same object for equal, so the model counts as parsed
    (G3). A pickled copy holds two separate NaNs and counts as changed (G4).
    Either way nothing is silent.
    """
    drawn_with = not parsed_with
    markdown = "---\nletter: {unused: [.nan]}\n---\n\n" + REVIEW_MD
    model = parse(markdown, parsed_with)
    if copied == "pickled":
        model = pickle.loads(pickle.dumps(model))
    result = render(model, drawn_with, False)
    reported = changed_message(parsed_with, drawn_with) in result[2]
    as_other_reading = result == render(parse(markdown, drawn_with), drawn_with, False)
    assert reported != as_other_reading
    if copied == "pickled":
        assert reported
