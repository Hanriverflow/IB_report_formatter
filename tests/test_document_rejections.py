"""Document validation retains messages and typed causes at each conversion boundary."""

from io import BytesIO

import pytest

from converters import get_default_registry
from document_profiles import RenderOptions, resolve_options
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser, parse_markdown_file
from md_to_word import IBReportConverter
from tools.portable_launcher import convert_document


def _convert(source, output, profile, route):
    registry = get_default_registry()
    if route in {"file", "cli", "portable"}:
        path = output.with_suffix(".md")
        path.write_text(source, encoding="utf-8")
        if route == "portable":
            return convert_document(str(path), str(output.parent))
        if route == "cli":
            return IBReportConverter(
                str(path), str(output), RenderOptions(profile=profile, strict=True)
            ).convert()
        model = parse_markdown_file(str(path), profile=profile)
    elif route == "stream":
        model = registry.convert(BytesIO(source.encode()), extension_hint="md", profile=profile)
    else:
        model = MarkdownParser(profile=profile).parse(source)
    if route == "render":
        document = IBDocumentRenderer(options=RenderOptions(profile=profile, strict=True)).render(
            model
        )
        document.save(output)
        return output
    return registry.convert(
        model, output_format="docx", output_path=output, strict=True, profile=profile
    )


def _strict_cause(error):
    while error is not None:
        if isinstance(error, StrictValidationError):
            return error
        error = error.__cause__
    pytest.fail("No StrictValidationError in the cause chain")


@pytest.mark.parametrize("profile", ["plain", "ib-report"])
@pytest.mark.parametrize("route", ["text", "stream", "file", "cli", "render"])
@pytest.mark.parametrize(
    "settings, message",
    [
        ("layout: {cover: wrong}", "layout.cover must be true or false"),
        ("layout: []", "layout must be a YAML mapping"),
        ("charts: wrong", "charts must be true or false"),
        ("tables: wrong", "tables must be a list with at most one specification per body table"),
        ("tables: [{landscape: wrong}]", "Table landscape must be true or false"),
        (
            "terms: []",
            'terms must be a YAML mapping of key: "text" pairs (write terms: {} for none)',
        ),
        ("checks: wrong", "checks must be a list of `left = right` relations"),
    ],
)
def test_document_validation_messages_and_causes(tmp_path, profile, route, settings, message):
    output = tmp_path / "example.docx"
    source = f"---\n{settings}\n---\n| 예시 |\n| --- |\n| 가상 |"
    with pytest.raises(Exception) as caught:
        _convert(source, output, profile, route)
    rejection = _strict_cause(caught.value)
    assert str(rejection) == message
    assert isinstance(rejection.__cause__, ValueError)
    assert not isinstance(rejection.__cause__, StrictValidationError)
    assert str(rejection.__cause__) == message
    assert not output.exists()


@pytest.mark.parametrize("profile", ["plain", "ib-report"])
@pytest.mark.parametrize("route", ["text", "stream", "file", "cli", "render"])
@pytest.mark.parametrize(
    "setting",
    [
        'title: "예시\\0"',
        'subtitle: "예시\\x0b"',
        'date: "예시\\uFFFF"',
        'tables: [{caption: "예시\\0"}]',
    ],
)
def test_decoded_metadata_xml_characters_reject_before_save(tmp_path, profile, route, setting):
    output = tmp_path / "example.docx"
    with pytest.raises(Exception) as caught:
        _convert(f"---\n{setting}\n---\n| 예시 |\n| --- |\n| 가상 |", output, profile, route)
    rejection = _strict_cause(caught.value)
    assert "XML-incompatible character U+" in str(rejection)
    assert not output.exists()


@pytest.mark.parametrize(
    "kind, content, message",
    [
        ("theme", "body_size: wrong", "body_size must be between 6 and 30 points"),
        ("house", "style: {header_rule: wrong}", "style.header_rule must be true or false"),
    ],
)
@pytest.mark.parametrize("operator", [False, True])
def test_named_configuration_content_respects_provenance(
    tmp_path, kind, content, message, operator
):
    config = tmp_path / "example.yaml"
    config.write_text(content, encoding="utf-8")
    profile = "term-sheet" if kind == "house" else "plain"
    source = f"---\nprofile: {profile}\ntitle: 예시 문서\nprepared_by: 가상 기관\ndisclaimer: 예시 문구\n"
    if not operator:
        source += f"{kind}: '{config.as_posix()}'\n"
    source += "---\n가상 본문"
    model = MarkdownParser(profile=profile).parse(source)
    if not operator:
        assert model.metadata.extra[kind] == config.as_posix()
    options = RenderOptions(
        profile=profile, strict=True, **({kind: str(config)} if operator else {})
    )
    with pytest.raises(ValueError, match=message) as caught:
        IBDocumentRenderer(options=options).render(model)
    assert isinstance(caught.value, StrictValidationError) is not operator
    if not operator:
        assert str(caught.value.__cause__) == message


@pytest.mark.parametrize(
    "options",
    [
        RenderOptions(profile="wrong"),
        RenderOptions(include_cover="wrong"),
        RenderOptions(preset="wrong"),
        RenderOptions(charts="wrong"),
        RenderOptions(theme=[]),
        RenderOptions(house=[]),
        RenderOptions(separator_mode="wrong"),
    ],
)
def test_operator_validation_is_not_a_document_rejection(options):
    model = MarkdownParser(profile="plain").parse("예시 본문")
    with pytest.raises(ValueError) as caught:
        resolve_options(model.metadata, options)
    assert not isinstance(caught.value, StrictValidationError)


def test_reparse_requirement_remains_operator_error():
    model = MarkdownParser(profile="plain").parse("예시 본문")
    with pytest.raises(ValueError, match="reparse the Markdown") as caught:
        IBDocumentRenderer(options=RenderOptions(profile="ib-report", strict=True)).render(model)
    assert not isinstance(caught.value, StrictValidationError)


@pytest.mark.parametrize(
    "settings, message",
    [
        ("layout: {separator_mode: []}", "unhashable type: 'list'"),
        ("tables: [{type: []}]", "unhashable type: 'list'"),
        (
            "layout: {0: null, wrong: null}",
            "'<' not supported between instances of 'str' and 'int'",
        ),
        ("tables: [{0: null}]", "sequence item 0: expected str instance, int found"),
    ],
)
def test_wrong_yaml_types_keep_original_cause(tmp_path, settings, message):
    with pytest.raises(Exception) as caught:
        _convert(
            f"---\n{settings}\n---\n| 예시 |\n| --- |\n| 가상 |",
            tmp_path / "example.docx",
            "plain",
            "text",
        )
    rejection = _strict_cause(caught.value)
    assert str(rejection) == str(rejection.__cause__) == message
    assert isinstance(rejection.__cause__, TypeError)


def test_existing_raw_xml_diagnosis_unchanged():
    model = MarkdownParser(profile="plain").parse("# \x00")
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert (
        str(caught.value)
        == "Input validation failed: XML-incompatible character U+0000 in Markdown input"
    )


CONSTRUCTOR_FAILURES = ("2026-99-01", "2026-01-01T25:00:00Z", "!!int example", "!!float example")


@pytest.mark.parametrize("scalar", CONSTRUCTOR_FAILURES)
@pytest.mark.parametrize("profile", ["plain", "ib-report"])
@pytest.mark.parametrize("route", ["text", "stream", "file", "cli", "render", "portable"])
def test_yaml_constructor_rejections(tmp_path, scalar, profile, route):
    output = tmp_path / "example.docx"
    with pytest.raises(Exception) as caught:
        _convert(f"---\nprofile: {profile}\ndate: {scalar}\n---\n예시 본문", output, profile, route)
    rejection = _strict_cause(caught.value)
    assert str(rejection) == "Invalid YAML frontmatter: " + str(rejection.__cause__)
    assert type(rejection.__cause__) is ValueError
    assert not list(tmp_path.glob("*.docx"))


@pytest.mark.parametrize("scalar", [*CONSTRUCTOR_FAILURES, "["])
@pytest.mark.parametrize("simple", [True, False])
def test_undeclared_yaml_failures_preserve_fallback(scalar, simple):
    from md_parser import FrontmatterParser

    source = (
        f"---\ndate: {scalar}\n"
        + ("" if simple else "example:\n  - 예시 중첩\n")
        + "---\n가상 본문"
    )
    metadata, remaining = FrontmatterParser.parse(source.splitlines(), "plain")
    if simple:
        assert metadata.extra["date"] == scalar
        assert remaining == ["가상 본문"]
    else:
        assert remaining == source.splitlines()
        assert "date" not in metadata.extra


@pytest.mark.parametrize("kind", ["house", "theme"])
@pytest.mark.parametrize("scalar", CONSTRUCTOR_FAILURES)
@pytest.mark.parametrize("operator", [False, True])
def test_named_yaml_constructor_failure_provenance(tmp_path, kind, scalar, operator):
    config = tmp_path / "example.yaml"
    config.write_text(f"prepared_by: {scalar}", encoding="utf-8")
    profile = "term-sheet" if kind == "house" else "plain"
    source = f"---\nprofile: {profile}\ntitle: 예시 문서\nsubtitle: 가상 조건\n"
    source += "prepared_by: 가상 기관\ndisclaimer: 예시 문구\n"
    if not operator:
        source += f"{kind}: '{config.as_posix()}'\n"
    model = MarkdownParser().parse(source + "---\n예시 본문")
    if not operator:
        assert model.metadata.extra[kind] == config.as_posix()
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer(
            options=RenderOptions(strict=True, **({kind: str(config)} if operator else {}))
        ).render(model)
    assert isinstance(caught.value, StrictValidationError) is not operator
    if not operator:
        assert type(caught.value.__cause__) is ValueError


@pytest.mark.parametrize(
    "profile",
    [
        "plain",
        "ib-report",
        "ib-memo",
        "office-letter",
        "business-report",
        "meeting-minutes",
        "term-sheet",
    ],
)
@pytest.mark.parametrize("route", ["render", "cli", "portable"])
def test_wrong_sender_text_is_document_rejection(tmp_path, profile, route):
    source = f"---\nprofile: {profile}\ntitle: 예시 문서\nsender: {{organization: 12}}\n"
    source += (
        "recipients: [가상 수신처]\nprepared_by: 가상 기관\ndisclaimer: 예시 문구\n---\n가상 본문"
    )
    output = tmp_path / "example.docx"
    with pytest.raises(Exception) as caught:
        _convert(source, output, profile, route)
    assert "sender" in str(_strict_cause(caught.value))
    assert not list(tmp_path.glob("*.docx"))


@pytest.mark.parametrize(
    "location", ["paragraph", "heading", "run", "table", "list", "footnote", "metadata"]
)
@pytest.mark.parametrize("profile", ["plain", "ib-report"])
def test_handbuilt_xml_strings_reject_before_save(tmp_path, location, profile):
    from document_model import (
        DocumentModel,
        Element,
        ElementType,
        Heading,
        ListItem,
        Paragraph,
        Table,
        TableCell,
        TableRow,
        TextRun,
    )
    from document_profiles import default_metadata

    text = "가상\x00예시"
    model = DocumentModel(metadata=default_metadata(profile))
    content = {
        "paragraph": Element(ElementType.PARAGRAPH, Paragraph(text)),
        "heading": Element(ElementType.HEADING_1, Heading(1, text)),
        "run": Element(ElementType.PARAGRAPH, Paragraph("예시", [TextRun(text)])),
        "table": Element(
            ElementType.TABLE, Table(rows=[TableRow(cells=[TableCell(text)])], col_count=1)
        ),
        "list": Element(ElementType.BULLET_LIST, ListItem(text)),
        "footnote": Element(ElementType.PARAGRAPH, Paragraph("예시", [TextRun("", footnote_id=1)])),
        "metadata": Element(ElementType.PARAGRAPH, Paragraph("예시")),
    }
    model.elements = [content[location]]
    if location == "footnote":
        model.footnotes[1] = text
    elif location == "metadata":
        model.metadata.title = text
    output = tmp_path / "example.docx"
    with pytest.raises(StrictValidationError, match="XML-incompatible character U\\+0000"):
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model).save(output)
    assert not output.exists()


@pytest.mark.parametrize(
    "profile",
    [
        "plain",
        "ib-report",
        "ib-memo",
        "office-letter",
        "business-report",
        "meeting-minutes",
        "term-sheet",
    ],
)
def test_existing_metadata_scalar_coercions_render(tmp_path, profile):
    from docx import Document

    source = f"---\nprofile: {profile}\ntitle: 예시 문서\nsubtitle: 가상 조건\n"
    source += "date: 2026-11-11\nanalyst: 12\ncompany: 34\nmeeting_time: 56\nlocation: 78\n"
    source += "document_no: 90\nversion: 2\nsender: {organization: 가상 기관}\nrecipients: [예시 수신처]\n"
    source += "prepared_by: 가상 기관\ndisclaimer: 예시 문구\n---\n가상 본문"
    output = tmp_path / "example.docx"
    _convert(source, output, profile, "render")
    assert Document(output).element.body is not None


@pytest.mark.parametrize("value", [12, False, None, ["예시"], {"예시": 12}])
def test_ib_department_keeps_existing_display_coercion(value):
    model = MarkdownParser(profile="ib-report").parse("# 예시 문서\n\n가상 본문")
    model.metadata.extra["sender"] = {"department": value}
    document = IBDocumentRenderer(options=RenderOptions(strict=True, include_cover=False)).render(
        model
    )
    if value:
        assert any(f"작성부서: {value}" == paragraph.text for paragraph in document.paragraphs)


@pytest.mark.parametrize(
    "extra, message, cause_type",
    [
        ({"checks": "wrong"}, "checks must be a list of `left = right` relations", ValueError),
        (
            {"checks": [], "terms": ["wrong"]},
            "'list' object has no attribute 'items'",
            AttributeError,
        ),
    ],
)
def test_renderer_declaration_boundary_keeps_typed_cause(extra, message, cause_type):
    model = MarkdownParser(profile="plain").parse("가상 본문")
    model.metadata.extra.update(extra)
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert str(caught.value) == message
    assert isinstance(caught.value.__cause__, cause_type)
    assert str(caught.value.__cause__) == message


@pytest.mark.parametrize("kind", ["house", "theme"])
@pytest.mark.parametrize("failure", ["missing", "syntax", "shape", "decoded"])
def test_document_named_configuration_preserves_original_failure(tmp_path, kind, failure):
    from document_profiles import get_profile, load_style
    from term_sheet import load_house

    path = tmp_path / "example.yaml"
    if failure != "missing":
        content = {"syntax": "example: [", "shape": "[]", "decoded": '"예시\\0"'}[failure]
        if failure == "decoded":
            content = ("prepared_by: " if kind == "house" else "body_font: ") + content
        path.write_text(content, encoding="utf-8")
    profile = "term-sheet" if kind == "house" else "plain"
    with pytest.raises(Exception) as original:
        if kind == "house":
            load_house(path)
        else:
            load_style(get_profile(profile), str(path))
    model = MarkdownParser(profile=profile).parse(
        f"---\nprofile: {profile}\ntitle: 예시 문서\n{kind}: '{path.as_posix()}'\n---\n가상 본문"
    )
    assert model.metadata.extra[kind] == path.as_posix()
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert str(caught.value) == str(original.value)
    assert type(caught.value.__cause__) is type(original.value)
    assert str(caught.value.__cause__) == str(original.value)


@pytest.mark.parametrize("profile", ["plain", "ib-report"])
@pytest.mark.parametrize(
    "block",
    [
        "가상 문단 \x01",
        "- 가상 항목 \x01",
        "> 가상 인용 \x01",
        "```\n가상 코드 \x01\n```",
        "| 예시 |\n| --- |\n| 가상\x01 |",
    ],
)
def test_non_strict_keeps_element_resilience_for_xml_characters(tmp_path, profile, block):
    """A body element the XML writer rejects becomes a marker, as before, not a rejection."""
    model = MarkdownParser(profile=profile).parse(f"---\ntitle: 예시 문서\n---\n\n{block}\n")
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile, strict=False))
    output = tmp_path / "example.docx"
    renderer.render(model).save(output)
    assert output.exists()
    assert any("XML-incompatible character U+0001" in error for error in renderer.errors)


@pytest.mark.parametrize(
    "profile, body",
    [
        ("plain", "가상 문단[^1].\n\n[^1]: 가상 각주 \x01"),
        ("ib-report", "가상 문단[^1].\n\n[^1]: 가상 각주 \x01"),
        # The report profile also uses headings for its contents preview.
        ("ib-report", "# 가상 \x01 제목\n\n가상 문단"),
    ],
)
def test_non_strict_names_xml_characters_outside_element_resilience(tmp_path, profile, body):
    """Text rendered outside the per-element guard ends in the documented rejection."""
    model = MarkdownParser(profile=profile).parse(f"---\ntitle: 예시 문서\n---\n\n{body}\n")
    output = tmp_path / "example.docx"
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(profile=profile, strict=False)).render(model).save(
            output
        )
    assert str(caught.value) == "XML-incompatible character U+0001 in Markdown input"
    assert isinstance(caught.value.__cause__, ValueError)
    assert not output.exists()


def test_non_strict_operator_error_is_not_relabelled_by_body_characters():
    """An operator mistake keeps its type even when the body also has such a character."""
    model = MarkdownParser(profile="plain").parse("가상 문단 \x01")
    renderer = IBDocumentRenderer(options=RenderOptions(profile="business-report", strict=False))
    with pytest.raises(ValueError, match="reparse the Markdown") as caught:
        renderer.render(model)
    assert not isinstance(caught.value, StrictValidationError)


# Tagged scalars the safe loader fails on with AttributeError, KeyError or IndexError.
TAGGED_FAILURES = ("!!timestamp example", "!!bool example", '!!int ""', '!!float ""')


@pytest.mark.parametrize("scalar", TAGGED_FAILURES)
@pytest.mark.parametrize("profile", ["plain", "ib-report"])
@pytest.mark.parametrize("route", ["text", "stream", "file", "cli", "render", "portable"])
def test_tagged_yaml_constructor_rejections(tmp_path, scalar, profile, route):
    output = tmp_path / "example.docx"
    with pytest.raises(Exception) as caught:
        _convert(f"---\nprofile: {profile}\ndate: {scalar}\n---\n예시 본문", output, profile, route)
    rejection = _strict_cause(caught.value)
    assert str(rejection).startswith("Invalid YAML frontmatter: YAML value cannot be constructed: ")
    assert type(rejection.__cause__) is ValueError
    assert not list(tmp_path.glob("*.docx"))


def test_deeply_nested_yaml_is_a_document_rejection():
    source = "---\nprofile: plain\nexample: " + "[" * 5000 + "]" * 5000 + "\n---\n예시 본문"
    with pytest.raises(StrictValidationError, match="Invalid YAML frontmatter: "):
        MarkdownParser(profile="plain").parse(source)


@pytest.mark.parametrize("scalar", TAGGED_FAILURES)
@pytest.mark.parametrize("kind", ["house", "theme"])
def test_tagged_yaml_failure_in_document_named_file(tmp_path, kind, scalar):
    config = tmp_path / "example.yaml"
    config.write_text(f"prepared_by: {scalar}", encoding="utf-8")
    profile = "term-sheet" if kind == "house" else "plain"
    source = f"---\nprofile: {profile}\ntitle: 예시 문서\nprepared_by: 가상 기관\n"
    source += f"disclaimer: 예시 문구\n{kind}: '{config.as_posix()}'\n---\n예시 본문"
    model = MarkdownParser().parse(source)
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert "YAML value cannot be constructed" in str(caught.value)


@pytest.mark.parametrize("scalar", ["2026-99-01", "!!timestamp example", '!!int ""'])
def test_chart_fence_with_unconstructible_yaml_is_a_chart_error(scalar):
    """The parser records it on the chart, as it does for a YAML syntax error."""
    source = f"---\nprofile: plain\n---\n\n```chart\ntitle: 예시\nvalue: {scalar}\n```\n"
    model = MarkdownParser(profile="plain").parse(source)
    chart = model.elements[0].content
    assert chart.spec is None
    assert chart.error.startswith("YAML is invalid: ")


@pytest.mark.parametrize("strict", [True, False])
def test_metadata_shapes_that_rendered_before_still_render(tmp_path, strict):
    """Nothing is checked ahead of use: a list of header text and unused fields stay accepted."""
    listed = MarkdownParser(profile="ib-report").parse(
        "---\nprofile: ib-report\ntitle: 예시 문서\nsender: {organization: [가상, 기관]}\n---\n가상 본문"
    )
    unused = MarkdownParser(profile="ib-report").parse("---\nprofile: ib-report\n---\n가상 본문")
    unused.metadata.title = ""
    unused.metadata.subtitle = None
    for index, (model, cover) in enumerate(((listed, None), (unused, False))):
        output = tmp_path / f"example-{index}.docx"
        IBDocumentRenderer(
            options=RenderOptions(profile="ib-report", strict=strict, include_cover=cover)
        ).render(model).save(output)
        assert output.exists()


@pytest.mark.parametrize("strict", [True, False])
@pytest.mark.parametrize(
    "profile, settings, message",
    [
        (
            "ib-memo",
            "sender: 가상 기관",
            "Metadata for the title block has a value of the wrong type",
        ),
        ("ib-report", "sender: {organization: 12}", "Metadata for the header (sender.organization"),
        (
            "ib-memo",
            "attendees: [1, 2]",
            "attendees must be a string or a list of non-empty strings",
        ),
    ],
)
def test_wrong_metadata_type_is_reported_where_it_is_used(profile, settings, message, strict):
    model = MarkdownParser(profile=profile).parse(
        f"---\nprofile: {profile}\ntitle: 예시 문서\n{settings}\n---\n가상 본문"
    )
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(
            options=RenderOptions(profile=profile, strict=strict, include_cover=False)
        ).render(model)
    assert str(caught.value).startswith(message)
    assert isinstance(caught.value.__cause__, (TypeError, AttributeError, ValueError))


def test_non_strict_names_lone_surrogate_outside_element_resilience():
    from document_model import DocumentModel, Element, ElementType, Paragraph, TextRun
    from document_profiles import default_metadata

    model = DocumentModel(metadata=default_metadata("plain"))
    model.elements = [
        Element(ElementType.PARAGRAPH, Paragraph("예시", [TextRun("", footnote_id=1)]))
    ]
    model.footnotes[1] = "가상 \ud800 각주"
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(profile="plain", strict=False)).render(model)
    assert str(caught.value) == "XML-incompatible character U+D800 in Markdown input"
    assert isinstance(caught.value.__cause__, ValueError)


def test_operator_error_text_cannot_trigger_a_document_rejection():
    """An operator mistake is never relabelled, whatever its message says."""
    model = MarkdownParser().parse("가상 문단 \x01")
    renderer = IBDocumentRenderer(options=RenderOptions(profile="XML compatible", strict=False))
    with pytest.raises(ValueError) as caught:
        renderer.render(model)
    assert not isinstance(caught.value, StrictValidationError)
    assert "XML compatible" in str(caught.value)
