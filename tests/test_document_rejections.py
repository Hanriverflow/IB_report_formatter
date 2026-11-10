"""Document validation retains messages and typed causes at each conversion boundary."""

from io import BytesIO

import pytest

from converters import get_default_registry
from document_profiles import RenderOptions, resolve_options
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser, parse_markdown_file
from md_to_word import IBReportConverter


def _convert(source, output, profile, route):
    registry = get_default_registry()
    if route in {"file", "cli"}:
        path = output.with_suffix(".md")
        path.write_text(source, encoding="utf-8")
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
    source = "---\ntitle: 예시 문서\nprepared_by: 가상 기관\ndisclaimer: 예시 문구\n"
    if not operator:
        source += f"{kind}: '{config.as_posix()}'\n"
    source += "---\n가상 본문"
    model = MarkdownParser(profile=profile).parse(source)
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
        f"---\ntitle: 예시 문서\n{kind}: '{path.as_posix()}'\n---\n가상 본문"
    )
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert str(caught.value) == str(original.value)
    assert type(caught.value.__cause__) is type(original.value)
    assert str(caught.value.__cause__) == str(original.value)
