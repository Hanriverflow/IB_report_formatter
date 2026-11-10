"""Asset trust boundaries through parsing, rendering and saving (fictional data)."""

import base64
import logging
import os
from io import BytesIO
from pathlib import Path
from urllib.parse import quote
from zipfile import ZipFile

import pytest
import yaml
from docx import Document

from converters import get_default_registry
from document_model import DocumentModel, Element, ElementType, Image
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser, parse_markdown_file
from md_to_word import IBReportConverter, build_parser, run_conversion
from term_sheet import load_house, resolve_term_sheet_texts

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a9h8AAAAASUVORK5CYII="
)
PREFIX = "[asset_outside_root] "
SAMPLES = sorted(
    path
    for folder in ("profiles", "qa", "distribution")
    for path in (Path(__file__).resolve().parents[1] / "samples" / folder).rglob("*.md")
)


def _image(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(PNG)
    return path


def _source(folder: Path, body: str, **metadata) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "example.md"
    data = {"profile": "plain", "title": "가상 문서", **metadata}
    path.write_text(
        "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body, encoding="utf-8"
    )
    return path


def _assert_rejection(error: BaseException) -> None:
    from input_safety import StrictValidationError

    assert isinstance(error, StrictValidationError)
    assert isinstance(error, ValueError)
    assert error.codes == ["asset_outside_root"]
    assert PREFIX in str(error)


@pytest.mark.parametrize("form", ["parent", "absolute", "unc", "link"])
def test_escape_forms_rejected_before_save(tmp_path: Path, form: str) -> None:
    outside = _image(tmp_path / "outside" / "example.png")
    folder = tmp_path / "document"
    folder.mkdir()
    if form == "parent":
        target = "../outside/example.png"
    elif form == "absolute":
        target = outside.as_posix()
    elif form == "unc":
        target = "//localhost/asset-policy-example-share/example.png"
    else:
        link = folder / "link"
        try:
            link.symlink_to(outside.parent, target_is_directory=True)
        except OSError as error:
            pytest.skip(f"Directory symlinks unavailable: {error}")
        target = "link/example.png"
    source = _source(folder, f"![가상 이미지]({target})")
    output = tmp_path / "rejected.docx"
    model = parse_markdown_file(str(source))
    with pytest.raises(ValueError) as caught:
        get_default_registry().convert(model, output_format="docx", output_path=output, strict=True)
    _assert_rejection(caught.value)
    assert not output.exists()


@pytest.mark.parametrize(
    "body",
    [
        "![가상 이미지]({path})",
        "앞 ![가상 이미지]({path}) 뒤",
        "![가상 이미지][example]\n\n[example]: {path}",
        '<img src="{path}" alt="가상 이미지">',
        "| 항목 | 내용 |\n|---|---|\n| 가상 | ![가상 이미지]({path}) |",
        '<table><tr><td><img src="{path}" alt="가상 이미지"></td></tr></table>',
    ],
)
def test_all_body_image_forms_leave_placeholder_and_coded_warning(
    tmp_path: Path, caplog: pytest.LogCaptureFixture, body: str
) -> None:
    outside = _image(tmp_path / "outside.png")
    source = _source(tmp_path / "doc", body.format(path=outside.as_posix()))
    renderer = IBDocumentRenderer(options=RenderOptions(strict=False))
    with caplog.at_level(logging.WARNING):
        doc = renderer.render(parse_markdown_file(str(source)))
    output = tmp_path / "continued.docx"
    doc.save(output)
    assert output.exists()
    assert not doc.inline_shapes
    assert "[Image:" in "".join(doc.element.body.itertext())
    assert any(issue.startswith(PREFIX) for issue in renderer.errors)
    assert any(record.message.startswith(PREFIX) for record in caplog.records)


@pytest.mark.parametrize("exists", [False, True])
def test_outside_existence_never_changes_code(tmp_path: Path, exists: bool, monkeypatch) -> None:
    outside = tmp_path / "outside.png"
    if exists:
        _image(outside)
    source = _source(tmp_path / "doc", f"![가상]({outside.as_posix()})")
    model = parse_markdown_file(str(source))
    original = Path.is_file

    def guarded_is_file(path: Path) -> bool:
        assert path != outside, "outside paths must be rejected before existence checks"
        return original(path)

    monkeypatch.setattr(Path, "is_file", guarded_is_file)
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    _assert_rejection(caught.value)


def test_absolute_inside_and_default_folder_cannot_be_removed(tmp_path: Path) -> None:
    inside = _image(tmp_path / "doc" / "example.png")
    source = _source(inside.parent, f"![가상]({inside.as_posix()})")
    model = parse_markdown_file(str(source), asset_roots=())
    doc = IBDocumentRenderer(options=RenderOptions(strict=True, asset_roots=())).render(model)
    assert len(doc.inline_shapes) == 1


@pytest.mark.skipif(os.name != "nt", reason="Windows drive spelling")
def test_absolute_drive_path_with_repeated_separator_is_inside(tmp_path: Path) -> None:
    inside = _image(tmp_path / "example.png")
    destination = inside.as_posix().replace(":/", "://", 1)
    source = _source(tmp_path, f"![가상]({destination})")
    doc = IBDocumentRenderer(options=RenderOptions(strict=True)).render(
        parse_markdown_file(str(source))
    )
    assert len(doc.inline_shapes) == 1


@pytest.mark.parametrize(
    "entry", ["parse", "render", "options", "registry_parse", "registry_render"]
)
def test_api_roots_are_additive(tmp_path: Path, entry: str) -> None:
    outside = _image(tmp_path / "assets" / "example.png")
    inside = _image(tmp_path / "doc" / "local.png")
    source = _source(
        inside.parent, "![가상 내부](local.png)\n\n![가상 외부](../assets/example.png)"
    )
    roots = (outside.parent,)
    registry = get_default_registry()
    if entry == "registry_parse":
        model = registry.convert(source, asset_roots=roots)
    else:
        model = parse_markdown_file(str(source), asset_roots=roots if entry == "parse" else ())
    if entry == "registry_render":
        doc = registry.convert(model, output_format="docx", strict=True, asset_roots=roots)
    else:
        options = RenderOptions(strict=True, asset_roots=roots if entry == "options" else ())
        doc = IBDocumentRenderer(options=options).render(
            model, asset_roots=roots if entry == "render" else ()
        )
    assert len(doc.inline_shapes) == 2


def test_cli_repeatable_asset_root(tmp_path: Path) -> None:
    first = _image(tmp_path / "one" / "example.png")
    second = _image(tmp_path / "two" / "example.png")
    source = _source(
        tmp_path / "doc", "![가상 하나](../one/example.png)\n\n![가상 둘](../two/example.png)"
    )
    output = tmp_path / "cli.docx"
    args = build_parser().parse_args(
        [
            str(source),
            str(output),
            "--strict",
            "--asset-root",
            str(first.parent),
            "--asset-root",
            str(second.parent),
        ]
    )
    assert run_conversion(source, args) == 0
    assert len(Document(output).inline_shapes) == 2


@pytest.mark.parametrize("kind", ["text", "stream", "model"])
def test_no_source_requires_roots(tmp_path: Path, kind: str) -> None:
    outside = _image(tmp_path / "example.png")
    text = f"![가상]({outside.as_posix()})"
    if kind == "text":
        model = MarkdownParser(profile="plain").parse(text)
    elif kind == "stream":
        model = parse_markdown_file(BytesIO(text.encode()), profile="plain")
    else:
        model = DocumentModel(
            elements=[Element(ElementType.IMAGE, Image("가상", path=str(outside)))]
        )
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    with pytest.raises(ValueError) as caught:
        renderer.render(model)
    _assert_rejection(caught.value)
    assert len(renderer.render(model, asset_roots=(tmp_path,)).inline_shapes) == 1
    with pytest.raises(ValueError):
        renderer.render(model)


def test_text_and_stream_accept_parse_roots(tmp_path: Path) -> None:
    asset = _image(tmp_path / "example.png")
    text = f"![가상]({asset.as_posix()})"
    models = [
        MarkdownParser(asset_roots=(tmp_path,)).parse(text),
        parse_markdown_file(BytesIO(text.encode()), asset_roots=(tmp_path,)),
    ]
    for model in models:
        assert (
            len(IBDocumentRenderer(options=RenderOptions(strict=True)).render(model).inline_shapes)
            == 1
        )


@pytest.mark.parametrize("kind", ["text", "stream", "model"])
def test_relative_paths_do_not_grant_cwd_access(tmp_path: Path, monkeypatch, kind: str) -> None:
    _image(tmp_path / "example.png")
    monkeypatch.chdir(tmp_path)
    text = "![가상](example.png)"
    if kind == "text":
        model = MarkdownParser(base_dir=tmp_path).parse(text)
    elif kind == "stream":
        model = parse_markdown_file(BytesIO(text.encode()))
    else:
        model = DocumentModel(
            source_dir=tmp_path,
            elements=[Element(ElementType.IMAGE, Image("가상", path="example.png"))],
        )
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    with pytest.raises(ValueError) as caught:
        renderer.render(model)
    _assert_rejection(caught.value)
    assert len(renderer.render(model, asset_roots=(tmp_path,)).inline_shapes) == 1


def test_registry_options_and_explicit_roots_are_combined(tmp_path: Path) -> None:
    first = _image(tmp_path / "one" / "example.png")
    second = _image(tmp_path / "two" / "example.png")
    source = _source(
        tmp_path / "doc", "![가상 하나](../one/example.png)\n\n![가상 둘](../two/example.png)"
    )
    registry = get_default_registry()
    options = RenderOptions(strict=True, asset_roots=(first.parent,))
    model = registry.convert(source, render_options=options, asset_roots=(second.parent,))
    assert len(registry.convert(model, output_format="docx", strict=True).inline_shapes) == 2
    model = registry.convert(source)
    doc = registry.convert(
        model, output_format="docx", render_options=options, asset_roots=(second.parent,)
    )
    assert len(doc.inline_shapes) == 2


@pytest.mark.parametrize("kind", ["text", "stream", "model"])
def test_embedded_images_need_no_roots(kind: str) -> None:
    data = base64.b64encode(PNG).decode()
    text = f"![가상](data:image/png;base64,{data})"
    if kind == "text":
        model = MarkdownParser().parse(text)
    elif kind == "stream":
        model = parse_markdown_file(BytesIO(text.encode()))
    else:
        model = DocumentModel(
            elements=[
                Element(
                    ElementType.IMAGE,
                    Image("가상", path="", base64_data=data, mime_type="image/png"),
                )
            ]
        )
    assert (
        len(IBDocumentRenderer(options=RenderOptions(strict=True)).render(model).inline_shapes) == 1
    )


def _house(folder: Path, logo: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    house = folder / "house.yaml"
    house.write_text(
        yaml.safe_dump(
            {"prepared_by": "가상 금융", "disclaimer": "가상 검토용", "style": {"logo": logo}},
            allow_unicode=True,
        ),
        encoding="utf-8",
    )
    return house


@pytest.mark.parametrize("operator", ["frontmatter", "api", "cli"])
def test_house_logo_trust_depends_on_who_named_house(tmp_path: Path, operator: str) -> None:
    logo = _image(tmp_path / "house" / "example.png")
    house = _house(logo.parent, logo.name)
    source = _source(tmp_path / "doc", "가상 본문", profile="term-sheet", house=str(house))
    output = tmp_path / "house.docx"
    model = parse_markdown_file(str(source))
    if operator == "frontmatter":
        with pytest.raises(ValueError) as caught:
            get_default_registry().convert(
                model, output_format="docx", output_path=output, strict=True
            )
        _assert_rejection(caught.value)
        assert not output.exists()
        renderer = IBDocumentRenderer()
        assert not renderer.render(model).inline_shapes
        assert any(issue.startswith(PREFIX) for issue in renderer.errors)
    elif operator == "api":
        doc = get_default_registry().convert(
            model, output_format="docx", strict=True, house=str(house)
        )
        assert len(doc.inline_shapes) == 1
    else:
        args = build_parser().parse_args(
            [str(source), str(output), "--strict", "--house", str(house)]
        )
        assert run_conversion(source, args) == 0
        assert len(Document(output).inline_shapes) == 1


def test_frontmatter_logo_cannot_borrow_operator_house_trust(tmp_path: Path) -> None:
    logo = _image(tmp_path / "outside.png")
    house = _house(tmp_path / "house", "")
    source = _source(tmp_path / "doc", "가상 본문", profile="term-sheet", style={"logo": str(logo)})
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True, house=str(house)))
    with pytest.raises(ValueError) as caught:
        renderer.render(parse_markdown_file(str(source)))
    _assert_rejection(caught.value)
    assert (
        len(
            renderer.render(parse_markdown_file(str(source)), asset_roots=(tmp_path,)).inline_shapes
        )
        == 1
    )


@pytest.mark.parametrize("kind", ["text", "stream"])
def test_no_source_frontmatter_logo_requires_roots(tmp_path: Path, kind: str) -> None:
    logo = _image(tmp_path / "example.png")
    source = _source(
        tmp_path,
        "가상 본문",
        profile="term-sheet",
        prepared_by="가상 금융",
        disclaimer="가상 검토용",
        style={"logo": str(logo)},
    )
    model = (
        MarkdownParser().parse(source.read_text(encoding="utf-8"))
        if kind == "text"
        else parse_markdown_file(BytesIO(source.read_bytes()))
    )
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    with pytest.raises(ValueError) as caught:
        renderer.render(model)
    _assert_rejection(caught.value)
    assert len(renderer.render(model, asset_roots=(tmp_path,)).inline_shapes) == 1


def test_data_logo_and_local_hyperlink_do_not_need_roots(tmp_path: Path) -> None:
    source = _source(
        tmp_path,
        f"[가상 링크](<{(tmp_path / 'outside.pdf').as_posix()}>)",
        profile="term-sheet",
        prepared_by="가상 금융",
        disclaimer="가상 검토용",
        style={"logo": "data:image/png;base64," + base64.b64encode(PNG).decode()},
    )
    model = MarkdownParser().parse(source.read_text(encoding="utf-8"))
    doc = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert len(doc.inline_shapes) == 1
    assert doc.element.body.xpath(".//w:hyperlink")


def test_render_rejection_keeps_prefix_on_each_issue(tmp_path: Path) -> None:
    source = _source(
        tmp_path / "doc",
        "![가상 하나](../one.png)\n\n본문 ![가상 둘](../two.png)\n\n![가상 누락](missing.png)",
    )
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    with pytest.raises(ValueError) as caught:
        renderer.render(parse_markdown_file(str(source)))
    _assert_rejection(caught.value)
    assert len(renderer.errors) == 3
    assert all(issue.startswith(PREFIX) for issue in renderer.errors[:2])
    assert renderer.errors[2] == "Image could not be rendered"
    assert str(caught.value) == "Document validation failed: " + "; ".join(renderer.errors)


def test_percent_candidates_preserve_order_and_check_decoded_escape(tmp_path: Path) -> None:
    raw = _image(tmp_path / "doc" / "example%20image.png")
    raw.with_name("example image.png").write_bytes(PNG + b"decoded candidate")
    source = _source(raw.parent, "![가상](example%20image.png)")
    doc = IBDocumentRenderer(options=RenderOptions(strict=True)).render(
        parse_markdown_file(str(source))
    )
    assert (
        doc.part.related_parts[
            doc.inline_shapes[0]._inline.graphic.graphicData.pic.blipFill.blip.embed
        ].blob
        == PNG
    )
    outside = _image(tmp_path / "outside image.png")
    source = _source(raw.parent, f"![가상]({quote('../' + outside.name, safe='')})")
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(
            parse_markdown_file(str(source))
        )
    _assert_rejection(caught.value)


@pytest.mark.parametrize("allowed", [False, True])
@pytest.mark.parametrize("strict", [False, True])
def test_registry_two_steps_match_cli_converter(
    tmp_path: Path, caplog: pytest.LogCaptureFixture, allowed: bool, strict: bool
) -> None:
    image = _image(tmp_path / "assets" / "example.png")
    local = _image(tmp_path / "doc" / "local.png")
    source = _source(local.parent, "![Local](local.png)\n\n![External](../assets/example.png)")
    roots = (image.parent,) if allowed else ()
    one = tmp_path / "one.docx"
    two = tmp_path / "two.docx"
    cli = IBReportConverter(str(source), str(one), RenderOptions(strict=strict, asset_roots=roots))
    with caplog.at_level(logging.WARNING):
        if strict and not allowed:
            with pytest.raises(RuntimeError, match="Rendering failed:") as cli_error:
                cli.convert()
            _assert_rejection(cli_error.value.__cause__)
        else:
            assert cli.convert() == str(one)
    cli_warnings = {record.message for record in caplog.records}
    caplog.clear()

    registry = get_default_registry()
    with caplog.at_level(logging.WARNING):
        model = registry.convert(source, asset_roots=roots)
        if strict and not allowed:
            with pytest.raises(ValueError) as registry_error:
                registry.convert(model, output_format="docx", output_path=two, strict=strict)
            _assert_rejection(registry_error.value)
            assert str(registry_error.value) == str(cli_error.value.__cause__)
        else:
            assert registry.convert(
                model, output_format="docx", output_path=two, strict=strict
            ) == str(two)
    # The CLI also logs renderer diagnostics; compare distinct warning messages.
    assert {record.message for record in caplog.records} == cli_warnings
    if allowed:
        assert not cli_warnings
    else:
        assert any(message.startswith(PREFIX) for message in cli_warnings)
    if strict and not allowed:
        assert not one.exists()
        assert not two.exists()
        return

    assert (
        len(Document(one).inline_shapes)
        == len(Document(two).inline_shapes)
        == (2 if allowed else 1)
    )
    with ZipFile(one) as first, ZipFile(two) as second:
        assert first.namelist() == second.namelist()
        for name in first.namelist():
            if name != "docProps/core.xml":
                assert first.read(name) == second.read(name), name


@pytest.mark.parametrize("profile", [None, "business-report"])
def test_registry_markdown_with_output_format_only_parses(
    tmp_path: Path, profile: str | None
) -> None:
    source = _source(tmp_path, "Example body")
    output = tmp_path / "unused.docx"
    model = get_default_registry().convert(
        source,
        output_format="docx",
        output_path=output,
        profile=profile,
        render_options=RenderOptions(profile="ib-report"),
    )
    assert isinstance(model, DocumentModel)
    assert model == parse_markdown_file(str(source), profile=profile)
    assert not output.exists()


def test_document_settings_cannot_grant_roots(tmp_path: Path) -> None:
    image = _image(tmp_path / "outside.png")
    source = _source(
        tmp_path / "doc",
        f"![가상]({image.as_posix()})",
        asset_roots=[str(tmp_path)],
        asset_root=str(tmp_path),
    )
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(
            parse_markdown_file(str(source))
        )
    _assert_rejection(caught.value)


def test_strict_exception_preserves_uncoded_issues_and_wrapper_cause(tmp_path: Path) -> None:
    from input_safety import StrictValidationError

    model = MarkdownParser(profile="plain").parse("가상 본문")
    model.warnings = ["Existing warning sentence", PREFIX + "First issue", PREFIX + "Second issue"]
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert str(caught.value) == "Input validation failed: " + "; ".join(model.warnings)
    assert caught.value.codes == ["asset_outside_root"]
    model.warnings = ["Existing warning sentence"]
    with pytest.raises(StrictValidationError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    assert caught.value.codes == []
    source = _source(tmp_path / "doc", "![가상](../missing.png)")
    output = tmp_path / "wrapped.docx"
    with pytest.raises(RuntimeError, match="Rendering failed:") as wrapped:
        IBReportConverter(str(source), str(output), RenderOptions(strict=True)).convert()
    _assert_rejection(wrapped.value.__cause__)
    assert not output.exists()


def test_existing_missing_image_warning_unchanged(tmp_path: Path, caplog) -> None:
    source = _source(tmp_path, "![가상](missing.png)")
    renderer = IBDocumentRenderer()
    renderer.render(parse_markdown_file(str(source)))
    assert renderer.errors == ["Image could not be rendered"]
    assert any(
        record.message
        == f"Image file not found: {tmp_path / 'missing.png'} — inserting placeholder"
        for record in caplog.records
    )


def test_existing_missing_logo_warning_unchanged(tmp_path: Path) -> None:
    missing = tmp_path / "missing.png"
    source = _source(
        tmp_path,
        "가상 본문",
        profile="term-sheet",
        prepared_by="가상 금융",
        disclaimer="가상 검토용",
        style={"logo": str(missing)},
    )
    try:
        Document().add_picture(str(missing))
    except FileNotFoundError as error:
        expected = f"House logo could not be rendered: {error}"
    renderer = IBDocumentRenderer()
    renderer.render(parse_markdown_file(str(source)))
    assert renderer.errors == [expected]


@pytest.mark.parametrize("source", SAMPLES, ids=lambda path: path.name)
def test_samples_still_convert_strictly(source: Path, tmp_path: Path) -> None:
    output = tmp_path / "sample.docx"
    IBReportConverter(str(source), str(output), RenderOptions(strict=True)).convert()
    assert output.is_file()


@pytest.mark.skipif(os.name != "nt", reason="Windows junctions")
def test_junction_escape(tmp_path: Path) -> None:
    import subprocess

    outside = _image(tmp_path / "outside" / "example.png")
    folder = tmp_path / "doc"
    folder.mkdir()
    junction = folder / "junction"
    result = subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(junction), str(outside.parent)], capture_output=True
    )
    if result.returncode:
        pytest.skip("Junction creation is not permitted")
    source = _source(folder, "![가상](junction/example.png)")
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(
            parse_markdown_file(str(source))
        )
    _assert_rejection(caught.value)


@pytest.mark.parametrize("origin", ["body", "frontmatter", "house"])
@pytest.mark.parametrize("escaped_candidate", ["written", "decoded"])
def test_logo_candidates_resolve_independently_through_links(tmp_path, origin, escaped_candidate):
    import subprocess

    outside = _image(tmp_path / "outside" / "example.png")
    folder = tmp_path / "doc"
    base = folder / "config" if origin == "house" else folder
    base.mkdir(parents=True)
    names = {"written": "example%20logo", "decoded": "example logo"}
    junction = base / names[escaped_candidate]
    if os.name == "nt":
        result = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(junction), str(outside.parent)], capture_output=True
        )
        if result.returncode:
            pytest.skip("Junction creation is not permitted")
    else:
        junction.symlink_to(outside.parent, target_is_directory=True)
    if escaped_candidate == "written":
        _image(base / names["decoded"] / "example.png")
    destination = "example%20logo/example.png"
    metadata = {}
    body = "가상 본문"
    if origin == "body":
        body = f"![가상]({destination})"
    else:
        metadata = {
            "profile": "term-sheet",
            "prepared_by": "가상 금융",
            "disclaimer": "예시 검토용",
        }
        if origin == "house":
            metadata["house"] = str(_house(base, destination))
        else:
            metadata["style"] = {"logo": destination}
    source = _source(folder, body, **metadata)
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    model = parse_markdown_file(str(source))
    if escaped_candidate == "written":
        # The outside written candidate must not destroy the allowed decoded one.
        assert len(renderer.render(model).inline_shapes) == 1
    else:
        with pytest.raises(StrictValidationError) as caught:
            renderer.render(model)
        _assert_rejection(caught.value)


@pytest.mark.parametrize("origin", ["frontmatter", "house"])
def test_logos_already_try_written_then_decoded_filenames(tmp_path, origin):
    raw = _image(tmp_path / "example%20logo.png")
    decoded = raw.with_name("example logo.png")
    decoded.write_bytes(PNG + b"example decoded candidate")
    metadata = {"profile": "term-sheet", "prepared_by": "가상 금융", "disclaimer": "예시 검토용"}
    if origin == "house":
        metadata["house"] = str(_house(tmp_path, raw.name))
    else:
        metadata["style"] = {"logo": raw.name}
    source = _source(tmp_path, "가상 본문", **metadata)
    for expected in (raw, decoded):
        document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(
            parse_markdown_file(str(source))
        )
        assert len(document.inline_shapes) == 1
        relation = document.inline_shapes[0]._inline.graphic.graphicData.pic.blipFill.blip.embed
        assert document.part.related_parts[relation].blob == expected.read_bytes()
        if expected == raw:
            raw.unlink()


@pytest.mark.skipif(os.name != "nt", reason="Windows drive-relative paths")
@pytest.mark.parametrize("origin", ["frontmatter", "house"])
def test_drive_relative_logos_are_anchored_without_resolving(tmp_path, monkeypatch, origin):
    drive = "Z:" if tmp_path.drive.upper() != "Z:" else "Y:"
    # Supply a drive CWD without depending on a mapped drive on the test host.
    abspath = os.path.abspath
    monkeypatch.setattr(
        os.path, "abspath", lambda value: drive + "/" if value == drive else abspath(value)
    )
    logo = drive + "example-link/../example.png"
    expected = Path(logo).absolute()
    assert expected.is_absolute()
    assert ".." in expected.parts
    if origin == "house":
        config = _house(tmp_path, logo)
        stored = load_house(config)["style"]["logo"]
    else:
        source = _source(
            tmp_path,
            "가상 본문",
            profile="term-sheet",
            prepared_by="가상 기관",
            disclaimer="예시 문구",
            style={"logo": logo},
        )
        stored = parse_markdown_file(str(source)).metadata.extra["style"]["logo"]
    assert stored == str(expected)


@pytest.mark.parametrize("document_house", [False, True])
@pytest.mark.parametrize("override", [False, True])
def test_final_logo_path_validation_keeps_provenance(
    tmp_path, monkeypatch, document_house, override
):
    import term_sheet

    # Simulate a loader returning an unanchored path to exercise the last guard
    # independently of platform-specific absolute-path anchoring.
    monkeypatch.setattr(
        term_sheet,
        "load_house",
        lambda path: {
            "prepared_by": "가상 기관",
            "disclaimer": "예시 문구",
            "style": {"logo": "example.png"},
        },
    )
    model = MarkdownParser(profile="term-sheet").parse(
        "---\nprofile: term-sheet\ntitle: 예시 문서\n---\n가상 본문"
    )
    if override:
        model.metadata.extra["style"] = {"logo": "example-override.png"}
    with pytest.raises(
        ValueError, match="style.logo must be absolute for input without a source file"
    ) as caught:
        resolve_term_sheet_texts(
            model.metadata, str(tmp_path / "example.yaml"), document_house=document_house
        )
    assert isinstance(caught.value, StrictValidationError) is (document_house or override)


@pytest.mark.parametrize("origin", ["frontmatter", "document-house", "operator-house"])
def test_logo_anchoring_error_keeps_provenance(tmp_path, monkeypatch, origin):
    config = _house(tmp_path, "example.png")
    metadata = {"profile": "term-sheet", "prepared_by": "가상 기관", "disclaimer": "예시 문구"}
    if origin == "frontmatter":
        metadata["style"] = {"logo": "example.png"}
    elif origin == "document-house":
        metadata["house"] = str(config)
    source = _source(tmp_path, "가상 본문", **metadata)

    def invalid_anchor(path):
        raise OSError("예시 logo anchoring failure")

    monkeypatch.setattr(Path, "absolute", invalid_anchor)
    with pytest.raises(
        (OSError, StrictValidationError), match="예시 logo anchoring failure"
    ) as caught:
        model = parse_markdown_file(str(source))
        IBDocumentRenderer(
            options=RenderOptions(
                strict=True,
                house=str(config) if origin == "operator-house" else None,
            )
        ).render(model)
    assert isinstance(caught.value, StrictValidationError) is (origin != "operator-house")


@pytest.mark.parametrize("origin", ["frontmatter", "house"])
def test_logo_link_parent_is_not_lexically_collapsed(tmp_path, origin):
    import subprocess

    outside = _image(tmp_path / "outside" / "example.png")
    target = outside.parent / "nested"
    target.mkdir()
    folder = tmp_path / "document"
    _image(folder / "example.png")
    link = folder / "example-link"
    if os.name == "nt":
        result = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(link), str(target)], capture_output=True
        )
        if result.returncode:
            pytest.skip("Junction creation is not permitted")
    else:
        link.symlink_to(target, target_is_directory=True)
    logo = "example-link/../example.png"
    metadata = {"profile": "term-sheet", "prepared_by": "가상 기관", "disclaimer": "예시 문구"}
    if origin == "house":
        metadata["house"] = str(_house(folder, logo))
    else:
        metadata["style"] = {"logo": logo}
    source = _source(folder, "가상 본문", **metadata)
    model = parse_markdown_file(str(source))
    stored = (
        load_house(model.metadata.extra["house"])["style"]["logo"]
        if origin == "house"
        else model.metadata.extra["style"]["logo"]
    )
    # Anchoring must leave the components for the shared resolver. Windows and
    # POSIX differ in how their final path resolver handles a junction plus '..'.
    assert stored == str(folder / logo)
    assert ".." in Path(stored).parts
