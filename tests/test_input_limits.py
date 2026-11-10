"""Input limits reject before saving, independently of strict-mode diagnostics."""

import base64
from dataclasses import FrozenInstanceError, replace
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock

import pytest
from PIL import Image as PillowImage

from converters import get_default_registry
from document_model import DocumentModel, Element, ElementType, ListItem, Table, TableCell, TableRow
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file
from md_to_word import IBReportConverter, build_parser, run_batch_conversion, run_conversion


def limits(**values):
    from input_safety import InputLimits

    return InputLimits(**values)


def rejection(error, code, limit, measured):
    from input_safety import ConversionRejection, InputLimitError

    while error.__cause__ is not None:
        error = error.__cause__
    assert isinstance(error, InputLimitError)
    assert isinstance(error, (ConversionRejection, ValueError))
    assert error.code == code
    assert error.limit == limit
    assert error.measured == measured
    assert str(error).startswith(f"[{code}] ")
    assert str(limit) in str(error)
    assert str(measured) in str(error)


def png(path):
    PillowImage.new("RGB", (3, 2), "white").save(path)
    return path.read_bytes()


def case(tmp_path, kind):
    if kind == "input_bytes":
        text = "가상 본문"
        measured = len(text.encode("utf-8"))
    elif kind == "table_cells":
        text = "| 가상 | 예시 |\n| --- | --- |\n| a | b |"
        measured = 4
    elif kind == "list_depth":
        text = "- 가상\n    - 예시"
        measured = 2
    else:
        data = png(tmp_path / "image.png")
        text = "![가상](image.png)"
        measured = len(data) if kind == "image_bytes" else 6
    source = tmp_path / "input.md"
    source.write_bytes(text.encode("utf-8"))
    return source, measured


KINDS = ["input_bytes", "table_cells", "list_depth", "image_bytes", "image_pixels"]


@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("strict", [False, True])
def test_each_limit_stops_without_output_and_can_be_raised(tmp_path, kind, strict):
    source, measured = case(tmp_path, kind)
    output = tmp_path / "output.docx"
    options = RenderOptions(
        profile="plain", strict=strict, limits=limits(**{f"max_{kind}": measured - 1})
    )
    with pytest.raises((ValueError, RuntimeError)) as caught:
        IBReportConverter(str(source), str(output), options).convert()
    rejection(caught.value, f"limit_{kind}", measured - 1, measured)
    assert not output.exists()
    raised = replace(options, limits=limits(**{f"max_{kind}": measured}))
    IBReportConverter(str(source), str(output), raised).convert()
    assert output.is_file()


@pytest.mark.parametrize("kind", KINDS)
@pytest.mark.parametrize("value", ["0", "-1", "1.5", "unlimited"])
def test_cli_invalid_limits_are_argument_errors(kind, value):
    with pytest.raises(SystemExit) as caught:
        build_parser().parse_args([f"--max-{kind.replace('_', '-')}", value])
    assert caught.value.code == 2


@pytest.mark.parametrize("kind", KINDS)
def test_cli_adjusts_each_limit(tmp_path, kind):
    source, measured = case(tmp_path, kind)
    output = tmp_path / "output.docx"
    option = f"--max-{kind.replace('_', '-')}"
    args = build_parser().parse_args(
        [str(source), str(output), "--profile", "plain", option, str(measured - 1)]
    )
    assert run_conversion(source, args) == 1
    assert not output.exists()
    setattr(args, f"max_{kind}", measured)
    assert run_conversion(source, args) == 0


@pytest.mark.parametrize("entry", ["text", "stream", "registry", "nonseekable"])
def test_input_bytes_all_entrypoints(entry):
    class Stream(BytesIO):
        def seekable(self):
            return False

    text = "가상"
    policy = limits(max_input_bytes=5)
    stream = Stream(text.encode("utf-8"))
    with pytest.raises(ValueError) as caught:
        if entry == "text":
            MarkdownParser(limits=policy).parse(text)
        elif entry == "stream":
            parse_markdown_file(stream, limits=policy)
        else:
            get_default_registry().convert(
                stream if entry == "nonseekable" else BytesIO(text.encode()), limits=policy
            )
    rejection(caught.value, "limit_input_bytes", 5, 6)


def test_file_size_checked_before_read(tmp_path, monkeypatch):
    source = tmp_path / "input.md"
    source.write_bytes(b"12345")
    read = Mock(side_effect=AssertionError("must not open oversized input"))
    monkeypatch.setattr(Path, "open", read)
    with pytest.raises(ValueError) as caught:
        parse_markdown_file(str(source), limits=limits(max_input_bytes=4))
    rejection(caught.value, "limit_input_bytes", 4, 5)
    read.assert_not_called()


def test_stream_stops_at_first_excess_byte():
    class Stream(BytesIO):
        def read(self, size=-1):
            assert 0 < size <= 5
            return super().read(size)

    stream = Stream(b"x" * 100)
    with pytest.raises(ValueError) as caught:
        parse_markdown_file(stream, limits=limits(max_input_bytes=4))
    rejection(caught.value, "limit_input_bytes", 4, 5)
    assert stream.tell() == 5


def test_registry_detection_does_not_read_past_limit():
    class Stream(BytesIO):
        def read(self, size=-1):
            assert 0 < size <= 5
            return super().read(size)

    with pytest.raises(ValueError) as caught:
        get_default_registry().convert(Stream(b"x" * 100), limits=limits(max_input_bytes=4))
    rejection(caught.value, "limit_input_bytes", 4, 5)


def test_raw_file_bytes_not_reencoded_utf8_size(tmp_path):
    source = tmp_path / "input.md"
    source.write_bytes("가상".encode("cp949"))
    model = parse_markdown_file(str(source), profile="plain", limits=limits(max_input_bytes=4))
    assert model.elements[0].content.text == "가상"


@pytest.mark.parametrize("kind", ["table_cells", "list_depth"])
@pytest.mark.parametrize("strict", [False, True])
def test_hand_built_models_cannot_bypass_limits(tmp_path, kind, strict):
    if kind == "table_cells":
        content = Table(
            rows=[TableRow(cells=[TableCell(content="가상"), TableCell(content="예시")])],
            col_count=2,
        )
        element = Element(ElementType.TABLE, content)
    else:
        element = Element(ElementType.NUMBERED_LIST, ("1", ListItem(text="가상", indent_level=1)))
    model = DocumentModel(elements=[element])
    output = tmp_path / "output.docx"
    with pytest.raises(ValueError) as caught:
        get_default_registry().convert(
            model,
            output_format="docx",
            output_path=str(output),
            strict=strict,
            limits=limits(**{f"max_{kind}": 1}),
        )
    rejection(caught.value, f"limit_{kind}", 1, 2)
    assert not output.exists()


def test_html_expanded_grid_counts_header_and_spans():
    text = '<table><tr><th colspan="2" rowspan="2">가상</th></tr><tr></tr></table>'
    with pytest.raises(ValueError) as caught:
        MarkdownParser(limits=limits(max_table_cells=3)).parse(text)
    rejection(caught.value, "limit_table_cells", 3, 4)
    model = MarkdownParser(limits=limits(max_table_cells=4)).parse(text)
    assert len(model.elements[0].content.rows) == 2


def test_html_limit_precedes_existing_word_span_clamp():
    text = '<table><tr><th colspan="50001">가상</th></tr></table>'
    with pytest.raises(ValueError) as caught:
        MarkdownParser().parse(text)
    rejection(caught.value, "limit_table_cells", 50000, 50001)


@pytest.mark.parametrize("kind", ["image_bytes", "image_pixels"])
@pytest.mark.parametrize(
    "where",
    [
        "body",
        "inline",
        "cell",
        "html",
        "html_cell",
        "frontmatter",
        "house",
        "operator_house",
        "base64",
        "data_logo",
    ],
)
@pytest.mark.parametrize("strict", [False, True])
def test_image_limits_cover_all_origins(tmp_path, kind, where, strict):
    import yaml

    image = tmp_path / "image.png"
    data = png(image)
    uri = "data:image/png;base64," + base64.b64encode(data).decode()
    bodies = {
        "body": "![가상](image.png)",
        "inline": "가상 ![예시](image.png) 본문",
        "cell": "| 가상 |\n| --- |\n| ![예시](image.png) |",
        "html": '<img src="image.png" alt="가상">',
        "html_cell": '<table><tr><td><img src="image.png" alt="가상"></td></tr></table>',
        "base64": f"![가상]({uri})",
    }
    metadata = {"profile": "plain"}
    options = {}
    if where not in bodies:
        metadata = {
            "profile": "term-sheet",
            "title": "가상 조건",
            "prepared_by": "가상 작성자",
            "disclaimer": "예시 안내",
            "confidential_label": "가상 기밀",
        }
        if where in {"frontmatter", "data_logo"}:
            metadata["style"] = {"logo": uri if where == "data_logo" else "image.png"}
        else:
            house = tmp_path / "house.yaml"
            house.write_text(yaml.safe_dump({"style": {"logo": "image.png"}}), encoding="utf-8")
            if where == "house":
                metadata["house"] = "house.yaml"
            else:
                options["house"] = str(house)
    source = tmp_path / "input.md"
    source.write_text(
        "---\n"
        + yaml.safe_dump(metadata, allow_unicode=True)
        + "---\n"
        + bodies.get(where, "가상 본문"),
        encoding="utf-8",
    )
    if where == "operator_house":
        folder = tmp_path / "document"
        folder.mkdir()
        source = source.rename(folder / "input.md")
    measured = len(data) if kind == "image_bytes" else 6
    policy = limits(**{f"max_{kind}": measured - 1})
    registry = get_default_registry()
    model = registry.convert(source, limits=policy)
    output = tmp_path / "output.docx"
    with pytest.raises(ValueError) as caught:
        registry.convert(
            model, output_format="docx", output_path=str(output), strict=strict, **options
        )
    rejection(caught.value, f"limit_{kind}", measured - 1, measured)
    assert not output.exists()
    registry.convert(
        model,
        output_format="docx",
        output_path=str(output),
        strict=True,
        limits=limits(**{f"max_{kind}": measured}),
        **options,
    )
    assert output.exists()


def test_base64_excess_is_not_decoded(monkeypatch):
    decode = Mock(side_effect=AssertionError("oversized payload was decoded"))
    monkeypatch.setattr(base64, "b64decode", decode)
    model = MarkdownParser(limits=limits(max_image_bytes=2)).parse(
        "![가상](data:image/png;base64,YWJj)"
    )
    with pytest.raises(ValueError) as caught:
        IBDocumentRenderer().render(model)
    rejection(caught.value, "limit_image_bytes", 2, 3)
    decode.assert_not_called()


@pytest.mark.parametrize("size", [1, 2, 3, 4, 5])
def test_base64_padding_and_whitespace_measure_decoded_bytes(size, monkeypatch):
    from input_safety import decode_image

    data = b"x" * size
    encoded = " \n".join(base64.b64encode(data).decode())
    assert decode_image(encoded, limits(max_image_bytes=size)).getvalue() == data
    if size > 1:
        decode = Mock(side_effect=AssertionError("must reject before decoding"))
        monkeypatch.setattr(base64, "b64decode", decode)
        with pytest.raises(ValueError) as caught:
            decode_image(encoded, limits(max_image_bytes=size - 1))
        rejection(caught.value, "limit_image_bytes", size - 1, size)
        decode.assert_not_called()


def test_huge_pixel_dimensions_are_read_without_decoding(tmp_path, monkeypatch):
    import struct
    import zlib

    from docx.image.image import Image as DocxImage

    from input_safety import check_image_source

    path = tmp_path / "image.png"
    data = bytearray(png(path))
    data[16:24] = struct.pack(">II", 50000, 50000)
    data[29:33] = struct.pack(">I", zlib.crc32(data[12:29]))
    path.write_bytes(data)
    load = Mock(side_effect=AssertionError("header inspection must not load the image"))
    monkeypatch.setattr(PillowImage.Image, "load", load)
    monkeypatch.setattr(DocxImage, "from_file", load)
    with pytest.raises(ValueError) as caught:
        check_image_source(str(path))
    rejection(caught.value, "limit_image_pixels", 50000000, 2500000000)
    check_image_source(str(path), limits(max_image_pixels=2500000000))
    load.assert_not_called()


def test_image_bytes_checked_before_opening(tmp_path, monkeypatch):
    from input_safety import check_image_source

    path = tmp_path / "image.png"
    data = png(path)
    read = Mock(side_effect=AssertionError("oversized image must not be opened"))
    monkeypatch.setattr(Path, "open", read)
    with pytest.raises(ValueError) as caught:
        check_image_source(str(path), limits(max_image_bytes=len(data) - 1))
    rejection(caught.value, "limit_image_bytes", len(data) - 1, len(data))
    read.assert_not_called()


@pytest.mark.parametrize(
    "flag,helper",
    [
        ("--format", "auto_format_markdown"),
        ("--deepresearch-cleaner", "clean_deepresearch_markdown_file"),
    ],
)
def test_cleanup_checks_original_then_derived(tmp_path, monkeypatch, flag, helper):
    import md_to_word

    source = tmp_path / "input.md"
    source.write_text("x" * 10, encoding="utf-8")
    derived = tmp_path / "derived.md"
    derived.write_text("x", encoding="utf-8")
    cleanup = Mock(return_value=derived)
    monkeypatch.setattr(md_to_word, helper, cleanup)
    output = tmp_path / "output.docx"
    flags = [flag] if flag == "--format" else [flag, "on"]
    args = build_parser().parse_args(
        [str(source), str(output), "--profile", "plain", "--max-input-bytes", "5", *flags]
    )
    assert run_conversion(source, args) == 1
    cleanup.assert_not_called()
    assert not output.exists()
    source.write_text("x", encoding="utf-8")
    derived.write_text("x" * 10, encoding="utf-8")
    assert run_conversion(source, args) == 1
    cleanup.assert_called_once()
    assert not output.exists()


def test_batch_continues_after_limit_failure(tmp_path):
    (tmp_path / "bad.md").write_text("x" * 10, encoding="utf-8")
    (tmp_path / "good.md").write_text("가상", encoding="utf-8")
    output = tmp_path / "output"
    args = build_parser().parse_args(
        [str(tmp_path), str(output), "--batch", "--profile", "plain", "--max-input-bytes", "6"]
    )
    assert run_batch_conversion(tmp_path, args) == 1
    assert not (output / "bad_Report_Pro.docx").exists()
    assert (output / "good_Report_Pro.docx").exists()


def test_frontmatter_cannot_raise_limits():
    text = "---\nlimits:\n  max_table_cells: 9999\nmax_table_cells: 9999\n---\n| 가상 | 예시 |\n| --- | --- |\n"
    with pytest.raises(ValueError) as caught:
        MarkdownParser(limits=limits(max_table_cells=1)).parse(text)
    rejection(caught.value, "limit_table_cells", 1, 2)


def test_limits_are_immutable_and_reject_nonpositive_nonintegers():
    policy = limits()
    assert [getattr(policy, f"max_{kind}") for kind in KINDS] == [
        16 * 1024**2,
        50000,
        32,
        20 * 1024**2,
        50000000,
    ]
    with pytest.raises(FrozenInstanceError):
        policy.max_input_bytes = 1
    for value in [0, -1, True, 1.5, "2", None]:
        with pytest.raises(ValueError):
            limits(max_input_bytes=value)


def test_cleanup_reader_keeps_existing_universal_newlines(tmp_path):
    from md_to_word import _read_with_encoding

    path = tmp_path / "input.md"
    path.write_bytes(b"a\r\nb\rc\n")
    assert _read_with_encoding(path, limits(max_input_bytes=7)) == "a\nb\nc\n"


def test_render_override_is_request_scoped(tmp_path):
    data = png(tmp_path / "image.png")
    uri = "data:image/png;base64," + base64.b64encode(data).decode()
    policy = limits(max_image_pixels=5)
    model = MarkdownParser(profile="plain", limits=policy).parse(f"![가상]({uri})")
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    assert len(renderer.render(model, limits=limits(max_image_pixels=6)).inline_shapes) == 1
    with pytest.raises(ValueError) as caught:
        renderer.render(model)
    rejection(caught.value, "limit_image_pixels", 5, 6)
    assert model.limits is policy


def test_registry_render_options_carry_limits_on_both_stages(tmp_path):
    source, measured = case(tmp_path, "image_pixels")
    policy = limits(max_image_pixels=measured - 1)
    registry = get_default_registry()
    options = RenderOptions(limits=policy)
    model = registry.convert(source, render_options=options)
    assert model.limits is policy
    with pytest.raises(ValueError) as caught:
        registry.convert(model, output_format="docx")
    rejection(caught.value, "limit_image_pixels", measured - 1, measured)
