"""A1 term-sheet contracts, boilerplate validation and shared entry points."""

import subprocess
import sys
from dataclasses import FrozenInstanceError
from io import BytesIO
from pathlib import Path

import pytest
import yaml
from docx import Document
from docx.shared import Inches, Pt

import ib_renderer
import term_sheet
from converters import get_default_registry
from document_model import CodeBlock, DocumentMetadata, DocumentModel, ElementType
from document_profiles import RenderOptions, get_profile, load_style, resolve_options
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file

TITLE = "가나다머티리얼즈㈜ 조건 검토"
PREPARED_BY = "라마바은행 자본시장부"
DISCLAIMER = "본 문서는 가상 조건 검토용이며 거래 확약이 아닙니다."


def markdown(body="검토 본문.", **fields):
    """Build entirely fictional term-sheet input through normal frontmatter."""
    metadata = {
        "profile": "term-sheet",
        "title": TITLE,
        "prepared_by": PREPARED_BY,
        "disclaimer": DISCLAIMER,
    }
    metadata.update(fields)
    return "---\n" + yaml.safe_dump(metadata, allow_unicode=True) + "---\n" + body


def render(content, **options):
    model = MarkdownParser().parse(content)
    renderer = IBDocumentRenderer(options=RenderOptions(**options))
    return renderer, renderer.render(model)


def reopen(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return Document(buffer)


def house_file(tmp_path, **fields):
    house = tmp_path / "house.yaml"
    data = {"prepared_by": PREPARED_BY, "disclaimer": DISCLAIMER}
    data.update(fields)
    house.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    return house


def house_markdown(path, body="검토 본문.", **fields):
    data = {"profile": "term-sheet", "title": TITLE, "house": str(path)}
    data.update(fields)
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def test_term_sheet_profile_defaults_and_subtitle():
    model = MarkdownParser().parse(markdown())
    options = resolve_options(model.metadata, RenderOptions())
    assert get_profile("term-sheet").a4
    assert not options.profile.is_ib
    assert options.confidential
    assert not options.cover and not options.toc and not options.disclaimer
    assert model.metadata.subtitle == "Term Sheet"
    renderer, document = render(markdown(), strict=True)
    saved = reopen(document)
    assert renderer.term_sheet_texts.prepared_by == PREPARED_BY
    assert saved.sections[0].page_width.mm == pytest.approx(210, abs=0.01)
    assert "검토 본문." in saved.element.xpath(".//w:t/text()")


@pytest.mark.parametrize("title", [None, "", " \t ", "Document", "IB Report", "가" * 256])
def test_term_sheet_rejects_invalid_title_before_heading_inference(title):
    with pytest.raises(ValueError, match="title"):
        MarkdownParser().parse(markdown("# 본문에서 추론하면 안 되는 제목", title=title))


def test_term_sheet_requires_explicit_title_before_heading_inference():
    content = f"---\nprofile: term-sheet\nprepared_by: {PREPARED_BY}\ndisclaimer: {DISCLAIMER}\n---\n# 추론 제목"
    with pytest.raises(ValueError, match="title"):
        MarkdownParser().parse(content)
    with pytest.raises(ValueError, match="title"):
        MarkdownParser(profile="term-sheet").parse("# 추론 제목")


@pytest.mark.parametrize("title", ["", "Document", "IB Report", "가" * 256])
def test_handbuilt_term_sheet_validates_title_in_render_preflight(title):
    model = DocumentModel(metadata=DocumentMetadata(title=title, profile="term-sheet", extra={"prepared_by": PREPARED_BY, "disclaimer": DISCLAIMER}))
    with pytest.raises(ValueError, match="title"):
        IBDocumentRenderer().render(model)


@pytest.mark.parametrize("subtitle", ["", " \t "])
def test_term_sheet_blank_subtitle_gets_default(subtitle):
    assert MarkdownParser().parse(markdown(subtitle=subtitle)).metadata.subtitle == "Term Sheet"


def test_term_sheet_checks_blank_subtitle_length_before_defaulting():
    assert MarkdownParser().parse(markdown(subtitle=" " * 255)).metadata.subtitle == "Term Sheet"
    with pytest.raises(ValueError, match="subtitle"):
        MarkdownParser().parse(markdown(subtitle=" " * 256))


@pytest.mark.parametrize("field", ["title", "subtitle"])
def test_term_sheet_title_and_subtitle_length_boundary(field):
    model = MarkdownParser().parse(markdown(**{field: "가" * 255}))
    IBDocumentRenderer().render(model)
    with pytest.raises(ValueError, match=field):
        MarkdownParser().parse(markdown(**{field: "가" * 256}))


@pytest.mark.parametrize("field", ["title", "subtitle"])
@pytest.mark.parametrize("value", [None, False, 37, ["문자열"], {"text": "문자열"}])
def test_term_sheet_rejects_non_text_display_metadata(field, value):
    with pytest.raises(ValueError, match=field):
        MarkdownParser().parse(markdown(**{field: value}))


def test_term_sheet_style_values_and_immutable_defaults():
    style = load_style(get_profile("term-sheet"))
    assert str(style.NAVY) == style.NAVY_HEX == "1A2270"
    assert style.HEADING_FONT == style.BODY_FONT == style.KOREAN_FONT == "Malgun Gothic"
    expected = {
        "BODY_SIZE": 9, "TABLE_HEADER_SIZE": 9, "TABLE_BODY_SIZE": 9,
        "SMALL_SIZE": 8, "H1_SIZE": 14, "H2_SIZE": 13, "H3_SIZE": 10,
        "H2_SPACE_BEFORE": 15, "H2_SPACE_AFTER": 7,
        "H3_SPACE_BEFORE": 12, "H3_SPACE_AFTER": 5, "BODY_SPACE_AFTER": 3,
        "TS_TITLE_SIZE": 20, "TS_SUBTITLE_SIZE": 16, "TS_META_SIZE": 10,
        "TS_NOTE_SIZE": 8, "TS_DISCLAIMER_SIZE": 7, "TS_HEADER_FOOTER_SIZE": 7.5,
    }
    assert {key: getattr(style, key).pt for key in expected} == expected
    assert style.BODY_LINE_SPACING == 1.05
    assert [getattr(style, key).mm for key in ("TOP_MARGIN", "BOTTOM_MARGIN", "LEFT_MARGIN", "RIGHT_MARGIN")] == pytest.approx([18, 16, 15, 15], abs=0.001)
    assert style.TABLE_HEADER_BG == "DCE3F5"
    assert not style.TABLE_ZEBRA and not style.BODY_JUSTIFY and not style.HEADING_BORDER
    assert style.TS_LABEL_BG_HEX == "F2F5FC"
    assert style.TS_BORDER_HEX == "9AA5C4"
    assert style.TS_MUTED_HEX == "555555"
    assert style.TS_CONFIDENTIAL_HEX == "888888"
    assert style.TS_LABEL_WIDTH == Inches(33.5 / 25.4)
    assert style.TS_SUBLABEL_WIDTH == Inches(30 / 25.4)
    assert term_sheet.ROW_SPLIT_THRESHOLD == 12
    with pytest.raises(FrozenInstanceError):
        style.TS_NOTE_SIZE = Pt(9)


def test_term_sheet_theme_overrides_are_local(tmp_path):
    theme = tmp_path / "theme.yaml"
    theme.write_text("primary_color: '345678'\nTS_LABEL_BG_HEX: 'EEDDCC'\nTS_TITLE_SIZE: 22", encoding="utf-8")
    style = load_style(get_profile("term-sheet"), str(theme))
    assert str(style.NAVY) == style.NAVY_HEX == "345678"
    assert style.TS_LABEL_BG_HEX == "EEDDCC" and style.TS_TITLE_SIZE.pt == 22
    assert load_style(get_profile("term-sheet")).TS_LABEL_BG_HEX == "F2F5FC"
    assert load_style(get_profile("plain")).NAVY_HEX == "202020"
    mono = load_style(get_profile("term-sheet"), "mono")
    assert str(mono.NAVY) == mono.NAVY_HEX == "202020"
    assert mono.BODY_SIZE.pt == 9


def test_house_texts_are_frozen_and_confirmation_items_are_tuple(tmp_path):
    path = house_file(tmp_path, confirmation={"intro": "가상 확인 안내", "items": ["가상 조건을 확인했습니다."], "signature": "가상 서명"})
    renderer, _ = render(house_markdown(path))
    texts = renderer.term_sheet_texts
    assert isinstance(texts, term_sheet.TermSheetTexts)
    assert texts.confidential_label == "Strictly Confidential"
    assert texts.confirmation.items == ("가상 조건을 확인했습니다.",)
    assert isinstance(texts.confirmation, term_sheet.ConfirmationText)
    with pytest.raises(FrozenInstanceError):
        texts.prepared_by = "변경"
    with pytest.raises(FrozenInstanceError):
        texts.confirmation.intro = "변경"


def test_house_frontmatter_presence_wins_even_when_empty(tmp_path):
    path = house_file(tmp_path, confidential_label="가상 기밀", confirmation={"intro": "house 안내"})
    renderer, _ = render(house_markdown(path, prepared_by="가나다머티리얼즈㈜ 검토부", confidential_label="", confirmation={"signature": "frontmatter 서명"}))
    texts = renderer.term_sheet_texts
    assert texts.prepared_by == "가나다머티리얼즈㈜ 검토부"
    assert texts.disclaimer == DISCLAIMER
    assert texts.confidential_label == ""
    assert texts.confirmation.intro == "" and texts.confirmation.signature == "frontmatter 서명"
    for key in ("prepared_by", "disclaimer"):
        with pytest.raises(ValueError, match=key):
            render(house_markdown(path, **{key: ""}))


@pytest.mark.parametrize("key", ["prepared_by", "disclaimer"])
@pytest.mark.parametrize("value", [None, False, 12, [], {}, "", " \t "])
def test_term_sheet_requires_nonblank_text_boilerplate(key, value):
    with pytest.raises(ValueError, match=key):
        render(markdown(**{key: value}))


@pytest.mark.parametrize("key", ["prepared_by", "disclaimer"])
def test_term_sheet_preflight_rejects_missing_text_before_document_creation(monkeypatch, key):
    fields = {"profile": "term-sheet", "title": TITLE, "prepared_by": PREPARED_BY, "disclaimer": DISCLAIMER}
    del fields[key]
    model = MarkdownParser().parse("---\n" + yaml.safe_dump(fields, allow_unicode=True) + "---\n검토 본문.")
    renderer = IBDocumentRenderer()
    calls = []
    monkeypatch.setattr(ib_renderer, "Document", lambda: calls.append("created"))
    with pytest.raises(ValueError, match=key):
        renderer.render(model)
    assert not calls


@pytest.mark.parametrize("confirmation", [{}, {"typo": "값"}, {"intro": 4}, {"signature": None}, {"items": "항목"}, {"items": []}, {"items": [""]}, {"items": ["   "]}, {"items": [False]}, None])
def test_confirmation_schema_rejects_invalid_values(confirmation):
    with pytest.raises(ValueError, match="confirmation"):
        render(markdown(confirmation=confirmation))


@pytest.mark.parametrize("confirmation", [{"intro": "가상 안내"}, {"items": ["가상 항목"]}, {"signature": "가상 서명"}, {"intro": ""}, {"signature": ""}])
def test_confirmation_schema_accepts_each_supported_field(confirmation):
    renderer, _ = render(markdown(confirmation=confirmation))
    assert renderer.term_sheet_texts.confirmation is not None


@pytest.mark.parametrize("value", [False, 3, None, [], {}])
def test_confidential_label_requires_string(value):
    with pytest.raises(ValueError, match="confidential_label"):
        render(markdown(confidential_label=value))


@pytest.mark.parametrize("content", ["- item", "null", "[]", "unknown: text", "prepared_by: 9", "disclaimer: [text]", "confidential_label: false", "confirmation: {unknown: text}", "confirmation: {items: []}", "prepared_by: ["])
def test_house_schema_rejects_invalid_yaml_and_fields(tmp_path, content):
    path = tmp_path / "invalid-house.yaml"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(ValueError):
        render(markdown(house=str(path)))


def test_missing_house_file_is_rejected(tmp_path):
    with pytest.raises((ValueError, FileNotFoundError), match="house"):
        render(markdown(house=str(tmp_path / "missing-house.yaml")))


def test_partial_house_is_completed_by_frontmatter(tmp_path):
    path = tmp_path / "house.yaml"
    path.write_text("confidential_label: 가상 기밀", encoding="utf-8")
    renderer, _ = render(markdown(house=str(path)))
    assert renderer.term_sheet_texts.confidential_label == "가상 기밀"


def test_file_house_path_is_source_relative_and_registry_uses_it(tmp_path):
    house_file(tmp_path)
    source = tmp_path / "source.md"
    source.write_text(house_markdown("house.yaml"), encoding="utf-8")
    model = parse_markdown_file(str(source))
    assert Path(model.metadata.extra["house"]) == tmp_path / "house.yaml"
    assert Path(model.metadata.extra["house"]).is_absolute()
    renderer = IBDocumentRenderer()
    renderer.render(model)
    assert renderer.term_sheet_texts.prepared_by == PREPARED_BY
    registry = get_default_registry()
    registered_model = registry.convert(str(source))
    doc = registry.convert(registered_model, output_format="docx", strict=True)
    assert "검토 본문." in reopen(doc).element.xpath(".//w:t/text()")


@pytest.mark.parametrize("entrypoint", ["string", "stream", "registry-stream"])
def test_sourceless_relative_house_is_rejected(entrypoint):
    content = house_markdown("house.yaml")
    with pytest.raises(ValueError, match="house.*(absolute|source|relative)"):
        if entrypoint == "string":
            model = MarkdownParser().parse(content)
        elif entrypoint == "stream":
            model = parse_markdown_file(BytesIO(content.encode("utf-8")))
        else:
            model = get_default_registry().convert(BytesIO(content.encode("utf-8")), extension_hint="md")
        IBDocumentRenderer().render(model)


def test_absolute_house_on_stream_and_registry_override_skip_frontmatter_house(tmp_path):
    path = house_file(tmp_path)
    model = parse_markdown_file(BytesIO(house_markdown(path).encode("utf-8")))
    renderer = IBDocumentRenderer()
    renderer.render(model)
    assert renderer.term_sheet_texts.disclaimer == DISCLAIMER
    override_model = MarkdownParser().parse(house_markdown("unopened-house.yaml"))
    doc = get_default_registry().convert(override_model, output_format="docx", house=str(path), strict=True)
    assert "검토 본문." in reopen(doc).element.xpath(".//w:t/text()")


def test_cli_house_is_cwd_relative_and_overrides_unopened_frontmatter_house(tmp_path):
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    path = house_file(tmp_path)
    source = source_dir / "source.md"
    source.write_text(house_markdown("unopened-house.yaml"), encoding="utf-8")
    destination = tmp_path / "result.docx"
    script = Path(__file__).resolve().parents[1] / "md_to_word.py"
    result = subprocess.run([sys.executable, "-B", str(script), str(source), str(destination), "--house", path.name, "--strict"], cwd=tmp_path, capture_output=True, timeout=30)
    assert result.returncode == 0, result.stderr.decode("utf-8", errors="replace")
    assert "검토 본문." in Document(destination).element.xpath(".//w:t/text()")


def test_cli_invalid_house_fails_without_output(tmp_path):
    source = tmp_path / "source.md"
    source.write_text(markdown(), encoding="utf-8")
    house = tmp_path / "house.yaml"
    house.write_text("unknown: true", encoding="utf-8")
    destination = tmp_path / "result.docx"
    script = Path(__file__).resolve().parents[1] / "md_to_word.py"
    result = subprocess.run([sys.executable, "-B", str(script), str(source), str(destination), "--house", house.name], cwd=tmp_path, capture_output=True, timeout=30)
    assert result.returncode != 0
    assert not destination.exists()


def test_renderer_resets_resolved_texts_on_reuse_and_failed_preflight():
    renderer = IBDocumentRenderer()
    renderer.render(MarkdownParser().parse(markdown()))
    assert renderer.term_sheet_texts.prepared_by == PREPARED_BY
    renderer.render(MarkdownParser(profile="plain").parse("검토 본문."))
    assert renderer.term_sheet_texts is None
    renderer.render(MarkdownParser().parse(markdown()))
    with pytest.raises(ValueError, match="disclaimer"):
        renderer.render(MarkdownParser().parse(markdown(disclaimer="")))
    assert renderer.term_sheet_texts is None


@pytest.mark.parametrize("fence", ["```confirmation\n```", "~~~confirmation\n \n~~~"])
def test_closed_empty_confirmation_fence_has_payload_and_a1_lossless_fallback(fence):
    model = MarkdownParser().parse(markdown(fence, confirmation={"intro": "가상 확인"}))
    element = model.elements[0]
    assert element.element_type == ElementType.CONFIRMATION
    assert element.content.source == fence
    assert not model.warnings
    renderer = IBDocumentRenderer()
    saved = reopen(renderer.render(model))
    text = "\n".join(saved.element.xpath(".//w:t/text()"))
    assert "confirmation" in text
    assert renderer.errors
    with pytest.raises(ValueError, match="validation"):
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)


@pytest.mark.parametrize("fence", ["```confirmation\n가상 확인 내용\n```", "```confirmation\n", "```confirmation\n가상 확인 내용"])
def test_nonempty_or_unclosed_confirmation_fence_stays_code_with_warning(fence):
    model = MarkdownParser().parse(markdown(fence, confirmation={"intro": "가상 확인"}))
    element = model.elements[0]
    assert element.element_type == ElementType.CODE_BLOCK
    assert isinstance(element.content, CodeBlock)
    assert element.content.code == fence
    assert any("confirmation" in warning for warning in model.warnings)
    saved = reopen(IBDocumentRenderer().render(model))
    assert "confirmation" in "\n".join(saved.element.xpath(".//w:t/text()"))
    with pytest.raises(ValueError, match="validation"):
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)


def test_confirmation_fence_without_text_is_preserved_and_strict_rejects():
    model = MarkdownParser().parse(markdown("```confirmation\n```"))
    assert model.elements[0].element_type == ElementType.CONFIRMATION
    renderer = IBDocumentRenderer()
    saved = reopen(renderer.render(model))
    assert "confirmation" in "\n".join(saved.element.xpath(".//w:t/text()"))
    assert renderer.errors
    with pytest.raises(ValueError, match="validation"):
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)


@pytest.mark.parametrize("profile", ["plain", "ib-report", "business-report"])
def test_confirmation_fence_remains_regular_code_in_other_profiles(profile):
    model = MarkdownParser(profile=profile).parse("```confirmation\n가상 확인 내용\n```")
    assert model.elements[0].element_type == ElementType.CODE_BLOCK
    assert model.elements[0].content.code == "가상 확인 내용"
    assert not model.warnings
