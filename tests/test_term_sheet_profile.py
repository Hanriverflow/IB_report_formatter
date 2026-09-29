"""Term-sheet contracts: A1 boilerplate/entry points and A2 document rendering."""

import subprocess
import sys
from dataclasses import FrozenInstanceError
from io import BytesIO
from pathlib import Path

import pytest
import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from lxml import etree

import ib_renderer
import term_sheet
from converters import get_default_registry
from document_model import CodeBlock, DocumentMetadata, DocumentModel, ElementType, TextRun
from document_profiles import RenderOptions, get_profile, load_style, resolve_options
from docx_audit import inspect_document
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
def test_closed_empty_confirmation_fence_renders_the_box(fence):
    model = MarkdownParser().parse(markdown(fence, confirmation={"intro": "가상 확인"}))
    element = model.elements[0]
    assert element.element_type == ElementType.CONFIRMATION
    assert element.content.source == fence
    assert not model.warnings
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    saved = reopen(renderer.render(model))
    assert not renderer.errors
    text = "\n".join(saved.element.xpath(".//w:t/text()"))
    assert "가상 확인" in text and "confirmation" not in text


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


# ═══════════════════════════════════════════════════════════════════════════════
# A2 RENDERING: OPENING, BODY, STYLES, CONFIRMATION, HEADER/FOOTER
# ═══════════════════════════════════════════════════════════════════════════════

CONFIRMATION = {
    "intro": "아래 항목을 확인하시고 ( *자필* )로 기재하여 주시기 바랍니다.",
    "items": [
        "□ 본인은 본 Term Sheet의 가상 조건과 위험에 대하여 설명을 들었습니다. (          )",
        "□ 본인은 가상 수수료 조건을 확인하였습니다. (          )",
    ],
    "signature": "고객확인 : 20    .    .    .   (서명/인)",
}


W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def find_all(element, path):
    """Namespace-explicit XPath; unregistered elements (docDefaults) lack python-docx's map."""
    return etree.XPath(path, namespaces=W_NS)(element)


def xml_values(element, path, name="w:val"):
    return [node.get(qn(name)) for node in find_all(element, path)]


def printable_twips(section):
    return round((section.page_width - section.left_margin - section.right_margin) / 635)


def test_opening_block_order_typography_and_single_title():
    content = markdown(f"# {TITLE}\n\n## 1. 본건 개요\n\n본문.", date="2026. 09.", version="v3")
    renderer, document = render(content, strict=True)
    saved = reopen(document)
    paragraphs = saved.paragraphs
    assert [p.text for p in paragraphs[:6]] == [TITLE, "Term Sheet", "2026. 09.", PREPARED_BY, DISCLAIMER, "1. 본건 개요"]
    assert sum(TITLE in p.text for p in paragraphs) == 1
    title, subtitle, date, prepared, disclaimer = paragraphs[:5]
    assert title.style.name == "Title"
    assert not title._p.xpath("./w:pPr/w:outlineLvl") and not title.style.element.xpath("./w:pPr/w:outlineLvl")
    for paragraph, size in ((title, 20), (subtitle, 16)):
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        assert all(run.bold and run.font.size.pt == size and str(run.font.color.rgb) == "1A2270" for run in paragraph.runs)
    for paragraph in (date, prepared):
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        assert {run.font.size.pt for run in paragraph.runs} == {10}
    assert disclaimer.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    assert {run.font.size.pt for run in disclaimer.runs} == {7}
    assert {str(run.font.color.rgb) for run in disclaimer.runs} == {"555555"}
    borders = disclaimer._p.xpath("./w:pPr/w:pBdr/*")
    assert [(etree.QName(node).localname, node.get(qn("w:sz")), node.get(qn("w:color"))) for node in borders] == [
        ("top", "4", "9AA5C4"), ("bottom", "4", "9AA5C4"),
    ]
    assert saved.core_properties.title == TITLE and saved.core_properties.subject == "Term Sheet"
    assert not renderer.errors


def test_opening_prefers_display_runs_and_parses_house_emphasis():
    model = MarkdownParser().parse(markdown(
        date="2026. 09.", prepared_by="라마바은행 **자본시장부**",
        disclaimer="본 자료는 ( *주요내용* ) 요약입니다.\n둘째 줄입니다.\n",
    ))
    model.metadata.display_runs = {
        "title": [TextRun("가나다 표시 제목", italic=True)],
        "subtitle": [TextRun("가상 부제")],
        "date": [TextRun("2026. 10.")],
    }
    saved = reopen(IBDocumentRenderer().render(model))
    paragraphs = saved.paragraphs
    assert [p.text for p in paragraphs[:4]] == ["가나다 표시 제목", "가상 부제", "2026. 10.", "라마바은행 자본시장부"]
    assert paragraphs[0].runs[0].italic and paragraphs[0].runs[0].bold
    assert [run.text for run in paragraphs[3].runs if run.bold] == ["자본시장부"]
    first, second = paragraphs[4:6]
    assert first.text == "본 자료는 ( 주요내용 ) 요약입니다." and second.text == "둘째 줄입니다."
    assert [run.text for run in first.runs if run.italic] == ["주요내용"]
    assert all(p._p.xpath("./w:pPr/w:pBdr/w:top") and p._p.xpath("./w:pPr/w:pBdr/w:bottom") for p in (first, second))
    assert saved.core_properties.title == TITLE


def test_opening_disclaimer_is_kept_without_end_disclaimer():
    saved = reopen(render(markdown(), include_disclaimer=False)[1])
    assert DISCLAIMER in [p.text for p in saved.paragraphs]


def test_paragraph_breaks_become_marker_indented_paragraphs():
    body = "① 첫째 조건<br>※ 가상 주석입니다.<br>둘째 줄"
    saved = reopen(render(markdown(body), strict=True)[1])
    first, note, plain = saved.paragraphs[4:7]
    assert [first.text, note.text, plain.text] == ["① 첫째 조건", "※ 가상 주석입니다.", "둘째 줄"]
    assert (xml_values(first._p, "./w:pPr/w:ind", "w:left"), xml_values(first._p, "./w:pPr/w:ind", "w:hanging")) == (["255"], ["255"])
    assert (xml_values(note._p, "./w:pPr/w:ind", "w:left"), xml_values(note._p, "./w:pPr/w:ind", "w:hanging")) == (["227"], ["227"])
    assert {run.font.size.pt for run in note.runs} == {8}
    assert {run.font.size.pt for run in first.runs} == {9}
    assert not plain._p.xpath("./w:pPr/w:ind")
    legacy = reopen(IBDocumentRenderer().render(MarkdownParser(profile="plain").parse(body)))
    assert len(legacy.paragraphs) == 1 and len(legacy.paragraphs[0]._p.xpath(".//w:br")) == 2


def test_term_sheet_headings_use_the_accent_and_heading_runs():
    model = MarkdownParser().parse(markdown("## 1. 원문 제목\n\n### 가. 세부 조건\n\n본문."))
    model.elements[0].content.runs = [TextRun("1. 치환 제목", italic=True)]
    saved = reopen(IBDocumentRenderer(options=RenderOptions(strict=True)).render(model))
    h2 = next(p for p in saved.paragraphs if p.style.name == "Heading 2")
    h3 = next(p for p in saved.paragraphs if p.style.name == "Heading 3")
    assert h2.text == "1. 치환 제목" and all(run.italic and run.bold for run in h2.runs)
    assert h3.text == "가. 세부 조건"
    for paragraph in (h2, h3):
        assert {str(run.font.color.rgb) for run in paragraph.runs} == {"1A2270"}
    heading2 = saved.styles["Heading 2"]
    assert str(heading2.font.color.rgb) == "1A2270"
    rule = heading2.element.xpath("./w:pPr/w:pBdr/w:bottom")[0]
    assert (rule.get(qn("w:val")), rule.get(qn("w:sz")), rule.get(qn("w:color"))) == ("single", "8", "1A2270")
    assert str(saved.styles["Heading 3"].font.color.rgb) == "1A2270"


def test_term_sheet_document_defaults_are_local_to_term_sheets():
    saved = reopen(render(markdown())[1])
    defaults = find_all(saved.styles.element, "./w:docDefaults")[0]
    paragraph_defaults = find_all(defaults, "./w:pPrDefault/w:pPr")[0]
    assert xml_values(paragraph_defaults, "./w:kinsoku") == ["1"]
    assert xml_values(paragraph_defaults, "./w:wordWrap") == ["0"]
    names = [etree.QName(child).localname for child in paragraph_defaults]
    assert names.index("kinsoku") < names.index("wordWrap") < names.index("spacing")
    assert xml_values(defaults, "./w:rPrDefault/w:rPr/w:lang", "w:eastAsia") == ["ko-KR"]
    assert xml_values(defaults, "./w:rPrDefault/w:rPr/w:sz") == ["18"]
    plain = reopen(IBDocumentRenderer().render(MarkdownParser(profile="plain").parse("본문.")))
    plain_defaults = find_all(plain.styles.element, "./w:docDefaults")[0]
    assert not find_all(plain_defaults, ".//w:wordWrap") and not find_all(plain_defaults, ".//w:kinsoku")
    assert xml_values(plain_defaults, "./w:rPrDefault/w:rPr/w:lang", "w:eastAsia") == ["en-US"]
    assert xml_values(plain_defaults, "./w:rPrDefault/w:rPr/w:sz") == ["22"]


def test_confirmation_box_has_two_rows_kept_on_one_page():
    content = markdown("## 고객 확인\n\n확인 안내 문단입니다.\n\n```confirmation\n```", confirmation=CONFIRMATION)
    renderer, document = render(content, strict=True)
    saved = reopen(document)
    assert not renderer.errors
    table = saved.element.body.xpath("./w:tbl")[0]
    rows = table.xpath("./w:tr")
    assert len(rows) == 2 and all(len(row.xpath("./w:tc")) == 1 for row in rows)
    assert all(row.xpath("./w:trPr/w:cantSplit") for row in rows)
    body_cell, signature_cell = (row.xpath("./w:tc")[0] for row in rows)
    intro, *items = body_cell.xpath("./w:p")
    assert len(items) == 2 and all(p.xpath("./w:pPr/w:keepNext") for p in [intro, *items])
    assert set(xml_values(intro, ".//w:r/w:rPr/w:sz")) == {"16"}
    assert set(xml_values(intro, ".//w:r/w:rPr/w:color")) == {"555555"}
    assert intro.xpath(".//w:r[w:rPr/w:i[not(@w:val)]]/w:t/text()") == ["자필"]
    for item in items:
        assert set(xml_values(item, ".//w:r/w:rPr/w:sz")) == {"20"}
        assert (xml_values(item, "./w:pPr/w:ind", "w:left"), xml_values(item, "./w:pPr/w:ind", "w:hanging")) == (["255"], ["255"])
    signature = signature_cell.xpath("./w:p")[0]
    assert "".join(signature.xpath(".//w:t/text()")) == CONFIRMATION["signature"]
    assert xml_values(signature, "./w:pPr/w:jc") == ["center"]
    assert len(signature.xpath(".//w:r/w:rPr/w:b[not(@w:val)]")) == len(signature.xpath(".//w:r"))
    assert set(xml_values(signature, ".//w:r/w:rPr/w:sz")) == {"20"}
    assert xml_values(signature_cell, "./w:tcPr/w:shd", "w:fill") == ["F2F2F2"]
    assert not signature.xpath("./w:pPr/w:keepNext")
    previous = table.getprevious()
    assert "".join(previous.xpath(".//w:t/text()")) == "확인 안내 문단입니다." and previous.xpath("./w:pPr/w:keepNext")
    borders = table.xpath("./w:tblPr/w:tblBorders/*")
    assert {(node.get(qn("w:sz")), node.get(qn("w:color"))) for node in borders} == {("4", "9AA5C4")}


def test_confirmation_box_omits_an_empty_row():
    content = markdown("```confirmation\n```", confirmation={"signature": "가상 서명"})
    renderer, document = render(content, strict=True)
    rows = reopen(document).element.body.xpath("./w:tbl/w:tr")
    assert len(rows) == 1 and "".join(rows[0].xpath(".//w:t/text()")) == "가상 서명"


@pytest.mark.parametrize("confirmation", [{"intro": " "}, {"signature": ""}])
def test_blank_confirmation_texts_keep_the_lossless_fallback(confirmation):
    model = MarkdownParser().parse(markdown("```confirmation\n```", confirmation=confirmation))
    renderer = IBDocumentRenderer()
    saved = reopen(renderer.render(model))
    assert "confirmation" in "\n".join(saved.element.xpath(".//w:t/text()"))
    assert any("confirmation" in error.lower() for error in renderer.errors)
    with pytest.raises(ValueError, match="validation"):
        IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)


def test_header_and_footer_follow_each_section_width():
    body = "앞 문단.\n\n| 구 분 | 내 용 |\n|---|---|\n| 가 | 나 |\n\n뒤 문단."
    saved = reopen(render(markdown(body, version="v3", tables=[{"landscape": True}]), strict=True)[1])
    assert len(saved.sections) == 3
    positions = []
    for section in saved.sections:
        header = section.header.paragraphs
        assert len(header) == 1 and header[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
        label = header[0].runs[0]
        assert label.text == "Strictly Confidential" and label.italic
        assert label.font.size.pt == 7.5 and str(label.font.color.rgb) == "888888"
        footer = section.footer.paragraphs[0]
        assert footer.text.startswith("Term Sheet v3\t")
        assert footer._p.xpath(".//w:instrText/text()") == ["PAGE", "NUMPAGES"]
        stops = [(stop.get(qn("w:val")), int(stop.get(qn("w:pos")))) for stop in footer._p.xpath("./w:pPr/w:tabs/w:tab")]
        assert stops == [("right", printable_twips(section))]
        assert set(xml_values(footer._p, ".//w:r/w:rPr/w:sz")) == {"15"}
        assert set(xml_values(footer._p, ".//w:r/w:rPr/w:color")) == {"888888"}
        positions.append(stops[0][1])
    assert positions[0] == positions[2] < positions[1]


@pytest.mark.parametrize(
    ("fields", "options"),
    [({"layout": {"confidential": False}}, {}), ({}, {"confidential": False}), ({"confidential_label": ""}, {})],
)
def test_confidential_header_can_be_suppressed(fields, options):
    saved = reopen(render(markdown(**fields), **options)[1])
    assert all(not section.header.paragraphs[0].text for section in saved.sections)


def test_footer_without_version_starts_with_page_numbers():
    saved = reopen(render(markdown())[1])
    footer = saved.sections[0].footer.paragraphs[0]
    assert footer.text.startswith("\t") and "Term Sheet" not in footer.text


def test_complete_term_sheet_renders_strict_without_audit_issues():
    body = (
        f"# {TITLE}\n\n## 1. 본건 개요\n\n| 구 분 | 내 용 |\n|---|---|\n| 발행금액 | • 500억원 |\n\n"
        "## 2. 주요 금융조건\n\n| 구 분 | << | 내 용 |\n|---|---|---|\n"
        "| ABCP | 금액 | • 500억원 |\n| ^^ | CAP | • 변동<br>※ 가상 기준 |\n\n"
        "※ 심사 과정에서 상기 조건은 변경될 수 있음\n\n### 가. 참고\n\n"
        "① 가상 설명<br>② 가상 설명[^1]\n\n- 가상 목록\n\n```confirmation\n```\n\n[^1]: 가상 각주입니다.\n"
    )
    content = markdown(
        body, date="2026. 09.", version="v3", confirmation=CONFIRMATION,
        tables=[{"caption": "개요", "unit": "억원"}, {"note": "(VAT 별도)", "landscape": True}],
    )
    renderer, document = render(content, strict=True)
    saved = reopen(document)
    assert not renderer.errors
    assert inspect_document(saved).issues == []
    assert saved.element.xpath(".//w:footnoteReference")
