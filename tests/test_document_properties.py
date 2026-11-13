"""Document properties and the DRAFT watermark written from frontmatter (roadmap E8).

`version` becomes the custom property `Version`, shown through `DOCPROPERTY`
fields whose stored result is the version text, so the document reads correctly
before Word updates any field. `status: draft` puts Word's own watermark shape
into every header so Word's Remove Watermark clears all pages. Every case goes
through actual parsing, strict rendering and saving, then reads the saved package.
"""

import re
import shutil
import zipfile
from datetime import UTC, datetime
from pathlib import Path

import pytest
from lxml import etree

from document_profiles import RenderOptions
from ib_render_document_style import (
    VERSION_PROPERTY,
    WATERMARK_SHAPE_PREFIX,
    make_property_field,
)
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import MarkdownParser, parse_markdown_file

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
V = "urn:schemas-microsoft-com:vml"
OFFICE = "urn:schemas-microsoft-com:office:office"
FIELD_CODE = " DOCPROPERTY Version \\* MERGEFORMAT "
GENERAL_PROFILES = ["ib-report", "ib-memo", "plain", "business-report", "meeting-minutes"]

LANDSCAPE_BODY = """# 예시 문서

세로 쪽 문단입니다.

| 지표 | 2025 | 2026 |
|---|---|---|
| 매출 | 100 | 120 |

## 다시 세로

마지막 문단입니다.
"""


def _render(source: str, profile: str, output: Path) -> Path:
    """Parse and strictly render Markdown text, then save it."""
    model = MarkdownParser(profile=profile).parse(source)
    document = IBDocumentRenderer(options=RenderOptions(profile=profile, strict=True)).render(model)
    document.save(str(output))
    return output


def _render_file(path: Path, output: Path) -> Path:
    """Parse a Markdown file (house files resolve beside it) and strictly render it."""
    model = parse_markdown_file(str(path))
    document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    document.save(str(output))
    return output


def _custom_properties(docx: Path) -> dict[str, str]:
    """Name to text of every custom document property in the saved package."""
    with zipfile.ZipFile(docx) as package:
        root = etree.fromstring(package.read("docProps/custom.xml"))
    return {
        prop.get("name"): "".join(prop.itertext()) for prop in root.iter(f"{{{CUSTOM}}}property")
    }


def _story_parts(docx: Path, kind: str) -> dict[str, etree._Element]:
    """Header or footer parts of the saved package by name."""
    with zipfile.ZipFile(docx) as package:
        return {
            name: etree.fromstring(package.read(name))
            for name in package.namelist()
            if re.fullmatch(rf"word/{kind}\d+\.xml", name)
        }


def _property_fields(part: etree._Element) -> list[tuple[str, str]]:
    """(field code, stored result text) of each complex field in a part."""
    fields: list[tuple[str, str]] = []
    for paragraph in part.iter(f"{{{W}}}p"):
        state, code, result = None, "", ""
        for run in paragraph.iter(f"{{{W}}}r"):
            char = run.find(f"{{{W}}}fldChar")
            if char is not None:
                kind = char.get(f"{{{W}}}fldCharType")
                if kind == "begin":
                    state, code, result = "code", "", ""
                elif kind == "separate":
                    state = "result"
                elif kind == "end":
                    fields.append((code, result))
                    state = None
                continue
            if state == "code":
                code += "".join(t.text or "" for t in run.iter(f"{{{W}}}instrText"))
            elif state == "result":
                result += "".join(t.text or "" for t in run.iter(f"{{{W}}}t"))
    return [field for field in fields if "DOCPROPERTY" in field[0]]


def _visible_text(part: etree._Element) -> str:
    """Text a reader sees in a part: text and tabs, never field codes."""
    pieces = []
    for node in part.iter(f"{{{W}}}t", f"{{{W}}}tab"):
        if node.tag == f"{{{W}}}tab" and node.getparent().tag != f"{{{W}}}r":
            continue
        pieces.append("\t" if node.tag == f"{{{W}}}tab" else node.text or "")
    return "".join(pieces)


def _source(profile: str, extra: str = "", body: str = "# 예시 문서\n\n가상 본문입니다.\n") -> str:
    return f"---\nprofile: {profile}\ntitle: 예시 문서\ncompany: 예시금융\n{extra}---\n{body}"


@pytest.mark.parametrize("profile", GENERAL_PROFILES)
def test_version_is_property_and_header_field_after_company(tmp_path, profile):
    output = _render(_source(profile, 'version: "v1.2"\n'), profile, tmp_path / "out.docx")

    assert _custom_properties(output)[VERSION_PROPERTY] == "v1.2"
    headers = _story_parts(output, "header")
    assert headers
    for part in headers.values():
        assert _property_fields(part) == [(FIELD_CODE, "v1.2")]
        assert _visible_text(part).startswith("예시금융 | v1.2")
    for part in _story_parts(output, "footer").values():
        assert _property_fields(part) == []


@pytest.mark.parametrize("profile", ["ib-report", "plain"])
def test_version_alone_when_header_has_no_company(tmp_path, profile):
    source = f'---\nprofile: {profile}\ntitle: 예시 문서\ncompany: ""\nversion: "v1.2"\n---\n본문\n'
    output = _render(source, profile, tmp_path / "out.docx")

    for part in _story_parts(output, "header").values():
        assert _visible_text(part).startswith("v1.2")
        assert _property_fields(part) == [(FIELD_CODE, "v1.2")]


def test_every_section_header_holds_the_version_field(tmp_path):
    extra = 'version: "v2.0"\ntables:\n  - {type: financial, landscape: true}\n'
    output = _render(_source("ib-memo", extra, LANDSCAPE_BODY), "ib-memo", tmp_path / "out.docx")

    with zipfile.ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    sections = list(document.iter(f"{{{W}}}sectPr"))
    headers = _story_parts(output, "header")
    assert len(sections) == len(headers) == 3
    for part in headers.values():
        assert _property_fields(part) == [(FIELD_CODE, "v2.0")]


def test_office_letter_writes_the_property_but_shows_no_version(tmp_path):
    source = (REPOSITORY_ROOT / "samples/profiles/office-letter.md").read_text(encoding="utf-8")
    source = source.replace("---\n", '---\nversion: "v1.2"\n', 1)
    output = _render(source, "office-letter", tmp_path / "out.docx")

    assert _custom_properties(output)[VERSION_PROPERTY] == "v1.2"
    for kind in ("header", "footer"):
        for part in _story_parts(output, kind).values():
            assert _property_fields(part) == []
            assert "v1.2" not in _visible_text(part)


def _term_sheet_copy(tmp_path: Path, version_line: str | None) -> Path:
    """The minimal term-sheet sample beside its house file, with a chosen version line."""
    folder = REPOSITORY_ROOT / "samples/distribution"
    shutil.copy2(folder / "minimal-house.yaml", tmp_path / "minimal-house.yaml")
    lines = (folder / "minimal-term-sheet.md").read_text(encoding="utf-8").split("\n")
    kept: list[str] = []
    for line in lines:
        if not line.startswith("version:"):
            kept.append(line)
        elif version_line is not None:
            kept.append(version_line)
    path = tmp_path / "term-sheet.md"
    path.write_text("\n".join(kept), encoding="utf-8")
    return path


def test_term_sheet_footer_version_is_a_field_with_the_same_visible_text(tmp_path):
    source = _term_sheet_copy(tmp_path, 'version: "v1.1 / 가상 예시"')
    output = _render_file(source, tmp_path / "out.docx")

    assert _custom_properties(output)[VERSION_PROPERTY] == "v1.1 / 가상 예시"
    for part in _story_parts(output, "footer").values():
        assert _property_fields(part) == [(FIELD_CODE, "v1.1 / 가상 예시")]
        assert _visible_text(part).startswith("Term Sheet v1.1 / 가상 예시\t")
    for part in _story_parts(output, "header").values():
        assert _property_fields(part) == []


def test_term_sheet_without_version_keeps_an_empty_footer_label(tmp_path):
    output = _render_file(_term_sheet_copy(tmp_path, None), tmp_path / "out.docx")

    assert VERSION_PROPERTY not in _custom_properties(output)
    for part in _story_parts(output, "footer").values():
        assert _property_fields(part) == []
        assert _visible_text(part).startswith("\t")


@pytest.mark.parametrize("profile", [*GENERAL_PROFILES, "office-letter"])
def test_no_version_writes_no_property_and_no_field(tmp_path, profile):
    source = _source(profile)
    if profile == "office-letter":
        source = (REPOSITORY_ROOT / "samples/profiles/office-letter.md").read_text(encoding="utf-8")
    output = _render(source, profile, tmp_path / "out.docx")

    assert VERSION_PROPERTY not in _custom_properties(output)
    for kind in ("header", "footer"):
        for part in _story_parts(output, kind).values():
            assert _property_fields(part) == []


@pytest.mark.parametrize(("written", "shown"), [("1.10", "1.1"), ("2", "2"), ('" v3 "', "v3")])
def test_version_keeps_the_existing_metadata_coercion(tmp_path, written, shown):
    """An unquoted number renders as YAML read it; strict mode does not reject it."""
    output = _render(
        _source("ib-report", f"version: {written}\n"), "ib-report", tmp_path / "out.docx"
    )

    assert _custom_properties(output)[VERSION_PROPERTY] == shown
    for part in _story_parts(output, "header").values():
        assert _property_fields(part) == [(FIELD_CODE, shown)]


def test_blank_version_is_absent(tmp_path):
    output = _render(_source("ib-report", 'version: "  "\n'), "ib-report", tmp_path / "out.docx")

    assert VERSION_PROPERTY not in _custom_properties(output)


@pytest.mark.parametrize("name", ["Deal Name", "버전", "1st", ""])
def test_property_field_accepts_only_plain_names(name):
    from docx import Document

    run = Document().add_paragraph().add_run("값")
    with pytest.raises(ValueError, match="Unsupported document property name"):
        make_property_field(run, name)


# ═══════════════════════════════════════════════════════════════════════════════
# DRAFT WATERMARK
# ═══════════════════════════════════════════════════════════════════════════════


def _watermarks(part: etree._Element) -> list[tuple[str, str, str]]:
    """(VML id, o:spid, text) of each text-path shape in a part."""
    return [
        (shape.get("id"), shape.get(f"{{{OFFICE}}}spid"), textpath.get("string"))
        for shape in part.iter(f"{{{V}}}shape")
        for textpath in shape.iter(f"{{{V}}}textpath")
    ]


def _all_parts(docx: Path) -> dict[str, etree._Element]:
    """Every header, footer and the body part of the saved package."""
    parts = {**_story_parts(docx, "header"), **_story_parts(docx, "footer")}
    with zipfile.ZipFile(docx) as package:
        parts["word/document.xml"] = etree.fromstring(package.read("word/document.xml"))
    return parts


def test_draft_marks_every_section_header_with_words_shape(tmp_path):
    extra = "status: draft\ntables:\n  - {type: financial, landscape: true}\n"
    output = _render(_source("ib-memo", extra, LANDSCAPE_BODY), "ib-memo", tmp_path / "out.docx")

    headers = _story_parts(output, "header")
    assert len(headers) == 3
    shapes = [shape for part in headers.values() for shape in _watermarks(part)]
    assert len(shapes) == 3
    assert all(len(_watermarks(part)) == 1 for part in headers.values())
    assert all(shape_id.startswith(WATERMARK_SHAPE_PREFIX) for shape_id, _, _ in shapes)
    assert {text for _, _, text in shapes} == {"DRAFT"}
    assert len({shape_id for shape_id, _, _ in shapes}) == 3
    assert len({spid for _, spid, _ in shapes}) == 3
    for name in ("word/document.xml", *_story_parts(output, "footer")):
        assert _watermarks(_all_parts(output)[name]) == []


def _letter_source(extra: str) -> str:
    """The office-letter sample with extra frontmatter lines."""
    source = (REPOSITORY_ROOT / "samples/profiles/office-letter.md").read_text(encoding="utf-8")
    return source.replace("---\n", "---\n" + extra, 1)


@pytest.mark.parametrize("profile", [*GENERAL_PROFILES, "office-letter"])
def test_draft_is_marked_in_every_profile(tmp_path, profile):
    extra = "status: draft\n"
    source = _letter_source(extra) if profile == "office-letter" else _source(profile, extra)
    output = _render(source, profile, tmp_path / "out.docx")

    headers = _story_parts(output, "header")
    assert headers
    for part in headers.values():
        assert [text for _, _, text in _watermarks(part)] == ["DRAFT"]


def test_draft_term_sheet_is_marked(tmp_path):
    source = _term_sheet_copy(tmp_path, 'version: "v1.0 / 가상 예시"\nstatus: draft')
    output = _render_file(source, tmp_path / "out.docx")

    for part in _story_parts(output, "header").values():
        assert [text for _, _, text in _watermarks(part)] == ["DRAFT"]


def test_draft_cover_page_shares_the_marked_header(tmp_path):
    """The IB cover is on the first page of the only section, so it is marked too."""
    output = _render(_source("ib-report", "status: draft\n"), "ib-report", tmp_path / "out.docx")

    with zipfile.ZipFile(output) as package:
        document = package.read("word/document.xml").decode("utf-8")
    assert "titlePg" not in document
    headers = _story_parts(output, "header")
    assert len(headers) == 1
    assert [text for _, _, text in _watermarks(next(iter(headers.values())))] == ["DRAFT"]


@pytest.mark.parametrize("extra", ["status: final\n", ""])
@pytest.mark.parametrize("profile", [*GENERAL_PROFILES, "office-letter"])
def test_final_or_absent_status_adds_no_watermark(tmp_path, profile, extra):
    source = _letter_source(extra) if profile == "office-letter" else _source(profile, extra)
    output = _render(source, profile, tmp_path / "out.docx")

    for part in _all_parts(output).values():
        assert _watermarks(part) == []
        assert part.find(f".//{{{W}}}pict") is None


@pytest.mark.parametrize("value", ["Draft", "DRAFT", "초안", "yes", "true", "''", ""])
@pytest.mark.parametrize("strict", [True, False])
def test_unknown_status_is_rejected_in_both_modes(value, strict):
    source = _source("ib-report", f"status: {value}\n")
    with pytest.raises(StrictValidationError, match="status must be draft or final"):
        model = MarkdownParser(profile="ib-report").parse(source)
        IBDocumentRenderer(options=RenderOptions(strict=strict)).render(model)


@pytest.mark.parametrize("strict", [True, False])
def test_unknown_status_in_a_hand_built_model_is_rejected(strict):
    model = MarkdownParser(profile="plain").parse(_source("plain"))
    model.metadata.extra["status"] = "DRAFT"
    with pytest.raises(StrictValidationError, match="status must be draft or final"):
        IBDocumentRenderer(options=RenderOptions(strict=strict)).render(model)


# ═══════════════════════════════════════════════════════════════════════════════
# CORE PROPERTIES (FILE > INFO)
# ═══════════════════════════════════════════════════════════════════════════════


def _core(docx: Path):
    from docx import Document

    return Document(str(docx)).core_properties


@pytest.mark.parametrize("profile", [*GENERAL_PROFILES, "office-letter"])
def test_file_information_is_this_documents(tmp_path, profile):
    source = _letter_source("") if profile == "office-letter" else _source(profile)
    before = datetime.now(UTC).replace(tzinfo=None, microsecond=0)
    output = _render(source, profile, tmp_path / "out.docx")
    after = datetime.now(UTC).replace(tzinfo=None)

    core = _core(output)
    assert core.author == ""
    assert core.comments == ""
    assert core.last_modified_by == ""
    expected = "2026년 하반기 운영자료 제출 요청" if profile == "office-letter" else "예시 문서"
    assert core.title == expected
    assert core.created == core.modified
    assert before <= core.created.replace(tzinfo=None) <= after


def test_term_sheet_file_information_keeps_title_and_subject(tmp_path):
    output = _render_file(_term_sheet_copy(tmp_path, 'version: "v1"'), tmp_path / "out.docx")

    core = _core(output)
    assert core.title == "가나다머티리얼즈㈜ 운전자금대출 주요조건(안)"
    assert core.subject == "Term Sheet"
    assert core.author == "" and core.comments == ""


@pytest.mark.parametrize(
    ("profile", "source"),
    [
        ("plain", "본문만 있는 문서입니다.\n"),
        ("ib-report", "---\nprofile: ib-report\n---\n본문만 있는 문서입니다.\n"),
    ],
)
def test_placeholder_title_is_not_written_as_the_file_title(tmp_path, profile, source):
    """A profile's stand-in title ("Document", "IB Report") is not the document's title."""
    output = _render(source, profile, tmp_path / "out.docx")

    assert _core(output).title == ""
