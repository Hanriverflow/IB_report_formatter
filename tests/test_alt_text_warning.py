"""Warnings for pictures without a description (alt text); synthetic data only.

An author image (Markdown or HTML, file or embedded, block or inline) that is
inserted without a description adds one rendering warning. The warning never
rejects, strict or not, and writes nothing to the document. `docx-audit` reports
every picture in the main body without a description, whoever made it.
Every test parses Markdown with the real parser and renders through
`IBDocumentRenderer.render`.
"""

import base64
import logging
import re
from pathlib import Path

import pytest
import yaml
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from document_profiles import RenderOptions
from docx_audit import inspect_document, inspect_document_issues
from ib_render_media import ImageRenderer
from ib_renderer import IBDocumentRenderer
from md_parser import parse_markdown_file
from md_to_word import IBReportConverter
from scripts.agent_render import render_document
from scripts.sample_corpus import PROJECT_ROOT, sample_cases
from tests.snapshot_helpers import xml_parts

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Z8eQAAAAASUVORK5CYII="
)
B64 = base64.b64encode(PNG).decode()
RENDER_WARNING = "Image without a description (alt text)"
AUDIT_WARNING = "Picture without a description (alt text)"

# Author images whose description is empty: (case, body, name shown in the warning).
UNDESCRIBED = [
    ("block", "![](pic.png)", "'pic.png'"),
    ("block-blank", "![ ](pic.png)", "'pic.png'"),
    ("html-line-no-alt", '<img src="pic.png">', "'pic.png'"),
    ("html-line-empty-alt", '<img src="pic.png" alt="">', "'pic.png'"),
    ("block-base64", f"![](data:image/png;base64,{B64})", "embedded image"),
    ("paragraph", "앞 ![](pic.png) 뒤", "'pic.png'"),
    ("paragraph-data-uri", f"앞 ![](data:image/png;base64,{B64}) 뒤", "embedded image"),
    ("list-item", "- 항목 ![](pic.png)", "'pic.png'"),
    ("pipe-cell", "| 구분 |\n|---|\n| ![](pic.png) |", "'pic.png'"),
    ("html-cell", '<table><tr><td><img src="pic.png"></td></tr></table>', "'pic.png'"),
    (
        "html-cell-data-uri",
        f'<table><tr><td><img src="data:image/png;base64,{B64}"></td></tr></table>',
        "embedded image",
    ),
]

# The same places with a description: no warning anywhere.
DESCRIBED = [
    ("block", "![가상 그림](pic.png)"),
    ("html-line", '<img src="pic.png" alt="가상 그림">'),
    ("block-base64", f"![가상 그림](data:image/png;base64,{B64})"),
    ("paragraph", "앞 ![가상 그림](pic.png) 뒤"),
    ("list-item", "- 항목 ![가상 그림](pic.png)"),
    ("pipe-cell", "| 구분 |\n|---|\n| ![가상 그림](pic.png) |"),
    ("html-cell", '<table><tr><td><img src="pic.png" alt="가상 그림"></td></tr></table>'),
    ("numbered-figure", "![가상 그림](pic.png){#fig-a}\n\n@fig-a 참조"),
]

DIAGRAM = "```diagram:flow\ntitle: 가상 흐름\nboxes:\n  - {id: a, label: 가, pos: [0, 0]}\n```"
CHART = "```chart\ntype: bar\nlabels: [A, B]\nseries:\n  - {name: s, values: [1, 2]}\n```"

# Pictures the engine makes, whose description the author cannot write.
ENGINE_MADE = [
    ("latex-block", "$$x^2$$"),
    ("latex-inline", "식 $x^2$ 끝"),
    ("diagram", DIAGRAM),
]


def write_source(folder: Path, body: str, **metadata) -> Path:
    (folder / "pic.png").write_bytes(PNG)
    data = {"title": "가상 문서", "profile": "plain", **metadata}
    source = folder / "doc.md"
    source.write_text(
        "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n\n" + body + "\n",
        encoding="utf-8",
    )
    return source


def render(source: Path, *, strict: bool = True, charts: bool = False):
    model = parse_markdown_file(str(source))
    renderer = IBDocumentRenderer(options=RenderOptions(strict=strict, charts=charts))
    return renderer.render(model), renderer


def pictures(doc) -> list:
    return list(doc.element.body.iter(qn("wp:docPr")))


def alt_warnings(warnings: list[str]) -> list[str]:
    return [warning for warning in warnings if RENDER_WARNING in warning]


@pytest.mark.parametrize("strict", [True, False], ids=["strict", "default"])
@pytest.mark.parametrize(("case", "body", "shown"), UNDESCRIBED, ids=[c[0] for c in UNDESCRIBED])
def test_an_author_image_without_a_description_warns_once(
    tmp_path: Path, case: str, body: str, shown: str, strict: bool
) -> None:
    doc, renderer = render(write_source(tmp_path, body), strict=strict)
    (picture,) = pictures(doc)
    assert not (picture.get("descr") or "").strip()
    name = picture.get("name")
    assert renderer.warnings == [
        f"{RENDER_WARNING}: {shown}, Word name '{name}'. "
        'Write one as ![description](...) or <img alt="...">.'
    ]
    # Never a rejection or a structural issue.
    assert renderer.errors == []
    assert inspect_document_issues(doc) == []
    # The audit names the same picture.
    assert [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w] == [
        f"{AUDIT_WARNING}: Word name '{name}'"
    ]


@pytest.mark.parametrize(("case", "body"), DESCRIBED, ids=[c[0] for c in DESCRIBED])
def test_an_image_with_a_description_gives_no_warning(tmp_path: Path, case: str, body: str) -> None:
    doc, renderer = render(write_source(tmp_path, body))
    assert pictures(doc) and all(p.get("descr") for p in pictures(doc))
    assert renderer.warnings == []
    assert not [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]


def test_a_chart_is_described_by_its_title_or_label(tmp_path: Path) -> None:
    doc, renderer = render(write_source(tmp_path, CHART), charts=True)
    assert [p.get("descr") for p in pictures(doc)] == ["Chart 1"]
    assert renderer.warnings == []
    assert not [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]


@pytest.mark.parametrize(("case", "body"), ENGINE_MADE, ids=[c[0] for c in ENGINE_MADE])
def test_an_engine_made_picture_warns_only_in_the_audit(
    tmp_path: Path, case: str, body: str
) -> None:
    doc, renderer = render(write_source(tmp_path, body))
    (picture,) = pictures(doc)
    assert picture.get("descr") is None
    assert renderer.warnings == [] and renderer.errors == []
    assert [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w] == [
        f"{AUDIT_WARNING}: Word name '{picture.get('name')}'"
    ]


def test_a_term_sheet_logo_warns_only_in_the_audit(tmp_path: Path) -> None:
    source = write_source(
        tmp_path,
        "본문",
        profile="term-sheet",
        prepared_by="가상은행 자본시장부",
        disclaimer="가상 조건 검토용 문서입니다.",
        style={"logo": "pic.png"},
    )
    doc, renderer = render(source)
    (picture,) = pictures(doc)
    assert picture.get("descr") is None
    assert renderer.warnings == [] and renderer.errors == []
    assert len([w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]) == 1


def test_each_undescribed_picture_warns_once_and_the_warning_holds_no_path_or_data(
    tmp_path: Path,
) -> None:
    body = (
        "![](pic.png)\n\n![](pic.png)\n\n![가상 그림](pic.png)\n\n"
        f"![](data:image/png;base64,{B64})\n\n| 구분 |\n|---|\n| ![](pic.png) |\n"
    )
    doc, renderer = render(write_source(tmp_path, body))
    names = [p.get("name") for p in pictures(doc) if not p.get("descr")]
    assert len(names) == 4 and len(set(names)) == 4
    assert len(renderer.warnings) == 4
    assert [w.split("Word name '")[1].split("'")[0] for w in renderer.warnings] == names
    for warning in renderer.warnings:
        assert str(tmp_path) not in warning and str(tmp_path.as_posix()) not in warning
        assert B64 not in warning and "base64" not in warning
    audit = [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]
    assert audit == [f"{AUDIT_WARNING}: Word name '{name}'" for name in names]


def test_a_missing_image_file_is_an_error_not_an_alt_text_warning(tmp_path: Path) -> None:
    source = write_source(tmp_path, "![](missing.png)\n\n앞 ![](missing.png) 뒤")
    doc, renderer = render(source, strict=False)
    assert pictures(doc) == []
    assert renderer.warnings == [] and len(renderer.errors) == 2


def test_warnings_start_afresh_on_every_render(tmp_path: Path) -> None:
    model = parse_markdown_file(str(write_source(tmp_path, "![](pic.png)")))
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    renderer.render(model)
    renderer.render(model)
    assert len(renderer.warnings) == 1


def test_the_warning_changes_nothing_in_the_saved_document(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    body = "\n\n".join(body for _, body, _ in UNDESCRIBED) + "\n\n" + DIAGRAM
    source = write_source(tmp_path, body)
    warned = tmp_path / "warned.docx"
    silent = tmp_path / "silent.docx"
    IBReportConverter(str(source), str(warned), RenderOptions(strict=True)).convert()
    monkeypatch.setattr(ImageRenderer, "warn_missing_description", lambda *args: None)
    IBReportConverter(str(source), str(silent), RenderOptions(strict=True)).convert()

    def parts(document: Path) -> dict[str, str]:
        # Embedded and diagram images keep their random temporary file name (existing).
        return {
            name: re.sub(r'name="tmp[^"]*"', 'name="tmp"', text)
            for name, text in xml_parts(document).items()
        }

    assert parts(warned) == parts(silent)


def test_strict_cli_conversion_saves_and_logs_the_warning(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    source = write_source(tmp_path, "![](pic.png)")
    output = tmp_path / "out.docx"
    with caplog.at_level(logging.WARNING):
        saved = IBReportConverter(str(source), str(output), RenderOptions(strict=True)).convert()
    assert Path(saved).is_file()
    logged = [r for r in caplog.records if RENDER_WARNING in r.getMessage()]
    assert len(logged) == 1 and logged[0].levelno == logging.WARNING


def test_the_audit_skips_decorative_and_deleted_pictures(tmp_path: Path) -> None:
    source = write_source(tmp_path, "![](pic.png)\n\n![](pic.png)\n\n![](pic.png)")
    doc, _ = render(source)
    decorative, deleted, plain = pictures(doc)
    # Word's "Mark as decorative" (Office 2019+) writes this extension.
    extension = etree.fromstring(
        '<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
        '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative"'
        ' val="1"/></a:ext></a:extLst>'
    )
    decorative.append(extension)
    run = next(a for a in deleted.iterancestors() if a.tag == qn("w:r"))
    removal = etree.SubElement(run.getparent(), qn("w:del"))
    removal.set(qn("w:id"), "90")
    removal.set(qn("w:author"), "가상")
    removal.append(run)
    assert [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w] == [
        f"{AUDIT_WARNING}: Word name '{plain.get('name')}'"
    ]


def test_agent_render_lists_render_warnings_and_stays_ok(tmp_path: Path) -> None:
    folder = tmp_path / "src"
    folder.mkdir()
    table = "| 구분 | 금액 |\n|---|---|\n| 분할 | -**1200** |"
    source = write_source(
        folder,
        "![](pic.png)\n\n" + table,
        tables=[{"type": "financial", "columns": ["text", "money"], "negative": "paren"}],
    )
    result = render_document(str(source), str(tmp_path / "render-01"))
    assert result["ok"] and result["diagnostics"] == []
    warnings = result["render_warnings"]
    assert len(warnings) == 2
    assert RENDER_WARNING in warnings[0]
    assert "split across differently formatted text" in warnings[1]
    assert len([w for w in result["audit"]["warnings"] if AUDIT_WARNING in w]) == 1


def test_agent_render_has_no_render_warnings_for_a_described_image(tmp_path: Path) -> None:
    folder = tmp_path / "src"
    folder.mkdir()
    source = write_source(folder, "![가상 그림](pic.png)")
    result = render_document(str(source), str(tmp_path / "render-01"))
    assert result["ok"] and result["render_warnings"] == []


@pytest.mark.parametrize(
    "case",
    sample_cases(),
    ids=lambda case: case.output_relative.with_suffix("").as_posix().replace("/", "--"),
)
def test_repository_samples_give_no_alt_text_warning(case) -> None:
    model = parse_markdown_file(
        str(PROJECT_ROOT / case.source),
        profile=case.options.profile,
        asset_roots=case.options.asset_roots,
    )
    renderer = IBDocumentRenderer(options=case.options)
    doc = renderer.render(model)
    assert alt_warnings(renderer.warnings) == []
    assert not [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]


def test_an_anchored_picture_is_audited_too(tmp_path: Path) -> None:
    doc, _ = render(write_source(tmp_path, "![](pic.png)"))
    (picture,) = pictures(doc)
    inline = picture.getparent()
    inline.tag = qn("wp:anchor")
    saved = tmp_path / "anchored.docx"
    doc.save(str(saved))
    reopened = Document(str(saved))
    assert [w for w in inspect_document(reopened).warnings if AUDIT_WARNING in w] == [
        f"{AUDIT_WARNING}: Word name '{picture.get('name')}'"
    ]
