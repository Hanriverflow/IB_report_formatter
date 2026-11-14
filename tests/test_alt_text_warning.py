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
import os
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import quote

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


# ── Independent review (2026-11-14) ──────────────────────────────────────────


def _folder_free(warning: str, *folders: str) -> None:
    """A warning names the file only: no folder, drive, separator or encoding."""
    shown = warning.split(": ", 1)[1].split(", Word name")[0]
    for folder in folders:
        assert folder not in warning
    assert "/" not in shown and "\\" not in shown and ":" not in shown and ".." not in shown


@pytest.mark.parametrize("inline", [False, True], ids=["block", "inline"])
@pytest.mark.parametrize(
    ("written", "stored", "shown"),
    [
        ("가상폴더%2Fpic.png", ("가상폴더", "pic.png"), "pic.png"),
        ("가상폴더%2F..%2Fpic.png", ("pic.png",), "pic.png"),
        ("my%20pic.png", ("my pic.png",), "my pic.png"),
        ("pic%41.png", ("pic%41.png",), "pic%41.png"),
    ],
    ids=["encoded-slash", "encoded-dots", "encoded-space", "literal-percent"],
)
def test_the_warning_names_the_loaded_file_not_the_written_path(
    tmp_path: Path, written: str, stored: tuple[str, ...], shown: str, inline: bool
) -> None:
    # Review finding [1]: the name comes from the file actually loaded.
    (tmp_path / "가상폴더").mkdir()
    body = f"앞 ![]({written}) 뒤" if inline else f"![]({written})"
    source = write_source(tmp_path, body)
    tmp_path.joinpath(*stored).write_bytes(PNG)
    doc, renderer = render(source)
    (picture,) = pictures(doc)
    (warning,) = renderer.warnings
    assert f"{RENDER_WARNING}: '{shown}', Word name '{picture.get('name')}'" in warning
    _folder_free(warning, "가상폴더", str(tmp_path), tmp_path.name)


# An encoded absolute path loads only inside text: a block image line is joined to
# the source folder before decoding and is rejected as outside (existing behaviour).
@pytest.mark.parametrize(
    ("absolute", "inline"),
    [(True, True), (False, False), (False, True)],
    ids=["absolute-inline", "relative-block", "relative-inline"],
)
def test_an_encoded_backslash_path_shows_no_folder(
    tmp_path: Path, inline: bool, absolute: bool
) -> None:
    folder = tmp_path / "가상 딜 폴더"
    folder.mkdir()
    if os.name == "nt":
        (folder / "pic.png").write_bytes(PNG)
    else:  # POSIX has no backslash separator: the decoded name is one file name.
        (tmp_path / "가상 딜 폴더\\pic.png").write_bytes(PNG)
    written = str(folder) + "\\pic.png" if absolute else "가상 딜 폴더\\pic.png"
    destination = quote(written, safe="")
    body = f"앞 ![]({destination}) 뒤" if inline else f"![]({destination})"
    doc, renderer = render(write_source(tmp_path, body))
    assert len(pictures(doc)) == 1
    (warning,) = renderer.warnings
    assert f"{RENDER_WARNING}: 'pic.png', Word name" in warning
    _folder_free(warning, "가상 딜 폴더", "%5C", "%3A", str(tmp_path), tmp_path.name)


def _alternate(doc, branches: list[tuple[str, str | None]]) -> None:
    """Replace the only picture run's drawing with an mc:AlternateContent copy."""
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    (picture,) = pictures(doc)
    drawing = next(a for a in picture.iterancestors() if a.tag == qn("w:drawing"))
    run = drawing.getparent()
    alternate = etree.SubElement(run, f"{{{mc}}}AlternateContent")
    for kind, description in branches:
        branch = etree.SubElement(alternate, f"{{{mc}}}{kind}")
        if kind == "Choice":
            branch.set("Requires", "wps")
        copy = etree.fromstring(etree.tostring(drawing))
        if description is not None:
            next(copy.iter(qn("wp:docPr"))).set("descr", description)
        branch.append(copy)
    run.remove(drawing)


@pytest.mark.parametrize(
    ("branches", "warned"),
    [
        ([("Choice", None), ("Choice", None), ("Fallback", None)], 1),
        ([("Choice", "가상 그림"), ("Choice", None), ("Fallback", None)], 0),
        ([("Choice", None), ("Choice", "가상 그림"), ("Fallback", "가상 그림")], 1),
        ([("Fallback", None)], 1),
        ([("Fallback", "가상 그림")], 0),
    ],
    ids=["two-choices", "first-described", "first-undescribed", "fallback-only", "fallback-ok"],
)
def test_the_audit_reads_one_branch_of_alternate_content(
    tmp_path: Path, branches: list[tuple[str, str | None]], warned: int
) -> None:
    # Review finding [2]: only the first Choice (or the Fallback without one) is read.
    doc, _ = render(write_source(tmp_path, "![](pic.png)"))
    _alternate(doc, branches)
    audit = [w for w in inspect_document(doc).warnings if AUDIT_WARNING in w]
    assert audit == [f"{AUDIT_WARNING}: Word name 'Picture 1'"] * warned


def test_one_renderer_keeps_no_warnings_from_an_earlier_document(tmp_path: Path) -> None:
    first, second = tmp_path / "a", tmp_path / "b"
    first.mkdir()
    second.mkdir()
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True))
    renderer.render(parse_markdown_file(str(write_source(first, "![](pic.png)\n\n![](pic.png)"))))
    assert len(renderer.warnings) == 2
    renderer.render(parse_markdown_file(str(write_source(second, "![가상 그림](pic.png)"))))
    assert renderer.warnings == []


def test_concurrent_conversions_keep_their_own_warnings(tmp_path: Path) -> None:
    def job(number: int) -> tuple[int, list[str]]:
        folder = tmp_path / f"job{number}"
        folder.mkdir()
        (folder / f"job{number}.png").write_bytes(PNG)
        body = "\n\n".join([f"![](job{number}.png)"] * (number % 3 + 1))
        _, renderer = render(write_source(folder, body))
        return number, renderer.warnings

    with ThreadPoolExecutor(max_workers=8) as pool:
        results = list(pool.map(job, range(16)))
    for number, warnings in results:
        assert len(warnings) == number % 3 + 1
        assert all(f"'job{number}.png'" in warning for warning in warnings)


@pytest.mark.parametrize(
    ("body", "metadata"),
    [
        ("## 가상 제목 ![](pic.png)\n\n본문", {}),
        ("> **참고** 가상 내용 ![](pic.png)", {}),
        (
            "| 구분 | 값 |\n|---|---|\n| 가 | 1 |",
            {"tables": [{"caption": "가상 표 ![](pic.png)"}]},
        ),
    ],
    ids=["heading", "callout", "table-caption"],
)
def test_an_image_in_a_heading_callout_or_caption_warns_per_picture(
    tmp_path: Path, body: str, metadata: dict
) -> None:
    doc, renderer = render(write_source(tmp_path, body, **metadata))
    undescribed = [p.get("name") for p in pictures(doc) if not p.get("descr")]
    assert undescribed
    assert [w.split("Word name '")[1].split("'")[0] for w in renderer.warnings] == undescribed
    assert all(f"{RENDER_WARNING}: 'pic.png'" in w for w in renderer.warnings)
    assert renderer.errors == []


def test_an_image_after_a_failed_element_still_warns(tmp_path: Path) -> None:
    broken = "```chart\ntype: bar\nlabels: [A]\nseries: not-a-list\n```"
    doc, renderer = render(
        write_source(tmp_path, broken + "\n\n![](pic.png)"), strict=False, charts=True
    )
    assert renderer.errors and not any(RENDER_WARNING in e for e in renderer.errors)
    (picture,) = [p for p in pictures(doc) if not p.get("descr")]
    assert renderer.warnings == [
        f"{RENDER_WARNING}: 'pic.png', Word name '{picture.get('name')}'. "
        'Write one as ![description](...) or <img alt="...">.'
    ]
