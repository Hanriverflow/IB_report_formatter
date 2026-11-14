"""Numbered captions (E5): `SEQ` fields with cached numbers and hidden bookmarks.

Every test goes through the production parser and renderer. Tables are numbered
by a table spec `id` (`tbl-…`), figures by `{#fig-…}` after an image file line.
"""

import shutil
from pathlib import Path

import pytest
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn

from document_profiles import PROFILES, RenderOptions
from ib_renderer import IBDocumentRenderer
from input_safety import StrictValidationError
from md_parser import parse_markdown_file

SAMPLE_IMAGE = Path(__file__).resolve().parent.parent / "samples" / "qa" / "captions-structure.png"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Required document information per profile, beyond a title.
PROFILE_FRONTMATTER = {
    "office-letter": "recipients:\n  - 가상 수신처\nsender:\n  organization: 주식회사 예시\n",
    "term-sheet": "prepared_by: 예시증권\ndisclaimer: 시험용 가상 자료입니다.\n",
}


def write(tmp_path: Path, body: str, tables: str = "", profile: str = "business-report") -> Path:
    """Write a Markdown file with frontmatter and the sample image beside it."""
    shutil.copy(SAMPLE_IMAGE, tmp_path / "s.png")
    front = f"---\nprofile: {profile}\ntitle: 시험 문서\n{PROFILE_FRONTMATTER.get(profile, '')}"
    if tables:
        front += "tables:\n" + tables
    path = tmp_path / "doc.md"
    path.write_text(front + "---\n\n# 시험 문서\n\n" + body, encoding="utf-8")
    return path


def render(path: Path, strict: bool = True) -> tuple[DocxDocument, list[str]]:
    """Parse and render through the single composition path."""
    renderer = IBDocumentRenderer(options=RenderOptions(strict=strict))
    document = renderer.render(parse_markdown_file(str(path)))
    return document, renderer.errors


def fields(document: DocxDocument) -> list[tuple[str, str]]:
    """Each SEQ, REF and PAGEREF field of the body in order: (instruction, cached result)."""
    found: list[tuple[str, str]] = []
    instruction: list[str] = []
    result: list[str] = []
    state = None
    for node in document.element.body.iter(f"{W}fldChar", f"{W}instrText", f"{W}t"):
        if node.tag == f"{W}fldChar":
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                state, instruction, result = "code", [], []
            elif kind == "separate":
                state = "result"
            else:
                code = "".join(instruction).strip()
                if code.split(" ", 1)[0] in {"SEQ", "REF", "PAGEREF"}:
                    found.append((code, "".join(result)))
                state = None
        elif node.tag == f"{W}instrText" and state == "code":
            instruction.append(node.text or "")
        elif node.tag == f"{W}t" and state == "result":
            result.append(node.text or "")
    return found


def bookmarks(document: DocxDocument) -> dict[str, tuple[str, str]]:
    """Bookmark name -> (id, visible text between its start and end)."""
    names: dict[str, str] = {}
    texts: dict[str, list[str]] = {}
    open_ids: set[str] = set()
    for node in document.element.body.iter(
        f"{W}bookmarkStart", f"{W}bookmarkEnd", f"{W}t", f"{W}instrText"
    ):
        if node.tag == f"{W}bookmarkStart":
            names[node.get(qn("w:id"))] = node.get(qn("w:name"))
            texts[node.get(qn("w:id"))] = []
            open_ids.add(node.get(qn("w:id")))
        elif node.tag == f"{W}bookmarkEnd":
            open_ids.discard(node.get(qn("w:id")))
        elif node.tag == f"{W}t":
            for bookmark_id in open_ids:
                texts[bookmark_id].append(node.text or "")
    return {name: (bookmark_id, "".join(texts[bookmark_id])) for bookmark_id, name in names.items()}


def caption_texts(document: DocxDocument) -> list[str]:
    """Visible text of every body paragraph that holds a SEQ field."""
    return [
        paragraph.text
        for paragraph in document.paragraphs
        if any("SEQ" in (code.text or "") for code in paragraph._p.iter(f"{W}instrText"))
    ]


TABLES = """  - id: tbl-perf
    caption: 가상 실적 요약
    unit: 억원
  - caption: 번호 없는 표
  - id: tbl-debt
    caption: 가상 차입 내역
"""
BODY = """## 1. 개요

| 구분 | 2025 |
|---|---:|
| 매출 | 1,000 |

![가상 거래 구조도](s.png){#fig-structure}

| 가 | 나 |
|---|---|
| 1 | 2 |

![번호 없는 그림](s.png)

![가상 현금흐름도](s.png){#fig-flow}

| 차입처 | 금액 |
|---|---:|
| 가상은행 | 500 |
"""


def test_tables_and_figures_are_numbered_separately_in_document_order(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, BODY, TABLES))
    assert errors == []
    assert fields(document) == [
        ("SEQ 표 \\* ARABIC", "1"),
        ("SEQ 그림 \\* ARABIC", "1"),
        ("SEQ 그림 \\* ARABIC", "2"),
        ("SEQ 표 \\* ARABIC", "2"),
    ]
    assert caption_texts(document) == [
        "표 1. 가상 실적 요약",
        "그림 1. 가상 거래 구조도",
        "그림 2. 가상 현금흐름도",
        "표 2. 가상 차입 내역",
    ]
    texts = [paragraph.text for paragraph in document.paragraphs]
    assert "번호 없는 표" in texts and "번호 없는 그림" in texts


def test_bookmark_spans_label_and_number_with_stable_names_and_ids(tmp_path: Path) -> None:
    path = write(tmp_path, BODY, TABLES)
    first, _ = render(path)
    second, _ = render(path)
    marks = bookmarks(first)
    captions = {name: value for name, value in marks.items() if name.startswith("_Ref_")}
    assert {name: text for name, (_, text) in captions.items()} == {
        "_Ref_tbl_perf": "표 1",
        "_Ref_fig_structure": "그림 1",
        "_Ref_fig_flow": "그림 2",
        "_Ref_tbl_debt": "표 2",
    }
    assert bookmarks(second) == marks
    ids = [bookmark_id for bookmark_id, _ in marks.values()]
    assert len(ids) == len(set(ids))


def test_bookmark_is_a_sibling_of_the_caption_runs(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, BODY, TABLES))
    start = next(
        node
        for node in document.element.body.iter(f"{W}bookmarkStart")
        if node.get(qn("w:name")) == "_Ref_tbl_perf"
    )
    paragraph = start.getparent()
    assert paragraph.tag == f"{W}p"
    children = [child.tag.split("}")[1] for child in paragraph]
    begin = children.index("bookmarkStart")
    end = children.index("bookmarkEnd")
    # label run, begin, instruction, separate, result, end
    assert end - begin == 7


def test_caption_number_keeps_the_caption_formatting(tmp_path: Path) -> None:
    document, _ = render(write(tmp_path, BODY, TABLES))
    caption = next(p for p in document.paragraphs if p.text == "표 1. 가상 실적 요약")
    sizes = {run.font.size for run in caption.runs}
    assert len(sizes) == 1
    figure = next(p for p in document.paragraphs if p.text == "그림 1. 가상 거래 구조도")
    assert all(run.font.italic for run in figure.runs)
    plain = next(p for p in document.paragraphs if p.text == "번호 없는 그림")
    assert {run.font.size for run in figure.runs} == {run.font.size for run in plain.runs}
    assert {str(run.font.color.rgb) for run in figure.runs} == {
        str(run.font.color.rgb) for run in plain.runs
    }


def test_document_without_ids_has_no_caption_fields_or_bookmarks(tmp_path: Path) -> None:
    tables = "  - caption: 가상 실적 요약\n"
    body = "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n\n![가상 그림](s.png)\n"
    document, _ = render(write(tmp_path, body, tables))
    xml = document.element.xml
    assert "SEQ" not in xml
    assert "_Ref_" not in xml
    assert [p.text for p in document.paragraphs if "가상" in p.text] == [
        "가상 실적 요약",
        "가상 그림",
    ]


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_every_profile_numbers_tables_and_figures(tmp_path: Path, profile: str) -> None:
    tables = "  - id: tbl-a\n    caption: 가상 표\n    unit: 억원\n"
    body = "문단.\n\n| 가 | 나 |\n|---|---:|\n| 1 | 2 |\n\n![가상 그림](s.png){#fig-a}\n"
    document, errors = render(write(tmp_path, body, tables, profile=profile))
    assert errors == []
    assert fields(document) == [("SEQ 표 \\* ARABIC", "1"), ("SEQ 그림 \\* ARABIC", "1")]
    captions = caption_texts(document)
    if profile == "term-sheet":
        # Caption at left, unit at the right tab stop, on one line.
        assert captions == ["표 1. 가상 표\t(단위 : 억원)", "그림 1. 가상 그림"]
    else:
        assert captions == ["표 1. 가상 표", "그림 1. 가상 그림"]


def test_term_sheet_caption_number_is_bold_like_the_caption(tmp_path: Path) -> None:
    tables = "  - id: tbl-a\n    caption: 가상 표\n    unit: 억원\n"
    body = "| 가 | 나 |\n|---|---:|\n| 1 | 2 |\n"
    document, _ = render(write(tmp_path, body, tables, profile="term-sheet"))
    caption = next(p for p in document.paragraphs if p.text.startswith("표 1. "))
    caption_runs = [run for run in caption.runs if run.text and run.text != "\t"][:-1]
    assert caption_runs and all(run.bold for run in caption_runs)


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        ("  - id: tbl-Perf\n    caption: 표\n", "Table id 'tbl-Perf' is not a table id"),
        ("  - id: fig-a\n    caption: 표\n", "Table id 'fig-a' is not a table id"),
        ("  - id: tbl_a\n    caption: 표\n", "is not a table id"),
        ("  - id: tbl-\n    caption: 표\n", "is not a table id"),
        ("  - id: 3\n    caption: 표\n", "Table id must be text"),
        ("  - id: tbl-a\n", "needs a caption"),
        ("  - id: tbl-a\n    caption: '  '\n", "needs a caption"),
        ("  - id: tbl-" + "a" * 32 + "\n    caption: 표\n", "longer than 35"),
        ("  - id: tbl-a\n    caption: 가\n  - id: tbl-a\n    caption: 나\n", "more than one table"),
    ],
)
def test_bad_table_ids_reject_the_document_in_both_modes(
    tmp_path: Path, spec: str, message: str
) -> None:
    body = "| 가 |\n|---|\n| 1 |\n\n| 나 |\n|---|\n| 2 |\n"
    path = write(tmp_path, body, spec)
    with pytest.raises(StrictValidationError, match=message):
        parse_markdown_file(str(path))


@pytest.mark.parametrize(
    ("line", "message"),
    [
        ("![가상 그림](s.png){#Fig-A}", "'Fig-A' is not a figure id"),
        ("![가상 그림](s.png){#tbl-a}", "'tbl-a' is not a figure id"),
        ("![가상 그림](s.png){#}", "'' is not a figure id"),
        ("![](s.png){#fig-a}", "needs an image description"),
        ("![가상 그림](data:image/png;base64,iVBORw0KGgo=){#fig-a}", "on an embedded image"),
    ],
)
def test_unusable_figure_ids_are_rejected_in_strict_and_kept_as_text_otherwise(
    tmp_path: Path, line: str, message: str
) -> None:
    path = write(tmp_path, line + "\n")
    with pytest.raises(StrictValidationError, match=message):
        render(path)
    document, errors = render(path, strict=False)
    assert any(message in error for error in errors)
    assert "SEQ" not in document.element.xml
    assert any(line[line.index("{") :] in paragraph.text for paragraph in document.paragraphs)


def test_repeated_figure_id_is_rejected_and_keeps_the_count_otherwise(tmp_path: Path) -> None:
    body = "![가](s.png){#fig-a}\n\n![나](s.png){#fig-a}\n"
    path = write(tmp_path, body)
    with pytest.raises(StrictValidationError, match="'fig-a' is used by more than one"):
        render(path)
    document, errors = render(path, strict=False)
    assert any("used by more than one" in error for error in errors)
    assert fields(document) == [("SEQ 그림 \\* ARABIC", "1"), ("SEQ 그림 \\* ARABIC", "2")]
    assert [name for name in bookmarks(document) if name.startswith("_Ref_")] == ["_Ref_fig_a"]


def test_figure_line_with_id_starts_its_own_block(tmp_path: Path) -> None:
    body = "앞 문단 첫 줄\n![가상 그림](s.png){#fig-a}\n뒤 문단\n"
    document, _ = render(write(tmp_path, body))
    texts = [paragraph.text for paragraph in document.paragraphs if paragraph.text]
    assert "앞 문단 첫 줄" in texts and "그림 1. 가상 그림" in texts and "뒤 문단" in texts
