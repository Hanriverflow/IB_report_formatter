"""Numbered captions (E5): `SEQ` fields with cached numbers and hidden bookmarks.

Every test goes through the production parser and renderer. Tables are numbered
by a table spec `id` (`tbl-…`), figures by `{#fig-…}` after an image file line.
"""

import random
import re
import shutil
from collections.abc import Iterator
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


# ═══════════════════════════════════════════════════════════════════════════════
# CROSS-REFERENCES (@tbl-…, @fig-…)
# ═══════════════════════════════════════════════════════════════════════════════

REF_TABLES = (
    "  - id: tbl-perf\n    caption: 가상 실적 요약\n  - id: tbl-debt\n    caption: 가상 차입 내역\n"
)
REF_TABLE_BODY = "| 가 | 나 |\n|---|---:|\n| 1 | 2 |\n\n| 다 | 라 |\n|---|---:|\n| 3 | 4 |\n"
REF_INSTRUCTION = "REF _Ref_{} \\h \\* CHARFORMAT"


def references(document: DocxDocument) -> list[tuple[str, str]]:
    """The REF fields of the body: (instruction, cached result)."""
    return [field for field in fields(document) if field[0].startswith("REF ")]


def body_text(document: DocxDocument) -> list[str]:
    """Visible text of every non-empty body paragraph (tables excluded)."""
    return [paragraph.text for paragraph in document.paragraphs if paragraph.text]


def test_references_become_ref_fields_with_cached_numbers(tmp_path: Path) -> None:
    body = (
        "앞에서 @tbl-debt를 먼저 가리키고 @fig-flow 그림을 본다.\n\n"
        + REF_TABLE_BODY
        + "\n![가상 흐름도](s.png){#fig-flow}\n\n뒤에서 @tbl-perf에서 보듯, @tbl-perf.\n"
    )
    document, errors = render(write(tmp_path, body, REF_TABLES))
    assert errors == []
    assert references(document) == [
        (REF_INSTRUCTION.format("tbl_debt"), "표 2"),
        (REF_INSTRUCTION.format("fig_flow"), "그림 1"),
        (REF_INSTRUCTION.format("tbl_perf"), "표 1"),
        (REF_INSTRUCTION.format("tbl_perf"), "표 1"),
    ]
    texts = body_text(document)
    assert "앞에서 표 2를 먼저 가리키고 그림 1 그림을 본다." in texts
    assert "뒤에서 표 1에서 보듯, 표 1." in texts


def test_reference_keeps_the_body_formatting_not_the_caption_formatting(tmp_path: Path) -> None:
    body = "**굵게 @tbl-perf 안** *기울임 @tbl-debt*\n\n" + REF_TABLE_BODY
    document, _ = render(write(tmp_path, body, REF_TABLES))
    paragraph = next(p for p in document.paragraphs if p.text.startswith("굵게"))
    shown = {run.text: run for run in paragraph.runs if run.text in {"표 1", "표 2"}}
    assert shown["표 1"].bold and not shown["표 1"].italic
    assert shown["표 2"].italic and not shown["표 2"].bold
    body_size = next(run.font.size for run in paragraph.runs if run.text.startswith("굵게"))
    instruction_runs = [run for run in paragraph.runs if run._r.find(f"{W}instrText") is not None]
    assert {run.font.size for run in instruction_runs} == {body_size}
    assert all(run.bold for run in instruction_runs[:1])


@pytest.mark.parametrize(
    ("markdown", "shown"),
    [
        ("문의: kim@tbl-desk.example.com", "문의: kim@tbl-desk.example.com"),
        ("담당 @kim 과 @tbl 은 참조가 아니다", "담당 @kim 과 @tbl 은 참조가 아니다"),
        ("코드 `@tbl-perf` 그대로", "코드 @tbl-perf 그대로"),
        ("이스케이프 \\@tbl-perf 그대로", "이스케이프 @tbl-perf 그대로"),
        ("[자료](https://example.com/@tbl-perf) 링크", "자료 링크"),
        ("경로 C:/data/@fig-perf.png 그대로", "경로 C:/data/@fig-perf.png 그대로"),
        ("이중 @@tbl-perf 그대로", "이중 @@tbl-perf 그대로"),
        ("수식 $a@tbl-perf$ 뒤", None),
    ],
)
def test_at_signs_that_are_not_references_stay_literal(
    tmp_path: Path, markdown: str, shown: str | None
) -> None:
    document, errors = render(write(tmp_path, markdown + "\n\n" + REF_TABLE_BODY, REF_TABLES))
    assert errors == []
    assert references(document) == []
    if shown is not None:
        assert shown in body_text(document)


def test_reference_in_a_code_block_stays_literal(tmp_path: Path) -> None:
    body = "```text\n@tbl-perf\n```\n\n" + REF_TABLE_BODY
    document, errors = render(write(tmp_path, body, REF_TABLES))
    assert errors == []
    assert references(document) == []
    assert "@tbl-perf" in document.element.xml


@pytest.mark.parametrize(
    ("markdown", "message"),
    [
        ("없는 @tbl-nope 참조", "Cross-reference @tbl-nope has no table or figure with that id"),
        ("없는 @fig-perf 참조", "Cross-reference @fig-perf has no table or figure with that id"),
        ("모양 @tbl-Perf 오류", "Cross-reference @tbl-Perf: 'tbl-Perf' is not a table id"),
        ("모양 @tbl-perf_x 오류", "Cross-reference @tbl-perf_x: 'tbl-perf_x' is not a table id"),
        ("[링크 @tbl-perf 글자](https://example.com)", "@tbl-perf is inside link text"),
    ],
)
def test_unresolvable_references_reject_in_strict_and_stay_text_otherwise(
    tmp_path: Path, markdown: str, message: str
) -> None:
    path = write(tmp_path, markdown + "\n\n" + REF_TABLE_BODY, REF_TABLES)
    with pytest.raises(StrictValidationError, match=message.replace("(", r"\(")):
        render(path)
    document, errors = render(path, strict=False)
    assert message in "\n".join(errors)
    written = markdown.split(" ")[1] if not markdown.startswith("[") else "@tbl-perf"
    assert written in document.element.xml
    assert references(document) == []


@pytest.mark.parametrize(
    ("frontmatter", "body", "place"),
    [
        ("subtitle: 부제 @tbl-perf\n", "", "subtitle"),
        ("", "![설명 @tbl-perf](s.png)\n\n", "image description"),
        ("", "본문[^1]\n\n[^1]: 각주 @tbl-perf\n\n", "footnote"),
    ],
)
def test_references_where_no_field_can_go_are_reported(
    tmp_path: Path, frontmatter: str, body: str, place: str
) -> None:
    path = write(tmp_path, body + REF_TABLE_BODY, REF_TABLES)
    text = path.read_text(encoding="utf-8").replace(
        "title: 시험 문서\n", "title: 시험 문서\n" + frontmatter
    )
    path.write_text(text, encoding="utf-8")
    message = f"Cross-reference @tbl-perf in the {place} stays text"
    with pytest.raises(StrictValidationError, match=message):
        render(path)
    _, errors = render(path, strict=False)
    assert message in errors


def test_reference_in_the_title_is_reported(tmp_path: Path) -> None:
    path = write(tmp_path, REF_TABLE_BODY, REF_TABLES)
    text = path.read_text(encoding="utf-8").replace("시험 문서", "제목 @tbl-perf")
    path.write_text(text, encoding="utf-8")
    with pytest.raises(StrictValidationError, match="@tbl-perf in the document title stays text"):
        render(path)


def test_references_work_in_lists_quotes_headings_cells_and_captions(tmp_path: Path) -> None:
    tables = (
        "  - id: tbl-perf\n    caption: 가상 실적 요약\n"
        "  - id: tbl-debt\n    caption: 차입 (@tbl-perf 참조)\n    note: 주석 @tbl-perf\n"
    )
    body = (
        "### 제목 @tbl-perf 검토\n\n- 목록 @tbl-perf\n\n1. 번호 @tbl-debt\n\n> 인용 @tbl-perf\n\n"
        "| 가 | 나 |\n|---|---|\n| 칸 @tbl-debt | 1 |\n\n| 다 | 라 |\n|---|---:|\n| 3 | 4 |\n"
    )
    document, errors = render(write(tmp_path, body, tables))
    assert errors == []
    shown = [result for _, result in references(document)]
    # heading, bullet, quote, caption and note point to 표 1; numbered item and cell to 표 2
    assert shown.count("표 1") == 5 and shown.count("표 2") == 2
    assert "표 2. 차입 (표 1 참조)" in caption_texts(document)


def test_report_toc_preview_shows_heading_references_as_numbers(tmp_path: Path) -> None:
    body = "## 개요 @tbl-perf\n\n" + REF_TABLE_BODY
    document, _ = render(write(tmp_path, body, REF_TABLES, profile="ib-report"))
    assert "개요 표 1" in body_text(document)
    assert all("@tbl-perf" not in text for text in body_text(document))


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_every_profile_writes_references(tmp_path: Path, profile: str) -> None:
    tables = "  - id: tbl-a\n    caption: 가상 표\n    unit: 억원\n"
    body = "문단 @tbl-a와 @fig-a.\n\n| 가 | 나 |\n|---|---:|\n| 1 | 2 |\n\n![가상 그림](s.png){#fig-a}\n"
    document, errors = render(write(tmp_path, body, tables, profile=profile))
    assert errors == []
    assert references(document) == [
        (REF_INSTRUCTION.format("tbl_a"), "표 1"),
        (REF_INSTRUCTION.format("fig_a"), "그림 1"),
    ]


# ── docx_audit ────────────────────────────────────────────────────────────────


def test_fresh_output_has_no_reference_findings(tmp_path: Path) -> None:
    from docx_audit import inspect_document

    body = "앞 @tbl-debt, 뒤 @tbl-perf.\n\n" + REF_TABLE_BODY
    document, _ = render(write(tmp_path, body, REF_TABLES))
    audit = inspect_document(document)
    assert audit.issues == [] and audit.warnings == []


def test_audit_reports_a_reference_whose_caption_was_deleted(tmp_path: Path) -> None:
    from docx_audit import inspect_document, inspect_document_issues

    body = "앞 @tbl-debt.\n\n" + REF_TABLE_BODY
    document, _ = render(write(tmp_path, body, REF_TABLES))
    for node in list(document.element.body.iter(f"{W}bookmarkStart", f"{W}bookmarkEnd")):
        if node.get(qn("w:name")) == "_Ref_tbl_debt" or node.get(qn("w:id")) == "2":
            node.getparent().remove(node)
    expected = "Cross-reference field REF _Ref_tbl_debt has no bookmark named _Ref_tbl_debt"
    assert inspect_document_issues(document) == [expected]
    assert expected in inspect_document(document).issues


def test_audit_warns_about_stale_references_and_numbers(tmp_path: Path) -> None:
    from docx_audit import inspect_document

    body = "앞 @tbl-debt.\n\n" + REF_TABLE_BODY
    document, _ = render(write(tmp_path, body, REF_TABLES))
    # As after inserting a table in Word without updating fields: the second
    # caption now reads 3 while the reference still shows its old number.
    seq = [
        code for code in document.element.body.iter(f"{W}instrText") if "SEQ" in (code.text or "")
    ][-1]
    result_run = seq.getparent().getnext().getnext()  # instruction -> separate -> result
    result_run.find(f"{W}t").text = "3"
    warnings = inspect_document(document).warnings
    assert (
        "Cross-reference REF _Ref_tbl_debt shows '표 2' but its target reads '표 3'; "
        "update fields in Word (select all, F9)"
    ) in warnings
    assert (
        "Caption numbers SEQ 표 are not 1, 2, 3... in document order; "
        "update fields in Word (select all, F9)"
    ) in warnings


# ═══════════════════════════════════════════════════════════════════════════════
# INDEPENDENT REVIEW FIXES (2026-11-14)
# ═══════════════════════════════════════════════════════════════════════════════
#
# The reviewer's inputs are used as given. `NOT_REFERENCES` pins principle P1:
# text that is not a resolvable reference renders exactly as the same text
# with the `@tbl-` / `@fig-` prefix disguised (which no engine version reads as
# a reference), i.e. as before the feature. The same inputs were also compared
# with the engine at 84cde43 (docs/verification-captions-crossrefs-20261114.md).

_TEMP_IMAGE_NAME_RE = re.compile(r'name="tmp[^"]*\.png"')


def document_xml(document: DocxDocument) -> str:
    """The body XML with the random names of temporary equation images removed."""
    return _TEMP_IMAGE_NAME_RE.sub('name="tmp.png"', document.element.xml)


def disguised(markdown: str) -> str:
    """The same text with every reference prefix changed so nothing can match."""
    return markdown.replace("@tbl-", "@tbX-").replace("@fig-", "@fiX-")


NOT_REFERENCES = [
    "before @tbl-missing after",
    "user+@tbl-x.example",
    "x**@tbl-missing**y",
    "user@tbl-x.example a@fig-1.com doc-team@tbl-example.com",
    "> \\@tbl-missing",
    "> **note** \\@tbl-x and @tbl-y",
    "> $@tbl-a$ and **@tbl-missing** here",
    "@tbl-a- @tbl-a-- @tbl- @fig- @tbl-Results @tbl-a_b @tbl-a**b**",
    "<table><tr><td>@tbl-missing</td><td>x</td></tr></table>",
    "prefix ![Alt @tbl-missing](s.png) suffix",
    "[see @tbl-missing](https://x.example/@tbl-a) end",
    "`@tbl-a` and \\@tbl-a",
    "text[^1]\n\n[^1]: note @tbl-missing",
    "| a | b |\n|---|---|\n| @tbl-missing | `@tbl-a` |",
    "(@tbl-x.example) and @fig-1.5",
    "### 제목 @tbl-missing\n\n1. 개요 \\@tbl-x",
]


@pytest.mark.parametrize("markdown", NOT_REFERENCES)
def test_text_that_is_not_a_resolvable_reference_renders_as_before(
    tmp_path: Path, markdown: str
) -> None:
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    document, _ = render(write(tmp_path / "a", markdown), strict=False)
    control, _ = render(write(tmp_path / "b", disguised(markdown)), strict=False)
    expected = document_xml(control).replace("@tbX-", "@tbl-").replace("@fiX-", "@fig-")
    assert document_xml(document) == expected
    assert references(document) == []


def test_unresolved_reference_stays_in_one_run(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, "before @tbl-missing after"), strict=False)
    paragraph = next(p for p in document.paragraphs if "@tbl-missing" in p.text)
    assert [run.text for run in paragraph.runs] == ["before @tbl-missing after"]
    assert errors == ["Cross-reference @tbl-missing has no table or figure with that id"]


@pytest.mark.parametrize(
    "markdown",
    [
        "user+@tbl-x.example",
        "x**@tbl-missing**y",
        "user@tbl-x.example a@fig-1.com",
        "(@tbl-x.example) and @fig-1.5",
        "> \\@tbl-missing",
        "`@tbl-missing` and $@tbl-missing$",
    ],
)
def test_at_signs_outside_the_reference_grammar_are_neither_read_nor_reported(
    tmp_path: Path, markdown: str
) -> None:
    document, errors = render(write(tmp_path, markdown))
    assert errors == []
    assert references(document) == []


@pytest.mark.parametrize(
    "markdown",
    ["(@tbl-a) 참조", "「@tbl-a」 참조", "**@tbl-a**의 값", "끝 @tbl-a.", "줄 처음\n@tbl-a에서"],
)
def test_references_after_openers_and_before_particles_resolve(
    tmp_path: Path, markdown: str
) -> None:
    tables = "  - id: tbl-a\n    caption: 가상 표\n"
    document, errors = render(write(tmp_path, markdown + "\n\n| 가 |\n|---|\n| 1 |\n", tables))
    assert errors == []
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]


# [2] A table that is not drawn takes no number.
EMPTY_TABLE_SPECS = "  - id: tbl-empty\n    caption: Empty\n  - id: tbl-a\n    caption: A\n"
EMPTY_TABLE_BODY = "|\n\n| x |\n|---|\n| 1 |\n\nsee @tbl-a and @tbl-empty\n"


def test_table_that_is_not_drawn_takes_no_number(tmp_path: Path) -> None:
    path = write(tmp_path, EMPTY_TABLE_BODY, EMPTY_TABLE_SPECS)
    with pytest.raises(StrictValidationError, match="'tbl-empty' is on a table with no columns"):
        render(path)
    document, errors = render(path, strict=False)
    assert caption_texts(document) == ["표 1. A"]
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]
    assert "Cross-reference @tbl-empty has no table or figure with that id" in errors


def test_caption_numbers_not_written_as_worked_out_are_reported(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from ib_render_tables import TableRenderer
    from md_parser import Table

    original = TableRenderer.render
    calls = {"count": 0}

    def fail_first(self: TableRenderer, table: Table) -> None:
        calls["count"] += 1
        if calls["count"] == 1:
            raise RuntimeError("simulated failure")
        original(self, table)

    monkeypatch.setattr(TableRenderer, "render", fail_first)
    path = write(tmp_path, REF_TABLE_BODY, REF_TABLES)
    _, errors = render(path, strict=False)
    assert (
        "Caption numbers 표 were written as [2] but worked out as [1, 2]; "
        "update fields in Word (select all, F9)"
    ) in errors
    calls["count"] = 0
    with pytest.raises(StrictValidationError, match="Caption numbers 표 were written as"):
        render(path)


# [3] Every path protects code, escapes and math the same way.
DEFINED_A = "  - id: tbl-a\n    caption: A\n"
TABLE_ONLY = "\n\n| x |\n|---|\n| 1 |\n"


def test_escaped_reference_in_a_callout_is_text(tmp_path: Path) -> None:
    document, errors = render(
        write(tmp_path, "> \\@tbl-missing\n\n> \\@tbl-a" + TABLE_ONLY, DEFINED_A)
    )
    assert errors == []
    assert references(document) == []
    assert "@tbl-a" in document_xml(document)


@pytest.mark.parametrize("subtitle", ['"`@tbl-a`"', '"$@tbl-a$"', '"\\\\@tbl-a"'])
def test_code_math_and_escapes_in_the_subtitle_are_not_references(
    tmp_path: Path, subtitle: str
) -> None:
    path = write(tmp_path, "body" + TABLE_ONLY, DEFINED_A)
    text = path.read_text(encoding="utf-8").replace(
        "title: 시험 문서\n", f"title: 시험 문서\nsubtitle: {subtitle}\n"
    )
    path.write_text(text, encoding="utf-8")
    _, errors = render(path)
    assert errors == []


def test_math_in_a_callout_is_not_a_reference(tmp_path: Path) -> None:
    document, errors = render(write(tmp_path, "> $@tbl-a$" + TABLE_ONLY, DEFINED_A))
    assert errors == []
    assert references(document) == []
    assert "$@tbl-a$" in document_xml(document)


def test_toc_preview_reads_the_heading_like_its_renderer(tmp_path: Path) -> None:
    body = "## `@tbl-a` heading\n\n## real @tbl-a\n" + TABLE_ONLY
    document, errors = render(write(tmp_path, body, DEFINED_A, profile="ib-report"))
    assert errors == []
    texts = body_text(document)
    assert "`@tbl-a` heading" in texts  # the preview line, as before
    assert "real 표 1" in texts
    assert not any("`표 1`" in text for text in texts)
    assert len(references(document)) == 1  # only the second heading


# [4] A malformed ID is reported whole, never shortened to a valid prefix.
@pytest.mark.parametrize("token", ["@tbl-a-", "@tbl-a--", "@tbl-", "@fig-"])
def test_malformed_references_are_reported_whole(tmp_path: Path, token: str) -> None:
    path = write(tmp_path, f"see {token} here" + TABLE_ONLY, DEFINED_A)
    message = f"Cross-reference {token}: '{token[1:]}' is not a"
    with pytest.raises(StrictValidationError, match=re.escape(message)):
        render(path)
    document, errors = render(path, strict=False)
    assert any(error.startswith(message) for error in errors)
    assert references(document) == []
    assert f"see {token} here" in body_text(document)


def test_id_split_by_formatting_is_reported(tmp_path: Path) -> None:
    path = write(tmp_path, "see @tbl-a**b** here" + TABLE_ONLY, DEFINED_A)
    with pytest.raises(StrictValidationError, match="@tbl-ab is split by formatting"):
        render(path)


# [5] `{#…}` after an image that is not a numbered figure line is reported.
@pytest.mark.parametrize(
    ("markdown", "message"),
    [
        ("prefix\n![A](s.png){#Fig-A}", "Figure id 'Fig-A' is not a figure id"),
        ("![A](s.png){#Fig-A}", "Figure id 'Fig-A' is not a figure id"),
        ("- ![A](s.png){#fig-a}", "Figure id 'fig-a' works only on an image line"),
        ("| ![A](s.png){#fig-a} |\n|---|\n| 1 |", "Figure id 'fig-a' works only on an image line"),
        ("prefix\n![](s.png){#fig-a}", "Figure id 'fig-a' needs an image description"),
    ],
)
def test_figure_id_that_numbers_nothing_is_reported_wherever_it_stands(
    tmp_path: Path, markdown: str, message: str
) -> None:
    path = write(tmp_path, markdown)
    with pytest.raises(StrictValidationError, match=re.escape(message)):
        render(path)
    document, errors = render(path, strict=False)
    assert any(error.startswith(message) for error in errors)
    assert caption_texts(document) == []


# [6] HTML table cells read references like pipe table cells.
def test_html_table_cells_hold_references(tmp_path: Path) -> None:
    body = "<table><tr><td>@tbl-a</td><td>x</td></tr></table>" + TABLE_ONLY
    document, errors = render(write(tmp_path, body, DEFINED_A))
    assert errors == []
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]
    path = write(tmp_path, "<table><tr><td>@tbl-missing</td></tr></table>" + TABLE_ONLY, DEFINED_A)
    with pytest.raises(StrictValidationError, match="@tbl-missing has no table or figure"):
        render(path)


def test_html_code_cell_holds_no_reference(tmp_path: Path) -> None:
    body = "<table><tr><td><code>@tbl-a</code></td></tr></table>" + TABLE_ONLY
    document, errors = render(write(tmp_path, body, DEFINED_A))
    assert errors == []
    assert references(document) == []


# [7] A term-sheet title and subtitle keep references as text, as reported.
def test_term_sheet_title_and_subtitle_write_no_fields(tmp_path: Path) -> None:
    path = write(
        tmp_path, "| 구분 | 내용 |\n|---|---|\n| 1 | 2 |\n", DEFINED_A, profile="term-sheet"
    )
    text = path.read_text(encoding="utf-8").replace(
        "title: 시험 문서\n", 'title: "Title @tbl-a"\nsubtitle: "Sub @tbl-a"\n'
    )
    path.write_text(text, encoding="utf-8")
    document, errors = render(path, strict=False)
    assert references(document) == []
    assert "Cross-reference @tbl-a in the document title stays text" in errors
    assert "Cross-reference @tbl-a in the subtitle stays text" in errors
    assert "Title @tbl-a" in body_text(document) and "Sub @tbl-a" in body_text(document)
    with pytest.raises(StrictValidationError, match="in the document title stays text"):
        render(path)


# [8] An inline image's description is checked like a figure line's.
def test_inline_image_description_reference_is_reported(tmp_path: Path) -> None:
    path = write(tmp_path, "prefix ![Alt @tbl-missing](s.png) suffix")
    message = "Cross-reference @tbl-missing in the image description stays text"
    with pytest.raises(StrictValidationError, match=message):
        render(path)
    _, errors = render(path, strict=False)
    assert errors == [message]


def test_escaped_reference_in_an_inline_image_description_is_text(tmp_path: Path) -> None:
    _, errors = render(write(tmp_path, "prefix ![Alt \\@tbl-missing](s.png) suffix"))
    assert errors == []


# [9]–[11] Audit: nested and quoted fields, and SEQ forms the engine does not write.
def _field(instruction: str, result: str) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"<w:r><w:t>{result}</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def _audit_document(*paragraphs: str) -> DocxDocument:
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    document = Document()
    for inner in paragraphs:
        document.element.body.append(parse_xml(f"<w:p {nsdecls('w')}>{inner}</w:p>"))
    return document


TARGET = (
    '<w:bookmarkStart w:id="1" w:name="_Ref123456"/><w:r><w:t>Alpha</w:t></w:r>'
    '<w:bookmarkEnd w:id="1"/>'
)


def test_audit_reads_the_result_of_a_field_nested_in_an_instruction() -> None:
    from docx_audit import inspect_document

    nested = (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>'
        + _field("REF _Ref123456", "Alpha")
        + '<w:r><w:instrText xml:space="preserve"> = "Alpha" "YES" "NO" </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>YES</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
    audit = inspect_document(_audit_document(TARGET, nested))
    assert audit.issues == [] and audit.warnings == []


@pytest.mark.parametrize("kind", ["REF", "PAGEREF"])
def test_audit_reads_a_quoted_bookmark_name(kind: str) -> None:
    from docx_audit import inspect_document

    audit = inspect_document(_audit_document(TARGET, _field(f'{kind} "_Ref123456" \\h', "Alpha")))
    assert audit.issues == [] and audit.warnings == []
    missing = inspect_document(_audit_document(_field(f'{kind} "_RefMissing" \\h', "x")))
    assert missing.issues == [
        f"Cross-reference field {kind} _RefMissing has no bookmark named _RefMissing"
    ]


@pytest.mark.parametrize(
    "sequence",
    [
        [("SEQ Figure \\* ROMAN", "I"), ("SEQ Figure \\* ROMAN", "II")],
        [("SEQ Figure \\r 5", "5"), ("SEQ Figure", "6")],
        [("SEQ Figure \\c", "1"), ("SEQ Figure", "1")],
        [("SEQ Figure \\s 1", "1"), ("SEQ Figure", "1")],
        [("SEQ Figure \\h", ""), ("SEQ Figure", "2")],
    ],
)
def test_audit_does_not_judge_sequences_with_other_switches(
    sequence: list[tuple[str, str]],
) -> None:
    from docx_audit import inspect_document

    paragraphs = [_field(instruction, result) for instruction, result in sequence]
    assert inspect_document(_audit_document(*paragraphs)).warnings == []


def test_audit_still_judges_plain_and_arabic_sequences() -> None:
    from docx_audit import inspect_document

    paragraphs = [_field("SEQ 표 \\* ARABIC", "1"), _field("SEQ 표", "3")]
    assert inspect_document(_audit_document(*paragraphs)).warnings == [
        "Caption numbers SEQ 표 are not 1, 2, 3... in document order; "
        "update fields in Word (select all, F9)"
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# SECOND INDEPENDENT REVIEW FIXES (2026-11-14)
# ═══════════════════════════════════════════════════════════════════════════════
#
# A run is split only for a reference that resolves, so nothing is joined back.
# `UNRESOLVED` pins principle P1 structurally: in a document that numbers a
# table (the feature is in use), each input renders with the same paragraphs,
# runs, pictures and content controls as the same input with every `@` written
# as a full-width `＠` (which the grammar never reads), and its text differs
# only in that character.

FULLWIDTH_AT = "＠"
IMAGE = "![A](s.png)"
SPAN = '<span style="color:#123456">'
TERMS = "terms:\n  x: VALUE\nterm_tags: true\n"


def write_front(tmp_path: Path, body: str, front: str, tables: str = DEFINED_A) -> Path:
    """Write a document with extra frontmatter lines and a numbered table `tbl-a`."""
    path = write(tmp_path, body + TABLE_ONLY, tables)
    text = path.read_text(encoding="utf-8").replace(
        "title: 시험 문서\n", "title: 시험 문서\n" + front
    )
    path.write_text(text, encoding="utf-8")
    return path


def structure(document: DocxDocument) -> tuple[list[list[str]], int, int, int]:
    """Run texts of every body paragraph, and the number of pictures, controls and fields."""
    paragraphs = [
        [
            "".join(node.text or "" for node in run.iter(f"{W}t")).replace(FULLWIDTH_AT, "@")
            for run in paragraph.iter(f"{W}r")
        ]
        for paragraph in document.element.body.iter(f"{W}p")
    ]
    body = document.element.body
    return (
        paragraphs,
        len(list(body.iter(f"{W}drawing"))),
        len(list(body.iter(f"{W}sdt"))),
        len(list(body.iter(f"{W}fldChar"))),
    )


UNRESOLVED = [
    f"before @tbl-missing {IMAGE}{IMAGE}",
    f"- item @tbl-missing {IMAGE}{IMAGE}\n    - nested @fig-x {IMAGE}{IMAGE}",
    f"> quote @tbl-missing {IMAGE}{IMAGE}",
    f"## heading @tbl-missing {IMAGE}{IMAGE}",
    f"| a | b |\n|---|---|\n| cell @tbl-missing {IMAGE}{IMAGE} | y |",
    "before @tbl-missing {{x}}{{x}}",
    "- item @tbl-missing {{x}}{{x}}\n\n> quote @tbl-missing {{x}}{{x}}",
    "## heading @tbl-missing {{x}}{{x}}",
    f"@tbl-a{SPAN}b</span> and {SPAN}@tbl-a</span>b",
    f"x{SPAN}</span>@tbl-a and @tbl-a{SPAN}</span>x",
    "**bold @tbl-missing** *italic @fig-x* @tbl-a**b** @tbl-**a**",
    "`@tbl-a` code, [link @tbl-a](https://x.example/@tbl-a) text",
    "> callout $x@tbl-a$ math and **@tbl-missing** \\@tbl-a",
    "line one<br>@tbl-missing after a break",
    "<table><tr><td>@tbl-missing <b>@fig-x</b></td><td><code>@tbl-a</code></td></tr></table>",
    "1. one @tbl-missing\n2. two \\@tbl-a",
    "## @tbl-**a**\n\n## **@tbl-**a",
    "see @tbl-a**b**.com and (@fig-x.example)",
]


@pytest.mark.parametrize("markdown", UNRESOLVED)
def test_unresolved_references_keep_the_structure_of_text_without_them(
    tmp_path: Path, markdown: str
) -> None:
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    document, _ = render(write_front(tmp_path / "a", markdown, TERMS), strict=False)
    control_source = markdown.replace("\\@", FULLWIDTH_AT).replace("@", FULLWIDTH_AT)
    control, _ = render(write_front(tmp_path / "b", control_source, TERMS), strict=False)
    assert references(document) == []
    assert structure(document) == structure(control)


# [13] Unresolved references never merge the runs after them.
def test_unresolved_reference_keeps_every_picture_and_control(tmp_path: Path) -> None:
    body = f"before @tbl-missing {IMAGE}{IMAGE}\n\nagain @tbl-missing {{{{x}}}}{{{{x}}}}"
    document, errors = render(write_front(tmp_path, body, TERMS), strict=False)
    assert len(list(document.element.body.iter(f"{W}drawing"))) == 2
    controls = list(document.element.body.iter(f"{W}sdt"))
    assert ["".join(t.text or "" for t in c.iter(f"{W}t")) for c in controls] == ["VALUE", "VALUE"]
    assert errors == ["Cross-reference @tbl-missing has no table or figure with that id"]


# [14] Reading an ID across a colour span always ends.
_COLOUR_PROBE = """
import sys
from md_parser import ReferenceScope, TextParser
with TextParser.reference_scope(ReferenceScope(targets=frozenset({"tbl-a"}), report=print)):
    for text in sys.argv[1:]:
        print([run.text for run in TextParser.parse_runs(text)])
        print([run.text for run in TextParser.parse_runs_plain(text)])
"""


def test_reading_an_id_across_a_colour_span_ends() -> None:
    import subprocess
    import sys

    texts = [f"@tbl-a{SPAN}b</span>", f"{SPAN}@tbl-a</span>b", f"@tbl-a{SPAN}b</span>.com"]
    try:
        probe = subprocess.run(
            [sys.executable, "-c", _COLOUR_PROBE, *texts],
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=Path(__file__).resolve().parents[1],
            timeout=60,
            check=False,
        )
    except subprocess.TimeoutExpired:
        pytest.fail("parsing an ID across a colour span did not finish")
    assert probe.returncode == 0, probe.stderr[-300:]
    lines = probe.stdout.splitlines()
    assert lines.count("['@tbl-a', 'b']") == 4
    assert lines.count("Cross-reference @tbl-ab is split by formatting and stays text") == 4


# [15] An empty colour span is not the start or end of the text.
def test_empty_colour_span_is_no_reference_edge(tmp_path: Path) -> None:
    path = write(tmp_path, f"x{SPAN}</span>@tbl-a" + TABLE_ONLY, DEFINED_A)
    document, errors = render(path)
    assert errors == []
    assert references(document) == []
    path = write(tmp_path, f"@tbl-a{SPAN}</span>x" + TABLE_ONLY, DEFINED_A)
    with pytest.raises(StrictValidationError, match="@tbl-ax is split by formatting"):
        render(path)
    document, _ = render(path, strict=False)
    assert references(document) == []


# [16] A heading is judged on its source, with the emphasis it renders without.
@pytest.mark.parametrize("profile", ["business-report", "ib-report", "term-sheet"])
@pytest.mark.parametrize("heading", ["## @tbl-**a**", "## **@tbl-**a"])
def test_heading_reference_split_by_emphasis_is_reported(
    tmp_path: Path, heading: str, profile: str
) -> None:
    body = heading + "\n\n| 구분 | 내용 |\n|---|---|\n| 1 | 2 |\n"
    path = write(tmp_path, body, DEFINED_A, profile=profile)
    with pytest.raises(StrictValidationError, match="@tbl-a is split by formatting"):
        render(path)
    document, _ = render(path, strict=False)
    assert references(document) == []


def test_heading_with_terms_is_judged_on_its_source(tmp_path: Path) -> None:
    path = write_front(tmp_path, "## {{x}} @tbl-**a**\n\n## {{x}} **@tbl-a**", TERMS)
    with pytest.raises(StrictValidationError, match="@tbl-a is split by formatting"):
        render(path)
    document, _ = render(path, strict=False)
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]


# [17] The domain exclusion reads the whole visible ID first.
@pytest.mark.parametrize("markdown", ["see @tbl-a**b**.com", f"see @tbl-a{SPAN}b</span>.org"])
def test_domain_split_by_formatting_is_not_a_reference(tmp_path: Path, markdown: str) -> None:
    document, errors = render(write(tmp_path, markdown + TABLE_ONLY, DEFINED_A))
    assert errors == []
    assert references(document) == []


# [18] Term values are never read as references.
@pytest.mark.parametrize("key", ["title", "subtitle", "date"])
@pytest.mark.parametrize("profile", ["business-report", "ib-report", "term-sheet"])
def test_term_values_in_metadata_are_not_references(tmp_path: Path, key: str, profile: str) -> None:
    path = write(tmp_path, "body\n\n| 구분 | 내용 |\n|---|---|\n| 1 | 2 |\n", profile=profile)
    front = f'terms:\n  x: "@tbl-missing"\n{key}: "{{{{x}}}}"\n'
    if key != "title":
        front = "title: 시험 문서\n" + front
    text = path.read_text(encoding="utf-8").replace("title: 시험 문서\n", front)
    path.write_text(text, encoding="utf-8")
    document, errors = render(path)
    assert errors == []
    assert references(document) == []


# [19] Every frontmatter value a document shows is checked; none holds a field.
METADATA_KEYS = [
    "subtitle",
    "company",
    "ticker",
    "sector",
    "analyst",
    "date",
    "recipient",
    "report_type",
    "subject_company",
    "analysis_period",
    "analysis_basis",
    "document_no",
    "meeting_time",
    "location",
    "version",
]
SENDER_KEYS = ["organization", "department", "signatory", "contact"]
LIST_KEYS = ["recipients", "cc", "attachments", "attendees"]


def _metadata_front(profile: str) -> tuple[str, dict[str, str]]:
    """Frontmatter with a distinct `@fig-k…` in every value; key -> its ID."""
    ids: dict[str, str] = {}
    lines = [f"profile: {profile}", 'title: "Title @fig-ktitle"']
    for index, key in enumerate(METADATA_KEYS):
        ids[key] = f"fig-k{index}"
        lines.append(f'{key}: "v @fig-k{index}"')
    lines.append("sender:")
    for index, key in enumerate(SENDER_KEYS):
        ids["sender." + key] = f"fig-s{index}"
        lines.append(f'  {key}: "v @fig-s{index}"')
    for index, key in enumerate(LIST_KEYS):
        ids[key] = f"fig-l{index}"
        lines.append(f'{key}:\n  - "v @fig-l{index}"')
    lines += ['prepared_by: "House @fig-kp"', 'disclaimer: "Test @fig-kd only."']
    ids["prepared_by"], ids["disclaimer"] = "fig-kp", "fig-kd"
    return "---\n" + "\n".join(lines) + "\n---\n\n# Title @fig-ktitle\n\nbody\n", ids


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_metadata_references_are_reported_exactly_where_they_are_shown(
    tmp_path: Path, profile: str
) -> None:
    import io
    import zipfile

    source, ids = _metadata_front(profile)
    path = tmp_path / "doc.md"
    path.write_text(source, encoding="utf-8")
    document, errors = render(path, strict=False)
    assert references(document) == []
    buffer = io.BytesIO()
    document.save(buffer)
    with zipfile.ZipFile(buffer) as archive:
        shown = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        )
    reported = {
        key for key, ref_id in ids.items() if any(f"@{ref_id} in the " in error for error in errors)
    }
    printed = {
        key
        for key, ref_id in ids.items()
        if re.search(re.escape("@" + ref_id) + "(?![A-Za-z0-9_-])", shown)
    }
    # The subtitle is checked in every profile, as before; the rest exactly where shown.
    assert reported - {"subtitle"} == printed - {"subtitle"}
    assert "subtitle" in reported
    assert any("@fig-ktitle in the document title stays text" in error for error in errors)
    with pytest.raises(StrictValidationError, match="stays text"):
        render(path)


def test_metadata_reference_in_business_report_rejects(tmp_path: Path) -> None:
    front = 'date: "Date @tbl-missing"\nsender:\n  organization: Org\n  department: "Team @fig-missing"\n'
    path = write_front(tmp_path, "body", front)
    with pytest.raises(StrictValidationError, match="@tbl-missing in the frontmatter date"):
        render(path)
    _, errors = render(path, strict=False)
    assert "Cross-reference @fig-missing in the frontmatter sender.department stays text" in errors


def test_term_sheet_metadata_never_holds_a_field(tmp_path: Path) -> None:
    front = 'date: "2026 @tbl-a"\nprepared_by: "House @tbl-a"\n'
    path = write_front(tmp_path, "body @tbl-a", front)
    text = path.read_text(encoding="utf-8").replace("business-report", "term-sheet")
    text = text.replace("title: 시험 문서\n", "title: 시험 문서\ndisclaimer: 시험용.\n", 1)
    path.write_text(text, encoding="utf-8")
    document, errors = render(path, strict=False)
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]  # body only
    assert "Cross-reference @tbl-a in the frontmatter date stays text" in errors
    assert "Cross-reference @tbl-a in the term-sheet prepared_by stays text" in errors
    assert "2026 @tbl-a" in body_text(document) and "House @tbl-a" in body_text(document)


# [20] A table that is not drawn: rejected in strict mode, reported and unnumbered otherwise.
def test_table_that_is_not_drawn_is_reported_in_the_default_mode(tmp_path: Path) -> None:
    path = write(tmp_path, EMPTY_TABLE_BODY.replace(" and @tbl-empty", ""), EMPTY_TABLE_SPECS)
    with pytest.raises(StrictValidationError, match="'tbl-empty' is on a table with no columns"):
        render(path)
    document, errors = render(path, strict=False)
    assert "Reference id 'tbl-empty' is on a table with no columns, which is not drawn" in errors
    assert caption_texts(document) == ["표 1. A"]


# [21] A simple field's result skips the instruction of a field nested in it.
def test_audit_reads_a_simple_field_with_a_nested_field() -> None:
    from docx_audit import inspect_document

    alpha = (
        '<w:bookmarkStart w:id="2" w:name="_Alpha"/><w:r><w:t>Alpha</w:t></w:r>'
        '<w:bookmarkEnd w:id="2"/>'
    )
    yes = '<w:bookmarkStart w:id="3" w:name="_Yes"/><w:r><w:t>YES</w:t></w:r><w:bookmarkEnd w:id="3"/>'
    nested = (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>'
        + _field("REF _Alpha", "Alpha")
        + '<w:r><w:instrText xml:space="preserve"> = "Alpha" "YES" "NO" </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>YES</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
    simple = f'<w:fldSimple w:instr=" REF _Yes ">{nested}</w:fldSimple>'
    audit = inspect_document(_audit_document(alpha, yes, simple))
    assert audit.issues == [] and audit.warnings == []
    stale = '<w:fldSimple w:instr=" REF _Yes "><w:r><w:t>NO</w:t></w:r></w:fldSimple>'
    assert inspect_document(_audit_document(yes, stale)).warnings == [
        "Cross-reference REF _Yes shows 'NO' but its target reads 'YES'; "
        "update fields in Word (select all, F9)"
    ]


def test_hand_built_model_metadata_is_checked_when_rendered() -> None:
    from md_parser import DocumentMetadata, DocumentModel, Element, ElementType, Paragraph, TextRun

    metadata = DocumentMetadata(title="Title @tbl-a", profile="business-report")
    metadata.extra["date"] = "Date @fig-b"
    body = Element(ElementType.PARAGRAPH, Paragraph(text="body", runs=[TextRun(text="body")]))
    renderer = IBDocumentRenderer(options=RenderOptions(strict=False))
    renderer.render(DocumentModel(metadata=metadata, elements=[body]))
    assert "Cross-reference @tbl-a in the document title stays text" in renderer.errors
    assert "Cross-reference @fig-b in the frontmatter date stays text" in renderer.errors


def test_term_values_in_table_texts_are_not_references(tmp_path: Path) -> None:
    tables = '  - id: tbl-a\n    caption: "가상 {{x}}"\n    note: "주석 {{x}}"\n'
    path = write_front(tmp_path, "body", 'terms:\n  x: "@tbl-a"\n', tables)
    document, errors = render(path)
    assert errors == []
    assert references(document) == []
    assert "표 1. 가상 @tbl-a" in caption_texts(document)
    assert "주석 @tbl-a" in body_text(document)


# ═══════════════════════════════════════════════════════════════════════════════
# THIRD REVIEW (2026-11-14): RANDOMIZED P1 CHECK
# ═══════════════════════════════════════════════════════════════════════════════
#
# Documents with no numbered table or figure are built from random inline
# fragments with a fixed seed. Each renders three ways: as it is; with every
# reference scope switched off (no text is ever read as a reference, as before
# the feature); and with each private-use probe character replaced by an
# ordinary symbol that no rule treats specially, mapped back afterwards. All
# three must give the same XML, so neither reading references nor any internal
# marker may change a character, a run, an image or a content control.

# Private-use characters, including those the engine uses internally, and the
# ordinary symbols (same East Asian width, not word characters) that stand in.
_PROBES = {
    "\ue000": "§",
    "\ue001": "¶",
    "\ue010": "※",
    "\ue011": "†",
    "\ue012": "‡",
    "\U000f0000": "‖",
}
_FRAGMENTS = [
    " ",
    " ",
    " ",
    "본문",
    "text",
    "a",
    "x",
    "1",
    "@",
    "@",
    "@",
    "tbl-",
    "fig-",
    "tbl",
    "-",
    "-",
    "a1",
    ".",
    ".com",
    "x.example",
    "(",
    ")",
    "「",
    "」",
    "**",
    "**",
    "*",
    "_",
    "~",
    "^",
    "`",
    "`",
    "``",
    "\\@",
    "\\@",
    "\\$",
    "\\*",
    "\\\\",
    "\\-",
    "\\`",
    "[",
    "]",
    "](https://x.example/@tbl-a)",
    "](s.png)",
    "![",
    "![alt @tbl-q](s.png)",
    '<span style="color:#C00000">',
    "</span>",
    "<br>",
    "&#64;",
    "&amp;",
    "{{x}}",
    "{{y}}",
    "@tbl-a",
    "@fig-b",
    "@tbl-missing",
    "user@tbl-x.example",
    "[^1]",
]
_DOLLARS = ["$", "$", "$x^2$", "$@tbl-a$", "\\$"]
_P1_TERMS = 'terms:\n  x: "@tbl-missing"\n  y: "값 \\ue010 \\U000f0000 @fig-b"\n'


def _random_inline(rng: random.Random, probes: bool) -> str:
    """A run of random fragments; with probes or with dollar signs, never both.

    Math is drawn as an image whose size depends on its glyphs, so a probe
    must not end up inside it.
    """
    pool = _FRAGMENTS + (list(_PROBES) * 3 if probes else _DOLLARS * 2)
    return "".join(rng.choice(pool) for _ in range(rng.randint(3, 12)))


def _random_document(seed: int, blocks: int) -> tuple[str, str]:
    """Frontmatter and body of a random document, and its profile."""
    rng = random.Random(seed)
    profile = "ib-report" if seed % 2 else "business-report"
    lines: list[str] = []
    for _ in range(blocks):
        probes = rng.random() < 0.6
        text = _random_inline(rng, probes)
        kind = rng.choice(
            ["paragraph", "paragraph", "heading", "list", "numbered", "quote", "table", "html"]
            + (["numbered-heading"] if profile == "ib-report" else [])
        )
        if kind == "heading":
            lines.append("## " + text)
        elif kind == "numbered-heading":
            lines.append("**1. 개요 " + text + "**")
        elif kind == "list":
            lines.append("- " + text + "\n- " + _random_inline(rng, probes))
        elif kind == "numbered":
            lines.append("1. " + text)
        elif kind == "quote":
            lines.append("> " + text + "\n> " + _random_inline(rng, probes))
        elif kind == "table":
            cells = [text.replace("|", "/")] + [
                _random_inline(rng, probes).replace("|", "/") for _ in range(3)
            ]
            lines.append(f"| {cells[0]} | {cells[1]} |\n|---|---|\n| {cells[2]} | {cells[3]} |")
        elif kind == "html":
            lines.append(
                f"<table><tr><td>{text}</td><td>{_random_inline(rng, probes)}</td></tr></table>"
            )
        else:
            lines.append(text)
    lines.append("[^1]: 각주 @tbl-missing \ue010 \\@tbl-a")
    front = f"---\nprofile: {profile}\ntitle: 시험 문서\n{_P1_TERMS}---\n\n"
    return front + "\n\n".join(lines) + "\n", profile


def _rendered_xml(path: Path) -> str:
    """Body and footnote XML of a non-strict render, temporary image names removed."""
    renderer = IBDocumentRenderer(options=RenderOptions(strict=False))
    document = renderer.render(parse_markdown_file(str(path)))
    parts = [document_xml(document)]
    for part in document.part.package.iter_parts():
        if str(part.partname) == "/word/footnotes.xml":
            parts.append(_TEMP_IMAGE_NAME_RE.sub('name="tmp.png"', part.blob.decode("utf-8")))
    return "\n".join(parts)


def _first_difference(left: str, right: str) -> str:
    """Where two XML texts first differ, for the assertion message."""
    index = next(
        (i for i, (a, b) in enumerate(zip(left, right, strict=False)) if a != b),
        min(len(left), len(right)),
    )
    return f"at {index}: {left[index - 80 : index + 80]!r} != {right[index - 80 : index + 80]!r}"


@pytest.mark.parametrize("seed", range(6))
def test_random_text_without_targets_renders_as_with_no_reference_scope(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, seed: int
) -> None:
    from contextlib import contextmanager

    from md_parse_text import TextParser

    markdown, _ = _random_document(seed, blocks=50)
    assert "{#" not in markdown and "id:" not in markdown  # nothing is numbered
    probed = markdown
    for probe, stand_in in _PROBES.items():
        probed = probed.replace(probe, stand_in)
    for name, text in (("on", markdown), ("off", markdown), ("probe", probed)):
        (tmp_path / name).mkdir()
        shutil.copy(SAMPLE_IMAGE, tmp_path / name / "s.png")
        (tmp_path / name / "doc.md").write_text(text, encoding="utf-8")

    scope_on = _rendered_xml(tmp_path / "on" / "doc.md")
    stand_ins = _rendered_xml(tmp_path / "probe" / "doc.md")
    for probe, stand_in in _PROBES.items():
        stand_ins = stand_ins.replace(stand_in, probe)
    original = TextParser.reference_scope

    @contextmanager
    def no_scope(scope: object) -> Iterator[None]:
        with original(None):
            yield None

    monkeypatch.setattr(TextParser, "reference_scope", staticmethod(no_scope))
    scope_off = _rendered_xml(tmp_path / "off" / "doc.md")
    assert scope_on == scope_off, _first_difference(scope_on, scope_off)
    assert scope_on == stand_ins, _first_difference(scope_on, stand_ins)


# [22] Private-use characters, including those the engine once used as marks,
# are text like any other: no mark stands for a character of the input.
def all_text(document: DocxDocument) -> str:
    """Every `w:t` of the body, tables included."""
    return "".join(node.text or "" for node in document.element.body.iter(f"{W}t"))


@pytest.mark.parametrize(
    "markdown",
    [
        "a\ue010b\ue011c",
        "`\ue010` and `\ue011`",
        "## 제목 a\ue010b\ue011c @tbl-a",
        "- 목록 a\ue010b @tbl-a\n- `\ue011`",
        "> 인용 \\@tbl-a \ue010 @tbl-a \ue011",
        "| a\ue010 @tbl-a | `\ue011` |\n|---|---|\n| \\@tbl-a\ue010 | 2 |",
        "<table><tr><td>\ue010 @tbl-a \ue011</td></tr></table>",
        "본문 \\@tbl-a \\$ \ue010\ue011 @tbl-a \U000f0000",
    ],
)
def test_private_use_characters_are_kept_with_and_without_targets(
    tmp_path: Path, markdown: str
) -> None:
    for name, tables in (("none", ""), ("defined", DEFINED_A)):
        (tmp_path / name).mkdir()
        document, _ = render(write(tmp_path / name, markdown + TABLE_ONLY, tables), strict=False)
        text = all_text(document)
        for character in ("\ue010", "\ue011", "\U000f0000"):
            assert text.count(character) == markdown.count(character), (name, character)
        if tables:
            # Each unescaped `@tbl-a` resolves; the escaped ones stay text.
            expected = markdown.count("@tbl-a") - markdown.count("\\@tbl-a")
            assert len(references(document)) == expected


def test_private_use_characters_in_a_numbered_heading_are_kept(tmp_path: Path) -> None:
    body = "**1. 개요 \ue010 \\@tbl-a @tbl-a \ue011**\n\n본문"
    document, errors = render(write(tmp_path, body + TABLE_ONLY, DEFINED_A, "ib-report"))
    assert errors == []
    heading = next(p.text for p in document.paragraphs if "개요" in p.text)
    assert heading == "1. 개요 \ue010 @tbl-a 표 1 \ue011"
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]


def test_escaped_at_offsets_ignore_every_other_character() -> None:
    from md_parse_text import TextParser

    text = "\ue010 0 1 \\@a @b \\\\@c `\\@d` \\@"
    cleaned = TextParser.cleanup_text(text)
    offsets = TextParser.escaped_ats(text)
    assert [cleaned[index] for index in offsets] == ["@", "@"]
    assert cleaned[offsets[0] + 1] == "a" and offsets[1] == len(cleaned) - 1


def test_text_holding_every_mark_is_not_read_and_says_so(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from md_parse_text import TextParser

    monkeypatch.setattr(TextParser, "_MARK_POOL", (range(0xF0000, 0xF0002),))
    body = "본문 \U000f0000\U000f0001 @tbl-a \\@tbl-a"
    path = write(tmp_path, body + TABLE_ONLY, DEFINED_A)
    with pytest.raises(StrictValidationError, match="Cross-references are not read in text"):
        render(path)
    document, errors = render(path, strict=False)
    assert references(document) == []
    assert "본문 \U000f0000\U000f0001 @tbl-a @tbl-a" in body_text(document)
    assert any(error.startswith("Cross-references are not read") for error in errors)


# [23] Whether references are read never depends on the raw Markdown: the
# recognizer reads the text as parsed and decoded.
@pytest.mark.parametrize(
    ("body", "tables", "message"),
    [
        ("see @**tbl-missing**", "", "@tbl-missing is split by formatting"),
        ("see @**tbl-a**", DEFINED_A, "@tbl-a is split by formatting"),
        ("see @tbl**-a**", DEFINED_A, "@tbl-a is split by formatting"),
        (
            "<table><tr><td>&#64;tbl-missing</td></tr></table>",
            "",
            "@tbl-missing has no table or figure with that id",
        ),
    ],
)
def test_references_the_raw_text_does_not_spell_are_read(
    tmp_path: Path, body: str, tables: str, message: str
) -> None:
    path = write(tmp_path, body + (TABLE_ONLY if tables else ""), tables)
    with pytest.raises(StrictValidationError, match=re.escape(message)):
        render(path)
    document, errors = render(path, strict=False)
    assert references(document) == []
    assert "Cross-reference " + message in " ".join(errors)


def test_a_reference_decoded_from_a_yaml_escape_is_read(tmp_path: Path) -> None:
    path = tmp_path / "doc.md"
    path.write_text(
        '---\nprofile: business-report\ntitle: "Title \\u0040tbl-missing"\n---\n\nbody\n',
        encoding="utf-8",
    )
    with pytest.raises(StrictValidationError, match="in the document title stays text"):
        render(path)


# [26] Each diagnostic is reported once, in the order first met.
def test_many_distinct_diagnostics_are_reported_once_in_order(tmp_path: Path) -> None:
    body = "\n\n".join(f"@tbl-x{index} and @tbl-x{index}" for index in range(300))
    model = parse_markdown_file(str(write(tmp_path, body)))
    expected = [
        f"Cross-reference @tbl-x{index} has no table or figure with that id" for index in range(300)
    ]
    assert [w for w in model.warnings if w.startswith("Cross-reference")] == expected
    renderer = IBDocumentRenderer(options=RenderOptions(strict=False))
    renderer.render(model)
    assert [e for e in renderer.errors if e.startswith("Cross-reference")] == expected


# [25] A term value is never read as a reference, wherever it is inserted,
# including a heading that becomes the document title.
@pytest.mark.parametrize(
    ("profile", "body", "shown"),
    [
        ("business-report", "# {{x}}\n\nbody", True),
        ("business-report", "# 제목 @{{x}} \\@tbl-a\n\nbody", True),
        ("business-report", "## {{x}}\n\n{{x}} and (@{{x}})\n\n- {{x}}\n\n> {{x}}", True),
        (
            "business-report",
            "| {{x}} |\n|---|\n| 1 |\n\n<table><tr><td>{{x}}</td></tr></table>",
            True,
        ),
        ("business-report", "![{{x}}](s.png)\n\ntext[^1]\n\n[^1]: {{x}}", False),
        ("ib-report", "**1. 개요 {{x}}**\n\n본문", True),
        ("ib-report", "**1. 개요 \\@tbl-a {{x}}**\n\n본문", True),
    ],
)
def test_term_values_are_never_read_as_references(
    tmp_path: Path, profile: str, body: str, shown: bool
) -> None:
    shutil.copy(SAMPLE_IMAGE, tmp_path / "s.png")
    path = tmp_path / "doc.md"
    front = f"---\nprofile: {profile}\n{PROFILE_FRONTMATTER.get(profile, '')}"
    path.write_text(front + 'terms:\n  x: "@tbl-missing"\n---\n\n' + body, encoding="utf-8")
    document, errors = render(path)
    assert errors == []
    assert references(document) == []
    assert ("@tbl-missing" in all_text(document)) is shown


# [24] A description or footnote is checked on every path: an HTML `data:`
# image, and a model built without the parser.
def test_the_description_of_an_html_data_image_is_checked(tmp_path: Path) -> None:
    import base64

    payload = base64.b64encode(SAMPLE_IMAGE.read_bytes()).decode()
    body = (
        f'<table><tr><td><img src="data:image/png;base64,{payload}" alt="Alt @tbl-missing">'
        "</td></tr></table>"
    )
    path = write(tmp_path, body)
    with pytest.raises(StrictValidationError, match="in the image description stays text"):
        render(path)
    document, _ = render(path, strict=False)
    assert len(list(document.element.body.iter(f"{W}drawing"))) == 1


def _hand_built(elements: list, footnotes: dict | None = None) -> list[str]:
    """Render a model built without the parser and return its diagnostics."""
    from md_parser import DocumentMetadata, DocumentModel

    model = DocumentModel(
        metadata=DocumentMetadata(title="Review", profile="business-report"),
        elements=elements,
        footnotes=footnotes or {},
    )
    renderer = IBDocumentRenderer(options=RenderOptions(strict=False))
    renderer.render(model)
    return renderer.errors


def test_hand_built_footnotes_and_descriptions_are_checked() -> None:
    from md_parser import Element, ElementType, Image, Paragraph, TextRun

    footnote = Element(ElementType.PARAGRAPH, Paragraph("body", runs=[TextRun("1", footnote_id=1)]))
    assert "Cross-reference @tbl-missing in the footnote stays text" in _hand_built(
        [footnote], {1: "Note @tbl-missing"}
    )
    image = Image(alt_text="Alt @tbl-a", path=str(SAMPLE_IMAGE))
    inline = Element(
        ElementType.PARAGRAPH,
        Paragraph("x", runs=[TextRun("x "), TextRun("", image=Image("In @fig-b", ""))]),
    )
    errors = _hand_built([Element(ElementType.IMAGE, image), inline])
    assert "Cross-reference @tbl-a in the image description stays text" in errors
    assert "Cross-reference @fig-b in the image description stays text" in errors


def test_parsed_descriptions_and_footnotes_are_reported_once(tmp_path: Path) -> None:
    body = "![Alt @tbl-missing](s.png)\n\ntext[^1]\n\n[^1]: note @tbl-missing\n"
    _, errors = render(write(tmp_path, body), strict=False)
    assert errors.count("Cross-reference @tbl-missing in the image description stays text") == 1
    assert errors.count("Cross-reference @tbl-missing in the footnote stays text") == 1


# ═══════════════════════════════════════════════════════════════════════════════
# AFTER THE MERGE WITH E6 (TABLE RULES), E7 (OFFICE LETTER), E8 (PROPERTIES)
# ═══════════════════════════════════════════════════════════════════════════════


def _saved_parts(document: DocxDocument) -> dict[str, str]:
    """Every XML part of the saved package, by name."""
    import io
    import zipfile

    buffer = io.BytesIO()
    document.save(buffer)
    with zipfile.ZipFile(buffer) as archive:
        return {
            name: archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.endswith(".xml")
        }


PL_TABLE = """| 구분 | 2025 | 2026 | 비고 |
|---|---|---|---|
| 매출 | 1000 | 1200 | NOTE |
| 비용 | -800 | (900) | |
| 이익 | 200 | -1100 | 합계 |
"""
PL_SPEC = """    type: financial
    columns: [text, money, money, text]
    caption: 가상 손익
    unit: 백만원
    total_rows: [3]
    negative: paren
"""


def test_a_numbered_financial_table_keeps_its_rules_and_restyled_negatives(
    tmp_path: Path,
) -> None:
    (tmp_path / "ref").mkdir()
    (tmp_path / "plain").mkdir()
    body = "본문은 @tbl-pl에서 본다.\n\n" + PL_TABLE.replace("NOTE", "@tbl-pl 참조")
    document, errors = render(write(tmp_path / "ref", body, "  - id: tbl-pl\n" + PL_SPEC))
    control, _ = render(
        write(tmp_path / "plain", PL_TABLE.replace("NOTE", "참조"), "  - " + PL_SPEC[4:])
    )
    assert errors == []
    assert caption_texts(document) == ["표 1. 가상 손익"]
    assert references(document) == [(REF_INSTRUCTION.format("tbl_pl"), "표 1")] * 2
    table, plain = document.tables[0], control.tables[0]
    texts = [[cell.text for cell in row.cells] for row in table.rows]
    assert [row[:3] for row in texts] == [[c.text for c in r.cells][:3] for r in plain.rows]
    assert [row[1:3] for row in texts[2:]] == [["(800)", "(900)"], ["200", "(1,100)"]]
    assert texts[1][3] == "표 1 참조"
    total = table.rows[3]._tr.xpath("./w:tc/w:tcPr/w:tcBorders/w:bottom/@w:val")
    assert total and set(total) == {"double"}
    assert table.rows[3]._tr.xpath("./w:tc/w:tcPr/w:tcBorders/w:bottom/@w:val") == (
        plain.rows[3]._tr.xpath("./w:tc/w:tcPr/w:tcBorders/w:bottom/@w:val")
    )


def test_an_office_letter_ends_a_partial_numbered_table_and_refers_to_it(
    tmp_path: Path,
) -> None:
    source = (
        Path(__file__).resolve().parent.parent / "samples/qa/office-letter-partial-table.md"
    ).read_text(encoding="utf-8")
    source = source.replace(
        "letter:\n  table_end: partial\n",
        "letter:\n  table_end: partial\ntables:\n  - id: tbl-plan\n    caption: 회신 일정\n",
    ).replace("아래 표와 같으며", "@tbl-plan과 같으며")
    path = tmp_path / "letter.md"
    path.write_text(source, encoding="utf-8")
    document, errors = render(path)
    assert errors == []
    assert caption_texts(document) == ["표 1. 회신 일정"]
    assert references(document) == [(REF_INSTRUCTION.format("tbl_plan"), "표 1")]
    rows = [[cell.text for cell in row.cells] for row in document.tables[-1].rows]
    assert ["이하 빈칸", "", ""] in rows
    assert sum(row.count("이하 빈칸") for row in rows) == 1


def test_a_draft_with_a_version_keeps_its_watermark_property_and_numbers(
    tmp_path: Path,
) -> None:
    path = write_front(tmp_path, "@tbl-a를 본다.", 'version: "v0.3"\nstatus: draft\n')
    document, errors = render(path)
    assert errors == []
    assert caption_texts(document) == ["표 1. A"]
    assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]
    parts = _saved_parts(document)
    headers = [text for name, text in parts.items() if name.startswith("word/header")]
    assert any('string="DRAFT"' in header for header in headers)
    assert any("DOCPROPERTY" in header and "Version" in header for header in headers)
    assert re.search(
        r'name="Version"[^>]*>\s*<vt:lpwstr>v0\.3</vt:lpwstr>', parts["docProps/custom.xml"]
    )


@pytest.mark.parametrize("written", [True, False])
def test_a_written_document_title_stays_and_is_the_one_judged(
    tmp_path: Path, written: bool
) -> None:
    from md_parser import parse_markdown_file

    front = "---\nprofile: business-report\n" + ("title: Document\n" if written else "")
    front += "tables:\n" + DEFINED_A + "---\n\n"
    path = tmp_path / "doc.md"
    path.write_text(front + "# 개요 @tbl-a\n\n본문" + TABLE_ONLY, encoding="utf-8")
    model = parse_markdown_file(str(path))
    if written:
        # E8: a written stand-in title is kept; the heading is body text and its reference resolves.
        assert model.metadata.title == "Document"
        document, errors = render(path)
        assert errors == []
        assert references(document) == [(REF_INSTRUCTION.format("tbl_a"), "표 1")]
        assert "<dc:title>Document</dc:title>" in _saved_parts(document)["docProps/core.xml"]
    else:
        # The first H1 becomes the title, which cannot hold a field.
        assert model.metadata.title == "개요 @tbl-a"
        with pytest.raises(StrictValidationError, match="@tbl-a in the document title stays text"):
            render(path)
