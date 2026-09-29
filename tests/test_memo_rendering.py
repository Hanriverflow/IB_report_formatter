"""Improvements found on a real internal memo: inline code, local file links, Korean spacing."""

import logging
import sys
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, TextParser, parse_markdown_file


def _render(markdown: str, profile: str = "business-report", strict: bool = True):
    model = MarkdownParser(profile=profile).parse(markdown)
    doc = IBDocumentRenderer(options=RenderOptions(profile=profile, strict=strict)).render(model)
    payload = BytesIO()
    doc.save(payload)
    return Document(payload)


def _hyperlink_targets(doc) -> list:
    targets = []
    for link in doc.element.body.iter(qn("w:hyperlink")):
        relationship = doc.part.rels[link.get(qn("r:id"))]
        assert relationship.reltype == RT.HYPERLINK and relationship.is_external
        # python-docx's run elements repeat text in itertext(); read w:t nodes instead.
        targets.append(("".join(link.xpath(".//w:t/text()")), relationship.target_ref))
    return targets


# ═══════════════════════════════════════════════════════════════════════════════
# INLINE CODE
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("source,expected", [
    ("계좌 `022-1` 참고", [("계좌 ", False), ("022-1", True), (" 참고", False)]),
    ("`` a`b ``", [("a`b", True)]),
    (r"\`literal\`", [("`literal`", False)]),
])
def test_code_spans_become_literal_code_runs(source: str, expected: list) -> None:
    assert [(run.text, run.code) for run in TextParser.parse_runs(source)] == expected


def test_code_run_keeps_surrounding_emphasis_and_link() -> None:
    bold = TextParser.parse_runs("**`코드`** 끝")[0]
    assert (bold.text, bold.code, bold.bold) == ("코드", True, True)
    link = TextParser.parse_runs("[`x`](https://example.com)")[0]
    assert (link.text, link.code, link.hyperlink) == ("x", True, "https://example.com")


def test_rendered_code_has_no_backticks_and_uses_the_code_font() -> None:
    doc = _render("# 메모\n\n계좌번호는 `022-9200-0000-000`입니다.\n")
    paragraph = next(p for p in doc.paragraphs if "022-9200-0000-000" in p.text)
    assert "`" not in paragraph.text
    code_run = next(run for run in paragraph.runs if run.text == "022-9200-0000-000")
    assert code_run.font.name == "Consolas"


def _code_texts(element) -> list:
    return [
        "".join(run.xpath("./w:t/text()"))
        for run in element.iter(qn("w:r"))
        if run.xpath('./w:rPr/w:rFonts[@w:ascii="Consolas"]')
    ]


def test_plain_callout_renders_code_and_links_like_body_text() -> None:
    doc = _render("# 메모\n\n> 계좌 `022-1` 와 [안내](https://example.com/a) 참고\n")
    assert "`" not in "".join(doc.element.body.xpath(".//w:t/text()"))
    assert _code_texts(doc.element.body) == ["022-1"]
    assert _hyperlink_targets(doc) == [("안내", "https://example.com/a")]


def test_term_sheet_line_splitting_keeps_code_whitespace() -> None:
    markdown = (
        "---\nprofile: term-sheet\ntitle: 가나다머티리얼즈㈜ 조건 검토\n"
        "prepared_by: 라마바은행 자본시장부\ndisclaimer: 가상 조건 검토용입니다.\n---\n\n"
        "## 1. 개요\n\n앞<br>`  b  `<br>끝\n\n"
        "| 구 분 | 내 용 |\n|---|---|\n| 코드 | 앞<br>`  c  `<br>`   `<br>끝 |\n"
    )
    doc = _render(markdown, profile="term-sheet")
    assert _code_texts(doc.element.body) == [" b ", " c ", "   "]


def test_code_in_a_money_column_keeps_its_literal_digits() -> None:
    markdown = (
        "---\nprofile: business-report\ntitle: 가상 메모\n"
        "tables:\n  - {columns: [text, money]}\n---\n\n"
        "| 항목 | 금액 |\n|---|---:|\n| 코드 | `1234567` |\n| 숫자 | 1234567 |\n"
    )
    cells = [cell.text for cell in _render(markdown).tables[0].columns[1].cells]
    assert cells[1:] == ["1234567", "1,234,567"]


# ═══════════════════════════════════════════════════════════════════════════════
# LOCAL FILE LINKS
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("source,target", [
    ("[웹](https://example.com/a)", "https://example.com/a"),
    ("[정리](03_정리.md)", "03_%EC%A0%95%EB%A6%AC.md"),
    ("[꺾쇠](<C:/가상 폴더/정리.md>)", "file:///C:/%EA%B0%80%EC%83%81%20%ED%8F%B4%EB%8D%94/%EC%A0%95%EB%A6%AC.md"),
    (r"[윈도](C:\docs\a.pdf)", "file:///C:/docs/a.pdf"),
    ("[상위](../자료/표.xlsx)", "../%EC%9E%90%EB%A3%8C/%ED%91%9C.xlsx"),
])
def test_local_and_web_links_become_hyperlinks(source: str, target: str) -> None:
    runs = TextParser.parse_runs(source)
    assert len(runs) == 1 and runs[0].hyperlink == target


@pytest.mark.parametrize("source", [
    "[참고](별첨)", "[주](Co.Ltd)", "[v](1.0)", "[도메인](www.example.com/a)", "[공백](./a b.md)",
])
def test_non_path_parentheses_stay_literal(source: str) -> None:
    runs = TextParser.parse_runs(source)
    assert "".join(run.text for run in runs) == source
    assert not any(run.hyperlink for run in runs)


def test_parsed_links_keep_absolute_file_targets_until_saved(tmp_path: Path) -> None:
    inside = tmp_path / "sub" / "가상 정리.md"
    source = tmp_path / "memo.md"
    source.write_text(
        "# 메모\n\n[안쪽][in] · [바깥][out] · [인라인](<" + inside.as_posix() + ">)\n\n"
        "[in]: <" + inside.as_posix() + ">\n[out]: <C:/elsewhere/other.md>\n",
        encoding="utf-8",
    )
    model = parse_markdown_file(str(source))
    assert model.source_dir == tmp_path.resolve()
    runs = model.elements[-1].content.runs
    assert [(run.text, run.hyperlink) for run in runs if run.hyperlink] == [
        ("안쪽", inside.as_uri()),
        ("바깥", "file:///C:/elsewhere/other.md"),
        ("인라인", inside.as_uri()),
    ]
    assert not model.warnings


def _write_memo(folder: Path, body: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    source = folder / "memo.md"
    source.write_text(
        "---\nprofile: business-report\ntitle: 가상 메모\n---\n\n# 메모\n\n" + body + "\n",
        encoding="utf-8",
    )
    return source


def _convert(entrypoint: str, source: Path, output: Path) -> None:
    """Save through the converter (CLI/API) or the registry, the two save paths."""
    if entrypoint == "converter":
        from md_to_word import IBReportConverter

        IBReportConverter(str(source), str(output), render_options=RenderOptions(strict=True)).convert()
    else:
        from converters import get_default_registry

        registry = get_default_registry()
        registry.convert(registry.convert(str(source)), output_format="docx", output_path=str(output), strict=True)


@pytest.mark.parametrize("entrypoint", ["converter", "registry"])
def test_saved_links_are_rebased_on_the_docx_folder(tmp_path: Path, entrypoint: str, caplog) -> None:
    source_dir = tmp_path / "source"
    attachment = source_dir / "첨부.pdf"
    source = _write_memo(
        source_dir,
        "[절대](<" + attachment.as_posix() + ">) · [상대](sub/정리.md) · [웹](https://example.com/a)",
    )
    same, export = source_dir / "same.docx", tmp_path / "export" / "export.docx"
    _convert(entrypoint, source, same)
    with caplog.at_level(logging.WARNING):
        _convert(entrypoint, source, export)
    assert _hyperlink_targets(Document(str(same))) == [
        ("절대", "%EC%B2%A8%EB%B6%80.pdf"),
        ("상대", "sub/%EC%A0%95%EB%A6%AC.md"),
        ("웹", "https://example.com/a"),
    ]
    assert _hyperlink_targets(Document(str(export))) == [
        ("절대", attachment.as_uri()),
        ("상대", "../source/sub/%EC%A0%95%EB%A6%AC.md"),
        ("웹", "https://example.com/a"),
    ]
    assert any("outside the output folder" in record.message for record in caplog.records)


def test_file_uris_angle_destinations_and_file_hashes_survive_saving(tmp_path: Path) -> None:
    source = _write_memo(
        tmp_path,
        "[URI](" + (tmp_path / "b.pdf").as_uri() + ") · [안내](<README>) · [정의][d] · "
        "[데이터](<" + (tmp_path / "data.json").as_posix() + ">) · "
        "[해시](<" + (tmp_path / "a#b.pdf").as_posix() + ">) · [조각](정리.md#개요)\n\n[d]: <README>",
    )
    output = tmp_path / "memo.docx"
    _convert("registry", source, output)
    assert _hyperlink_targets(Document(str(output))) == [
        ("URI", "b.pdf"),
        ("안내", "README"),
        ("정의", "README"),
        ("데이터", "data.json"),
        ("해시", "a%23b.pdf"),
        ("조각", "%EC%A0%95%EB%A6%AC.md#%EA%B0%9C%EC%9A%94"),
    ]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows path syntax")
def test_windows_path_with_markdown_escape_characters_is_rebased(tmp_path: Path) -> None:
    target = tmp_path / "_a.pdf"
    source = _write_memo(tmp_path, "[윈도](" + str(target) + ")")
    output = tmp_path / "memo.docx"
    _convert("registry", source, output)
    assert _hyperlink_targets(Document(str(output))) == [("윈도", "_a.pdf")]


@pytest.mark.parametrize("parse", [TextParser.parse_runs, TextParser.parse_runs_plain])
def test_code_span_inside_a_destination_is_restored_before_encoding(parse) -> None:
    assert [(run.text, run.hyperlink) for run in parse("[x](a`b`.pdf)")] == [("x", "a%60b%60.pdf")]


def test_rendered_local_link_is_an_external_hyperlink_showing_only_its_label() -> None:
    doc = _render("# 메모\n\n자세한 내용은 [계좌 정리](03_계좌_정리.md)를 참고.\n")
    assert _hyperlink_targets(doc) == [("계좌 정리", "03_%EA%B3%84%EC%A2%8C_%EC%A0%95%EB%A6%AC.md")]
    assert "03_계좌_정리.md" not in "".join(p.text for p in doc.paragraphs)


# ═══════════════════════════════════════════════════════════════════════════════
# KOREAN SPACING
# ═══════════════════════════════════════════════════════════════════════════════


def _auto_spacing(doc) -> dict:
    defaults = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:pPrDefault")).find(qn("w:pPr"))
    return {
        tag: (defaults.find(qn("w:" + tag)).get(qn("w:val")) if defaults.find(qn("w:" + tag)) is not None else None)
        for tag in ("autoSpaceDE", "autoSpaceDN")
    }


@pytest.mark.parametrize("profile", ["business-report", "meeting-minutes", "plain"])
def test_general_profiles_turn_off_korean_auto_spacing(profile: str) -> None:
    doc = _render("# 메모\n\nSPC는 제2종 300억원을 보유합니다.\n", profile=profile)
    assert _auto_spacing(doc) == {"autoSpaceDE": "0", "autoSpaceDN": "0"}


@pytest.mark.parametrize("profile", ["ib-report", "ib-memo"])
def test_ib_profiles_keep_word_default_spacing(profile: str) -> None:
    doc = _render("# Memo\n\nBody.\n", profile=profile, strict=False)
    assert _auto_spacing(doc) == {"autoSpaceDE": None, "autoSpaceDN": None}
