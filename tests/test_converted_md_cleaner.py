"""Converted term-sheet input cleanup (`md-format --converted-term-sheet`); all data fictional."""

import logging
import sys
from pathlib import Path

import pytest
import yaml

import md_formatter
from converted_md_cleaner import clean_converted_term_sheet
from deep_md_cleaner import CleanerConfig
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_formatter import format_file_with_options, format_markdown, main
from md_parser import MarkdownParser

# Converter-shaped input: cover lines ahead of the body, a bold chapter title, an
# escaped numbered title before a table, a numbered list and an appendix band.
CONVERTED = """**Strictly Confidential**

![](media/image1.png){width="1.5in"}

**가나다제일차(유) 유동화증권**\\
**Term Sheet**

2026. 9. 29.

본 자료는 가상의 거래 조건을 설명하기 위해 작성되었습니다. 기재된 금액과 수익률은 예시이며 자금 제공이나
거래 성립을 약속하지 않습니다. 실제 조건은 별도 심사와 계약으로 정해집니다.

라마바은행 자본시장부

Strictly Confidential

**1. 본건 개요**

| 구 분 | 내 용 |
|---|---|
| 발행인 | 가나다제일차(유) |
| 기초자산 | 가나다머티리얼즈㈜ 매출채권 |

2\\. 발행 조건

| 구 분 | 내 용 |
|---|---|
| 발행금액 | 100억원 |

3. 기타 사항

1. 첫째 항목
2. 둘째 항목

| 별첨 | | 상환스케줄 |
|---|---|---|

| 회차 | 상환액 | 잔액 |
|---|---|---|
| 1 | 50 | 50 |
| 2 | 50 | 0 |
"""

DISCLAIMER = (
    "본 자료는 가상의 거래 조건을 설명하기 위해 작성되었습니다. 기재된 금액과 수익률은 예시이며 "
    "자금 제공이나 거래 성립을 약속하지 않습니다. 실제 조건은 별도 심사와 계약으로 정해집니다."
)

BODY = """라마바은행 자본시장부

## 1. 본건 개요

| 구 분 | 내 용 |
|---|---|
| 발행인 | 가나다제일차(유) |
| 기초자산 | 가나다머티리얼즈㈜ 매출채권 |

## 2. 발행 조건

| 구 분 | 내 용 |
|---|---|
| 발행금액 | 100억원 |

3. 기타 사항

1. 첫째 항목
2. 둘째 항목

## 별첨 | 상환스케줄

| 회차 | 상환액 | 잔액 |
|---|---|---|
| 1 | 50 | 50 |
| 2 | 50 | 0 |
"""

FRONT = "---\nprofile: term-sheet\ntitle: 가나다제일차(유) 유동화증권\n---\n"


def _split(cleaned: str):
    _, front, body = cleaned.split("---\n", 2)
    return yaml.safe_load(front), body.lstrip("\n")


def test_converted_input_becomes_term_sheet_markdown() -> None:
    cleaned, _ = clean_converted_term_sheet(CONVERTED)
    front, body = _split(cleaned)
    assert front == {
        "profile": "term-sheet",
        "title": "가나다제일차(유) 유동화증권",
        "subtitle": "Term Sheet",
        "date": "2026. 9. 29.",
        "confidential_label": "Strictly Confidential",
        "disclaimer": DISCLAIMER,
    }
    assert list(front) == ["profile", "title", "subtitle", "date", "confidential_label", "disclaimer"]
    assert body == BODY


def test_every_change_and_undecided_line_is_reported() -> None:
    _, report = clean_converted_term_sheet(CONVERTED)
    assert report.lines() == [
        'line 1: frontmatter confidential_label <- "Strictly Confidential"',
        'line 3: removed cover image "media/image1.png": set it as the house style.logo if it is the logo',
        'line 5: frontmatter title <- "가나다제일차(유) 유동화증권"',
        'line 6: frontmatter subtitle <- "Term Sheet"',
        'line 8: frontmatter date <- "2026. 9. 29."',
        'lines 10-11: frontmatter disclaimer <- "' + DISCLAIMER[:59] + '…"',
        'line 13: kept unclassified cover text: "라마바은행 자본시장부"',
        'line 15: removed repeated confidentiality line "Strictly Confidential"',
        'line 17: heading "**1. 본건 개요**" -> "## 1. 본건 개요" (bold numbered line)',
        'line 24: heading "2\\. 발행 조건" -> "## 2. 발행 조건" (numbered line before a table)',
        'line 30: kept numbered line "3. 기타 사항": not bold and no table follows; '
        "mark it as a heading if it is a chapter",
        'lines 35-36: band table -> "## 별첨 | 상환스케줄"',
        'frontmatter profile <- "term-sheet"',
        "prepared_by is not inferred: add it to the house file or the frontmatter",
    ]


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("**1. 본건 개요**\n\n본문입니다.\n", "## 1. 본건 개요"),
        ("1. 본건 개요\n\n| 가 | 나 |\n|---|---|\n| 1 | 2 |\n", "## 1. 본건 개요"),
        ("1\\. 본건 개요\n\n<table><tr><td>가</td></tr><tr><td>1</td></tr></table>\n", "## 1. 본건 개요"),
        ("**12. 기타**\n", "## 12. 기타"),
    ],
)
def test_numbered_lines_that_become_chapter_headings(source: str, expected: str) -> None:
    cleaned, _ = clean_converted_term_sheet(FRONT + source)
    assert expected + "\n" in cleaned


@pytest.mark.parametrize(
    "source",
    [
        "1. 본건 개요\n\n본문입니다.\n",  # a list item without a table after it
        "1. 본건 개요\n2. 발행 조건\n\n| 가 | 나 |\n|---|---|\n| 1 | 2 |\n",  # a two-item list
        "**123. 번호가 큼**\n",
        "**1. " + "가" * 41 + "**\n",
        "**1. 개요** 뒤에 이어지는 문장\n",
        "```\n**1. 코드 안**\n```\n",
        "    **1. 들여쓴 줄**\n",
        "- 상위 항목\n\n  **1. 목록 안의 줄**\n",  # a list continuation stays in its list
        "> **1. 인용 안의 줄**\n",
        "    | 코드 | 예시 |\n    |---|---|\n",  # an indented table is code, not a band
        "<table><tr><td>\n<table><tr><td>안쪽</td></tr></table>\n\n**1. 바깥 셀 안**\n\n</td></tr></table>\n",
    ],
)
def test_other_numbered_text_is_unchanged(source: str) -> None:
    cleaned, report = clean_converted_term_sheet(FRONT + source)
    assert cleaned == FRONT + "\n" + source
    assert not any("## " in line for line in report.lines())


def test_a_lone_numbered_line_left_as_text_is_reported() -> None:
    _, report = clean_converted_term_sheet(FRONT + "1. 본건 개요\n\n본문입니다.\n")
    assert (
        'line 5: kept numbered line "1. 본건 개요": not bold and no table follows; '
        "mark it as a heading if it is a chapter"
    ) in report.lines()


def test_band_tables_become_headings_and_other_tables_stay() -> None:
    source = (
        "| **별첨** | | 상환스케줄 |\n|:--|---|--:|\n\n"
        "<table>\n<tr><td>참고</td><td> </td><td>용어 &amp; 정의</td></tr>\n</table>\n\n"
        "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n\n"
        "<table><tr><td>가</td></tr><tr><td>1</td></tr></table>\n\n"
        "| | |\n|---|---|\n"
    )
    cleaned, report = clean_converted_term_sheet(FRONT + source)
    assert cleaned == FRONT + "\n" + (
        "## 별첨 | 상환스케줄\n\n"
        "## 참고 | 용어 & 정의\n\n"
        "| 가 | 나 |\n|---|---|\n| 1 | 2 |\n\n"
        "<table><tr><td>가</td></tr><tr><td>1</td></tr></table>\n\n"
        "| | |\n|---|---|\n"
    )
    assert "lines 18-19: kept a one-row table without text" in report.lines()


def test_an_html_one_row_table_with_text_outside_its_cells_stays_a_table() -> None:
    source = "<table>\n<caption>지급 조건 요약</caption>\n<tr><th>조건</th></tr>\n</table>\n"
    cleaned, report = clean_converted_term_sheet(FRONT + source)
    assert cleaned == FRONT + "\n" + source
    assert not any("band table" in line for line in report.lines())


def test_a_heading_followed_directly_by_prose_is_a_chapter_not_cover() -> None:
    prose = "지급은 청구일로부터 삼십 일 이내에 전액으로 하며 상계나 공제 없이 지정 계좌로 이행하여야 합니다. 이는 가상 예시입니다."
    source = f"**초안**\n\n## 1. 지급 조건\n{prose}\n\n## 2. 기타\n"
    cleaned, _ = clean_converted_term_sheet(source)
    front, body = _split(cleaned)
    assert front == {"profile": "term-sheet", "title": "초안"}
    assert body == f"## 1. 지급 조건\n{prose}\n\n## 2. 기타\n"


def test_escaped_hard_break_markers_stay_literal_text() -> None:
    source = (
        "본 가상 문서는 줄바꿈 표시를 글자로 보여 줍니다 \\<br>\n"
        "이 표시는 줄바꿈이 아니며 고지문에 그대로 남아야 합니다. 두 줄은 한 단락입니다.\n\n"
        "**1. 개요**\n"
    )
    front, _ = _split(clean_converted_term_sheet(source)[0])
    assert front["disclaimer"] == (
        "본 가상 문서는 줄바꿈 표시를 글자로 보여 줍니다 \\<br> "
        "이 표시는 줄바꿈이 아니며 고지문에 그대로 남아야 합니다. 두 줄은 한 단락입니다."
    )


def test_hard_breaks_in_a_disclaimer_become_line_breaks() -> None:
    source = (
        "첫째 고지 문장은 가상의 거래 조건을 설명하기 위한 것입니다.<br>\n"
        "둘째 고지 문장은 자금 제공을 약속하지 않는다는 내용입니다.\\\n"
        "셋째 줄은 가상 예시의 마지막 문장입니다.\n\n**1. 개요**\n"
    )
    front, _ = _split(clean_converted_term_sheet(source)[0])
    assert front["disclaimer"] == (
        "첫째 고지 문장은 가상의 거래 조건을 설명하기 위한 것입니다.\n"
        "둘째 고지 문장은 자금 제공을 약속하지 않는다는 내용입니다.\n셋째 줄은 가상 예시의 마지막 문장입니다."
    )


def test_flow_style_frontmatter_is_rewritten_as_block_yaml_and_reported() -> None:
    cleaned, report = clean_converted_term_sheet("---\n{title: 초안, version: v1}\n---\n\n**1. 개요**\n")
    front, _ = _split(cleaned)
    assert front == {"title": "초안", "version": "v1", "profile": "term-sheet"}
    assert "rewrote the frontmatter as block YAML to add keys; its comments and layout were not kept" in report.lines()


def test_a_yaml_document_end_closer_is_rewritten_for_the_parser() -> None:
    cleaned, report = clean_converted_term_sheet("---\ntitle: 초안\n...\n\n**1. 개요**\n\n본문입니다.\n")
    assert cleaned == "---\ntitle: 초안\nprofile: term-sheet\n---\n\n## 1. 개요\n\n본문입니다.\n"
    assert 'line 3: frontmatter closer "..." -> "---" (the parser closes frontmatter with ---)' in report.lines()
    assert MarkdownParser().parse(cleaned).metadata.title == "초안"


def test_existing_keys_are_matched_without_case() -> None:
    cleaned, report = clean_converted_term_sheet("---\nTitle: 기존 제목\n---\n\n**새 제목**\n\n**1. 개요**\n")
    assert "**새 제목**\n\n## 1. 개요" in cleaned
    assert "title: 새 제목" not in cleaned
    assert 'line 5: kept (frontmatter already sets title): "**새 제목**"' in report.lines()


def test_cover_lines_in_one_paragraph_are_classified_line_by_line() -> None:
    source = (
        "**대외비**\n**아자차케미칼㈜ 제1회 무보증 사모사채 인수 조건 (가상 예시용 긴 제목)**\n"
        "**Term Sheet (예비)**\n2026년 9월 29일\n\n## 1. 개요\n\n본문입니다.\n"
    )
    cleaned, report = clean_converted_term_sheet(source)
    front, body = _split(cleaned)
    assert front == {
        "profile": "term-sheet",
        "title": "아자차케미칼㈜ 제1회 무보증 사모사채 인수 조건 (가상 예시용 긴 제목)",
        "subtitle": "Term Sheet (예비)",
        "date": "2026년 9월 29일",
        "confidential_label": "대외비",
    }
    assert body == "## 1. 개요\n\n본문입니다.\n"
    assert "disclaimer not found: the house file or the frontmatter must supply it" in report.lines()


def test_unclear_cover_lines_stay_in_place_and_are_reported() -> None:
    source = (
        "**마바사전자㈜ 운영자금 대출**\n\n"
        "마바사전자㈜ 재무팀 앞\n\n"
        "**작성 부서 메모**\n\n"
        "대외비\n\nStrictly Confidential\n\n"
        "2026. 9. 29.\n\n2026. 10. 1.\n\n"
        "| 수신 | 마바사전자㈜ |\n|---|---|\n| 발신 | 라마바은행 |\n\n"
        "**1. 대출 개요**\n\n본문입니다.\n\nStrictly Confidential\n"
    )
    cleaned, report = clean_converted_term_sheet(source)
    front, body = _split(cleaned)
    assert front["title"] == "마바사전자㈜ 운영자금 대출"
    assert front["confidential_label"] == "대외비"
    assert front["date"] == "2026. 9. 29."
    assert "subtitle" not in front
    assert body == (
        "마바사전자㈜ 재무팀 앞\n\n**작성 부서 메모**\n\nStrictly Confidential\n\n2026. 10. 1.\n\n"
        "| 수신 | 마바사전자㈜ |\n|---|---|\n| 발신 | 라마바은행 |\n\n"
        "## 1. 대출 개요\n\n본문입니다.\n\nStrictly Confidential\n"
    )
    lines = report.lines()
    assert 'line 3: kept unclassified cover text: "마바사전자㈜ 재무팀 앞"' in lines
    assert 'line 5: kept a bold line after the title: "**작성 부서 메모**"' in lines
    assert 'line 9: kept another confidentiality line: "Strictly Confidential"' in lines
    assert 'line 13: kept a second date line: "2026. 10. 1."' in lines
    assert "lines 15-17: kept a table before the first chapter" in lines


def test_existing_frontmatter_lines_and_keys_are_kept() -> None:
    source = (
        "---\n# 기존 설정\nprofile: term-sheet\ntitle: 기존 제목\nhouse: house.yaml\n---\n\n"
        "**가나다제일차(유) 유동화증권**\n**Term Sheet**\n\n**1. 개요**\n\n본문입니다.\n"
    )
    cleaned, report = clean_converted_term_sheet(source)
    assert cleaned.startswith(
        "---\n# 기존 설정\nprofile: term-sheet\ntitle: 기존 제목\nhouse: house.yaml\nsubtitle: Term Sheet\n---\n"
    )
    assert "**가나다제일차(유) 유동화증권**\n\n## 1. 개요" in cleaned
    assert 'line 8: kept (frontmatter already sets title): "**가나다제일차(유) 유동화증권**"' in report.lines()
    assert not any("profile" in line for line in report.lines())


def test_without_a_chapter_heading_the_cover_is_not_moved() -> None:
    cleaned, report = clean_converted_term_sheet("**가나다제일차(유) 유동화증권**\n\n본문입니다.\n")
    assert cleaned == "---\nprofile: term-sheet\n---\n\n**가나다제일차(유) 유동화증권**\n\n본문입니다.\n"
    assert "no chapter heading found: cover lines were not moved" in report.lines()


def test_invalid_existing_frontmatter_is_an_error() -> None:
    with pytest.raises(ValueError, match="mapping"):
        clean_converted_term_sheet("---\n- 목록\n---\n본문\n")


def test_cleaned_output_renders_strictly_without_warnings(tmp_path: Path) -> None:
    house = tmp_path / "house.yaml"
    house.write_text("prepared_by: 라마바은행 구조화금융부\n", encoding="utf-8")
    cleaned, _ = clean_converted_term_sheet(CONVERTED)
    model = MarkdownParser().parse(cleaned)
    assert not model.warnings
    renderer = IBDocumentRenderer(options=RenderOptions(strict=True, house=str(house)))
    doc = renderer.render(model)
    assert not renderer.errors
    assert doc.core_properties.title == "가나다제일차(유) 유동화증권"
    texts = [paragraph.text for paragraph in doc.paragraphs]
    assert {"1. 본건 개요", "2. 발행 조건", "별첨 | 상환스케줄", "라마바은행 자본시장부"} <= set(texts)
    assert not doc.element.body.xpath(".//w:drawing")


def test_the_formatter_uses_the_cleanup_only_when_asked(tmp_path: Path) -> None:
    source = tmp_path / "converted.md"
    source.write_text(CONVERTED, encoding="utf-8")
    plain = Path(format_file_with_options(str(source), str(tmp_path / "plain.md")))
    assert plain.read_text(encoding="utf-8") == format_markdown(CONVERTED, cleaner_config=CleanerConfig())
    cleaned = Path(format_file_with_options(str(source), str(tmp_path / "ts.md"), converted_term_sheet=True))
    assert cleaned.read_text(encoding="utf-8") == clean_converted_term_sheet(CONVERTED)[0]
    for options in ({"cleaner_mode": "on"}, {"cite_mode": "strip"}, {"cleaner_report": True}):
        with pytest.raises(ValueError, match="DeepResearch"):
            format_file_with_options(str(source), str(tmp_path / "x.md"), converted_term_sheet=True, **options)


@pytest.mark.parametrize("extra", [["--check"], ["--cite-mode", "strip"], ["--deepresearch-cleaner", "on"]])
def test_the_cli_rejects_options_that_do_not_apply(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, extra: list
) -> None:
    source = tmp_path / "converted.md"
    source.write_text(CONVERTED, encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["md-format", str(source), "--converted-term-sheet", *extra])
    with pytest.raises(SystemExit) as exit_info:
        main()
    assert exit_info.value.code == 2
    assert not (tmp_path / "converted_formatted.md").exists()


def test_the_cli_flag_writes_the_cleaned_file_and_logs_the_report(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    source = tmp_path / "converted.md"
    source.write_text(CONVERTED, encoding="utf-8")
    target = tmp_path / "term-sheet.md"
    monkeypatch.setattr(sys, "argv", ["md-format", str(source), str(target), "--converted-term-sheet"])
    monkeypatch.setattr(md_formatter, "configure_logging", lambda: None)  # keep caplog's root handler
    with caplog.at_level(logging.INFO, logger="md_formatter"):
        main()
    assert target.read_text(encoding="utf-8") == clean_converted_term_sheet(CONVERTED)[0]
    assert "Converted term sheet: lines 35-36: band table -> \"## 별첨 | 상환스케줄\"" in caplog.messages
