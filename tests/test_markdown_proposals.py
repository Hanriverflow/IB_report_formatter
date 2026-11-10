"""Edges of the four Markdown reading changes: bullets, deep headings, fences, escapes."""

from io import BytesIO

import pytest
from docx import Document

from document_model import ElementType, Table
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

NEW_ESCAPES = "!#%&+,/:;<=>?@"


def _docx(source: str, profile: str = "plain", strict: bool = True):
    model = MarkdownParser(profile=profile).parse(source)
    document = IBDocumentRenderer(
        options=RenderOptions(
            profile=profile, strict=strict, include_cover=False, include_toc=False
        )
    ).render(model)
    stream = BytesIO()
    document.save(stream)
    return Document(BytesIO(stream.getvalue()))


def _texts(document) -> list[str]:
    texts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        texts.extend(cell.text for row in table.rows for cell in row.cells)
    return texts


def _cell_texts(source: str) -> list[str]:
    model = MarkdownParser(profile="plain").parse("| 구분 | 내용 |\n|---|---|\n" + source)
    table = model.elements[0].content
    assert isinstance(table, Table)
    return ["".join(run.text for run in cell.runs) for cell in table.rows[1].cells]


@pytest.mark.parametrize("char", ["%", "#", ",", ";", "!", ":", "&"])
@pytest.mark.parametrize("profile", ["plain", "ib-report"])
def test_callout_shows_text_between_dollar_signs_as_written(profile: str, char: str) -> None:
    """A callout does not render math; `\\%` between dollar signs is not a prose escape."""
    payload = "$12\\" + char + "$"
    source = f"> **Key Insight**\n> 가상 수익률 {payload} 와 끝"
    texts = _texts(_docx(source, profile, strict=False))
    assert any(payload in text for text in texts), texts


@pytest.mark.parametrize("char", ["%", "#"])
@pytest.mark.parametrize(
    "template",
    [
        "**1. 가상 금리 ${}$ 전망**",
        "가상 **굵게 ${}$ 안** 문단",
        "가상 ${}$ 문단",
        "| 구분 | 내용 |\n|---|---|\n| 가상 ${}$ | 예시 |",
    ],
)
def test_math_with_latex_commands_still_converts_in_strict_mode(template: str, char: str) -> None:
    """Where math renders, an unescaped `%` or `#` would fail the equation."""
    document = _docx(template.format("5\\" + char) + "\n\n가상 끝", "ib-report")
    assert any("가상 끝" in text for text in _texts(document))
    assert len(document.element.body.findall(".//{*}drawing")) == 1


@pytest.mark.parametrize("payload", ["5\\%", "\\#1"])
def test_numbered_heading_with_math_still_converts_in_strict_mode(payload: str) -> None:
    document = _docx(f"**1. 가상 금리 ${payload}$ 전망**\n\n가상 문단", "ib-report")
    assert any("전망" in text for text in _texts(document))


@pytest.mark.parametrize("char", NEW_ESCAPES)
def test_code_span_and_link_destination_in_a_quote_are_unchanged(char: str) -> None:
    code = _docx(f"> 가상 `코드 \\{char} 구간` 글")
    assert any(f"코드 \\{char} 구간" in text for text in _texts(code))
    model = MarkdownParser(profile="plain").parse(f"> [글](http://example.com/a\\{char}b)")
    assert f"a\\{char}b" in model.elements[0].content.text


@pytest.mark.parametrize("char", ["#", "!", "*", "_", "\\"])
def test_escaped_backslash_in_a_table_cell_is_resolved_once(char: str) -> None:
    """`\\\\#` is a backslash followed by `#`: the row splitter must not unescape it first."""
    assert _cell_texts(f"| x\\\\{char} y | 예시 |")[0] == f"x\\{char} y"


def test_table_cell_pipes_paths_and_numbers() -> None:
    assert _cell_texts("| a\\|b | c\\\\|") == ["a|b", "c\\"]
    assert _cell_texts("| C:\\\\Users\\\\data | a\\\\b |") == ["C:\\Users\\data", "a\\b"]
    assert _cell_texts("| \\-1,234 | 12\\% |") == ["-1,234", "12%"]


def test_escaped_html_break_and_double_backslash() -> None:
    texts = _texts(_docx("가상 \\<br> 와 \\<BR/> 글\n\n가상 \\\\! 와 \\\\\\! 끝"))
    assert "가상 <br> 와 <BR/> 글" in texts
    assert "가상 \\! 와 \\! 끝" in texts


def test_escaped_block_markers_are_text() -> None:
    texts = _texts(
        _docx("\\# 제목 아님\n\n\\+ 항목 아님\n\n가상 \\![대체](https://example.com/a) 글")
    )
    assert "# 제목 아님" in texts
    assert "+ 항목 아님" in texts
    assert "가상 !대체 글" in texts


@pytest.mark.parametrize("profile", ["plain", "ib-report", "business-report"])
def test_deep_headings_use_the_level_four_style(profile: str) -> None:
    document = _docx(
        "#### 가상 넷\n\n##### 가상 다섯\n\n###### 가상 여섯\n\n####### 가상 일곱", profile
    )
    styles = {paragraph.text: paragraph.style.name for paragraph in document.paragraphs}
    assert styles["가상 다섯"] == styles["가상 넷"] == styles["가상 여섯"]
    assert styles["####### 가상 일곱"] != styles["가상 넷"]


def test_deep_reference_header_is_a_reference_section() -> None:
    """Level 5 and 6 reference headers are consumed like levels 1 to 4."""
    deep = MarkdownParser(profile="ib-report").parse(
        "가상 문단 [1]\n\n##### References\n\n[1] 가상 출처"
    )
    flat = MarkdownParser(profile="ib-report").parse(
        "가상 문단 [1]\n\n#### References\n\n[1] 가상 출처"
    )
    assert [element.element_type for element in deep.elements] == [
        element.element_type for element in flat.elements
    ]
    assert deep.footnotes == flat.footnotes


def test_plus_bullets_nest_and_do_not_catch_signed_numbers() -> None:
    model = MarkdownParser(profile="plain").parse(
        "+ 가상 하나\n    + 가상 안쪽\n+ 가상 둘\n\n+82-10-0000-0000 번호\n\n+5% 증가"
    )
    kinds = [element.element_type for element in model.elements]
    assert kinds == [ElementType.BULLET_LIST] * 3 + [ElementType.PARAGRAPH] * 2
    levels = [element.content.indent_level for element in model.elements[:3]]
    assert levels[1] == levels[0] + 1 and levels[2] == levels[0]


@pytest.mark.parametrize("ending", ["", "\n"])
def test_closed_fence_at_end_of_input_keeps_trailing_blank_line(ending: str) -> None:
    model = MarkdownParser(profile="plain").parse("```text\n가상 줄\n\n```" + ending)
    assert model.elements[0].content.code == "가상 줄\n"
    unclosed = MarkdownParser(profile="plain").parse("```text\n가상 줄\n\n" + ending)
    assert unclosed.elements[0].content.code == "가상 줄"
