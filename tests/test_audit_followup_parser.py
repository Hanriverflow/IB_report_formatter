"""Audit follow-up (2026-11-15), group A: reading Markdown.

Each test goes through the real parser and renderer. Findings: area1 F-03
(HTML cells rewritten by Markdown preprocessing), F-04 (`</table>` inside a
quoted attribute ended the table), F-07 (a code span across lines lost its
`<br>`), F-08 (a closed chart fence lost its trailing spaces and blank lines);
area2 F-02 (callouts parsed escaped punctuation twice); and the removal of
public names nothing used.
"""

import pytest
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.text.run import Run

from document_model import Blockquote, Chart, CodeBlock, ElementType, Table
from document_profiles import RenderOptions
from ib_render_blocks import CalloutRenderer
from ib_renderer import IBDocumentRenderer
from md_parse_blocks import FootnoteParser, LaTeXParser
from md_parser import HtmlTableParser, MarkdownParser

BT = chr(96)
FENCE = BT * 3


def render(markdown: str, profile: str = "plain") -> tuple[DocxDocument, IBDocumentRenderer]:
    """Parse and render `markdown` with `profile` given to parser and renderer."""
    model = MarkdownParser(profile=profile).parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile))
    return renderer.render(model), renderer


def cell_texts(document: DocxDocument) -> list[list[str]]:
    """Text of every cell of every table, row by row."""
    return [[cell.text for cell in row.cells] for table in document.tables for row in table.rows]


def body_texts(document: DocxDocument) -> list[str]:
    """Non-empty body paragraph texts."""
    return [paragraph.text for paragraph in document.paragraphs if paragraph.text]


# ═══════════════════════════════════════════════════════════════════════════════
# A3: HTML table boundary counts real tags only (area1 F-04)
# ═══════════════════════════════════════════════════════════════════════════════


def test_table_close_inside_a_quoted_attribute_does_not_end_the_table() -> None:
    markdown = (
        '<table data-note="</table>">\n<tr><td>Value</td><td>k</td></tr>\n</table>\n\nAfter.\n'
    )
    model = MarkdownParser(profile="plain").parse(markdown)
    assert [element.element_type for element in model.elements] == [
        ElementType.TABLE,
        ElementType.PARAGRAPH,
    ]
    document, _ = render(markdown)
    assert cell_texts(document) == [["Value", "k"]]
    assert body_texts(document) == ["After."]


def test_quoted_attribute_across_the_first_read_window_is_read_whole() -> None:
    # The boundary is read in growing windows; a tag whose quoted value runs
    # past the first window must not be cut there.
    # Lines 0-63 are the first window; the quoted value opens on line 63.
    rows = "".join("<tr><td>r</td><td>s</td></tr>\n" for _ in range(62))
    split = '<tr><td title="a\n</table>\nb">x</td><td>y</td></tr>\n'
    lines = ("<table>\n" + rows + split + "</table>\nAfter.").split("\n")
    end = HtmlTableParser.block_end(lines, 0)
    assert end is not None and lines[end - 1] == "</table>" and lines[end] == "After."


def test_converted_cleaner_uses_the_same_table_boundary() -> None:
    import converted_md_cleaner
    from converted_md_cleaner import _html_table_end

    lines = ['<table data-note="</table>">', "<tr><td>v</td></tr>", "</table>", "text"]
    assert _html_table_end(lines, 0) == HtmlTableParser.block_end(lines, 0) == 3
    assert converted_md_cleaner._HTML_TOKEN_RE is HtmlTableParser.TOKEN_RE


def test_a_quote_outside_an_attribute_value_is_plain_text() -> None:
    # As in HTML, a quote opens a value only right after `=`.
    lines = ["<table>", "<tr><td>a <b it's</td></tr>", "</table>", "x's > y"]
    assert HtmlTableParser.block_end(lines, 0) == 3


def test_unquoted_values_and_a_bare_equals_sign_do_not_open_a_value() -> None:
    lines = ["<table border=1>", "<tr><td x= >v</td><td y=>w</td></tr>", "</table>", "t"]
    assert HtmlTableParser.block_end(lines, 0) == 3
    lines = ["<table data-note = '</table>'>", "</table>", "t"]
    assert HtmlTableParser.block_end(lines, 0) == 2


# ═══════════════════════════════════════════════════════════════════════════════
# A2: HTML cells reach the HTML reader as written (area1 F-03)
# ═══════════════════════════════════════════════════════════════════════════════


def test_reference_link_in_an_html_cell_is_not_expanded() -> None:
    markdown = "<table><tr><td>[x]</td><td>k</td></tr></table>\n\n[x]: target\n"
    document, _ = render(markdown)
    assert cell_texts(document) == [["[x]", "k"]]


def test_footnote_definition_line_in_an_html_cell_is_kept() -> None:
    markdown = "<table><tr><td>\n[^1]: Keep this text\n</td><td>k</td></tr></table>\n"
    model = MarkdownParser(profile="plain").parse(markdown)
    assert model.footnotes == {}
    document, _ = render(markdown)
    assert cell_texts(document) == [["[^1]: Keep this text", "k"]]


def test_link_definition_inside_an_html_table_defines_nothing_outside() -> None:
    markdown = "<table><tr><td>\n[x]: target\n</td></tr></table>\n\nSee [x].\n"
    document, _ = render(markdown)
    assert cell_texts(document) == [["[x]: target"]]
    assert body_texts(document) == ["See [x]."]


def test_reference_links_outside_html_tables_still_resolve() -> None:
    markdown = "<table><tr><td>a</td></tr></table>\n\nSee [x].\n\n[x]: https://example.com\n"
    model = MarkdownParser(profile="plain").parse(markdown)
    paragraph = model.elements[-1].content
    assert [run.hyperlink for run in paragraph.runs if run.hyperlink] == ["https://example.com"]


# ═══════════════════════════════════════════════════════════════════════════════
# A4: a code span across lines keeps its literal text (area1 F-07)
# ═══════════════════════════════════════════════════════════════════════════════


def code_runs(document: DocxDocument) -> list[str]:
    """Text of the runs drawn in the code font, paragraph by paragraph."""
    texts = []
    for paragraph in document.paragraphs:
        texts.extend(run.text for run in paragraph.runs if run.font.name == "Consolas")
    return texts


def test_code_span_across_lines_keeps_its_break_tag() -> None:
    document, _ = render(f"{BT}a<br>\nb{BT}\n")
    assert code_runs(document) == ["a<br> b"]


def test_code_span_across_lines_keeps_trailing_backslash_and_spacing() -> None:
    document, _ = render(f"x {BT}a  b\\\n( c{BT} y\n")
    assert code_runs(document) == ["a  b\\ ( c"]
    assert body_texts(document) == ["x a  b\\ ( c y"]


def test_break_tag_outside_code_is_still_a_hard_break() -> None:
    model = MarkdownParser(profile="plain").parse(f"one<br>\ntwo {BT}c{BT}<br>\nthree\n")
    assert model.elements[0].content.text == f"one\ntwo {BT}c{BT}\nthree"


# ═══════════════════════════════════════════════════════════════════════════════
# A1: callouts parse their body once (area2 F-02)
# ═══════════════════════════════════════════════════════════════════════════════


def callout_body(document: DocxDocument) -> list[tuple[str, bool | None, bool | None]]:
    """(text, bold, italic) of the body runs of the first callout box.

    The callout is the first one-cell table whose title line holds a callout
    icon; IB profiles draw other tables (cover, disclaimer) before it. Runs
    inside content controls (term values) are included, in document order.
    """
    icons = ("▌", "⚠", "ℹ")
    for table in document.tables:
        cell = table.rows[0].cells[0]
        if len(table.rows) == 1 and len(table.columns) == 1:
            if cell.paragraphs[0].text[:1] not in icons:
                continue
            return [
                (run.text, run.bold, run.italic)
                for paragraph in cell.paragraphs[1:]
                for element in paragraph._p.iter(qn("w:r"))
                for run in [Run(element, paragraph)]
            ]
    raise AssertionError("no callout box in the document")


@pytest.mark.parametrize("profile", ["plain", "ib-report", "business-report"])
def test_escaped_asterisks_in_a_callout_stay_literal(profile: str) -> None:
    document, _ = render("> \\*literal\\*\n\nplain \\*literal\\* para\n", profile=profile)
    assert "".join(text for text, _, _ in callout_body(document)) == "*literal*"
    assert not any(italic for _, _, italic in callout_body(document))
    assert "plain *literal* para" in body_texts(document)


def test_escaped_asterisks_in_a_callout_with_term_values_stay_literal() -> None:
    markdown = "---\nterms:\n  issuer: ACME\n---\n> Issuer {{issuer}} \\*lit\\* here.\n"
    model = MarkdownParser(profile="plain").parse(markdown)
    quote = model.elements[0].content
    assert isinstance(quote, Blockquote) and quote.runs  # the term-value branch
    document, _ = render(markdown)
    runs = callout_body(document)
    assert "".join(text for text, _, _ in runs) == "Issuer ACME *lit* here."
    assert not any(italic for _, _, italic in runs)


def test_unresolved_reference_in_a_rewritten_callout_is_reported_once() -> None:
    markdown = (
        "---\ntables:\n  - id: tbl-a\n    caption: Cap\n---\n"
        "| a | b |\n|---|---|\n| 1 | 2 |\n\n> \\*a\\* see @tbl-missing\n"
    )
    model = MarkdownParser(profile="plain").parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True))
    with pytest.raises(ValueError) as rejected:
        renderer.render(model)
    message = "Cross-reference @tbl-missing has no table or figure with that id"
    assert str(rejected.value).count(message) == 1
    assert model.warnings.count(message) + renderer.errors.count(message) == 1


def test_escaped_code_tick_and_break_in_a_callout_stay_literal() -> None:
    document, _ = render(f"> a \\{BT}b\\{BT} \\<br> c\n")
    assert "".join(text for text, _, _ in callout_body(document)) == f"a {BT}b{BT} <br> c"


def test_callout_with_real_emphasis_keeps_no_parse_time_runs() -> None:
    # No escape is read as syntax: the model and its rendering are as before.
    model = MarkdownParser(profile="plain").parse("> **bold** text [site](https://e.com)\n")
    quote = model.elements[0].content
    assert isinstance(quote, Blockquote) and quote.runs == []
    document, _ = render("> **bold** text\n")
    assert callout_body(document)[0][:2] == ("bold", True)


# ═══════════════════════════════════════════════════════════════════════════════
# A5: a closed chart fence keeps its literal content (area1 F-08)
# ═══════════════════════════════════════════════════════════════════════════════


def test_closed_chart_fence_code_is_literal_like_any_closed_fence() -> None:
    body = "x  \n\n"
    chart = MarkdownParser(profile="plain").parse(f"{FENCE}chart\n{body}{FENCE}\n")
    text = MarkdownParser(profile="plain").parse(f"{FENCE}text\n{body}{FENCE}\n")
    assert isinstance(chart.elements[0].content, Chart)
    assert isinstance(text.elements[0].content, CodeBlock)
    assert chart.elements[0].content.code == text.elements[0].content.code == "x  \n"

    chart_cells = cell_texts(render(f"{FENCE}chart\n{body}{FENCE}\n")[0])
    text_cells = cell_texts(render(f"{FENCE}text\n{body}{FENCE}\n")[0])
    assert chart_cells == text_cells


def test_chart_spec_is_read_from_the_trimmed_fence() -> None:
    markdown = (
        f"{FENCE}chart\ntype: bar\ntitle: T\nlabels: [A, B]\n"
        f"series: [{{name: S, values: [1, 2]}}]\n   \n\n{FENCE}\n"
    )
    chart = MarkdownParser(profile="plain").parse(markdown).elements[0].content
    assert isinstance(chart, Chart) and chart.error is None and chart.spec is not None
    assert chart.code.endswith("[1, 2]}]\n   \n")


def test_unclosed_chart_fence_is_trimmed_like_an_unclosed_code_fence() -> None:
    chart = MarkdownParser(profile="plain").parse(f"{FENCE}chart\nx  \n\n").elements[0]
    text = MarkdownParser(profile="plain").parse(f"{FENCE}text\nx  \n\n").elements[0]
    assert chart.content.code == text.content.code == "x"


# ═══════════════════════════════════════════════════════════════════════════════
# A6: public names nothing used are gone
# ═══════════════════════════════════════════════════════════════════════════════


def test_unused_public_names_are_removed() -> None:
    assert not hasattr(CalloutRenderer, "render_executive_summary")
    assert not hasattr(FootnoteParser, "find_inline_references")
    assert not hasattr(LaTeXParser, "INLINE_RE")
    assert not hasattr(LaTeXParser, "extract_inline_segments")
    assert not hasattr(MarkdownParser, "CODE_FENCE_PATTERN")


def test_html_table_is_still_a_table_element() -> None:
    model = MarkdownParser(profile="plain").parse("<table><tr><td>a</td></tr></table>\n")
    assert isinstance(model.elements[0].content, Table)
