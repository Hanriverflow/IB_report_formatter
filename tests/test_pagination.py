"""Engine pagination: page-break separators and short tables (synthetic data only).

A page-break separator becomes "page break before" on the next paragraph, so a
break at the foot of a page cannot leave a blank page. A table of at most
`SHORT_TABLE_MAX_ROWS` rows keeps every row but the last with the next row, so
it moves to the next page whole. Every test parses Markdown with the real
parser and renders through `IBDocumentRenderer.render`.
"""

import pytest
import yaml
from docx.oxml.ns import qn

from document_profiles import RenderOptions
from ib_render_tables import SHORT_TABLE_MAX_ROWS
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

PROFILES = [
    "ib-report",
    "ib-memo",
    "plain",
    "office-letter",
    "business-report",
    "meeting-minutes",
    "term-sheet",
]
PAGE_BREAK = './/w:br[@w:type="page"]'


def markdown(
    body: str,
    profile: str = "plain",
    separator_mode: str | None = "page-break",
    tables: list | None = None,
) -> str:
    data: dict = {"profile": profile, "title": "가상 검토 문서"}
    if profile == "term-sheet":
        data.update(prepared_by="가상은행 자본시장부", disclaimer="가상 조건 검토용 문서입니다.")
    if profile == "office-letter":
        data.update(sender={"organization": "가상회사"}, recipients=["가상 수신부서"])
    if separator_mode is not None:
        data["layout"] = {"separator_mode": separator_mode}
    if tables is not None:
        data["tables"] = tables
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n" + body


def render(content: str, profile: str = "plain"):
    model = MarkdownParser(profile=profile).parse(content)
    return IBDocumentRenderer(options=RenderOptions(profile=profile, strict=True)).render(model)


def text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def body_blocks(doc) -> list:
    return [child for child in doc.element.body if child.tag != qn("w:sectPr")]


def paragraph_with_text(doc, wanted: str):
    """The body paragraph whose text starts with `wanted` (a letter appends `  끝.`)."""
    return next(p for p in doc.element.body.iter(qn("w:p")) if text(p).startswith(wanted))


def author_table(doc):
    """The table written in the Markdown (an IB cover has a table of its own)."""
    return next(
        t for t in doc.element.body.iter(qn("w:tbl")) if text(t).startswith(("구분", "항목0"))
    )


def is_break_only(element) -> bool:
    return (
        element.tag == qn("w:p")
        and not text(element)
        and bool(element.xpath(PAGE_BREAK))
        and not element.xpath("./w:pPr/w:sectPr")
    )


def page_break_before(paragraph) -> bool:
    return bool(paragraph.xpath("./w:pPr/w:pageBreakBefore"))


def separator_bookmarks(element) -> list[str]:
    return [
        name
        for name in (node.get(qn("w:name")) for node in element.iter(qn("w:bookmarkStart")))
        if name.startswith("_ibrep_SEPARATOR_")
    ]


def keeps_with_next(row) -> list[bool]:
    return [bool(p.xpath("./w:pPr/w:keepNext")) for p in row.iter(qn("w:p"))]


def table_markdown(body_rows: int, columns: int = 3) -> str:
    lines = ["| " + " | ".join(f"항목{c}" for c in range(columns)) + " |"]
    lines.append("|" + "---|" * columns)
    for r in range(body_rows):
        lines.append("| " + " | ".join(f"값{r}-{c}" for c in range(columns)) + " |")
    return "\n".join(lines) + "\n"


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE-BREAK SEPARATORS
# ═══════════════════════════════════════════════════════════════════════════════


class TestSeparatorBreak:
    @pytest.mark.parametrize("profile", PROFILES)
    def test_paragraph_after_separator_breaks_before_itself(self, profile):
        doc = render(markdown("앞 문단입니다.\n\n---\n\n뒤 문단입니다.\n", profile), profile)

        after = paragraph_with_text(doc, "뒤 문단입니다.")
        before = after.getprevious()

        assert page_break_before(after)
        assert not is_break_only(before)
        assert text(before) == "앞 문단입니다."
        assert not after.xpath(PAGE_BREAK)

    def test_separator_bookmark_moves_to_the_next_paragraph(self):
        doc = render(markdown("앞 문단.\n\n---\n\n## 다음 절\n"))

        heading = paragraph_with_text(doc, "다음 절")

        assert len(separator_bookmarks(heading)) == 1
        assert sum(len(separator_bookmarks(block)) for block in body_blocks(doc)) == 1
        # The moved bookmark is empty and sits before the heading's own runs.
        start = heading.find(qn("w:bookmarkStart"))
        assert start.getprevious().tag == qn("w:pPr")
        assert start.getnext().tag == qn("w:bookmarkEnd")

    def test_spacer_before_the_separator_is_removed(self):
        doc = render(markdown(table_markdown(10) + "\n---\n\n## 다음 절\n"))

        heading = paragraph_with_text(doc, "다음 절")

        assert page_break_before(heading)
        # The table's empty spacer paragraph went with the separator.
        assert heading.getprevious().tag == qn("w:tbl")

    def test_paragraph_with_content_before_the_separator_stays(self):
        doc = render(markdown("앞 문단.\n\n---\n\n뒤 문단.\n"))

        assert [text(block) for block in body_blocks(doc)] == ["앞 문단.", "뒤 문단."]

    def test_separator_before_a_table_keeps_its_break(self):
        doc = render(markdown("앞 문단.\n\n---\n\n" + table_markdown(3)))

        table = author_table(doc)

        assert is_break_only(table.getprevious())
        assert len(separator_bookmarks(table.getprevious())) == 1
        assert not doc.element.xpath(".//w:pageBreakBefore")

    def test_separator_at_the_end_keeps_its_break(self):
        doc = render(markdown("앞 문단.\n\n---\n"))

        assert is_break_only(body_blocks(doc)[-1])
        assert not doc.element.xpath(".//w:pageBreakBefore")

    def test_two_separators_keep_the_blank_page_they_ask_for(self):
        doc = render(markdown("앞 문단.\n\n---\n\n---\n\n뒤 문단.\n"))

        after = paragraph_with_text(doc, "뒤 문단.")

        assert page_break_before(after)
        assert is_break_only(after.getprevious())
        assert text(after.getprevious().getprevious()) == "앞 문단."

    def test_separator_before_the_disclaimer_page_keeps_its_break(self):
        doc = render(markdown("앞 문단.\n\n## ---\n", "ib-report", None), "ib-report")

        disclaimer = paragraph_with_text(doc, "면책 조항")
        disclaimer_break = disclaimer.getprevious()
        separator = disclaimer_break.getprevious()

        assert is_break_only(disclaimer_break)
        assert is_break_only(separator)
        assert len(separator_bookmarks(separator)) == 1

    def test_cover_and_contents_breaks_are_untouched(self):
        baseline = render(
            markdown("# 제목\n\n앞 문단.\n\n뒤 문단.\n", "ib-report", None), "ib-report"
        )
        doc = render(
            markdown("# 제목\n\n앞 문단.\n\n## ---\n\n뒤 문단.\n", "ib-report", None), "ib-report"
        )

        breaks = [block for block in body_blocks(doc) if is_break_only(block)]
        baseline_breaks = [block for block in body_blocks(baseline) if is_break_only(block)]

        assert len(breaks) == len(baseline_breaks) > 0
        assert doc.element.xpath("count(.//w:pageBreakBefore)") == 1
        assert page_break_before(paragraph_with_text(doc, "뒤 문단."))

    def test_rule_separator_is_unchanged(self):
        doc = render(markdown("앞 문단.\n\n---\n\n뒤 문단.\n", separator_mode="rule"))

        assert not doc.element.xpath(".//w:pageBreakBefore")
        assert not doc.element.xpath(PAGE_BREAK)
        assert doc.element.xpath(".//w:pBdr/w:bottom")

    def test_separator_before_a_landscape_section_keeps_its_break(self):
        body = "앞 문단.\n\n---\n\n" + table_markdown(3)
        doc = render(markdown(body, tables=[{"landscape": True}]))

        section_break = next(block for block in body_blocks(doc) if block.xpath("./w:pPr/w:sectPr"))

        # The paragraph after the separator ends the portrait section: unchanged.
        assert is_break_only(section_break.getprevious())
        assert not doc.element.xpath(".//w:pageBreakBefore")


# ═══════════════════════════════════════════════════════════════════════════════
# SHORT TABLES
# ═══════════════════════════════════════════════════════════════════════════════


class TestShortTable:
    def test_threshold_is_seven_rows(self):
        assert SHORT_TABLE_MAX_ROWS == 7

    @pytest.mark.parametrize("profile", PROFILES)
    def test_short_table_keeps_every_row_but_the_last_with_the_next(self, profile):
        doc = render(markdown(table_markdown(SHORT_TABLE_MAX_ROWS - 1), profile), profile)

        rows = author_table(doc).findall(qn("w:tr"))

        assert len(rows) == SHORT_TABLE_MAX_ROWS
        for row in rows[:-1]:
            assert all(keeps_with_next(row))
        assert not any(keeps_with_next(rows[-1]))

    @pytest.mark.parametrize("profile", PROFILES)
    def test_long_table_body_rows_are_not_chained(self, profile):
        doc = render(markdown(table_markdown(SHORT_TABLE_MAX_ROWS), profile), profile)

        rows = author_table(doc).findall(qn("w:tr"))

        assert len(rows) == SHORT_TABLE_MAX_ROWS + 1
        for row in rows[1:]:
            assert not any(keeps_with_next(row))

    def test_term_sheet_long_table_header_row_still_keeps_with_next(self):
        doc = render(markdown(table_markdown(SHORT_TABLE_MAX_ROWS), "term-sheet"), "term-sheet")

        header = author_table(doc).find(qn("w:tr"))

        assert all(keeps_with_next(header))

    def test_every_paragraph_of_a_multi_line_cell_keeps_with_next(self):
        body = "| 구분 | 내용 |\n|---|---|\n| 조건 | 첫 줄<br>둘째 줄 |\n| 기타 | 끝 |\n"
        doc = render(markdown(body, "term-sheet"), "term-sheet")

        rows = author_table(doc).findall(qn("w:tr"))

        assert len(keeps_with_next(rows[1])) >= 3
        assert all(keeps_with_next(rows[1]))
        assert not any(keeps_with_next(rows[-1]))

    def test_one_row_table_is_unchanged(self):
        doc = render(markdown("| 구분 | 내용 |\n|---|---|\n"))

        rows = author_table(doc).findall(qn("w:tr"))

        assert len(rows) == 1
        assert not any(keeps_with_next(rows[0]))

    def test_keep_with_next_is_in_schema_order(self):
        from lxml import etree

        from ooxml_order import SEQUENCES

        body = "앞 문단.\n\n---\n\n" + table_markdown(3) + "\n---\n\n뒤 문단.\n"
        doc = render(markdown(body))
        order = SEQUENCES["pPr"]

        for paragraph_properties in doc.element.body.iter(qn("w:pPr")):
            names = [etree.QName(child).localname for child in paragraph_properties]
            positions = [order.index(name) for name in names]
            assert positions == sorted(positions), names
