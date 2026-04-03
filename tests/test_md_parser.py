"""
Tests for md_parser module.

Tests core parsing functionality including:
- Frontmatter parsing
- Element type detection
- Table parsing
- Text run formatting
"""

from pathlib import Path

from md_parser import (
    Blockquote,
    CodeBlock,
    DocumentMetadata,
    ElementType,
    FrontmatterParser,
    LaTeXEquation,
    MarkdownParser,
    Paragraph,
    TableParser,
    TableType,
    TextParser,
    parse_markdown_file,
)

# ═══════════════════════════════════════════════════════════════════════════════
# FRONTMATTER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestFrontmatterParser:
    """Tests for YAML frontmatter parsing."""

    def test_parse_basic_frontmatter(self):
        """Parse simple frontmatter with all fields."""
        lines = [
            "---",
            "title: Test Report",
            "subtitle: Q4 Analysis",
            "company: Korea Development Bank",
            "ticker: KDB",
            "sector: Banking",
            "analyst: John Doe",
            "---",
            "# Content starts here",
        ]

        metadata, remaining = FrontmatterParser.parse(lines)

        assert metadata.title == "Test Report"
        assert metadata.subtitle == "Q4 Analysis"
        assert metadata.company == "Korea Development Bank"
        assert metadata.ticker == "KDB"
        assert metadata.sector == "Banking"
        assert metadata.analyst == "John Doe"
        assert remaining == ["# Content starts here"]

    def test_parse_no_frontmatter(self):
        """Handle files without frontmatter."""
        lines = [
            "# Just a heading",
            "Some content here.",
        ]

        metadata, remaining = FrontmatterParser.parse(lines)

        assert metadata.title == "IB Report"  # Default
        assert remaining == lines

    def test_parse_partial_frontmatter(self):
        """Parse frontmatter with only some fields."""
        lines = [
            "---",
            "title: Partial Report",
            "---",
            "Content",
        ]

        metadata, remaining = FrontmatterParser.parse(lines)

        assert metadata.title == "Partial Report"
        assert metadata.company == "Korea Development Bank"  # Default
        assert remaining == ["Content"]

    def test_non_yaml_block_between_rules_is_not_frontmatter(self):
        """Do not consume markdown content wrapped by horizontal rules."""
        lines = [
            "---",
            "",
            "# Report Title",
            "",
            "**Label:** value",
            "",
            "---",
            "",
            "## 1. Section",
        ]

        metadata, remaining = FrontmatterParser.parse(lines)

        assert metadata == DocumentMetadata()
        assert remaining == lines


# ═══════════════════════════════════════════════════════════════════════════════
# TEXT PARSER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestTextParser:
    """Tests for inline text formatting parsing."""

    def test_parse_plain_text(self):
        """Parse text without formatting."""
        runs = TextParser.parse_runs("Hello world")

        assert len(runs) == 1
        assert runs[0].text == "Hello world"
        assert runs[0].bold is False

    def test_parse_bold_text(self):
        """Parse text with bold markers."""
        runs = TextParser.parse_runs("This is **bold** text")

        assert len(runs) == 3
        assert runs[0].text == "This is "
        assert runs[0].bold is False
        assert runs[1].text == "bold"
        assert runs[1].bold is True
        assert runs[2].text == " text"
        assert runs[2].bold is False

    def test_parse_bold_text_preserves_korean_spacing(self):
        """Keep spacing around bold boundaries in Korean text."""
        source = (
            "발행 주체 및 규제 체계에 따라 **기본자본(Tier 1)**과 "
            "**보완자본(Tier 2)**으로 구분됩니다."
        )

        runs = TextParser.parse_runs(source)
        reconstructed = "".join(run.text for run in runs)

        assert (
            reconstructed
            == "발행 주체 및 규제 체계에 따라 기본자본(Tier 1)과 보완자본(Tier 2)으로 구분됩니다."
        )

    def test_parse_multiple_bold(self):
        """Parse text with multiple bold sections."""
        runs = TextParser.parse_runs("**First** and **Second**")

        bold_runs = [r for r in runs if r.bold]
        assert len(bold_runs) == 2
        assert bold_runs[0].text == "First"
        assert bold_runs[1].text == "Second"

    def test_cleanup_text(self):
        """Test escape character removal."""
        result = TextParser.cleanup_text(r"Test \* escaped \~ chars")
        assert result == "Test * escaped ~ chars"

    def test_has_inline_latex(self):
        """Detect inline LaTeX expressions."""
        assert TextParser.has_inline_latex("The formula $x^2$ is simple")
        assert TextParser.has_inline_latex("WACC = $r_f + \\beta(r_m - r_f)$")
        assert not TextParser.has_inline_latex("Price is $100")
        assert not TextParser.has_inline_latex("No latex here")

    def test_parse_italic_text(self):
        """Parse italic markers conservatively."""
        runs = TextParser.parse_runs("This is *italic* and _also italic_")

        italic_runs = [run for run in runs if run.italic]
        assert len(italic_runs) == 2
        assert italic_runs[0].text == "italic"
        assert italic_runs[1].text == "also italic"

    def test_parse_superscript_runs(self):
        """Parse superscript markers into TextRun metadata."""
        runs = TextParser.parse_runs("Value^1^ adjusted")

        superscript_runs = [run for run in runs if run.superscript]
        assert len(superscript_runs) == 1
        assert superscript_runs[0].text == "1"

    def test_parse_subscript_runs(self):
        """Parse subscript markers into TextRun metadata."""
        runs = TextParser.parse_runs("H~2~O")

        subscript_runs = [run for run in runs if run.subscript]
        assert len(subscript_runs) == 1
        assert subscript_runs[0].text == "2"

    def test_parse_korean_range_tilde_is_not_subscript(self):
        """Range tildes in Korean date spans should not be consumed as subscript markers."""
        text = "제9기(2024년 1월~12월) → 제10기(2025년 1월~12월)"
        runs = TextParser.parse_runs(text)

        assert "".join(run.text for run in runs) == text
        assert not any(run.subscript for run in runs)

    def test_parse_color_spans(self):
        """Parse HTML color spans into TextRun color metadata."""
        runs = TextParser.parse_runs('<span style="color:#C00000">**Loss**</span> buffer')

        colored_runs = [run for run in runs if run.color_hex == "#C00000"]
        assert len(colored_runs) == 1
        assert colored_runs[0].text == "Loss"
        assert colored_runs[0].bold is True

    def test_parse_inline_latex_runs(self):
        """Parse inline LaTeX into dedicated TextRun metadata."""
        runs = TextParser.parse_runs("Formula $x^2$ done")

        latex_runs = [run for run in runs if run.is_latex]
        assert len(latex_runs) == 1
        assert latex_runs[0].text == "x^2"


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE PARSER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestTableParser:
    """Tests for markdown table parsing."""

    def test_parse_simple_table(self):
        """Parse a basic 2x2 table."""
        lines = [
            "| Header 1 | Header 2 |",
            "|----------|----------|",
            "| Cell 1   | Cell 2   |",
        ]

        table = TableParser.parse(lines)

        assert table.col_count == 2
        assert len(table.rows) == 2
        assert table.rows[0].cells[0].content == "Header 1"
        assert table.rows[1].cells[1].content == "Cell 2"

    def test_detect_financial_table(self):
        """Detect financial table type."""
        lines = [
            "| Metric | 2024A | 2025E |",
            "|--------|-------|-------|",
            "| Revenue | 100 | 120 |",
        ]

        table = TableParser.parse(lines)
        assert table.table_type == TableType.FINANCIAL

    def test_detect_risk_table(self):
        """Detect risk matrix table type."""
        lines = [
            "| Risk | Impact | Probability |",
            "|------|--------|-------------|",
            "| Market | High | Medium |",
        ]

        table = TableParser.parse(lines)
        assert table.table_type == TableType.RISK_MATRIX

    def test_parse_numeric_cells(self):
        """Detect numeric cells in table."""
        lines = [
            "| Item | Amount |",
            "|------|--------|",
            "| Revenue | 1234567 |",
        ]

        table = TableParser.parse(lines)
        data_row = table.rows[1]

        assert data_row.cells[0].is_numeric is False  # "Revenue"
        assert data_row.cells[1].is_numeric is True  # "1234567"

    def test_parse_negative_numbers(self):
        """Detect negative numbers in parentheses."""
        lines = [
            "| Item | Amount |",
            "|------|--------|",
            "| Loss | (500) |",
        ]

        table = TableParser.parse(lines)
        assert table.rows[1].cells[1].is_negative is True

    def test_preserve_blank_leading_header_cell(self):
        """Keep an intentionally blank first header cell so data columns stay aligned."""
        lines = [
            "|| Amount | Share |",
            "|---|---|---|",
            "| Tranche A | 7,000 | 63.6% |",
        ]

        table = TableParser.parse(lines)

        assert table.col_count == 3
        assert [cell.content for cell in table.rows[0].cells] == ["", "Amount", "Share"]
        assert [cell.content for cell in table.rows[1].cells] == ["Tranche A", "7,000", "63.6%"]

    def test_preserve_blank_interior_cells(self):
        """Keep empty interior cells instead of collapsing later columns leftward."""
        lines = [
            "| Item | Amount | Notes |",
            "|------|--------|-------|",
            "| Tranche B | 2,000 | |",
        ]

        table = TableParser.parse(lines)

        assert table.col_count == 3
        assert [cell.content for cell in table.rows[1].cells] == ["Tranche B", "2,000", ""]

    def test_parse_table_cell_inline_latex(self):
        """Balanced inline LaTeX in a table cell should preserve latex run metadata."""
        lines = [
            "| Metric | Formula |",
            "|--------|---------|",
            "| Value | $x^2$ |",
        ]

        table = TableParser.parse(lines)
        latex_runs = [run for run in table.rows[1].cells[1].runs if run.is_latex]

        assert len(latex_runs) == 1
        assert latex_runs[0].text == "x^2"

    def test_parse_fenced_code_block(self):
        """Fenced code blocks should become dedicated code block elements."""
        content = """## Diagram

```text
A
└── B
```
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        code_blocks = [e for e in model.elements if e.element_type == ElementType.CODE_BLOCK]
        assert len(code_blocks) == 1
        assert isinstance(code_blocks[0].content, CodeBlock)
        assert "└── B" in code_blocks[0].content.code


# ═══════════════════════════════════════════════════════════════════════════════
# MARKDOWN PARSER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestMarkdownParser:
    """Tests for main markdown parser."""

    def test_parse_headings(self):
        """Parse different heading levels."""
        content = """# Heading 1
## Heading 2
### Heading 3
#### Heading 4
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        heading_types = [e.element_type for e in model.elements]
        assert ElementType.HEADING_1 in heading_types
        assert ElementType.HEADING_2 in heading_types
        assert ElementType.HEADING_3 in heading_types
        assert ElementType.HEADING_4 in heading_types

    def test_parse_bullet_list(self):
        """Parse bullet list items."""
        content = """- First item
- Second item
- Third item
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        bullet_elements = [e for e in model.elements if e.element_type == ElementType.BULLET_LIST]
        assert len(bullet_elements) == 3

    def test_parse_nested_bullet_list_indent(self):
        """Nested bullet indentation should map to indent levels."""
        content = """- Parent
  - Child
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        bullet_elements = [e for e in model.elements if e.element_type == ElementType.BULLET_LIST]
        assert bullet_elements[0].content.indent_level == 0
        assert bullet_elements[1].content.indent_level == 1

    def test_parse_blockquote(self):
        """Parse blockquote as callout."""
        content = """> [시사점] This is an important insight.
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        blockquotes = [e for e in model.elements if e.element_type == ElementType.BLOCKQUOTE]
        assert len(blockquotes) == 1
        assert isinstance(blockquotes[0].content, Blockquote)
        assert blockquotes[0].content.title == "시사점"

    def test_parse_latex_block(self):
        """Parse LaTeX block equation."""
        content = """$$ E = mc^2 $$
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        latex_blocks = [e for e in model.elements if e.element_type == ElementType.LATEX_BLOCK]
        assert len(latex_blocks) == 1
        assert isinstance(latex_blocks[0].content, LaTeXEquation)
        assert "E = mc^2" in latex_blocks[0].content.expression

    def test_soft_wrapped_lines_merge_into_single_paragraph(self):
        """Soft-wrapped markdown lines should remain one paragraph."""
        content = """첫 문장입니다.
둘째 문장입니다.
셋째 문장입니다.
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        paragraphs = [e for e in model.elements if e.element_type == ElementType.PARAGRAPH]
        assert len(paragraphs) == 1
        assert isinstance(paragraphs[0].content, Paragraph)
        assert paragraphs[0].content.text == "첫 문장입니다. 둘째 문장입니다. 셋째 문장입니다."

    def test_two_trailing_spaces_do_not_force_hard_break_by_default(self):
        """Trailing spaces should not create hard break in default parser mode."""
        content = "첫 줄입니다.  \n둘째 줄입니다.\n"
        parser = MarkdownParser()
        model = parser.parse(content)

        paragraphs = [e for e in model.elements if e.element_type == ElementType.PARAGRAPH]
        assert len(paragraphs) == 1
        assert isinstance(paragraphs[0].content, Paragraph)
        assert "\n" not in paragraphs[0].content.text
        assert paragraphs[0].content.text == "첫 줄입니다. 둘째 줄입니다."

    def test_two_trailing_spaces_preserved_when_opted_in(self):
        """Legacy markdown hard break behavior can be enabled explicitly."""
        content = "첫 줄입니다.  \n둘째 줄입니다.\n"
        parser = MarkdownParser(preserve_trailing_double_space_break=True)
        model = parser.parse(content)

        paragraphs = [e for e in model.elements if e.element_type == ElementType.PARAGRAPH]
        assert len(paragraphs) == 1
        assert isinstance(paragraphs[0].content, Paragraph)
        assert "\n" in paragraphs[0].content.text
        assert paragraphs[0].content.text == "첫 줄입니다.\n둘째 줄입니다."

    def test_html_br_line_break_preserved_inside_paragraph(self):
        """Trailing <br> should become a hard line break inside paragraph."""
        content = """첫 줄입니다.<br>
둘째 줄입니다.
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        paragraphs = [e for e in model.elements if e.element_type == ElementType.PARAGRAPH]
        assert len(paragraphs) == 1
        assert isinstance(paragraphs[0].content, Paragraph)
        assert paragraphs[0].content.text == "첫 줄입니다.\n둘째 줄입니다."

    def test_excessive_internal_spaces_are_normalized_in_paragraphs(self):
        """Paragraph parser should normalize repeated spaces and parenthesis spacing."""
        content = """동서울터미널   ( 신세계동서울  PFV )
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        paragraphs = [e for e in model.elements if e.element_type == ElementType.PARAGRAPH]
        assert len(paragraphs) == 1
        assert isinstance(paragraphs[0].content, Paragraph)
        assert paragraphs[0].content.text == "동서울터미널 (신세계동서울 PFV)"


# ═══════════════════════════════════════════════════════════════════════════════
# INTEGRATION TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestIntegration:
    """Integration tests for full document parsing."""

    def test_full_document_parse(self):
        """Parse a complete markdown document."""
        content = """---
title: Test Report
company: Test Corp
---

# Executive Summary

This report analyzes **key metrics**.

## Financial Overview

| Metric | 2024 | 2025 |
|--------|------|------|
| Revenue | 100 | 120 |

> [시사점] Revenue growth is strong.

- Point one
- Point two
"""
        parser = MarkdownParser()
        model = parser.parse(content)

        assert model.metadata.title == "Test Report"
        assert model.metadata.company == "Test Corp"
        assert len(model.elements) > 0

        # Check we have various element types
        types = {e.element_type for e in model.elements}
        assert ElementType.HEADING_1 in types
        assert ElementType.HEADING_2 in types
        assert ElementType.PARAGRAPH in types
        assert ElementType.TABLE in types
        assert ElementType.BLOCKQUOTE in types
        assert ElementType.BULLET_LIST in types


def test_parse_markdown_file_infers_title_and_date_from_leading_content(tmp_path):
    """Without frontmatter, first H1 and leading bold metadata should backfill metadata."""
    md_path = Path(tmp_path) / "inferred_metadata.md"
    md_path.write_text(
        "# 일동제약 주식회사 수익성 변화 분석 보고서\n\n**분석 대상 기간:** 제9기→제10기\n\n**분석 기준:** 연결재무제표 기준\n\n**작성일:** 2026년 3월 20일\n\n## 본문\n\n내용\n",
        encoding="utf-8",
    )

    model = parse_markdown_file(str(md_path))

    assert model.metadata.title == "일동제약 주식회사 수익성 변화 분석 보고서"
    assert model.metadata.company == "Korea Development Bank"
    assert model.metadata.extra["subject_company"] == "일동제약 주식회사"
    assert model.metadata.extra["analysis_period"] == "제9기→제10기"
    assert model.metadata.extra["analysis_basis"] == "연결재무제표 기준"
    assert model.metadata.extra["date"] == "2026년 3월 20일"
