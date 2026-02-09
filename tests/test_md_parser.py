"""
Tests for md_parser module.

Tests core parsing functionality including:
- Frontmatter parsing
- Element type detection
- Table parsing
- Text run formatting
"""

import pytest
from md_parser import (
    FrontmatterParser,
    TextParser,
    TableParser,
    MarkdownParser,
    DocumentMetadata,
    Blockquote,
    ElementType,
    LaTeXEquation,
    TableType,
    TextRun,
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
