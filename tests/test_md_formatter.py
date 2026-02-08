"""
Tests for md_formatter module.

Tests the markdown pre-processor for single-line text conversion.
"""

import pytest
from md_formatter import (
    LaTeXProtector,
    BoldProtector,
    MetadataExtractor,
    ColonLabelDetector,
    StructureDetector,
    format_markdown,
    check_needs_formatting,
)


# ═══════════════════════════════════════════════════════════════════════════════
# LATEX PROTECTOR TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestLaTeXProtector:
    """Tests for LaTeX equation protection."""

    def test_protect_block_equation(self):
        """Protect block equations ($$...$$)."""
        protector = LaTeXProtector()
        text = "Before $$E = mc^2$$ after"

        protected = protector.protect(text)

        assert "$$" not in protected
        assert "__LATEX_BLOCK_" in protected

    def test_protect_inline_equation(self):
        """Protect inline equations ($...$)."""
        protector = LaTeXProtector()
        text = "The formula $x^2 + y^2$ is simple"

        protected = protector.protect(text)

        assert "$x^2 + y^2$" not in protected
        assert "__LATEX_INLINE_" in protected

    def test_restore_equations(self):
        """Restore protected equations."""
        protector = LaTeXProtector()
        original = "Block: $$a^2$$ and inline: $b^2$"

        protected = protector.protect(original)
        restored = protector.restore(protected)

        assert "$$a^2$$" in restored
        assert "$b^2$" in restored


# ═══════════════════════════════════════════════════════════════════════════════
# BOLD PROTECTOR TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestBoldProtector:
    """Tests for bold text protection."""

    def test_protect_bold_text(self):
        """Protect bold markers from sentence splitting."""
        protector = BoldProtector()
        text = "This is **very important** text"

        protected = protector.protect(text)

        assert "**" not in protected
        assert "__BOLD_" in protected

    def test_restore_bold_text(self):
        """Restore bold markers."""
        protector = BoldProtector()
        original = "Some **bold** and **more bold** text"

        protected = protector.protect(original)
        restored = protector.restore(protected)

        assert restored == original


# ═══════════════════════════════════════════════════════════════════════════════
# METADATA EXTRACTOR TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestMetadataExtractor:
    """Tests for report metadata extraction."""

    def test_extract_title(self):
        """Extract report title from [보고서] marker."""
        text = "[보고서] 네페스 기업분석작성일: 2024-01-15"

        metadata, remaining = MetadataExtractor.extract(text)

        assert "title" in metadata
        assert "네페스" in metadata["title"]

    def test_extract_metadata_fields(self):
        """Extract standard metadata fields."""
        text = "작성일: 2024-01-15수신: 투자본부주제: 기업분석"

        metadata, _ = MetadataExtractor.extract(text)

        assert "작성일" in metadata
        assert "수신" in metadata
        assert "주제" in metadata

    def test_to_frontmatter(self):
        """Convert metadata to YAML frontmatter."""
        metadata = {
            "title": "Test Report",
            "주제": "Analysis",
            "작성일": "2024-01-15",
        }

        frontmatter = MetadataExtractor.to_frontmatter(metadata)

        assert "---" in frontmatter
        assert 'title: "Test Report"' in frontmatter
        assert 'subtitle: "Analysis"' in frontmatter


# ═══════════════════════════════════════════════════════════════════════════════
# STRUCTURE DETECTOR TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestStructureDetector:
    """Tests for structure detection and insertion."""

    def test_detect_subsection_heading(self):
        """Detect sub-section headings (2.1. Title)."""
        text = "content.2.1. 시장 분석more content"

        result = StructureDetector.insert_structure(text)

        assert "### 2.1. 시장 분석more content" in result
        assert "## 1." not in result

    def test_detect_subsubsection(self):
        """Detect sub-sub-section headings (3.1.1. Title)."""
        text = "content3.1.1. 세부 분석more"

        result = StructureDetector.insert_structure(text)

        assert "#### 3.1.1. 세부 분석more" in result
        assert "### 1.1" not in result
        assert "## 1." not in result

    def test_detect_callout(self):
        """Detect callout labels."""
        text = "content[시사점] Important insight"

        result = StructureDetector.insert_structure(text)

        assert "> [시사점]" in result


# ═══════════════════════════════════════════════════════════════════════════════
# FORMAT MARKDOWN TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestFormatMarkdown:
    """Tests for main format_markdown function."""

    def test_already_formatted_passthrough(self):
        """Already formatted text gets light formatting only."""
        text = (
            """# Heading 1

Paragraph one.

## Heading 2

Paragraph two.
"""
            + "\n" * 20
        )  # Ensure line count > 20

        result = format_markdown(text)

        # Should not drastically change
        assert "# Heading 1" in result or "Heading 1" in result

    def test_single_line_gets_structured(self):
        """Single-line text gets structure added."""
        text = "1. 서론이 보고서는 분석을 제공합니다.2. 본론주요 내용입니다."

        result = format_markdown(text)

        # Should have newlines added
        assert result.count("\n") > text.count("\n")

    def test_colon_spacing_normalized_for_labels(self):
        """Known labels should normalize to ': ' spacing."""
        text = "보고서 작성일:2026년 2월 7일 가정:연 매출 2조 원"

        normalized = ColonLabelDetector.normalize_colon_spacing(text)
        assert "보고서 작성일: 2026년 2월 7일" in normalized
        assert "가정: 연 매출 2조 원" in normalized

    def test_colon_spacing_preserved_through_format(self):
        """Formatter output should keep one-space label separators."""
        text = "가정:연 매출 2조 원트리거 요건:DSCR 1.2x"

        result = format_markdown(text)

        assert "가정: 연 매출 2조 원" in result
        assert "트리거 요건: DSCR 1.2x" in result

    def test_colon_spacing_preserves_bold_markers(self):
        """Bold markers with trailing colons should not be altered."""
        text = "\n".join([f"line {i}" for i in range(21)])
        text += "\n**보고서 작성일:** 2026년 2월 7일 **분석 대상:** 주식회사"

        result = format_markdown(text)

        assert "**보고서 작성일:**" in result or "**보고서 작성일:** " in result
        assert "**보고서 작성일: **" not in result
        assert "**분석 대상:**" in result or "**분석 대상:** " in result
        assert "**분석 대상: **" not in result

    def test_colon_spacing_global_real_world_lines(self):
        """Global colon normalization should cover common report lines."""
        text = (
            "보고서 작성일:2026년 2월 7일 (초판 2026.01.30 대비 전면 개정)\n"
            "분석 대상:주식회사 네패스 (KOSDAQ: 033640)\n"
            "보고서 유형:심층 기업 분석 및 산업 전략 리포트\n"
            "Disclaimer:본 보고서는 공개 가용 정보를 기반으로 한 분석 자료"
        )

        result = format_markdown(text)

        assert 'date: "2026년 2월 7일 (초판 2026.01.30 대비 전면 개정)"' in result
        assert "분석 대상: 주식회사 네패스" in result
        assert "보고서 유형: 심층 기업 분석" in result
        assert "Disclaimer: 본 보고서는" in result

    def test_colon_spacing_global_safety_exceptions(self):
        """Do not break URL, drive path, time, ratio, or double-colon tokens."""
        text = "URL:https://example.com API::v1 시간 10:30 경로 C:\\temp 비율 1:1 라벨:값"

        result = ColonLabelDetector.normalize_colon_spacing_global(text)

        assert "https://example.com" in result
        assert "API::v1" in result
        assert "10:30" in result
        assert "C:\\temp" in result
        assert "1:1" in result
        assert "라벨: 값" in result


class TestCheckNeedsFormatting:
    """Tests for format check function."""

    def test_single_line_needs_formatting(self, tmp_path):
        """Single-line file needs formatting."""
        file = tmp_path / "test.md"
        file.write_text("A" * 500, encoding="utf-8")  # Long single line

        assert check_needs_formatting(str(file)) is True

    def test_multiline_no_formatting(self, tmp_path):
        """Multi-line file does not need formatting."""
        file = tmp_path / "test.md"
        content = "\n".join(["Line " + str(i) for i in range(50)])
        file.write_text(content, encoding="utf-8")

        assert check_needs_formatting(str(file)) is False
