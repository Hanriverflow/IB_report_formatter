"""
Tests for md_formatter module.

Tests the markdown pre-processor for single-line text conversion.
"""

import pytest

from deep_md_cleaner import MARKER_END, MARKER_SEPARATOR, MARKER_START, CleanerConfig
from md_formatter import (
    BoldProtector,
    ColonLabelDetector,
    LaTeXProtector,
    MetadataExtractor,
    StructureDetector,
    check_needs_formatting,
    format_markdown,
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
        text = "[보고서] 가나다전자 기업분석작성일: 2024-01-15"

        metadata, remaining = MetadataExtractor.extract(text)

        assert "title" in metadata
        assert "가나다전자" in metadata["title"]

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
            "분석 대상:주식회사 가나다전자 (KOSDAQ: 000000)\n"
            "보고서 유형:심층 기업 분석 및 산업 전략 리포트\n"
            "Disclaimer:본 보고서는 공개 가용 정보를 기반으로 한 분석 자료"
        )

        result = format_markdown(text)

        assert 'date: "2026년 2월 7일 (초판 2026.01.30 대비 전면 개정)"' in result
        assert "분석 대상: 주식회사 가나다전자" in result
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

    def test_light_format_keeps_markers_when_cleaner_off(self):
        """Light-format path should keep markers when cleaner is not enabled."""
        marker = f"{MARKER_START}cite{MARKER_SEPARATOR}turn1search0{MARKER_END}"
        text = "\n".join([f"line {i}" for i in range(25)]) + f"\nreference {marker}"

        result = format_markdown(text)
        assert marker in result

    def test_light_format_cleans_markers_when_cleaner_enabled(self):
        """Light-format path should clean markers when cleaner is enabled."""
        marker = f"{MARKER_START}cite{MARKER_SEPARATOR}turn1search0{MARKER_END}"
        text = "\n".join([f"line {i}" for i in range(25)]) + f"\nreference {marker}"

        result = format_markdown(
            text,
            cleaner_config=CleanerConfig(activation_mode="auto", cite_mode="strip"),
        )
        assert marker not in result
        assert "turn1search0" not in result


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


# ═══════════════════════════════════════════════════════════════════════════════
# AUDIT 2026-11-15 REGRESSIONS
# ═══════════════════════════════════════════════════════════════════════════════


def test_major_heading_replaces_its_number_instead_of_leaving_a_digit():
    """`1. Overview` became `1` + a heading: the first digit of the match was kept."""
    assert format_markdown("1. Overview") == "## 1. Overview"
    result = format_markdown("서론입니다. 1. 개요 본문입니다. 2. 배경 본문이다.")
    assert result == "서론입니다.\n\n## 1. 개요 본문입니다.\n\n## 2. 배경 본문이다."


def test_generated_frontmatter_keeps_backslashes_and_quotes():
    """A backslash in a double-quoted YAML scalar is an escape unless it is doubled."""
    import yaml

    title = 'C:\\temp "draft"'
    frontmatter = MetadataExtractor.to_frontmatter({"title": title, "수신": "a\\nb"})
    loaded = yaml.safe_load(frontmatter.strip("-\n"))
    assert loaded == {"title": title, "recipient": "a\\nb"}


def test_light_format_leaves_frontmatter_and_fenced_code_literal():
    """Prose rules (colon spacing, callout breaks, blank-line trimming) stop at literals."""
    literal = '---\ntitle: "a:b"\n---\n```text\na:b\n\n\n\n[참고] key:value\n```\n~~~\nx:y\n~~~'
    prose = "\n".join(f"항목{n}:값" for n in range(25))
    result = format_markdown(literal + "\n" + prose)
    assert result.startswith(literal + "\n")
    assert "항목0: 값" in result and "항목24: 값" in result
    # An unclosed fence runs to the end, as in the parser.
    unclosed = "\n".join(["line"] * 25) + "\n```\nk:v"
    assert format_markdown(unclosed).endswith("```\nk:v")
    tail = "\n" + "\n".join(["line"] * 25)
    # The parser's fences: any indentation; a backtick in the info string is no fence.
    assert format_markdown("- item\n    ```\n    a:b\n    ```" + tail).startswith(
        "- item\n    ```\n    a:b\n    ```\n"
    )
    assert format_markdown("``` a`b\nk:v\n```" + tail).startswith("``` a`b\nk: v\n")
    # A rule at the top followed by prose is not frontmatter, nor is a `...` closer.
    assert "본문: 값" in format_markdown("---\n\n본문:값\n\n---" + tail)
    assert "본문: 값" in format_markdown("---\n본문:값\n..." + tail)
    # A label rule that reaches across the line break does not pull the fence up.
    assert format_markdown("작성일:\n```\ncode:x\n```" + tail).startswith(
        "작성일:\n```\ncode:x\n```\n"
    )


def test_format_file_names_the_missing_path_as_given(tmp_path, monkeypatch):
    import md_formatter

    monkeypatch.setattr(md_formatter, "PARENT_DIR", tmp_path / "parent")
    monkeypatch.chdir(tmp_path)
    with pytest.raises(FileNotFoundError, match=r"^File not found: nowhere-xyz\.md$"):
        md_formatter.format_file_with_options("nowhere-xyz.md")


def test_format_file_reads_the_path_given_not_a_namesake_in_the_parent_folder(
    tmp_path, monkeypatch
):
    """`a/report.md` used to be replaced by `<project parent>/report.md` when both existed."""
    import md_formatter

    parent = tmp_path / "parent"
    work = tmp_path / "work"
    (work / "a").mkdir(parents=True)
    parent.mkdir()
    (parent / "report.md").write_text("WRONG\n", encoding="utf-8")
    (work / "a" / "report.md").write_text("RIGHT\n", encoding="utf-8")
    (work / "bare.md").write_text("CWD\n", encoding="utf-8")
    (parent / "bare.md").write_text("WRONG\n", encoding="utf-8")
    (parent / "only-parent.md").write_text("PARENT\n", encoding="utf-8")
    monkeypatch.setattr(md_formatter, "PARENT_DIR", parent)
    monkeypatch.chdir(work)

    def formatted(name: str) -> str:
        out = tmp_path / "out.md"
        md_formatter.format_file_with_options(name, str(out))
        return out.read_text(encoding="utf-8")

    assert formatted("a/report.md") == "RIGHT"
    assert formatted("bare.md") == "CWD"
    # A bare name that is not in the working directory still falls back.
    assert formatted("only-parent.md") == "PARENT"
