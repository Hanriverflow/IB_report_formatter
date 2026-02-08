"""
Tests for ib_renderer module.

Tests rendering functionality including:
- Table number formatting
- Callout box styling
- Image rendering
"""

import pytest


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE RENDERER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestTableRendererFormatting:
    """Tests for TableRenderer._format_financial_number static method."""

    @pytest.fixture
    def format_number(self):
        """Import the format method."""
        from ib_renderer import TableRenderer

        return TableRenderer._format_financial_number

    def test_format_plain_number(self, format_number):
        """Format plain integers with thousand separators."""
        assert format_number("1234567") == "1,234,567"
        assert format_number("1000") == "1,000"
        assert format_number("999") == "999"
        assert format_number("100") == "100"

    def test_format_decimal_number(self, format_number):
        """Format decimal numbers."""
        assert format_number("1234.56") == "1,234.56"
        assert format_number("1000000.99") == "1,000,000.99"

    def test_format_negative_parentheses(self, format_number):
        """Format negative numbers in parentheses."""
        assert format_number("(1234567)") == "(1,234,567)"
        assert format_number("(500)") == "(500)"

    def test_format_negative_minus(self, format_number):
        """Format negative numbers with minus sign."""
        assert format_number("-1234567") == "-1,234,567"
        assert format_number("-500") == "-500"

    def test_format_with_suffix(self, format_number):
        """Format numbers with Korean unit suffixes."""
        assert format_number("1234567억") == "1,234,567억"
        assert format_number("100조") == "100조"
        assert format_number("50%") == "50%"
        assert format_number("12.5%") == "12.5%"

    def test_skip_already_formatted(self, format_number):
        """Skip numbers already with thousand separators."""
        assert format_number("1,234,567") == "1,234,567"
        assert format_number("1,000") == "1,000"

    def test_skip_non_numeric(self, format_number):
        """Skip non-numeric content."""
        assert format_number("Revenue") == "Revenue"
        assert format_number("N/A") == "N/A"
        assert format_number("") == ""
        assert format_number("-") == "-"


# ═══════════════════════════════════════════════════════════════════════════════
# CALLOUT RENDERER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestCalloutStyles:
    """Tests for CalloutRenderer style configurations."""

    def test_callout_style_exists(self):
        """Verify callout styles are defined."""
        from ib_renderer import CalloutRenderer

        styles = CalloutRenderer._CALLOUT_STYLES

        # Check key styles exist
        assert "EXECUTIVE SUMMARY" in styles
        assert "요약" in styles
        assert "KEY INSIGHT" in styles
        assert "시사점" in styles
        assert "WARNING" in styles
        assert "주의" in styles
        assert "NOTE" in styles
        assert "참고" in styles

    def test_executive_summary_has_navy_background(self):
        """Executive Summary should have navy background."""
        from ib_renderer import CalloutRenderer, STYLE

        style = CalloutRenderer._CALLOUT_STYLES["EXECUTIVE SUMMARY"]
        bg_hex, _, _, _ = style

        assert bg_hex == STYLE.NAVY_HEX


# ═══════════════════════════════════════════════════════════════════════════════
# IMAGE RENDERER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestImageRendererMimeTypes:
    """Tests for ImageRenderer MIME type handling."""

    def test_mime_to_extension(self):
        """Convert MIME types to file extensions."""
        from ib_renderer import ImageRenderer

        assert ImageRenderer._mime_to_extension("image/png") == ".png"
        assert ImageRenderer._mime_to_extension("image/jpeg") == ".jpg"
        assert ImageRenderer._mime_to_extension("image/gif") == ".gif"
        assert ImageRenderer._mime_to_extension("image/webp") == ".webp"
        assert ImageRenderer._mime_to_extension("image/svg+xml") == ".svg"

    def test_unknown_mime_defaults_to_png(self):
        """Unknown MIME types default to .png."""
        from ib_renderer import ImageRenderer

        assert ImageRenderer._mime_to_extension("image/unknown") == ".png"
        assert ImageRenderer._mime_to_extension("application/pdf") == ".png"


# ═══════════════════════════════════════════════════════════════════════════════
# LATEX RENDERER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestLaTeXRenderer:
    """Tests for LaTeXRenderer."""

    def test_latex_availability_check(self):
        """Check matplotlib availability detection."""
        from ib_renderer import LaTeXRenderer

        # Should return True or False, not raise
        result = LaTeXRenderer.is_available()
        assert isinstance(result, bool)

    @pytest.mark.skipif(
        not __import__("importlib.util", fromlist=[""]).find_spec("matplotlib"),
        reason="matplotlib not installed",
    )
    def test_render_simple_equation(self):
        """Render a simple LaTeX equation."""
        from ib_renderer import LaTeXRenderer
        import os

        if not LaTeXRenderer.is_available():
            pytest.skip("matplotlib not available")

        path = LaTeXRenderer.render_to_image("x^2 + y^2 = z^2")

        assert path is not None
        assert os.path.exists(path)
        assert path.endswith(".png")

        # Cleanup
        os.unlink(path)


# ═══════════════════════════════════════════════════════════════════════════════
# STYLE CONFIGURATION TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestIBStyle:
    """Tests for IBStyle configuration."""

    def test_style_colors_defined(self):
        """Verify essential colors are defined."""
        from ib_renderer import STYLE

        assert STYLE.NAVY is not None
        assert STYLE.DARK_GRAY is not None
        assert STYLE.WHITE is not None
        assert STYLE.RED is not None

    def test_style_fonts_defined(self):
        """Verify fonts are defined."""
        from ib_renderer import STYLE

        assert STYLE.HEADING_FONT == "Arial"
        assert STYLE.BODY_FONT == "Calibri"
        assert STYLE.KOREAN_FONT == "Malgun Gothic"

    def test_style_hex_colors(self):
        """Verify hex color strings are valid."""
        from ib_renderer import STYLE

        # Should be 6-character hex strings
        assert len(STYLE.NAVY_HEX) == 6
        assert len(STYLE.LIGHT_GRAY_HEX) == 6
        assert all(c in "0123456789ABCDEF" for c in STYLE.NAVY_HEX.upper())


# ═══════════════════════════════════════════════════════════════════════════════
# DISCLAIMER RENDERER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestDisclaimerRendererFormatting:
    """Tests for disclaimer paragraph formatting behavior."""

    def test_split_content_lines_removes_blank_lines(self):
        """Split helper should keep only non-empty lines."""
        from ib_renderer import DisclaimerRenderer

        content = "첫 문장\n\n둘째 문장\n  \n셋째 문장"
        lines = DisclaimerRenderer._split_content_lines(content)

        assert lines == ["첫 문장", "둘째 문장", "셋째 문장"]

    def test_disclaimer_content_uses_plain_paragraph_text(self):
        """Disclaimer content should not inject manual newline characters."""
        from docx import Document
        from ib_renderer import DisclaimerRenderer, DocumentStyler

        doc = Document()
        DocumentStyler(doc).create_styles()
        renderer = DisclaimerRenderer(doc)
        renderer.render("Korea Development Bank")

        target_lines = [
            "본 자료는 해당 문서에 최대한 정확하고 완전한 정보를 담고자 노력하였으나,",
            "본 자료는 당행의 저작물로서 모든 저작권은 당행에게 있으며, 당행의 동의 없이",
            "본 제안서의 내용은 현재의 시장상황 및 발행구조에 대한 기초정보에 근거한 것으로",
        ]

        for line in target_lines:
            paragraph = next((p for p in doc.paragraphs if line in p.text), None)
            assert paragraph is not None
            assert "\n" not in paragraph.text
