"""Font styling for rendered runs and font choice for rasterized text."""

import logging
import platform

from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from render_styles import STYLE

logger = logging.getLogger("ib_renderer")


# ═══════════════════════════════════════════════════════════════════════════════
# FONT STYLER
# ═══════════════════════════════════════════════════════════════════════════════


class FontStyler:
    """Handles font styling including East Asian fonts"""

    @staticmethod
    def set_east_asian_font(element, font_name: str | None = None):
        """Set East Asian font (for Korean text) on a style or run element"""
        resolved_font = font_name or FontPolicy.resolve_korean_font()
        elm = element._element
        rPr = elm.get_or_add_rPr()
        if rPr.rFonts is None:
            rPr.get_or_add_rFonts()
        rPr.rFonts.set(qn("w:eastAsia"), resolved_font)
        # Theme attributes take precedence over explicit fonts in Word.
        for explicit, theme in (
            ("ascii", "asciiTheme"),
            ("hAnsi", "hAnsiTheme"),
            ("eastAsia", "eastAsiaTheme"),
        ):
            if rPr.rFonts.get(qn("w:" + explicit)):
                rPr.rFonts.attrib.pop(qn("w:" + theme), None)

    @staticmethod
    def apply_run_style(
        run,
        font_name: str | None = None,
        font_size: Pt | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        color: RGBColor | None = None,
        superscript: bool | None = None,
    ):
        """Apply styling to a run"""
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if color:
            run.font.color.rgb = color
        if superscript is not None:
            run.font.superscript = superscript
        FontStyler.set_east_asian_font(run, font_name=font_name)


class FontPolicy:
    """Declare Word fonts for the reader's machine, independently of raster fonts."""

    @classmethod
    def resolve_korean_font(cls, system_name: str | None = None) -> str:
        """Return the preferred Korean font for the current platform."""
        system = system_name or platform.system() or "Unknown"
        if system == "Darwin":
            candidates = ("Apple SD Gothic Neo", STYLE.KOREAN_FONT, "NanumGothic")
        elif system == "Windows":
            candidates = (STYLE.KOREAN_FONT, "NanumGothic", "Apple SD Gothic Neo")
        else:
            candidates = (STYLE.KOREAN_FONT, "NanumGothic", "Apple SD Gothic Neo")

        chosen = candidates[0]
        logger.debug(
            "Resolved Korean font for %s: %s (fallbacks: %s)",
            system,
            chosen,
            ", ".join(candidates[1:]) or "none",
        )
        return chosen
