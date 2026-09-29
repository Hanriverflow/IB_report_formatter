"""Immutable styles scoped to a render request; legacy helpers share no mutable state.

Changelog (raster fonts):
    - Resolve installed image fonts independently of Word font declarations.
    - Cache font inventory/coverage only; collect glyph-loss diagnostics per render.

Changelog (term-sheet foundation):
    - Add immutable, theme-configurable term-sheet colors, widths and text sizes.
    - Add the confirmation signature-row fill (TS_SIGNATURE_BG_HEX).
"""

import logging
import re
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from functools import lru_cache
from typing import Iterator

from docx.shared import Inches, Pt, RGBColor

logger = logging.getLogger(__name__)

_RASTER_FONT_DIAGNOSTICS: ContextVar[list[str] | None] = ContextVar(
    "raster_font_diagnostics", default=None,
)


@lru_cache(maxsize=128)
def _installed_raster_font(family: str) -> str | None:
    """Look up an installed family without Matplotlib's silent default fallback."""
    from matplotlib import font_manager

    try:
        return str(font_manager.findfont(
            font_manager.FontProperties(family=[family]), fallback_to_default=False,
        ))
    except (ValueError, OSError) as exc:
        logger.debug("Raster font %s is unavailable: %s", family, exc)
        return None


@lru_cache(maxsize=128)
def _raster_font_codepoints(path: str) -> frozenset[int]:
    """Cache actual glyph coverage, including custom theme fonts with Latin names."""
    from matplotlib.ft2font import FT2Font

    try:
        return frozenset(FT2Font(path).get_charmap())
    except (OSError, RuntimeError, ValueError) as exc:
        logger.debug("Cannot inspect raster font %s: %s", path, exc)
        return frozenset()


class RasterFontPolicy:
    """Choose installed fonts for images; never alter a Word font declaration."""

    FALLBACKS = (
        "Malgun Gothic", "Apple SD Gothic Neo", "AppleGothic", "NanumGothic",
        "NanumBarunGothic", "Noto Sans CJK KR", "Noto Sans KR", "Source Han Sans KR", "UnDotum",
    )
    _CJK_RE = re.compile(
        "[\u1100-\u11ff\u2e80-\ua4cf\ua960-\ua97f\uac00-\ud7ff"
        "\uf900-\ufaff\ufe30-\ufe4f\uff65-\uffdc"
        "\U0001b000-\U0001b2ff\U00020000-\U000323af]"
    )

    @classmethod
    def resolve(cls, preferred: str, text: str) -> str:
        """Resolve a family for this raster operation and diagnose CJK glyph loss.

        Args:
            preferred: Current theme's family, supplied anew by each caller.
            text: Text actually drawn into the image.

        Returns:
            First installed candidate covering the text's CJK characters, or
            DejaVu Sans with a diagnostic when none can render those characters.
        """
        candidates = tuple(dict.fromkeys((preferred, *cls.FALLBACKS)))
        required = {ord(character) for character in cls._CJK_RE.findall(text)}
        for family in candidates:
            path = _installed_raster_font(family)
            if path is not None and (not required or required <= _raster_font_codepoints(path)):
                return family
        if required:
            message = (
                "No installed CJK raster font covers this image's text; missing glyphs. "
                "Install a font with the required CJK glyphs. Tried: " + ", ".join(candidates)
            )
            diagnostics = _RASTER_FONT_DIAGNOSTICS.get()
            if diagnostics is None or message not in diagnostics:
                logger.warning("%s", message)
                if diagnostics is not None:
                    diagnostics.append(message)
        return "DejaVu Sans"


@contextmanager
def collect_raster_font_diagnostics() -> Iterator[list[str]]:
    """Collect image glyph-loss messages only for the current render request.

    Yields:
        Unique warning messages for the orchestrator's visible/strict diagnostics.
    """
    diagnostics: list[str] = []
    token = _RASTER_FONT_DIAGNOSTICS.set(diagnostics)
    try:
        yield diagnostics
    finally:
        _RASTER_FONT_DIAGNOSTICS.reset(token)


@dataclass(frozen=True)
class IBStyle:
    """IB Bank styling constants"""

    # ── Colors (RGBColor) ───────────────────────────────────────────────────
    NAVY: RGBColor = RGBColor(0, 51, 102)
    DARK_GRAY: RGBColor = RGBColor(64, 64, 64)
    LIGHT_GRAY: RGBColor = RGBColor(245, 245, 245)
    ACCENT_BLUE: RGBColor = RGBColor(230, 240, 250)
    WHITE: RGBColor = RGBColor(255, 255, 255)
    RED: RGBColor = RGBColor(192, 0, 0)
    GREEN: RGBColor = RGBColor(0, 128, 0)
    ORANGE: RGBColor = RGBColor(255, 165, 0)
    MEDIUM_GRAY: RGBColor = RGBColor(128, 128, 128)
    CODE_BG: RGBColor = RGBColor(248, 249, 250)
    CHART_NEGATIVE_COLOR: RGBColor = RGBColor(192, 0, 0)

    # ── Colors (Hex for OOXML) ──────────────────────────────────────────────
    NAVY_HEX: str = "003366"
    LIGHT_GRAY_HEX: str = "F5F5F5"
    ACCENT_BLUE_HEX: str = "E6F0FA"
    GRAY_BORDER_HEX: str = "C8C8C8"
    YELLOW_HEX: str = "FFFF00"

    # ── Fonts ───────────────────────────────────────────────────────────────
    HEADING_FONT: str = "Arial"
    BODY_FONT: str = "Calibri"
    KOREAN_FONT: str = "Malgun Gothic"
    CODE_FONT: str = "Consolas"
    COVER_FONT: str = "Malgun Gothic"
    TOC_FONT: str = "Malgun Gothic"

    # ── Sizes ───────────────────────────────────────────────────────────────
    H1_SIZE: Pt = Pt(14)
    H2_SIZE: Pt = Pt(12)
    H3_SIZE: Pt = Pt(11)
    H4_SIZE: Pt = Pt(10.5)
    BODY_SIZE: Pt = Pt(10.5)
    SMALL_SIZE: Pt = Pt(9)
    TABLE_HEADER_SIZE: Pt = Pt(10)
    TABLE_BODY_SIZE: Pt = Pt(10)

    # ── Term-sheet presentation ─────────────────────────────────────────────
    TS_LABEL_BG_HEX: str = "F2F5FC"
    TS_BORDER_HEX: str = "9AA5C4"
    TS_MUTED_HEX: str = "555555"
    TS_CONFIDENTIAL_HEX: str = "888888"  # header label and footer text
    TS_SIGNATURE_BG_HEX: str = "F2F2F2"  # confirmation box signature row
    TS_LABEL_WIDTH: Inches = Inches(33.5 / 25.4)
    TS_SUBLABEL_WIDTH: Inches = Inches(30 / 25.4)
    TS_TITLE_SIZE: Pt = Pt(20)
    TS_SUBTITLE_SIZE: Pt = Pt(16)
    TS_META_SIZE: Pt = Pt(10)  # opening date/prepared_by; confirmation items/signature
    TS_NOTE_SIZE: Pt = Pt(8)
    TS_DISCLAIMER_SIZE: Pt = Pt(7)
    TS_HEADER_FOOTER_SIZE: Pt = Pt(7.5)

    # ── Spacing ─────────────────────────────────────────────────────────────
    H1_SPACE_BEFORE: Pt = Pt(18)
    H1_SPACE_AFTER: Pt = Pt(6)
    H2_SPACE_BEFORE: Pt = Pt(12)
    H2_SPACE_AFTER: Pt = Pt(4)
    H3_SPACE_BEFORE: Pt = Pt(10)
    H3_SPACE_AFTER: Pt = Pt(2)
    BODY_SPACE_AFTER: Pt = Pt(8)
    BULLET_SPACE_AFTER: Pt = Pt(4)

    # ── Line spacing ────────────────────────────────────────────────────────
    BODY_LINE_SPACING: float = 1.15

    # ── Margins ─────────────────────────────────────────────────────────────
    TOP_MARGIN: Inches = Inches(1.0)
    BOTTOM_MARGIN: Inches = Inches(0.75)
    LEFT_MARGIN: Inches = Inches(1.0)
    RIGHT_MARGIN: Inches = Inches(0.8)

    # ── Bullet ──────────────────────────────────────────────────────────────
    BULLET_INDENT: Inches = Inches(0.25)
    DEEP_LIST_INDENT: Inches = Inches(0.125)
    FULL_LIST_INDENT_LEVELS: int = 4
    MAX_LIST_INDENT: Inches = Inches(1.5)
    BULLET_CHAR: str = "■"

    # ── Custom Style Names ──────────────────────────────────────────────────
    STYLE_IB_BODY: str = "IB Body"
    STYLE_IB_BULLET: str = "IB Bullet"
    STYLE_TABLE_GRID: str = "Table Grid"

    TABLE_HEADER_BG: str = "003366"
    TABLE_HEADER_COLOR: RGBColor = RGBColor(255, 255, 255)
    TABLE_ZEBRA: bool = True
    BODY_JUSTIFY: bool = True
    HEADING_BORDER: bool = True
    NATIVE_NUMBERING: bool = False
    TOC_TITLE: str = "TABLE OF CONTENTS"
    PAGE_LABEL: str = "Page "
    PAGE_OF_LABEL: str = " of "


_DEFAULT_STYLE = IBStyle()
_ACTIVE_STYLE: ContextVar[IBStyle] = ContextVar("document_style", default=_DEFAULT_STYLE)


class _StyleAccess:
    """Backward-compatible read-only access to the current render style."""

    __slots__ = ()

    def __getattr__(self, name: str):
        return getattr(_ACTIVE_STYLE.get(), name)


STYLE = _StyleAccess()


@contextmanager
def use_style(style: IBStyle) -> Iterator[None]:
    """Apply a style to this context and restore it even if rendering fails."""
    token = _ACTIVE_STYLE.set(style)
    try:
        yield
    finally:
        _ACTIVE_STYLE.reset(token)
