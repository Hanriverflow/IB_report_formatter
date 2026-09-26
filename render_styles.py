"""Immutable styles scoped to a render request; legacy helpers share no mutable state."""

from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Iterator

from docx.shared import Inches, Pt, RGBColor


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
