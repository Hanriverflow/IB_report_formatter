"""Term-sheet profile composition: house boilerplate, opening block and layout helpers.

Called from `IBDocumentRenderer.render` (the single composition path) when the
resolved profile is `term-sheet`. Design: docs/term-sheet-design-20260929.md.
This module depends only on the model, styles, YAML and python-docx primitives;
it must not import `ib_renderer` (the renderer injects run-rendering callbacks).

Changelog (A1 foundation):
    - Validate immutable house boilerplate with presence-based frontmatter precedence.

Changelog (A2 rendering):
    - Schema-ordered single cell fills shared with generic merged-table emission.
    - Per-line runs with marker hanging indents; row-split estimation; fixed label
      grid, table frame, caption/unit line, note, source and 4pt spacer.
"""

import math
import unicodedata
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple, Union

import yaml
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tc
from docx.shared import Emu, Mm, Pt, RGBColor, Twips

from document_model import DocumentMetadata, TextRun
from md_parser import TextParser
from render_styles import STYLE

RenderRuns = Callable[..., None]
"""Injected `TextRenderer.render_runs(paragraph, runs, default_color=, font_name=, font_size=)`."""

ROW_SPLIT_THRESHOLD = 12
CELL_MARGIN_VERTICAL = 70  # twips (3.5pt) above and below cell text
CELL_MARGIN_HORIZONTAL = 100  # twips (5pt) left and right of cell text
_BORDER_SIZE = "4"  # eighths of a point: 0.5pt
_BORDER_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")
_CELL_SPACE_AFTER = Pt(1.5)
_CELL_LINE_SPACING = 1.05
_SPACER_HEIGHT = Pt(4)
_MIN_CONTENT_WIDTH = Mm(30)
# Successor elements used to insert children in ECMA-376 schema order.
_TCPR_AFTER_SHD = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
    "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
)
_TBLPR_AFTER_IND = (
    "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
    "w:tblCaption", "w:tblDescription", "w:tblPrChange",
)
_TBLPR_AFTER_MARGINS = ("w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange")
_TRPR_AFTER_CANT_SPLIT = (
    "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del",
    "w:trPrChange",
)
_HOUSE_KEYS = frozenset({"prepared_by", "disclaimer", "confidential_label", "confirmation"})
_CONFIRMATION_KEYS = frozenset({"intro", "items", "signature"})


@dataclass(frozen=True)
class ConfirmationText:
    """Validated confirmation wording without mutable item collections."""

    intro: str = ""
    items: Tuple[str, ...] = ()
    signature: str = ""


@dataclass(frozen=True)
class TermSheetTexts:
    """Resolved boilerplate for one render; empty labels intentionally remain empty."""

    prepared_by: str
    disclaimer: str
    confidential_label: str = "Strictly Confidential"
    confirmation: Optional[ConfirmationText] = None


def _confirmation_text(value: Any) -> ConfirmationText:
    """Convert a supported confirmation mapping into an immutable payload."""
    if not isinstance(value, dict) or not value:
        raise ValueError("confirmation must be a mapping with intro, items or signature")
    unknown = set(value) - _CONFIRMATION_KEYS
    if unknown:
        raise ValueError("Unknown confirmation settings: " + ", ".join(sorted(map(str, unknown))))
    for key in ("intro", "signature"):
        if key in value and not isinstance(value[key], str):
            raise ValueError(f"confirmation.{key} must be a string")
    if "items" in value:
        items = value["items"]
        if (
            not isinstance(items, list)
            or not items
            or any(not isinstance(item, str) or not item.strip() for item in items)
        ):
            raise ValueError("confirmation.items must be a non-empty list of non-empty strings")
    return ConfirmationText(
        intro=value.get("intro", ""),
        items=tuple(value.get("items", [])),
        signature=value.get("signature", ""),
    )


def _validate_house_fields(data: Dict[str, Any]) -> None:
    """Reject unsupported keys and malformed values without requiring all fields."""
    unknown = set(data) - _HOUSE_KEYS
    if unknown:
        raise ValueError("Unknown house settings: " + ", ".join(sorted(map(str, unknown))))
    for key in ("prepared_by", "disclaimer", "confidential_label"):
        if key in data and not isinstance(data[key], str):
            raise ValueError(f"{key} must be a string")
    if "confirmation" in data:
        _confirmation_text(data["confirmation"])


def load_house(path: Union[str, Path]) -> Dict[str, Any]:
    """Load and validate the fields supplied by a house YAML file.

    Args:
        path: House file to read; callers resolve its source-relative location.

    Returns:
        Validated supplied fields, retaining which optional keys were present.

    Raises:
        ValueError: YAML syntax, mapping shape or a supplied field is invalid.
        OSError: The selected file cannot be read.
    """
    try:
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8-sig"))
    except yaml.YAMLError as exc:
        raise ValueError(f"Invalid house YAML: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("house must be a YAML mapping")
    _validate_house_fields(data)
    return data


def resolve_term_sheet_texts(
    metadata: DocumentMetadata, house_path: Optional[str] = None
) -> TermSheetTexts:
    """Resolve boilerplate before rendering, preserving explicit empty values.

    Args:
        metadata: Term-sheet metadata containing optional boilerplate overrides.
        house_path: Effective house path after caller/frontmatter precedence.
            Source-relative paths must already be absolute; absent paths use no file.

    Returns:
        Immutable texts for the renderer's current request.

    Raises:
        ValueError: A path is relative, a field is malformed, or required text is absent.
        OSError: The effective house file cannot be read.
    """
    house: Dict[str, Any] = {}
    if house_path is not None:
        if not isinstance(house_path, str) or not house_path.strip():
            raise ValueError("house must be a YAML file path")
        if not Path(house_path).is_absolute():
            raise ValueError("house path must be absolute for input without a source file")
        house = load_house(house_path)
    values = {
        key: metadata.extra[key] if key in metadata.extra else house[key]
        for key in _HOUSE_KEYS
        if key in metadata.extra or key in house
    }
    _validate_house_fields(values)
    for key in ("prepared_by", "disclaimer"):
        if key not in values or not values[key].strip():
            raise ValueError(f"term-sheet requires non-empty {key}")
    return TermSheetTexts(
        prepared_by=values["prepared_by"],
        disclaimer=values["disclaimer"],
        confidential_label=values.get("confidential_label", "Strictly Confidential"),
        confirmation=_confirmation_text(values["confirmation"])
        if "confirmation" in values else None,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE CELL PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════════


def set_cell_fill(tc: CT_Tc, hex_color: str) -> None:
    """Give one `w:tc` exactly one solid fill, placed in schema order.

    Args:
        tc: The cell element; a covered vertical-merge cell is a distinct `w:tc`.
        hex_color: Six-digit RGB fill without a leading `#`.
    """
    tc_pr = tc.get_or_add_tcPr()
    for shading in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(shading)
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), hex_color)
    tc_pr.insert_element_before(shading, *_TCPR_AFTER_SHD)


# ═══════════════════════════════════════════════════════════════════════════════
# LINES AND LIST-LIKE MARKERS
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class MarkerLayout:
    """Hanging indent for a line that starts with a list-like marker (plan §2-3).

    Word receives `left = start + hanging` and `first line = -hanging`, so a
    wrapped line aligns with the first character after the marker.
    """

    start_mm: float
    hanging_mm: float
    note: bool = False  # ※ lines also use the note text size


_MARKER_LAYOUTS = {
    "•": MarkerLayout(0.0, 3.0),
    "-": MarkerLayout(3.0, 3.0),
    "·": MarkerLayout(6.0, 3.0),
    "※": MarkerLayout(0.0, 4.0, note=True),
}
_CIRCLED_LAYOUT = MarkerLayout(0.0, 4.5)
_CIRCLED_NUMBERS = frozenset(chr(code) for code in range(0x2460, 0x2474))  # ①–⑳


def marker_layout(text: str) -> Optional[MarkerLayout]:
    """Return the layout for a line's leading marker; the text is never changed.

    A hyphen counts only when followed by whitespace, so negative amounts and a
    lone dash placeholder keep their ordinary layout.
    """
    stripped = text.lstrip()
    if not stripped:
        return None
    marker = stripped[0]
    if marker in _CIRCLED_NUMBERS:
        return _CIRCLED_LAYOUT
    if marker == "-" and not stripped[1:2].isspace():
        return None
    return _MARKER_LAYOUTS.get(marker)


def apply_marker_layout(paragraph: Any, text: str) -> Optional[Pt]:
    """Give a paragraph its marker hanging indent.

    Args:
        paragraph: python-docx paragraph that will hold exactly this line.
        text: The line's visible text, used only to detect the marker.

    Returns:
        The note size for `※` lines, otherwise None (keep the caller's size).
    """
    layout = marker_layout(text)
    if layout is None:
        return None
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = Mm(layout.start_mm + layout.hanging_mm)
    paragraph_format.first_line_indent = Mm(-layout.hanging_mm)
    return STYLE.TS_NOTE_SIZE if layout.note else None


def _is_plain(run: TextRun) -> bool:
    """Semantic runs (equations, footnote references) are never trimmed or split."""
    return not run.is_latex and run.footnote_id is None


def _strip_edge(line: List[TextRun], leading: bool) -> List[TextRun]:
    """Drop spaces and tabs beside a line break, as HTML does around `<br>`."""
    runs = list(line)
    index = 0 if leading else -1
    while runs and _is_plain(runs[index]):
        text = runs[index].text.lstrip(" \t") if leading else runs[index].text.rstrip(" \t")
        if text:
            runs[index] = replace(runs[index], text=text)
            break
        runs.pop(index)
    return runs


def split_run_lines(runs: Sequence[TextRun]) -> List[List[TextRun]]:
    """Split runs at line breaks into one run list per Word paragraph.

    Formatting, link targets, footnote references and term keys stay on every
    piece, and empty lines are kept as empty lists. Only spaces and tabs next to
    a break are dropped; the source runs are never mutated.

    Args:
        runs: Parsed runs whose text may contain newlines from `<br>` or hard breaks.

    Returns:
        At least one, possibly empty, run list.
    """
    lines: List[List[TextRun]] = [[]]
    for run in runs:
        if not _is_plain(run) or "\n" not in run.text:
            lines[-1].append(run)
            continue
        for index, piece in enumerate(run.text.split("\n")):
            if index:
                lines.append([])
            if piece:
                lines[-1].append(replace(run, text=piece))
    trimmed: List[List[TextRun]] = []
    for index, line in enumerate(lines):
        if index > 0:
            line = _strip_edge(line, leading=True)
        if index < len(lines) - 1:
            line = _strip_edge(line, leading=False)
        trimmed.append(line)
    return trimmed


def line_text(line: Sequence[TextRun]) -> str:
    """Visible text of one split line."""
    return "".join(run.text for run in line)


def _em_width(text: str) -> float:
    """Approximate advance in ems: East Asian wide/full-width 1.0, others 0.55."""
    return sum(1.0 if unicodedata.east_asian_width(char) in ("W", "F") else 0.55 for char in text)


def estimate_cell_lines(texts: Sequence[str], width: int, font_size: int, markers: bool) -> int:
    """Estimate one cell's wrapped line count for row pagination (plan §4 A2).

    Args:
        texts: One entry per cell paragraph, already split at line breaks.
        width: Merged cell width in EMU, including the cell margins.
        font_size: Body text size of the cell in EMU.
        markers: Whether marker lines are indented (content cells).

    Returns:
        Sum over paragraphs of ceil(text advance / inner width), at least one each.
    """
    inner = int(width) - 2 * int(Twips(CELL_MARGIN_HORIZONTAL))
    total = 0
    for text in texts:
        size, indent = int(font_size), 0
        layout = marker_layout(text) if markers else None
        if layout is not None:
            indent = int(Mm(layout.start_mm + layout.hanging_mm))
            if layout.note:
                size = int(STYLE.TS_NOTE_SIZE)
        available = max(inner - indent, 1)
        total += max(1, math.ceil(_em_width(text) * size / available))
    return total


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE LAYOUT
# ═══════════════════════════════════════════════════════════════════════════════


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color)


def key_value_widths(label_columns: int, available: int) -> Optional[List[int]]:
    """Fixed label grid for key-value tables, so their first vertical line is common.

    Args:
        label_columns: One or two leading label columns.
        available: Printable width of the table's own section in EMU (portrait
            or landscape); the content column receives the remainder.

    Returns:
        Column widths in EMU, or None when the section leaves no usable content width.
    """
    fixed = [int(STYLE.TS_LABEL_WIDTH)]
    if label_columns == 2:
        fixed.append(int(STYLE.TS_SUBLABEL_WIDTH))
    remaining = int(available) - sum(fixed)
    if remaining < int(_MIN_CONTENT_WIDTH):
        return None
    return fixed + [remaining]


def _measure(tag: str, twips: int) -> Any:
    element = OxmlElement(tag)
    element.set(qn("w:w"), str(twips))
    element.set(qn("w:type"), "dxa")
    return element


def apply_table_frame(word_table: Any, width: int) -> None:
    """Apply the term-sheet table frame in schema order.

    The grid is fixed and spans `width`. In compatibility mode 14 Word draws the
    left border at the table indent minus the left cell margin, so an indent equal
    to that margin puts both outer borders on the text margins. All six border
    edges are 0.5pt `TS_BORDER_HEX`; cell margins are 70 (top/bottom) and 100
    (left/right) twips.

    Args:
        word_table: python-docx table whose grid widths are already set.
        width: Total grid width in EMU.
    """
    word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    word_table.autofit = False
    tbl_pr = word_table._tbl.tblPr
    for tag in ("w:tblInd", "w:tblBorders", "w:tblCellMar"):
        for existing in tbl_pr.findall(qn(tag)):
            tbl_pr.remove(existing)
    table_width = tbl_pr.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        tbl_pr.insert_element_before(table_width, "w:jc", "w:tblCellSpacing", "w:tblInd", *_TBLPR_AFTER_IND)
    table_width.set(qn("w:w"), str(Emu(width).twips))
    table_width.set(qn("w:type"), "dxa")
    tbl_pr.insert_element_before(_measure("w:tblInd", CELL_MARGIN_HORIZONTAL), *_TBLPR_AFTER_IND)
    borders = OxmlElement("w:tblBorders")
    for edge in _BORDER_EDGES:
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), _BORDER_SIZE)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), STYLE.TS_BORDER_HEX)
        borders.append(border)
    tbl_pr.insert_element_before(borders, *_TBLPR_AFTER_IND[1:])
    margins = OxmlElement("w:tblCellMar")
    for edge, value in (
        ("top", CELL_MARGIN_VERTICAL), ("left", CELL_MARGIN_HORIZONTAL),
        ("bottom", CELL_MARGIN_VERTICAL), ("right", CELL_MARGIN_HORIZONTAL),
    ):
        margins.append(_measure(f"w:{edge}", value))
    tbl_pr.insert_element_before(margins, *_TBLPR_AFTER_MARGINS)


def configure_cell_paragraph(paragraph: Any) -> None:
    """Cell paragraph rhythm: no space before, 1.5pt after, 1.05 line spacing."""
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = _CELL_SPACE_AFTER
    paragraph_format.line_spacing = _CELL_LINE_SPACING


def set_row_pagination(word_row: Any, keep_together: bool, repeat_header: bool) -> None:
    """Add `w:cantSplit`/`w:tblHeader` to a row in schema order, only when needed."""
    if not keep_together and not repeat_header:
        return
    tr_pr = word_row._tr.get_or_add_trPr()
    if keep_together:
        tr_pr.insert_element_before(OxmlElement("w:cantSplit"), *_TRPR_AFTER_CANT_SPLIT)
    if repeat_header:
        tr_pr.insert_element_before(OxmlElement("w:tblHeader"), *_TRPR_AFTER_CANT_SPLIT[2:])


def add_table_heading(
    doc: Any, caption: str, unit: str, as_of: str, render_runs: RenderRuns, width: int
) -> None:
    """Caption at left and `(단위 : …, 기준일 : …)` at right, kept with the table.

    Args:
        doc: Document receiving the paragraph before the table.
        caption: Bold accent-coloured caption text; may be empty.
        unit: Unit label; may be empty.
        as_of: Reference date text; may be empty.
        render_runs: Injected run renderer.
        width: Printable width of the current section in EMU (right tab stop).
    """
    details = ", ".join(
        part for part in (f"단위 : {unit}" if unit else "", f"기준일 : {as_of}" if as_of else "") if part
    )
    if not caption and not details:
        return
    paragraph = doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
    paragraph.paragraph_format.keep_with_next = True
    if caption:
        runs = [replace(run, bold=True) for run in TextParser.parse_runs(caption)]
        render_runs(paragraph, runs, default_color=STYLE.NAVY, font_size=STYLE.BODY_SIZE)
    if details:
        if caption:
            paragraph.paragraph_format.tab_stops.add_tab_stop(Emu(width), WD_TAB_ALIGNMENT.RIGHT)
            render_runs(paragraph, [TextRun(text="\t")], font_size=STYLE.TS_NOTE_SIZE)
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        render_runs(
            paragraph, TextParser.parse_runs(f"({details})"),
            default_color=_rgb(STYLE.TS_MUTED_HEX), font_size=STYLE.TS_NOTE_SIZE,
        )


def add_table_note(doc: Any, note: str, render_runs: RenderRuns) -> None:
    """Right-aligned muted note directly below the table."""
    paragraph = doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    render_runs(
        paragraph, TextParser.parse_runs(note),
        default_color=_rgb(STYLE.TS_MUTED_HEX), font_size=STYLE.TS_NOTE_SIZE,
    )


def add_table_source(doc: Any, source: str, render_runs: RenderRuns) -> None:
    """Muted `출처 : …` line below the table (and its note)."""
    paragraph = doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
    render_runs(
        paragraph, TextParser.parse_runs("출처 : " + source),
        default_color=_rgb(STYLE.TS_MUTED_HEX), font_size=STYLE.TS_NOTE_SIZE,
    )


def add_table_spacer(doc: Any) -> None:
    """Empty paragraph of exactly 4pt after a table instead of a body-size line."""
    paragraph_format = doc.add_paragraph().paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    paragraph_format.line_spacing = _SPACER_HEIGHT
