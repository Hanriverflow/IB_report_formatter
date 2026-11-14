"""Table rendering for every profile, including spans, financial formatting and term-sheet tables."""

import logging
import re
from collections.abc import Callable, Set
from dataclasses import replace
from typing import Any

from docx.document import Document as DocxDocument
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor

from ib_render_document_style import TableStyler
from ib_render_fonts import FontStyler
from ib_render_text import TextRenderer, caption_prefix
from md_parser import Table, TableCell, TableRow, TableType, TextParser, TextRun
from render_styles import STYLE
from term_sheet import (
    ROW_SPLIT_THRESHOLD,
    TermSheetStyle,
    add_table_heading,
    add_table_note,
    add_table_source,
    add_table_spacer,
    apply_marker_layout,
    apply_table_frame,
    configure_cell_paragraph,
    estimate_cell_lines,
    header_token_width,
    key_value_widths,
    label_widths,
    line_text,
    natural_cell_width,
    positive_printable_width,
    set_cell_fill,
    set_row_pagination,
    split_run_lines,
)
from term_variables import term_cell_text

logger = logging.getLogger("ib_renderer")


_BLACK = RGBColor(0, 0, 0)


class TableRenderer:
    """Renders tables with type-specific styling"""

    _EMUS_PER_INCH = 914400
    _COLUMN_SAMPLE_SIZE = 3
    _TOKEN_SPLIT_RE = re.compile(r"\s+")
    _NUMERIC_LIKE_RE = re.compile(
        r"^\(?[+-]?\d[\d,]*(?:\.\d+)?\)?\s*"
        r"(?:%|bp|bps|x|배|원|천원|만원|백만원|억원|억|조|주|개|명|건)?$",
        re.IGNORECASE,
    )
    _PURE_NUMBER_RE = re.compile(r"^\(?[+-]?\d[\d,]*(?:\.\d+)?\)?$")
    # A dash placeholder ("nothing this period") says nothing about a column's kind.
    _DASH_PLACEHOLDER_RE = re.compile(r"[-‒–—―－]+")
    _MIN_COLUMN_WIDTH_INCHES = 0.65
    _MIN_TEXT_COLUMN_WIDTH_INCHES = 1.15
    _MAX_NUMERIC_COLUMN_WIDTH_INCHES = 1.35
    _MAX_TEXT_COLUMN_SHARE = 0.55

    def __init__(
        self,
        doc: DocxDocument,
        term_sheet: bool = False,
        house_style: TermSheetStyle | None = None,
    ):
        self.doc = doc
        # Profile-only layout; merge emission itself is shared by every profile.
        self.term_sheet = term_sheet
        self.house_style = house_style or TermSheetStyle()

    def render(self, table: Table):
        """Render a table"""
        if not table.rows or table.col_count == 0:
            return

        row_count = len(table.rows)
        col_count = table.col_count

        previous_geometry = self._begin_landscape(table)
        if self.term_sheet:
            try:
                self._render_term_sheet_table(table)
            finally:
                # A diagnosed failure must not leave later content in landscape.
                self._end_landscape(previous_geometry)
            return
        for text, prefix in [
            (table.caption, caption_prefix(table)),
            (
                " · ".join(
                    v
                    for v in [
                        "단위: " + table.unit if table.unit else "",
                        "기준일: " + table.as_of if table.as_of else "",
                    ]
                    if v
                ),
                [],
            ),
        ]:
            if text:
                paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
                paragraph.paragraph_format.keep_with_next = True
                TextRenderer.render_runs(
                    paragraph, [*prefix, *TextParser.parse_runs(text)], font_size=STYLE.SMALL_SIZE
                )
        word_table = self.doc.add_table(rows=row_count, cols=col_count)
        word_table.style = STYLE.STYLE_TABLE_GRID
        column_kinds = self._infer_column_kinds(table)
        self._apply_column_widths(word_table, table, column_kinds)
        # Merge the empty, sized grid so spanned widths add up and no content
        # is concatenated; each owner cell is then filled exactly once.
        rectangles = self._merge_cells(word_table, table)
        covered = self._covered_cells(rectangles)

        # Render header rows
        header_rows = min(table.header_rows, len(table.rows))
        for h_idx in range(header_rows):
            self._render_header_row(word_table, table.rows[h_idx], col_count, covered, h_idx)

        # Render data rows based on table type
        for r_idx, row in enumerate(table.rows[header_rows:], header_rows):
            self._render_data_row(
                word_table,
                row,
                r_idx,
                col_count,
                table.table_type,
                column_kinds,
                table.column_types,
                table.alignments,
                covered,
                header_rows,
            )
        self._mirror_vertical_fills(word_table, rectangles)

        for h_idx in range(header_rows):
            word_table.rows[h_idx]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for word_row in word_table.rows:
            word_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        # Apply borders
        TableStyler.set_table_borders(word_table)

        if table.source:
            paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
            TextRenderer.render_runs(
                paragraph,
                TextParser.parse_runs("출처: " + table.source),
                font_size=STYLE.SMALL_SIZE,
            )
        if table.note:
            paragraph = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            TextRenderer.render_runs(
                paragraph, TextParser.parse_runs(table.note), font_size=STYLE.SMALL_SIZE
            )
        # Spacer paragraph after table
        self.doc.add_paragraph()
        self._end_landscape(previous_geometry)

    def _begin_landscape(self, table: Table) -> tuple[Any, int, int] | None:
        """Open a landscape section for a landscape table; return the prior geometry.

        The new section already starts on a new page, so a lone page break
        right before it (a cover's or TOC's) is redundant and is dropped. Word
        ignores such a break, but other viewers may show it as a blank page.
        """
        if not table.landscape:
            return None
        section = self.doc.sections[-1]
        assert section.page_width is not None and section.page_height is not None
        previous_geometry = (section.orientation, section.page_width, section.page_height)
        self._drop_trailing_page_break()
        section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Emu(max(previous_geometry[1:]))
        section.page_height = Emu(min(previous_geometry[1:]))
        return previous_geometry

    def _drop_trailing_page_break(self) -> None:
        """Remove the last body paragraph when it holds nothing but one page break.

        Several breaks in one paragraph ask for blank pages and are kept.
        """
        body = self.doc.element.body
        last = body[-1] if len(body) else None
        if last is not None and last.tag == qn("w:sectPr"):
            last = last.getprevious()
        if (
            last is not None
            and last.tag == qn("w:p")
            and len(last.xpath('./w:r/w:br[@w:type="page"]')) == 1
            and not last.xpath("./*[not(self::w:pPr or self::w:r)] | ./w:pPr/w:sectPr")
            and not last.xpath('./w:r/*[not(self::w:rPr or self::w:br[@w:type="page"])]')
        ):
            body.remove(last)

    def _end_landscape(self, previous_geometry: tuple[Any, int, int] | None) -> None:
        """Restore the page geometry that preceded a landscape table."""
        if previous_geometry:
            section = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
            section.orientation = previous_geometry[0]
            section.page_width = Emu(previous_geometry[1])
            section.page_height = Emu(previous_geometry[2])

    def _render_term_sheet_table(self, table: Table) -> None:
        """Render a term-sheet table through the shared merge emission.

        Key-value tables (one or two label columns plus one content column) use
        the fixed label grid so their first vertical line is common; other tables
        keep content-based widths. Each owner cell is filled once with one
        paragraph per line, and a row stays unsplit when its estimated height is
        at most `ROW_SPLIT_THRESHOLD` lines (the header row always).

        Args:
            table: Parsed table; `label_columns` is inferred by the parser.

        Raises:
            ValueError: The section has no positive printable width. Unlike the
                legacy estimate there is no three-inch floor, so a term-sheet table
                never exceeds its section.
        """
        row_count, col_count = len(table.rows), table.col_count
        available = positive_printable_width(self.doc.sections[-1], "Term-sheet table")
        add_table_heading(
            self.doc,
            table.caption,
            table.unit,
            table.as_of,
            TextRenderer.render_runs,
            available,
            prefix=caption_prefix(table),
        )
        word_table = self.doc.add_table(rows=row_count, cols=col_count)
        word_table.style = STYLE.STYLE_TABLE_GRID
        column_kinds = self._infer_column_kinds(table)
        label_columns = (
            table.label_columns
            if table.label_columns is not None
            else min(1, max(col_count - 1, 0))
        )
        key_value = label_columns in (1, 2) and col_count == label_columns + 1
        widths = key_value_widths(label_columns, available) if key_value else None
        if key_value and widths is None:
            needed_mm, section_mm = sum(label_widths(label_columns)) / 36000, available / 36000
            self._report(
                f"Term-sheet label columns need {needed_mm:.1f} mm but the section is only "
                f"{section_mm:.1f} mm wide; column widths estimated from content"
            )
        if widths is None:
            widths = self._term_sheet_grid_widths(table, available, column_kinds)
        self._set_column_widths(word_table, widths)
        rectangles = self._merge_cells(word_table, table)
        covered = self._covered_cells(rectangles)
        extents = {(top, left): (bottom, right) for top, left, bottom, right in rectangles}
        apply_table_frame(
            word_table, sum(widths), open_sides=self.house_style.table_sides == "open"
        )

        line_counts = [0] * row_count
        for row_index, row in enumerate(table.rows):
            word_cells = word_table.rows[row_index].cells
            for column_index in range(col_count):
                if (row_index, column_index) in covered:
                    continue
                bottom, right = extents.get((row_index, column_index), (row_index, column_index))
                cell_data = (
                    row.cells[column_index]
                    if column_index < len(row.cells)
                    else TableCell(content="")
                )
                role = self._term_sheet_role(
                    row_index,
                    column_index,
                    right,
                    label_columns,
                    key_value,
                    table.header_rows,
                )
                texts, size = self._fill_term_sheet_cell(
                    word_cells[column_index],
                    cell_data,
                    role,
                    row_index,
                    column_index,
                    table,
                    column_kinds,
                )
                if bottom == row_index:  # multi-row owners are excluded from row estimates
                    lines = estimate_cell_lines(
                        texts,
                        sum(widths[column_index : right + 1]),
                        size,
                        markers=role == "content",
                    )
                    line_counts[row_index] = max(line_counts[row_index], lines)
        self._mirror_vertical_fills(word_table, rectangles)
        for row_index, word_row in enumerate(word_table.rows):
            header = row_index < table.header_rows
            set_row_pagination(
                word_row,
                keep_together=header or line_counts[row_index] <= ROW_SPLIT_THRESHOLD,
                repeat_header=header,
            )
        if table.note:
            add_table_note(self.doc, table.note, TextRenderer.render_runs)
        if table.source:
            add_table_source(self.doc, table.source, TextRenderer.render_runs)
        add_table_spacer(self.doc)

    def _term_sheet_grid_widths(
        self, table: Table, available: int, column_kinds: list[str]
    ) -> list[int]:
        """Widths (EMU) for a term-sheet grid table, filling the section width.

        When the columns' content fits side by side on single lines, each column
        gets its natural width (see `_term_sheet_natural_widths`) plus an equal
        share of the spare width, so short amount columns are not starved while
        date or label columns balloon. Otherwise widths come from the content
        estimate, whose minimums keep header words on one line: each
        single-column header cell requires room for its longest unbreakable token
        in the bold header size plus the cell margins, and merged header cells
        impose no minimum (see `_required_minimums` for the graceful fallback).
        """
        natural = self._term_sheet_natural_widths(table, column_kinds)
        if natural is not None and sum(natural) <= available:
            spare = (available - sum(natural)) // table.col_count
            fitted = [width + spare for width in natural]
            fitted[-1] += available - sum(fitted)
            return fitted
        required = [0.0] * table.col_count
        for header in table.rows[: table.header_rows]:
            for index in range(min(len(header.cells), table.col_count)):
                if self._is_span_cell(header, index):
                    continue
                cell = header.cells[index]
                text = line_text(cell.runs or TextParser.parse_runs_plain(cell.content))
                width = header_token_width(text, STYLE.TABLE_HEADER_SIZE) / self._EMUS_PER_INCH
                required[index] = max(required[index], width)
        widths = self._estimate_column_widths(
            table, available / self._EMUS_PER_INCH, column_kinds, required
        )
        return [int(Inches(width)) for width in widths]

    def _term_sheet_natural_widths(self, table: Table, column_kinds: list[str]) -> list[int] | None:
        """Per-column widths (EMU) that keep every cell line on one line.

        Header lines are measured bold in the header size and body lines as
        displayed (formatted amounts, marker indents). A cell merged across
        columns spreads any extra width it needs evenly over them, and every
        column keeps at least the estimator's numeric minimum.

        Args:
            table: Term-sheet grid table.
            column_kinds: Inferred column kinds (unused by the measure itself,
                kept for symmetry with the content estimate).

        Returns:
            One width per column, or None when a cell holds an image, whose
            width text measurement cannot know.
        """
        natural = [0] * table.col_count
        spanning: list[tuple[int, int, int]] = []
        for row_index, row in enumerate(table.rows):
            header = row_index < table.header_rows
            for index in range(min(len(row.cells), table.col_count)):
                cell = row.cells[index]
                if cell.merge is not None:
                    continue
                right = index
                while (
                    right + 1 < min(len(row.cells), table.col_count)
                    and row.cells[right + 1].merge == "left"
                ):
                    right += 1
                if header:
                    runs = cell.runs or TextParser.parse_runs_plain(cell.content)
                    size, bold, markers = STYLE.TABLE_HEADER_SIZE, True, False
                else:
                    role = table.column_types[index] if table.column_types else None
                    content, runs = self._display_content(cell, role, table.table_type)
                    runs = runs or TextParser.parse_runs_plain(content)
                    size, bold, markers = STYLE.TABLE_BODY_SIZE, False, True
                if any(run.image is not None for run in runs):
                    return None
                width = natural_cell_width(
                    [line_text(line) for line in split_run_lines(runs)],
                    size,
                    bold=bold,
                    markers=markers,
                )
                if right == index:
                    natural[index] = max(natural[index], width)
                else:
                    spanning.append((index, right, width))
        floor = int(Inches(self._MIN_COLUMN_WIDTH_INCHES))
        natural = [max(width, floor) for width in natural]
        for left, right, width in spanning:
            missing = width - sum(natural[left : right + 1])
            if missing > 0:
                share = -(-missing // (right - left + 1))
                for index in range(left, right + 1):
                    natural[index] += share
        return natural

    @staticmethod
    def _term_sheet_role(
        row_index: int,
        column_index: int,
        right: int,
        label_columns: int,
        key_value: bool,
        header_rows: int = 1,
    ) -> str:
        """Classify an owner cell by where it starts and how far it spans.

        A cell that starts in a second-tier label column but reaches the content
        columns is content; a label in the first column stays a label however
        far it spans (for example a totals row).
        """
        if row_index < header_rows:
            return "header"
        if column_index >= label_columns or (column_index > 0 and right >= label_columns):
            return "content"
        if not key_value:
            return "grid-label"
        return "sublabel" if column_index == 1 else "label"

    def _fill_term_sheet_cell(
        self,
        word_cell,
        cell_data: TableCell,
        role: str,
        row_index: int,
        column_index: int,
        table: Table,
        column_kinds: list[str],
    ) -> tuple[list[str], Pt]:
        """Fill and style one owner cell exactly once, one paragraph per line.

        Header: header fill, bold accent, centred. Key-value label: label fill,
        bold black, centred; second label tier: unshaded, regular muted, centred.
        Grid-table labels: label fill, regular black, centred. Content cells keep
        column kinds, alignment markers and numeric roles, and indent marker lines.

        Returns:
            The cell's line texts and base text size, for row-split estimation.
        """
        column_role = table.column_types[column_index] if table.column_types else None
        fill: str | None = None
        color: RGBColor | None = None
        alignment = WD_ALIGN_PARAGRAPH.CENTER
        if role == "header":
            runs = [
                replace(run, bold=True, color_hex=None)
                for run in cell_data.runs or TextParser.parse_runs_plain(cell_data.content)
            ]
            if self.house_style.table_header == "dark":
                fill, color = STYLE.NAVY_HEX, RGBColor(0xFF, 0xFF, 0xFF)
            else:
                fill, color = STYLE.TABLE_HEADER_BG, STYLE.NAVY
            font_name, size = STYLE.HEADING_FONT, STYLE.TABLE_HEADER_SIZE
        else:
            content, runs = self._display_content(cell_data, column_role, table.table_type)
            runs = runs or TextParser.parse_runs_plain(content)
            font_name, size = STYLE.BODY_FONT, STYLE.TABLE_BODY_SIZE
            if role == "label":
                fill, color = STYLE.TS_LABEL_BG_HEX, _BLACK
                runs = [replace(run, bold=True) for run in runs]
            elif role == "grid-label":
                fill, color = STYLE.TS_LABEL_BG_HEX, _BLACK
            elif role == "sublabel":
                color = RGBColor.from_string(STYLE.TS_MUTED_HEX)
            else:
                alignment = self._cell_alignment(column_index, column_kinds, table.alignments)
                body_index = row_index - table.header_rows
                if STYLE.TABLE_ZEBRA and body_index % 2 == 0 and not cell_data.is_base_case:
                    fill = STYLE.LIGHT_GRAY_HEX
        if fill:
            set_cell_fill(word_cell._tc, fill)
        word_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        texts: list[str] = []
        for index, line in enumerate(split_run_lines(runs)):
            paragraph = word_cell.paragraphs[0] if index == 0 else word_cell.add_paragraph()
            configure_cell_paragraph(paragraph)
            paragraph.alignment = alignment
            if role == "header":
                # Never strand the header row alone at the foot of a page.
                paragraph.paragraph_format.keep_with_next = True
            text = line_text(line)
            line_size = apply_marker_layout(paragraph, text) if role == "content" else None
            TextRenderer.render_runs(
                paragraph,
                line,
                default_color=color,
                font_name=font_name,
                font_size=line_size or size,
            )
            texts.append(text)
        if role != "header" and not (
            table.table_type == TableType.FINANCIAL and column_role in {"text", "code", "date"}
        ):
            self._apply_type_styling(
                word_cell, cell_data, row_index, table.table_type, fill=self._ordered_fill
            )
        return texts, size

    @staticmethod
    def _ordered_fill(cell, hex_color: str) -> None:
        """Term-sheet cell fill placed in `w:tcPr` schema order (before `w:vAlign`)."""
        set_cell_fill(cell._tc, hex_color)

    @staticmethod
    def _set_column_widths(word_table, widths: list[int]) -> None:
        """Set exact grid and cell widths (EMU) on a table that is not merged yet."""
        word_table.autofit = False
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        for index, width in enumerate(widths):
            word_table.columns[index].width = Emu(width)
            for cell in word_table.columns[index].cells:
                cell.width = Emu(width)

    def _merge_cells(self, word_table, table: Table) -> list[tuple[int, int, int, int]]:
        """Merge validated span rectangles of an empty table whose widths are set.

        python-docx concatenates the paragraphs of merged cells that have content
        and adds the `tcW` of horizontally merged cells, so merging must happen
        after sizing and before any cell is filled.

        Args:
            word_table: Freshly created Word table with grid and cell widths applied.
            table: Parsed table whose `TableCell.merge` directives define the spans.

        Returns:
            Inclusive zero-based (top, left, bottom, right) rectangles that were merged.
        """
        rectangles, problems = self._span_rectangles(table)
        for problem in problems:
            self._report(problem)
        for top, left, bottom, right in rectangles:
            word_table.cell(top, left).merge(word_table.cell(bottom, right))
        return rectangles

    def _report(self, message: str) -> None:
        """Log a rendering diagnostic and record it for strict validation."""
        logger.warning("%s", message)
        errors = getattr(self.doc.part, "_ib_render_errors", None)
        if errors is not None:
            errors.append(message)

    @staticmethod
    def _span_rectangles(table: Table) -> tuple[list[tuple[int, int, int, int]], list[str]]:
        """Group span directives into rectangles, rejecting malformed groups.

        Parsed tables are already validated by `TableSpanResolver`; this guard keeps
        hand-built models from producing ragged or overlapping Word merges. Header
        and body are classified as the renderer draws them (the first
        `Table.header_rows` rows repeat as the header), not by
        `TableRow.is_header`, which hand-built rows may omit.

        Args:
            table: Table whose cells may carry "up"/"left" merge directives.

        Returns:
            Valid rectangles and one diagnostic per rejected group, whose cells
            are then rendered unmerged with their own content.
        """
        roots: dict[tuple[int, int], tuple[int, int]] = {}
        groups: dict[tuple[int, int], list[tuple[int, int]]] = {}
        outside: set[tuple[int, int]] = set()
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells[: table.col_count]):
                position = root = (row_index, column_index)
                if cell.merge in ("up", "left"):
                    target = (
                        (row_index - 1, column_index)
                        if cell.merge == "up"
                        else (row_index, column_index - 1)
                    )
                    if target in roots:
                        root = roots[target]
                    else:
                        outside.add(position)
                roots[position] = root
                groups.setdefault(root, []).append(position)
        rectangles: list[tuple[int, int, int, int]] = []
        problems: list[str] = []
        for (top, left), members in groups.items():
            if all(table.rows[row].cells[column].merge is None for row, column in members):
                continue
            bottom = max(row for row, _ in members)
            right = max(column for _, column in members)
            if (
                not outside.intersection(members)
                and min(row for row, _ in members) == top
                and min(column for _, column in members) == left
                and table.rows[top].cells[left].merge is None
                and len(members) == (bottom - top + 1) * (right - left + 1)
                and not (top < table.header_rows <= bottom)
            ):
                rectangles.append((top, left, bottom, right))
            else:
                problems.append(
                    f"Invalid table span group at row {top + 1}, column {left + 1}; "
                    "cells rendered unmerged"
                )
        return rectangles, problems

    @staticmethod
    def _covered_cells(rectangles: list[tuple[int, int, int, int]]) -> set[tuple[int, int]]:
        """Return grid positions represented by another cell's merged owner."""
        return {
            (row, column)
            for top, left, bottom, right in rectangles
            for row in range(top, bottom + 1)
            for column in range(left, right + 1)
            if (row, column) != (top, left)
        }

    @staticmethod
    def _mirror_vertical_fills(word_table, rectangles: list[tuple[int, int, int, int]]) -> None:
        """Repeat each vertical owner's fill on its continuation `w:tc` elements.

        Continuation cells keep their own properties in OOXML, so the owner's fill
        is repeated to keep the whole merged area shaded in every consumer.
        """
        for top, left, bottom, _ in rectangles:
            if bottom == top:
                continue
            owner = word_table.rows[top]._tr.tc_at_grid_offset(left)
            shading = owner.tcPr.find(qn("w:shd")) if owner.tcPr is not None else None
            if shading is None or not shading.get(qn("w:fill")):
                continue
            for row_index in range(top + 1, bottom + 1):
                continuation = word_table.rows[row_index]._tr.tc_at_grid_offset(left)
                set_cell_fill(continuation, shading.get(qn("w:fill")))

    @staticmethod
    def _is_span_cell(row: TableRow, col_idx: int) -> bool:
        """True for covered cells and owners spanning columns (excluded from sizing)."""
        cells = row.cells
        return cells[col_idx].merge is not None or (
            col_idx + 1 < len(cells) and cells[col_idx + 1].merge == "left"
        )

    def _apply_column_widths(self, word_table, table: Table, column_kinds: list[str]) -> None:
        """Apply content-aware column widths for more readable report tables."""
        widths = self._estimate_column_widths(
            table,
            self._get_available_table_width_inches(),
            column_kinds=column_kinds,
        )
        if not widths:
            return

        word_table.autofit = False
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT

        for c_idx, width_inches in enumerate(widths):
            width = Inches(width_inches)
            word_table.columns[c_idx].width = width
            for cell in word_table.columns[c_idx].cells:
                cell.width = width

    def _estimate_column_widths(
        self,
        table: Table,
        available_width_inches: float,
        column_kinds: list[str] | None = None,
        required_widths: list[float] | None = None,
    ) -> list[float]:
        """Estimate table column widths from content density and column semantics.

        Args:
            required_widths: Optional per-column widths (inches) that should not be
                undercut, such as a header's longest word; see `_required_minimums`.
        """
        if table.col_count <= 0:
            return []

        resolved_column_kinds = column_kinds or self._infer_column_kinds(table)
        column_infos = [
            self._build_column_info(table, col_idx, resolved_column_kinds[col_idx])
            for col_idx in range(table.col_count)
        ]
        min_widths = [
            self._minimum_column_width(column_info["kind"]) for column_info in column_infos
        ]
        max_widths = [
            self._maximum_column_width(column_info["kind"], available_width_inches)
            for column_info in column_infos
        ]
        if required_widths is not None:
            min_widths = self._required_minimums(
                min_widths, required_widths, available_width_inches
            )
            max_widths = [max(high, low) for high, low in zip(max_widths, min_widths, strict=False)]
        preferred = [column_info["score"] for column_info in column_infos]
        widths = self._fit_widths_to_available_space(
            preferred,
            min_widths,
            max_widths,
            available_width_inches,
        )

        return widths

    @classmethod
    def _required_minimums(
        cls, minimums: list[float], required: list[float], available_width_inches: float
    ) -> list[float]:
        """Raise column minimums to required widths when the section can hold them.

        First the heuristic minimums are raised; if that overflows, the required
        widths are kept over the numeric floor only (relaxing the heuristic text
        minimum); if even that overflows, the heuristic minimums are returned
        unchanged, so the table still fits without an error.
        """
        raised = [max(low, need) for low, need in zip(minimums, required, strict=False)]
        if sum(raised) <= available_width_inches:
            return raised
        relaxed = [max(cls._MIN_COLUMN_WIDTH_INCHES, need) for need in required]
        if sum(relaxed) <= available_width_inches:
            return relaxed
        return minimums

    def _build_column_info(self, table: Table, col_idx: int, column_kind: str) -> dict[str, Any]:
        """Summarize the content profile of a single column."""
        texts = []

        for row in table.rows:
            if col_idx >= len(row.cells) or self._is_span_cell(row, col_idx):
                continue

            cell = row.cells[col_idx]
            text = self._cell_display_text(cell, table.table_type).strip()
            if not text:
                continue

            texts.append(text)

        representative_length = self._representative_length(texts)
        longest_token = self._longest_token_length(texts)
        is_numeric_column = column_kind == "numeric"
        score = self._column_score(is_numeric_column, representative_length, longest_token)

        return {
            "kind": column_kind,
            "score": score,
        }

    def _infer_column_kinds(self, table: Table) -> list[str]:
        if table.column_types:
            return [
                "text" if role in {"text", "code", "date"} else "numeric"
                for role in table.column_types
            ]
        """Infer each column's semantic type from the first few data cells."""
        return [self._infer_column_kind(table, col_idx) for col_idx in range(table.col_count)]

    def _infer_column_kind(self, table: Table, col_idx: int) -> str:
        """Classify a column as text or numeric using the first 2-3 body cells.

        Empty cells and dash placeholders (`-`, `–`) are not samples.
        """
        sample_texts = []

        for row in table.rows[table.header_rows :]:
            if col_idx >= len(row.cells) or self._is_span_cell(row, col_idx):
                continue

            text = self._cell_display_text(row.cells[col_idx], table.table_type).strip()
            if not text or self._DASH_PLACEHOLDER_RE.fullmatch(text):
                continue

            sample_texts.append(text)
            if len(sample_texts) >= self._COLUMN_SAMPLE_SIZE:
                break

        if not sample_texts:
            return "text"

        numeric_like_count = sum(1 for text in sample_texts if self._is_numeric_like(text))
        return "numeric" if numeric_like_count >= (len(sample_texts) + 1) // 2 else "text"

    @classmethod
    def _is_numeric_like(cls, text: str) -> bool:
        """Return True for financial/count strings that should align right."""
        normalized = text.strip()
        if not normalized or not any(char.isdigit() for char in normalized):
            return False
        return bool(cls._NUMERIC_LIKE_RE.match(normalized))

    @classmethod
    def _representative_length(cls, texts: list[str]) -> int:
        """Return a robust content length that ignores a few extreme outliers."""
        if not texts:
            return 0

        lengths = sorted(len(text) for text in texts)
        percentile_index = int(round((len(lengths) - 1) * 0.75))
        return lengths[percentile_index]

    @classmethod
    def _longest_token_length(cls, texts: list[str]) -> int:
        """Return the length of the longest unbroken token in the sample."""
        longest = 0
        for text in texts:
            tokens = [token for token in cls._TOKEN_SPLIT_RE.split(text) if token]
            if not tokens:
                continue
            longest = max(longest, max(len(token) for token in tokens))
        return longest

    @classmethod
    def _column_score(
        cls,
        is_numeric_column: bool,
        representative_length: int,
        longest_token: int,
    ) -> float:
        """Assign a width score to a column based on how hard it is to wrap cleanly."""
        if is_numeric_column:
            return 1.0 + min(representative_length, 12) * 0.05

        return 1.8 + min(representative_length, 48) * 0.06 + min(longest_token, 24) * 0.03

    @classmethod
    def _minimum_column_width(cls, kind_marker: str) -> float:
        """Return a minimum readable width for a column kind."""
        if kind_marker == "numeric":
            return cls._MIN_COLUMN_WIDTH_INCHES
        return cls._MIN_TEXT_COLUMN_WIDTH_INCHES

    @classmethod
    def _maximum_column_width(cls, kind_marker: str, available_width_inches: float) -> float:
        """Return an upper width bound for a column kind."""
        if kind_marker == "numeric":
            return cls._MAX_NUMERIC_COLUMN_WIDTH_INCHES
        return max(
            cls._MIN_TEXT_COLUMN_WIDTH_INCHES,
            available_width_inches * cls._MAX_TEXT_COLUMN_SHARE,
        )

    @classmethod
    def _fit_widths_to_available_space(
        cls,
        preferred: list[float],
        min_widths: list[float],
        max_widths: list[float],
        available_width_inches: float,
    ) -> list[float]:
        """Fit preferred widths into the available width with min/max guards."""
        col_count = len(preferred)
        if col_count == 0:
            return []

        min_total = sum(min_widths)
        if min_total >= available_width_inches:
            scale = available_width_inches / min_total
            return [width * scale for width in min_widths]

        total_preferred = sum(preferred) or float(col_count)
        widths = [
            max(min_widths[idx], available_width_inches * preferred[idx] / total_preferred)
            for idx in range(col_count)
        ]

        widths = [min(widths[idx], max_widths[idx]) for idx in range(col_count)]
        widths = cls._rebalance_widths(
            widths, min_widths, max_widths, preferred, available_width_inches
        )
        return widths

    @classmethod
    def _rebalance_widths(
        cls,
        widths: list[float],
        min_widths: list[float],
        max_widths: list[float],
        preferred: list[float],
        available_width_inches: float,
    ) -> list[float]:
        """Rebalance widths after clamping so the table uses the full line width."""
        for _ in range(8):
            total_width = sum(widths)
            gap = available_width_inches - total_width

            if abs(gap) < 0.01:
                break

            if gap > 0:
                growable = [
                    idx for idx, width in enumerate(widths) if width < max_widths[idx] - 0.01
                ]
                if not growable:
                    break
                grow_weight = sum(preferred[idx] for idx in growable) or float(len(growable))
                for idx in growable:
                    share = preferred[idx] / grow_weight if grow_weight else 1.0 / len(growable)
                    widths[idx] = min(max_widths[idx], widths[idx] + gap * share)
            else:
                shrinkable = [
                    idx for idx, width in enumerate(widths) if width > min_widths[idx] + 0.01
                ]
                if not shrinkable:
                    break
                shrink_capacity = sum(widths[idx] - min_widths[idx] for idx in shrinkable)
                if shrink_capacity <= 0:
                    break
                overflow = -gap
                for idx in shrinkable:
                    capacity = widths[idx] - min_widths[idx]
                    share = capacity / shrink_capacity
                    widths[idx] = max(min_widths[idx], widths[idx] - overflow * share)

        return widths

    def _get_available_table_width_inches(self) -> float:
        """Return the horizontal space available for body tables."""
        return float(self._get_available_table_width_emu()) / float(self._EMUS_PER_INCH)

    def _get_available_table_width_emu(self) -> int:
        """Return the printable width of the current (last) section in EMU."""
        if not self.doc.sections:
            return int(Inches(6.0))

        section = self.doc.sections[-1]
        if (
            section.page_width is None
            or section.left_margin is None
            or section.right_margin is None
        ):
            return int(Inches(6.0))
        available_width = (
            int(section.page_width) - int(section.left_margin) - int(section.right_margin)
        )
        return max(available_width, int(Inches(3.0)))

    def _cell_display_text(self, cell_data: TableCell, table_type: TableType) -> str:
        """Return display text used for width estimation and column-kind inference."""
        shown = term_cell_text(cell_data)
        if shown is not None:
            return shown  # term values are shown exactly as written in `terms:`
        if table_type == TableType.FINANCIAL and cell_data.is_numeric:
            return self._format_financial_number(cell_data.content)
        return cell_data.content

    def _render_header_row(
        self,
        word_table,
        row: TableRow,
        col_count: int,
        covered: Set[tuple[int, int]] = frozenset(),
        row_index: int = 0,
    ):
        """Render one header row with Navy background, skipping merged-away cells."""
        word_cells = word_table.rows[row_index].cells

        for c_idx, cell_data in enumerate(row.cells):
            if c_idx >= col_count:
                break
            if (row_index, c_idx) in covered:
                continue

            cell = word_cells[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            TableStyler.set_cell_background(cell, STYLE.TABLE_HEADER_BG)

            # Clear default paragraph and add styled run
            p = cell.paragraphs[0]
            self._configure_cell_paragraph(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.clear()

            # Parse header text: strip markdown markers but preserve structure
            # Headers are always bold+white, so we parse for content only
            if cell_data.runs:
                # Use structured runs but override style for header
                TextRenderer.render_runs(
                    p,
                    [replace(run, bold=True, color_hex=None) for run in cell_data.runs],
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.TABLE_HEADER_SIZE,
                    default_color=STYLE.TABLE_HEADER_COLOR,
                )
            else:
                clean_text = cell_data.content.replace("**", "").strip()
                run = p.add_run(clean_text)
                FontStyler.apply_run_style(
                    run,
                    font_name=STYLE.HEADING_FONT,
                    font_size=STYLE.TABLE_HEADER_SIZE,
                    bold=True,
                    color=STYLE.TABLE_HEADER_COLOR,
                )

    def _render_data_row(
        self,
        word_table,
        row: TableRow,
        row_idx: int,
        col_count: int,
        table_type: TableType,
        column_kinds: list[str],
        column_types: list[str] | None = None,
        alignments: list[str] | None = None,
        covered: Set[tuple[int, int]] = frozenset(),
        header_rows: int = 1,
    ):
        """Render a data row with type-specific styling, skipping merged-away cells.

        Zebra shading counts body rows only, so it starts on the first body row
        whatever the number of header rows.
        """
        word_cells = word_table.rows[row_idx].cells

        for c_idx, cell_data in enumerate(row.cells):
            if c_idx >= col_count:
                break
            if (row_idx, c_idx) in covered:
                continue

            cell = word_cells[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            self._configure_cell_paragraph(p)

            # ── Alignment ───────────────────────────────────────────────────
            p.alignment = self._cell_alignment(c_idx, column_kinds, alignments)

            # ── Format content (financial number formatting) ───────────────
            role = column_types[c_idx] if column_types else None
            display_content, display_runs = self._display_content(cell_data, role, table_type)

            # ── Render content ──────────────────────────────────────────────
            if cell_data.runs:
                # Use structured runs for full formatting fidelity
                TextRenderer.render_runs(
                    p,
                    display_runs,
                    font_name=STYLE.BODY_FONT,
                    font_size=STYLE.TABLE_BODY_SIZE,
                )
            else:
                TextRenderer.render_text_with_formatting(
                    p,
                    display_content,
                    font_name=STYLE.BODY_FONT,
                    font_size=STYLE.TABLE_BODY_SIZE,
                )

            # ── Type-specific styling ───────────────────────────────────────
            if not (table_type == TableType.FINANCIAL and role in {"text", "code", "date"}):
                self._apply_type_styling(cell, cell_data, row_idx, table_type)

            # ── Alternating row colors (unless special styling applied) ─────
            if (
                STYLE.TABLE_ZEBRA
                and (row_idx - header_rows) % 2 == 0
                and not cell_data.is_base_case
            ):
                TableStyler.set_cell_background(cell, STYLE.LIGHT_GRAY_HEX)

    @staticmethod
    def _cell_alignment(
        c_idx: int, column_kinds: list[str], alignments: list[str] | None
    ) -> WD_ALIGN_PARAGRAPH:
        """Numeric columns align right; Markdown alignment markers take precedence."""
        if alignments and c_idx < len(alignments) and alignments[c_idx] in {"center", "right"}:
            return (
                WD_ALIGN_PARAGRAPH.CENTER
                if alignments[c_idx] == "center"
                else WD_ALIGN_PARAGRAPH.RIGHT
            )
        if c_idx < len(column_kinds) and column_kinds[c_idx] == "numeric":
            return WD_ALIGN_PARAGRAPH.RIGHT
        return WD_ALIGN_PARAGRAPH.LEFT

    def _display_content(
        self, cell_data: TableCell, role: str | None, table_type: TableType
    ) -> tuple[str, list[TextRun]]:
        """Return display text and runs, formatting amounts only for numeric roles."""
        format_numbers = role not in {None, "text", "code", "date"} or (
            not role and table_type == TableType.FINANCIAL and cell_data.is_numeric
        )
        display_content = cell_data.content
        display_runs = cell_data.runs
        if format_numbers:
            display_content = self._format_numeric_text(display_content, role)
            # Semantic runs are independent of the amount. In particular, a
            # footnote's displayed number must never enter numeric formatting,
            # and a term value keeps the exact text written in `terms:`.
            display_runs = [
                run
                if (
                    run.footnote_id is not None
                    or run.is_latex
                    or run.superscript
                    or run.subscript
                    or run.term_key is not None
                    or run.code
                )
                else replace(run, text=self._format_numeric_text(run.text, role))
                for run in display_runs
            ]
        return display_content, display_runs

    @classmethod
    def _format_numeric_text(cls, text: str, role: str | None) -> str:
        """Format a numeric text run without changing its semantic metadata.

        Args:
            text: One text run, excluding footnote and equation runs.
            role: Explicit column presentation role, if configured.

        Returns:
            Formatted number with its original scale, or unchanged nonnumeric text.
        """
        if not cls._NUMERIC_LIKE_RE.fullmatch(text.strip()):
            return text
        formatted = cls._format_financial_number(text)
        suffix = {"percent": "%", "bps": " bps", "multiple": "x"}.get(role or "", "")
        if suffix and cls._PURE_NUMBER_RE.fullmatch(text.strip()):
            formatted += suffix
        leading = text[: len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()) :]
        return leading + formatted + trailing

    @staticmethod
    def _configure_cell_paragraph(paragraph) -> None:
        """Normalize paragraph spacing inside table cells for a tighter, cleaner grid."""
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0

    @staticmethod
    def _format_financial_number(text: str) -> str:
        """
        Format financial numbers with thousand separators.

        Handles:
            - Plain numbers: 1234567 → 1,234,567
            - Percentages: 12.5% → 12.5%
            - Negative parentheses: (1234) → (1,234)
            - Korean 억/조 units preserved
            - Already formatted numbers: 1,234 → 1,234

        Args:
            text: Cell text content

        Returns:
            Formatted text with thousand separators
        """
        import re

        text = text.strip()

        # Skip if empty or non-numeric looking
        if not text or not any(c.isdigit() for c in text):
            return text

        # Skip if already has thousand separators
        if "," in text and re.search(r"\d{1,3}(,\d{3})+", text):
            return text

        # Handle negative in parentheses: (1234567) → (1,234,567)
        paren_match = re.match(r"^\((\d+(?:\.\d+)?)\)(.*)$", text)
        if paren_match:
            num_str = paren_match.group(1)
            suffix = paren_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"({formatted}){suffix}"

        # Handle negative with minus: -1234567 → -1,234,567
        neg_match = re.match(r"^-(\d+(?:\.\d+)?)(.*)$", text)
        if neg_match:
            num_str = neg_match.group(1)
            suffix = neg_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"-{formatted}{suffix}"

        # Handle positive numbers with optional suffix (%, 억, 조, 원, etc.)
        pos_match = re.match(r"^(\d+(?:\.\d+)?)(.*)$", text)
        if pos_match:
            num_str = pos_match.group(1)
            suffix = pos_match.group(2)
            formatted = TableRenderer._add_thousand_sep(num_str)
            return f"{formatted}{suffix}"

        return text

    @staticmethod
    def _add_thousand_sep(num_str: str) -> str:
        """Add thousand separators to a numeric string."""
        if "." in num_str:
            integer_part, decimal_part = num_str.split(".", 1)
            formatted_int = f"{int(integer_part):,}"
            return f"{formatted_int}.{decimal_part}"
        else:
            return f"{int(num_str):,}"

    def _apply_type_styling(
        self,
        cell,
        cell_data: TableCell,
        row_idx: int,
        table_type: TableType,
        fill: Callable[[Any, str], None] = TableStyler.set_cell_background,
    ):
        """Apply table-type specific styling to every paragraph of one owner cell.

        Args:
            fill: Cell background setter. Legacy tables keep the historical append;
                term-sheet tables pass a schema-ordered setter.
        """
        runs = [run for paragraph in cell.paragraphs for run in paragraph.runs]
        if table_type == TableType.FINANCIAL:
            if cell_data.is_negative:
                for run in runs:
                    run.font.color.rgb = STYLE.RED

        elif table_type == TableType.BEP_SENSITIVITY:
            if cell_data.is_base_case:
                fill(cell, STYLE.YELLOW_HEX)
                for run in runs:
                    run.font.bold = True

        elif table_type == TableType.RISK_MATRIX:
            if cell_data.risk_level:
                color_map = {
                    "high": STYLE.RED,
                    "medium": STYLE.ORANGE,
                    "low": STYLE.GREEN,
                }
                color = color_map.get(cell_data.risk_level)
                if color:
                    for run in runs:
                        run.font.color.rgb = color
                        run.font.bold = True
