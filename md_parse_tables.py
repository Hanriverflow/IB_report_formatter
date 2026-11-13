"""Markdown pipe-table parsing and table span resolution."""

import logging
from collections.abc import Sequence
from dataclasses import replace

from document_model import Table, TableCell, TableRow, TableType, TextRun
from document_profiles import header_labels
from input_safety import DEFAULT_LIMITS, InputLimits, RejectionCode
from md_parse_text import TextParser, _parse_term_runs, _substitute_text
from term_variables import TermResolver

logger = logging.getLogger("md_parser")


class TableParser:
    """Parses markdown tables"""

    # Keywords for table type detection
    BEP_KEYWORDS = [
        "bep",
        "sensitivity",
        "cmr",
        "contribution margin",
        "fixed cost",
        "손익분기",
        "민감도",
        "고정비",
        "변동비",
    ]

    RISK_KEYWORDS = [
        "risk",
        "impact",
        "probability",
        "likelihood",
        "리스크",
        "위험",
        "영향",
        "확률",
    ]

    FINANCIAL_KEYWORDS = [
        "revenue",
        "income",
        "ebitda",
        "profit",
        "margin",
        "expense",
        "매출",
        "수익",
        "이익",
        "손익",
        "순이익",
        "영업",
        "비용",
    ]

    YEAR_INDICATORS = [
        "2024",
        "2025",
        "2026",
        "yoy",
        "cagr",
        "a)",
        "b)",
        "e)",
        "년도",
        "연도",
        "실적",
    ]

    UPSIDE_DOWNSIDE_KEYWORDS = [
        "upside",
        "downside",
        "상승",
        "하락",
        "요인",
    ]

    @staticmethod
    def parse(
        lines: list[str],
        financial_rules: bool = True,
        terms: TermResolver | None = None,
        limits: InputLimits = DEFAULT_LIMITS,
    ) -> Table:
        """
        Parse markdown table lines into a Table object.

        Args:
            lines: Lines that make up the table (starting with |)
            financial_rules: Whether IB table semantics are inferred.
            terms: Active term resolver; cell runs receive values, content stays raw.
            limits: Operator-only table-cell limit.
        """
        # Only the second line can be the structural delimiter row.
        data_lines = [
            line
            for index, line in enumerate(lines)
            if not (index == 1 and TableParser._is_delimiter_row(line))
        ]
        return TableParser.from_cells(
            [TableParser._split_row(line) for line in data_lines],
            TableParser._parse_alignments(lines),
            financial_rules=financial_rules,
            terms=terms,
            limits=limits,
        )

    @staticmethod
    def from_cells(
        rows: Sequence[Sequence[str | TableCell]],
        alignments: list[str],
        financial_rules: bool = True,
        terms: TermResolver | None = None,
        header_rows: int = 1,
        limits: InputLimits = DEFAULT_LIMITS,
    ) -> Table:
        """
        Build a Table from cell sources; the leading `header_rows` rows are the header.

        Args:
            rows: Each cell as inline Markdown (span markers allowed) or as a
                prepared cell whose runs are final (HTML tables).
            alignments: Column alignments; ignored unless one per column.
            financial_rules: Whether IB table semantics are inferred.
            terms: Active term resolver; cell runs receive values, content stays raw.
            header_rows: Number of header rows (at least one).
            limits: Operator-only table-cell limit.
        """
        table = Table()
        if not rows:
            return table
        limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(rows) * max(map(len, rows)))
        table.header_rows = max(1, min(header_rows, len(rows)))

        # Get column count from first row
        first_row_cells = rows[0]
        table.col_count = len(first_row_cells)
        table.alignments = (
            alignments if len(alignments) == table.col_count else ["left"] * table.col_count
        )

        # Header semantics follow the values the reader sees across every header
        # row, merged labels included; cells keep raw content.
        header_cells = header_labels(
            [
                [
                    "".join(run.text for run in source.runs)
                    if isinstance(source, TableCell)
                    else source
                    if terms is None
                    else _substitute_text(source, terms)
                    for source in header_row
                ]
                for header_row in rows[: table.header_rows]
            ]
        )

        # Detect table type
        header_text = " ".join(header_cells).lower()
        table.table_type = (
            TableParser._detect_type(header_text) if financial_rules else TableType.GENERIC
        )

        # Parse all rows — normalise column count per row
        for i, row_cells in enumerate(rows):
            cells = list(row_cells)
            is_header = i < table.header_rows

            # Pad short rows with empty cells
            while len(cells) < table.col_count:
                cells.append("")

            # Truncate extra cells with warning (v3: no silent data loss)
            if len(cells) > table.col_count:
                logger.warning(
                    "Table row %d has %d columns (expected %d) — extra columns dropped: %s",
                    i,
                    len(cells),
                    table.col_count,
                    cells[table.col_count :],
                )
                table.warnings.append(f"Table row {i} has extra cells; content was dropped")
                cells = cells[: table.col_count]

            row = TableRow(is_header=is_header)
            for j, cell_source in enumerate(cells):
                if isinstance(cell_source, TableCell):
                    cell = replace(cell_source, is_header=is_header)
                    TableParser._detect_flags(
                        cell,
                        "".join(run.text for run in cell.runs),
                        j,
                        table.table_type,
                        header_cells,
                    )
                else:
                    cell = TableParser._parse_cell(
                        cell_source,
                        is_header=is_header,
                        col_idx=j,
                        row_idx=i,
                        total_rows=len(rows),
                        table_type=table.table_type,
                        header_cells=header_cells,
                        terms=terms,
                    )
                cell.alignment = table.alignments[j] if j < len(table.alignments) else "left"
                row.cells.append(cell)

            table.rows.append(row)

        return table

    @staticmethod
    def _split_row(line: str) -> list[str]:
        """Split a table row into cell contents"""
        text = line.strip()
        cells: list[str] = []
        current: list[str] = []
        index = 0
        while index < len(text):
            char = text[index]
            if char == "\\" and index + 1 < len(text) and text[index + 1] in {"|", "\\"}:
                # An escaped pipe is cell text. An escaped backslash stays escaped:
                # the cell's inline parse resolves it, once, with every other escape.
                current.append("|" if text[index + 1] == "|" else "\\\\")
                index += 2
                continue
            if char == "|":
                cells.append("".join(current).strip())
                current = []
            else:
                current.append(char)
            index += 1
        cells.append("".join(current).strip())
        if text.startswith("|"):
            cells = cells[1:]
        if text.endswith("|") and cells and cells[-1] == "":
            backslashes = len(text[:-1]) - len(text[:-1].rstrip("\\"))
            if backslashes % 2 == 0:
                cells = cells[:-1]
        return cells

    @staticmethod
    def _is_delimiter_row(line: str) -> bool:
        """Recognize a structural table delimiter without classifying body rows."""
        return "-" in line and set(line.strip()).issubset({"|", "-", " ", "\t", ":"})

    @staticmethod
    def _parse_alignments(lines: list[str]) -> list[str]:
        """Parse column alignments from separator line"""
        alignments: list[str] = []
        for line in lines[1:2]:
            if TableParser._is_delimiter_row(line):
                cells = line.split("|")
                for cell in cells:
                    cell = cell.strip()
                    if not cell:
                        continue
                    if cell.startswith(":") and cell.endswith(":"):
                        alignments.append("center")
                    elif cell.endswith(":"):
                        alignments.append("right")
                    else:
                        alignments.append("left")
                break
        return alignments

    @staticmethod
    def _detect_type(header_text: str) -> TableType:
        """Detect table type from header row"""
        if any(kw in header_text for kw in TableParser.UPSIDE_DOWNSIDE_KEYWORDS):
            return TableType.UPSIDE_DOWNSIDE

        if any(kw in header_text for kw in TableParser.BEP_KEYWORDS):
            return TableType.BEP_SENSITIVITY

        if any(kw in header_text for kw in TableParser.RISK_KEYWORDS):
            return TableType.RISK_MATRIX

        has_financial = any(kw in header_text for kw in TableParser.FINANCIAL_KEYWORDS)
        has_year = any(yi in header_text for yi in TableParser.YEAR_INDICATORS)
        if has_financial or has_year:
            return TableType.FINANCIAL

        return TableType.GENERIC

    @staticmethod
    def _parse_cell(
        text: str,
        is_header: bool,
        col_idx: int,
        row_idx: int,
        total_rows: int,
        table_type: TableType,
        header_cells: list[str],
        terms: TermResolver | None = None,
    ) -> TableCell:
        """Parse a single table cell.

        The raw `content` is kept for span markers and width estimation; runs
        and content-derived flags see substituted term values as typed text.
        """
        cell = TableCell(content=text, is_header=is_header)

        # Prefer LaTeX-aware parsing only when a balanced inline expression exists.
        latex = TextParser.has_inline_latex(text)
        parse = TextParser.parse_runs if latex else TextParser.parse_runs_plain
        detected_text, cell.runs, _ = _parse_term_runs(text, parse, terms, latex=latex)
        TableParser._detect_flags(cell, detected_text, col_idx, table_type, header_cells)
        return cell

    @staticmethod
    def _detect_flags(
        cell: TableCell,
        detected_text: str,
        col_idx: int,
        table_type: TableType,
        header_cells: list[str],
    ) -> None:
        """Set a cell's numeric, negative and risk flags from its text.

        Args:
            cell: Cell with final runs.
            detected_text: Text the numeric and risk detection reads.
            col_idx: Zero-based column.
            table_type: Detected or configured table type.
            header_cells: Header texts, for risk columns.
        """
        # Detect numeric content
        cell.is_numeric = any(char.isdigit() for char in detected_text) and col_idx > 0

        # Detect negative numbers
        visible_text = "".join(run.text for run in cell.runs)
        cell.is_negative = (
            visible_text.startswith("(")
            and visible_text.endswith(")")
            and any(c.isdigit() for c in visible_text)
        ) or (visible_text.startswith("-") and any(c.isdigit() for c in visible_text))

        # Table-type specific detection
        if table_type == TableType.RISK_MATRIX:
            header_lower = header_cells[col_idx].lower() if col_idx < len(header_cells) else ""
            risk_header_kw = ("impact", "probability", "영향", "확률")
            if any(kw in header_lower for kw in risk_header_kw):
                text_lower = detected_text.lower()
                if "high" in text_lower or "높" in detected_text:
                    cell.risk_level = "high"
                elif (
                    any(kw in text_lower for kw in ("medium", "moderate")) or "중" in detected_text
                ):
                    cell.risk_level = "medium"
                elif "low" in text_lower or "낮" in detected_text:
                    cell.risk_level = "low"


class TableSpanResolver:
    """Resolve opt-in span markers while retaining invalid groups literally."""

    @staticmethod
    def resolve(table: Table, profile: str = "ib-report") -> None:
        """Validate each connected marker group and infer leading label columns.

        References point only above or left, so row-major traversal resolves a
        cell's group from an already visited target without recursive walks.
        Invalid groups never receive merge directives; neighboring groups remain
        independent even when their bounding rectangles overlap.

        Args:
            table: A parsed table after ordered specifications have been applied.
            profile: Selected document profile, including its default span policy.
        """
        active = table.spans if table.spans is not None else profile == "term-sheet"
        first_header_width = 1
        if active:
            markers: dict[tuple[int, int], str] = {}
            roots: dict[tuple[int, int], tuple[int, int]] = {}
            groups: dict[tuple[int, int], list[tuple[int, int]]] = {}
            outside: set[tuple[int, int]] = set()

            for row_index, row in enumerate(table.rows):
                for column_index, cell in enumerate(row.cells):
                    coordinate = (row_index, column_index)
                    raw = "" if cell.literal else cell.content.strip()
                    root = coordinate
                    if raw in {r"\^^", r"\<<"}:
                        # Classify raw source before unescaping either representation.
                        cell.content = raw[1:]
                        cell.runs = [TextRun(text=raw[1:])]
                    elif raw in {"^^", "<<"}:
                        markers[coordinate] = "up" if raw == "^^" else "left"
                        target = (
                            (row_index - 1, column_index)
                            if raw == "^^"
                            else (row_index, column_index - 1)
                        )
                        if target in roots:
                            root = roots[target]
                        else:
                            outside.add(coordinate)
                    roots[coordinate] = root
                    groups.setdefault(root, []).append(coordinate)

            for coordinates in groups.values():
                marker_cells = [position for position in coordinates if position in markers]
                if not marker_cells:
                    continue
                top = min(row for row, _ in coordinates)
                bottom = max(row for row, _ in coordinates)
                left = min(column for _, column in coordinates)
                right = max(column for _, column in coordinates)
                anchors = [position for position in coordinates if position not in markers]
                error = ""
                if any(position in outside for position in coordinates):
                    error = "a marker points outside the table"
                elif len({table.rows[row].is_header for row, _ in coordinates}) > 1:
                    error = "a merge crosses the header/body boundary"
                elif anchors != [(top, left)]:
                    error = "a merge requires one top-left anchor"
                elif len(coordinates) != (bottom - top + 1) * (right - left + 1):
                    error = "a merge must fill one complete rectangle"
                if error:
                    warning = (
                        f"Invalid table span group at row {top + 1}, column {left + 1}: "
                        f"{error}; markers kept literally"
                    )
                    table.warnings.append(warning)
                    logger.warning("%s", warning)
                    continue
                for marker_row, marker_column in marker_cells:
                    table.rows[marker_row].cells[marker_column].merge = markers[
                        (marker_row, marker_column)
                    ]
                if (top, left) == (0, 0) and table.rows[0].is_header:
                    first_header_width = right + 1

        if table.label_columns is None and (active or profile == "term-sheet"):
            table.label_columns = min(first_header_width, max(table.col_count - 1, 0))
