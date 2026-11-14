"""HTML table parsing for converter output, including spans and inline images."""

import re
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser

from document_model import CrossReference, DescribedImage, Image, TableCell, TextRun
from input_safety import DEFAULT_LIMITS, InputLimits, RejectionCode
from md_parse_text import TextParser
from term_variables import TermResolver, TokenMap, split_term_runs

# ═══════════════════════════════════════════════════════════════════════════════
# HTML TABLES (converter output)
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class _HtmlCell:
    """One `<td>`/`<th>` of the outermost table while it is being read."""

    colspan: int = 1
    rowspan: int = 1
    runs: list[TextRun] = field(default_factory=list)


def _html_image(source: str, alt: str) -> Image | None:
    """Return the image an HTML `src` names; a `data:` URI keeps its bytes.

    Args:
        source: The `src` attribute.
        alt: The `alt` attribute.

    Returns:
        The image, or None without a usable source.
    """
    source = source.strip()
    if source.lower().startswith("data:"):
        header, _, payload = source.partition(",")
        if not header.lower().endswith(";base64") or not payload.strip():
            return None
        return Image(
            alt_text=alt,
            path="",
            base64_data=payload.strip(),
            mime_type=header[5:-7] or "image/png",
        )
    if not source or any(character in source for character in "<>\n"):
        return None
    references = tuple(
        run.target
        for run in TextParser.mark_cross_references([TextRun(text=alt)])
        if isinstance(run, CrossReference) and run.target
    )
    if references:
        return DescribedImage(alt_text=alt, path=source, references=references)
    return Image(alt_text=alt, path=source)


class _HtmlTableReader(HTMLParser):
    """Collect the outermost table's cells as runs; HTML text is always literal.

    Formatting tags nest (`<b><i>x</i></b>` is bold italic) and end with their
    cell. Block tags (`p`, `div`, `li`, headings, nested rows and tables) start
    and end cell lines; every `<br>` breaks a line.
    """

    _STYLE_TAGS = {
        "b": "bold",
        "strong": "bold",
        "i": "italic",
        "em": "italic",
        "code": "code",
        "tt": "code",
        "kbd": "code",
        "sup": "superscript",
        "sub": "subscript",
    }
    _BLOCK_TAGS = frozenset(
        {
            "p",
            "div",
            "li",
            "ul",
            "ol",
            "tr",
            "blockquote",
            "pre",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
        }
    )
    _MAX_SPAN = 63  # Word's column limit bounds any sensible span

    def __init__(self, terms: TermResolver | None = None) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[_HtmlCell]] = []
        self.stray: list[str] = []
        self.extra_tables = 0  # further top-level tables in the same block (not read)
        self._terms = terms
        self._depth = 0
        self._cell: _HtmlCell | None = None
        self._styles: dict[str, int] = {}
        self._links: list[str] = []
        self._colors: list[str | None] = []
        self._line_break = False  # a block boundary: the next content starts a new line
        self._separator = False  # a nested cell boundary: the next content is spaced

    @classmethod
    def _span(cls, value: str | None) -> int:
        try:
            return max(1, int(str(value).strip()))
        except ValueError:
            return 1

    def _line_has_content(self) -> bool:
        for run in reversed(self._cell.runs if self._cell is not None else []):
            if run.text == "\n":
                return False
            if run.image is not None or run.text.strip():
                return True
        return False

    def _add(self, run: TextRun) -> None:
        """Append content, first realizing a pending line break or nested-cell space."""
        if self._cell is None:
            return
        if self._line_break:
            if self._line_has_content():
                self._cell.runs.append(TextRun(text="\n"))
        elif self._separator and self._line_has_content():
            self._cell.runs.append(TextRun(text=" "))
        self._line_break = self._separator = False
        self._cell.runs.append(run)

    def _formatted(self, text: str) -> TextRun:
        """A run of literal text under the open formatting tags."""
        link = next((href for href in reversed(self._links) if href), "")
        return TextRun(
            text=text,
            bold=self._styles.get("bold", 0) > 0,
            italic=self._styles.get("italic", 0) > 0,
            superscript=self._styles.get("superscript", 0) > 0,
            subscript=self._styles.get("subscript", 0) > 0,
            code=self._styles.get("code", 0) > 0,
            color_hex=next((color for color in reversed(self._colors) if color), None),
            hyperlink=TextParser.link_target(link) if link else None,
        )

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = {name: value or "" for name, value in attrs}
        if tag == "table":
            if self._depth == 0 and self.rows:
                self.extra_tables += 1
            elif self._cell is not None:
                self._line_break = True  # a nested table starts on its own line
            self._depth += 1
            return
        if self.extra_tables or self._depth == 0:
            return
        if self._depth == 1 and tag == "tr":
            self.rows.append([])
            self._cell = None
            return
        if self._depth == 1 and tag in ("td", "th"):
            if not self.rows:
                self.rows.append([])
            self._cell = _HtmlCell(
                self._span(attributes.get("colspan")), self._span(attributes.get("rowspan"))
            )
            self._styles, self._links, self._colors = {}, [], []
            self._line_break = self._separator = False
            self.rows[-1].append(self._cell)
            return
        if self._cell is None:
            return
        if tag in self._BLOCK_TAGS:
            self._line_break = True
            if tag == "li":
                self._add(TextRun(text="• "))
        elif tag in ("td", "th"):
            self._separator = True
        elif tag == "br":
            self._line_break = self._separator = False
            self._cell.runs.append(TextRun(text="\n"))
        elif tag in self._STYLE_TAGS:
            name = self._STYLE_TAGS[tag]
            self._styles[name] = self._styles.get(name, 0) + 1
        elif tag == "a":
            self._links.append(attributes.get("href", "").strip())
        elif tag in ("span", "font"):
            match = TextParser._COLOR_STYLE_RE.search(attributes.get("style", ""))
            color = match.group(1) if match else attributes.get("color", "")
            self._colors.append(color.upper() if re.fullmatch(r"#[0-9A-Fa-f]{6}", color) else None)
        elif tag == "img":
            image = _html_image(attributes.get("src", ""), attributes.get("alt", ""))
            if image is not None:
                self._add(TextRun(text="", image=image))

    def handle_endtag(self, tag: str) -> None:
        if tag == "table":
            self._depth = max(self._depth - 1, 0)
            if self._cell is not None:
                self._line_break = True  # text after a nested table starts a new line
            return
        if self.extra_tables or self._depth == 0:
            return
        if self._depth == 1 and tag in ("td", "th"):
            self._cell = None
            return
        if self._cell is None:
            return
        if tag in self._BLOCK_TAGS:
            self._line_break = True
        elif tag in self._STYLE_TAGS:
            name = self._STYLE_TAGS[tag]
            self._styles[name] = max(self._styles.get(name, 0) - 1, 0)
        elif tag == "a" and self._links:
            self._links.pop()
        elif tag in ("span", "font") and self._colors:
            self._colors.pop()

    def handle_data(self, data: str) -> None:
        if self.extra_tables:
            return
        if self._cell is None:
            if data.strip():
                self.stray.append(data.strip())
            return
        text = re.sub(r"[ \t\r\n\f]+", " ", data)
        if not text.strip() and (
            self._line_break or self._separator or not self._line_has_content()
        ):
            return  # layout whitespace between blocks
        self._add(self._formatted(text))

    def finish(self, runs: list[TextRun]) -> list[TextRun]:
        """Make one cell's final runs, as HTML shows them.

        Adjacent runs with the same formatting are merged first, so an amount
        or a `{{key}}` reference split by a neutral tag such as `<span>` is
        whole; then references are substituted, spaces HTML would not show are
        removed and line breaks at the cell edges dropped. Code and term values
        keep their text, including an empty value.

        Args:
            runs: Runs collected for one cell.

        Returns:
            The cell's runs.
        """
        merged: list[TextRun] = []
        for run in runs:
            previous = merged[-1] if merged else None
            if (
                previous is not None
                and previous.image is None
                and run.image is None
                and "\n" not in (previous.text, run.text)
                and replace(previous, text="") == replace(run, text="")
            ):
                tail = run.text.lstrip(" ") if previous.text.endswith(" ") else run.text
                merged[-1] = replace(previous, text=previous.text + tail)
            else:
                merged.append(run)
        substituted: list[TextRun] = []
        for run in merged:
            if self._terms is None or run.code or "{{" not in run.text:
                substituted.append(run)
                continue
            self._terms.reserve(run.text)  # literal token-shaped text never becomes a value
            tokens: TokenMap = {}
            tokenized = TextParser._tokenize_references(run.text, {}, {}, self._terms, tokens)
            substituted.extend(split_term_runs([replace(run, text=tokenized)], tokens))

        def fixed(run: TextRun) -> bool:
            return run.image is not None or run.term_key is not None or run.code or run.text == "\n"

        cleaned: list[TextRun] = []
        for run in substituted:
            text = run.text
            if not fixed(run):
                previous = cleaned[-1] if cleaned else None
                if (
                    previous is None
                    or previous.text == "\n"
                    or (previous.image is None and previous.text.endswith(" "))
                ):
                    text = text.lstrip(" ")
            if text or run.image is not None or run.term_key is not None:
                cleaned.append(replace(run, text=text))
        for index, run in enumerate(cleaned):
            following = cleaned[index + 1] if index + 1 < len(cleaned) else None
            if (following is None or following.text == "\n") and not fixed(run):
                cleaned[index] = replace(run, text=run.text.rstrip(" "))
        cleaned = [
            run for run in cleaned if run.text or run.image is not None or run.term_key is not None
        ]
        while cleaned and cleaned[0].text == "\n":
            cleaned.pop(0)
        while cleaned and cleaned[-1].text == "\n":
            cleaned.pop()
        # Cross-references follow the rules of Markdown text (code holds none).
        return TextParser.mark_cross_references(cleaned)


class HtmlTableParser:
    """Read an HTML `<table>` block, as document converters write them.

    `colspan`/`rowspan` become the engine's `<<`/`^^` span markers, so the
    table renders merged whatever the profile's span default. Cells are built
    as runs, so HTML text is literal (never re-read as Markdown): `<br>`, `<p>`,
    `<div>` and `<li>` start cell lines; `<b>`/`<strong>`, `<i>`/`<em>`,
    `<code>`, `<sup>`/`<sub>`, colour spans and `<a href>` format text; `<img>`
    is an inline image; `{{key}}` term references are substituted. A nested
    table is flattened into its cell, one line per row. Empty rows are kept
    while spans are placed, then dropped when nothing covers them.

    The header is the first row plus any rows its row spans cover; every header
    row is drawn and repeated as a header (`Table.header_rows`).
    """

    START_RE = re.compile(r"^\s*<table\b", re.IGNORECASE)
    _OPEN_RE = re.compile(r"<table\b", re.IGNORECASE)
    _CLOSE_RE = re.compile(r"</table\s*>", re.IGNORECASE)

    @classmethod
    def block_end(cls, lines: list[str], start: int) -> int | None:
        """Return the index after the line closing the table opened at `start`.

        Args:
            lines: Document lines.
            start: Index of the line that opens the table.

        Returns:
            The end index, or None when the table is never closed.
        """
        depth = 0
        for index in range(start, len(lines)):
            depth += len(cls._OPEN_RE.findall(lines[index])) - len(
                cls._CLOSE_RE.findall(lines[index])
            )
            if depth <= 0:
                return index + 1
        return None

    @staticmethod
    def _table_cell(runs: list[TextRun]) -> TableCell:
        """A prepared literal cell; `content` is its text with breaks as `<br>`."""
        content = "".join(run.text for run in runs).replace("\n", "<br>")
        return TableCell(content=content, runs=runs, literal=True)

    @classmethod
    def parse(
        cls,
        html: str,
        terms: TermResolver | None = None,
        limits: InputLimits = DEFAULT_LIMITS,
    ) -> tuple[list[list[str | TableCell]], int, list[str]]:
        """Convert a complete table block into rows of cells.

        Args:
            html: The `<table>...</table>` source.
            terms: Active term resolver, or None when the document has no `terms:`.
            limits: Operator-only expanded-grid limit, checked before allocating it.

        Returns:
            The rows (covered positions hold span markers), how many leading
            rows form the header, and warnings about content that could not be
            placed.
        """
        reader = _HtmlTableReader(terms)
        reader.feed(html)
        reader.close()
        warnings = [
            f"Text outside the HTML table cells was dropped: {text}" for text in reader.stray
        ]
        if reader.extra_tables:
            warnings.append(
                f"{reader.extra_tables} further HTML table(s) on the closing line were dropped; "
                "start each table on its own line"
            )

        runs: dict[tuple[int, int], list[TextRun]] = {}
        anchors: dict[tuple[int, int], tuple[int, int]] = {}
        spans: dict[tuple[int, int], int] = {}
        rows = reader.rows
        if any(
            cell.colspan > reader._MAX_SPAN or cell.rowspan > reader._MAX_SPAN
            for row in rows
            for cell in row
        ):
            cls._check_unclamped_grid(rows, limits)
        for row_index, html_row in enumerate(rows):
            position = 0
            for cell in html_row:
                while (row_index, position) in anchors:
                    position += 1
                # Retain the existing Word span cap, after checking the written grid.
                rowspan = min(cell.rowspan, reader._MAX_SPAN, len(rows) - row_index)
                colspan = min(cell.colspan, reader._MAX_SPAN)
                limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(rows) * (position + colspan))
                spans[(row_index, position)] = rowspan
                runs[(row_index, position)] = reader.finish(cell.runs)
                for down in range(rowspan):
                    for right in range(colspan):
                        anchors[(row_index + down, position + right)] = (row_index, position)
                position += colspan
        if not anchors:
            return [], 0, warnings
        col_count = max(covered for _, covered in anchors) + 1
        kept = [
            row
            for row in range(len(rows))
            if any((row, column) in anchors for column in range(col_count))
        ]

        header_end = kept[0] + 1
        grown = True
        while grown:
            grown = False
            for (top, _), rowspan in spans.items():
                if top < header_end < top + rowspan:
                    header_end, grown = top + rowspan, True

        def source(row: int, column: int) -> str | TableCell:
            anchor = anchors.get((row, column))
            if anchor is None:
                return ""
            if anchor == (row, column):
                return cls._table_cell(runs[anchor])
            return "^^" if column == anchor[1] else "<<"

        cells = [[source(row, column) for column in range(col_count)] for row in kept]
        return cells, sum(1 for row in kept if row < header_end), warnings

    @staticmethod
    def _check_unclamped_grid(rows: list[list[_HtmlCell]], limits: InputLimits) -> None:
        """Count written spans before Word's existing span cap can reduce them."""
        occupied: set[tuple[int, int]] = set()
        for row_index, row in enumerate(rows):
            column = 0
            for cell in row:
                while (row_index, column) in occupied:
                    column += 1
                limits.check(RejectionCode.LIMIT_TABLE_CELLS, len(rows) * (column + cell.colspan))
                occupied.update(
                    (row_index + down, column + right)
                    for down in range(min(cell.rowspan, len(rows) - row_index))
                    for right in range(cell.colspan)
                )
                column += cell.colspan

    @staticmethod
    def image(line: str) -> Image | None:
        """Return the image of a line that is a single HTML `<img>` tag.

        Args:
            line: Stripped document line.

        Returns:
            The image, or None when the line is not one `<img>` with a usable `src`.
        """
        found: list[dict[str, str]] = []

        class _Reader(HTMLParser):
            def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
                if tag == "img":
                    found.append({name: value or "" for name, value in attrs})

        reader = _Reader(convert_charrefs=True)
        reader.feed(line)
        reader.close()
        if len(found) != 1:
            return None
        return _html_image(found[0].get("src", ""), found[0].get("alt", ""))
