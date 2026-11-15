"""Regression tests for the 2026-11-15 audit follow-up, group B (Word writing).

B1  area2 F-01: a numeric column's unit is placed once, after an amount split
    across formatting runs, instead of inside it.
B2  area2 F-03: a fenced code block uses the theme's `CODE_FONT`, like inline code.
B3  area3 F-05: a diagram picture is no wider than the section's printable width.
B4  area3 F-02: `docx-audit` reads tabs, breaks and symbols in bookmarks and
    cached field results, so a stale reference that differs only by them is caught.
"""

from pathlib import Path

import pytest
from docx.document import Document as DocxDocument
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.table import _Cell

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FENCE = "`" * 3


def _render(markdown: str, profile: str, **options: object) -> tuple[DocxDocument, list[str]]:
    """Parse and render through the single composition path; return warnings too."""
    model = MarkdownParser(profile=profile).parse(markdown)
    renderer = IBDocumentRenderer(options=RenderOptions(profile=profile, **options))  # type: ignore[arg-type]
    document = renderer.render(model)
    assert renderer.errors == []
    return document, list(renderer.warnings)


def _body_cell(document: DocxDocument, row: int, column: int = 1) -> _Cell:
    return document.tables[0].rows[row].cells[column]


def _runs(cell: _Cell) -> list[tuple[str, bool]]:
    return [(run.text, bool(run.bold)) for run in cell.paragraphs[0].runs if run.text]


# ═══════════════════════════════════════════════════════════════════════════════
# B1: UNIT AFTER A SPLIT AMOUNT
# ═══════════════════════════════════════════════════════════════════════════════


def _role_table(role: str, *cells: str, extra: str = "") -> str:
    front = f"---\nprofile: plain\ntables:\n  - type: financial\n    columns: [text, {role}]\n{extra}---\n"
    rows = "".join(f"| R{index} | {cell} |\n" for index, cell in enumerate(cells))
    return front + "| Item | Value |\n|---|---|\n" + rows


@pytest.mark.parametrize(
    ("role", "shown"),
    [("percent", "12.5%"), ("bps", "12.5 bps"), ("multiple", "12.5x")],
)
def test_a_split_amount_gets_its_unit_once_after_the_amount(role: str, shown: str) -> None:
    document, warnings = _render(_role_table(role, "**12**.5", "12.5"), "plain")
    split, whole = _body_cell(document, 1), _body_cell(document, 2)
    assert split.text == whole.text == shown
    # Every run keeps its own format and characters; the unit follows the last one.
    assert _runs(split) == [("12", True), (shown[2:], False)]
    assert warnings == []


def test_a_split_amount_gets_thousands_separators_in_its_own_runs() -> None:
    markdown = _role_table("number", "**1234**567", "-**1234**.5")
    document, warnings = _render(markdown, "plain")
    assert _runs(_body_cell(document, 1)) == [("1,234", True), (",567", False)]
    assert _runs(_body_cell(document, 2)) == [("-", False), ("1,234", True), (".5", False)]
    assert warnings == []


def test_a_split_amount_written_with_its_unit_is_not_given_another() -> None:
    document, warnings = _render(_role_table("percent", "**12**%", "12.**5**%"), "plain")
    assert _body_cell(document, 1).text == "12%"
    assert _body_cell(document, 2).text == "12.5%"
    assert warnings == []


def test_a_split_amount_that_cannot_be_spread_back_keeps_its_text_and_warns() -> None:
    # Formatting drops the leading zeros, which belong to a differently
    # formatted run: the written text is kept and the cell is named.
    document, warnings = _render(_role_table("percent", "**00**12"), "plain")
    assert _runs(_body_cell(document, 1)) == [("00", True), ("12", False)]
    assert warnings == [
        "Table number formatting not applied to '0012': the amount is split across "
        "differently formatted text, so it keeps its written form"
    ]


def test_footnote_and_space_runs_around_a_split_amount_keep_their_place() -> None:
    markdown = _role_table("percent", "**12**.5[^1]") + "\n[^1]: Fictional note.\n"
    document, warnings = _render(markdown, "plain")
    paragraph = _body_cell(document, 1).paragraphs[0]
    assert _runs(_body_cell(document, 1)) == [("12", True), (".5%", False)]
    # The footnote reference mark stays after the amount.
    runs = paragraph._p.findall(f"{W}r")
    assert runs[-1].find(f"{W}footnoteReference") is not None
    assert warnings == []


def test_a_negative_style_table_still_restyles_and_warns_as_before() -> None:
    extra = "    negative: paren\n"
    markdown = _role_table("number", "-1234", "-**12**34", extra=extra)
    document, warnings = _render(markdown, "plain")
    assert _body_cell(document, 1).text == "(1,234)"
    # A split negative amount keeps its written form under the style and warns;
    # its number formatting is now read from the whole amount.
    assert _runs(_body_cell(document, 2)) == [("-", False), ("1,2", True), ("34", False)]
    assert warnings == [
        "Table negative style not applied to '-1234': the amount is split across "
        "differently formatted text, so it keeps its written form"
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# B2: CODE BLOCK FONT FROM THE THEME
# ═══════════════════════════════════════════════════════════════════════════════


def test_a_fenced_code_block_uses_the_theme_code_font(tmp_path: Path) -> None:
    theme = tmp_path / "theme.yaml"
    theme.write_text("CODE_FONT: Courier New\n", encoding="utf-8")
    markdown = f"---\nprofile: plain\n---\ninline `x` here\n\n{FENCE}\nroot\n├── a: 1\n{FENCE}\n"
    document, _ = _render(markdown, "plain", theme=str(theme))
    inline = [run for run in document.paragraphs[0].runs if run.text == "x"]
    assert [run.font.name for run in inline] == ["Courier New"]
    block = [
        run
        for paragraph in document.tables[0].rows[0].cells[0].paragraphs
        for run in paragraph.runs
    ]
    assert block and {run.font.name for run in block} == {"Courier New"}


def test_a_fenced_code_block_keeps_consolas_by_default() -> None:
    document, _ = _render(f"---\nprofile: plain\n---\n{FENCE}\nx\n{FENCE}\n", "plain")
    block = list(document.tables[0].rows[0].cells[0].paragraphs[0].runs)
    assert {run.font.name for run in block} == {"Consolas"}


# ═══════════════════════════════════════════════════════════════════════════════
# B3: DIAGRAM WIDTH
# ═══════════════════════════════════════════════════════════════════════════════

DIAGRAM = (
    f"---\nprofile: business-report\ntitle: T\n---\n{FENCE}diagram:flow\nboxes:\n"
    "  - {id: a, label: Start, pos: [0, 0]}\n  - {id: b, label: End, pos: [3, 0]}\n"
    f"arrows:\n  - {{from: a, to: b, label: Go}}\n{FENCE}\n"
)


def _diagram_widths(theme: Path | None) -> tuple[int, list[int]]:
    model = MarkdownParser().parse(DIAGRAM)
    options = RenderOptions(theme=str(theme)) if theme else RenderOptions()
    document = IBDocumentRenderer(options=options).render(model)
    section = document.sections[-1]
    printable = int(section.page_width) - int(section.left_margin) - int(section.right_margin)
    return printable, [int(shape.width) for shape in document.inline_shapes]


def test_a_diagram_is_no_wider_than_the_printable_width(tmp_path: Path) -> None:
    pytest.importorskip("matplotlib")
    theme = tmp_path / "theme.yaml"
    theme.write_text("margin_mm: 40\n", encoding="utf-8")
    printable, widths = _diagram_widths(theme)
    assert printable < 6 * 914400
    assert widths == [printable]


def test_a_diagram_in_a_wide_body_keeps_six_inches() -> None:
    pytest.importorskip("matplotlib")
    printable, widths = _diagram_widths(None)
    assert printable >= 6 * 914400
    assert widths == [6 * 914400]


# ═══════════════════════════════════════════════════════════════════════════════
# B4: AUDIT READS TABS, BREAKS AND SYMBOLS IN REFERENCES
# ═══════════════════════════════════════════════════════════════════════════════


def _audit_document(*paragraphs: str) -> DocxDocument:
    from docx import Document

    document = Document()
    for inner in paragraphs:
        document.element.body.append(parse_xml(f"<w:p {nsdecls('w')}>{inner}</w:p>"))
    return document


def _complex(instruction: str, result: str) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"{result}"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def _simple(instruction: str, result: str) -> str:
    return f'<w:fldSimple w:instr=" {instruction} ">{result}</w:fldSimple>'


def _target(content: str) -> str:
    return f'<w:bookmarkStart w:id="1" w:name="_Ref1"/>{content}<w:bookmarkEnd w:id="1"/>'


TAB_TARGET = "<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>"
STALE = (
    "Cross-reference REF _Ref1 shows 'AB' but its target reads 'A\\tB'; "
    "update fields in Word (select all, F9)"
)


@pytest.mark.parametrize("form", [_complex, _simple])
@pytest.mark.parametrize(
    ("target", "result", "stale"),
    [
        (TAB_TARGET, "<w:r><w:t>AB</w:t></w:r>", True),
        (TAB_TARGET, "<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>", False),
        ("<w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r>", "<w:r><w:t>AB</w:t></w:r>", True),
        (
            '<w:r><w:t>A</w:t><w:sym w:font="Symbol" w:char="F062"/></w:r>',
            "<w:r><w:t>A</w:t></w:r>",
            True,
        ),
        (
            "<w:r><w:t>AB</w:t></w:r>",
            "<w:r><w:t>A</w:t></w:r><w:del><w:r><w:tab/></w:r></w:del><w:r><w:t>B</w:t></w:r>",
            False,
        ),
    ],
)
def test_audit_compares_visible_run_content_of_targets_and_results(
    form, target: str, result: str, stale: bool
) -> None:
    from docx_audit import inspect_document

    audit = inspect_document(_audit_document(_target(target), form("REF _Ref1 \\h", result)))
    assert audit.issues == []
    assert bool(audit.warnings) is stale
    if target == TAB_TARGET and stale:
        assert audit.warnings == [STALE]


def test_audit_catches_a_rendered_reference_whose_caption_gained_a_tab(tmp_path: Path) -> None:
    from docx_audit import inspect_document

    path = tmp_path / "doc.md"
    path.write_text(
        "---\nprofile: business-report\ntitle: 시험 문서\ntables:\n"
        "  - id: tbl-debt\n    caption: 가상 차입 내역\n---\n\n# 시험 문서\n\n"
        "앞 @tbl-debt.\n\n| 가 | 나 |\n|---|---:|\n| 1 | 2 |\n",
        encoding="utf-8",
    )
    document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(
        parse_markdown_file(str(path))
    )
    assert inspect_document(document).warnings == []
    # As after typing a tab into the caption label in Word without updating fields.
    start = next(
        node
        for node in document.element.body.iter(f"{W}bookmarkStart")
        if node.get(qn("w:name")) == "_Ref_tbl_debt"
    )
    # After the label run, so the tab is inside the target text, not at its edge.
    start.getnext().addnext(parse_xml(f"<w:r {nsdecls('w')}><w:tab/></w:r>"))
    warnings = inspect_document(document).warnings
    assert len(warnings) == 1 and warnings[0].startswith("Cross-reference REF _Ref_tbl_debt shows")
