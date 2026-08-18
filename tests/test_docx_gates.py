"""
Replacement verification gates for the Markdown-to-Word pipeline.

Covered scope:
- Exact DocumentModel element snapshots for the two Korean report fixtures
- Embedded table and heading-family count invariants
- End-to-end DOCX rendering through IBReportConverter
- Rendered tables, semantic bookmarks, and Word heading styles
"""

from pathlib import Path
from typing import List

from docx import Document

from md_parser import parse_markdown_file
from md_to_word import IBReportConverter


HEADING_ELEMENT_TYPES = {
    "HEADING_1",
    "HEADING_2",
    "HEADING_3",
    "HEADING_4",
    "NUMBERED_HEADING",
}

FIXTURE_DIR = Path(__file__).resolve().parent
WOONGJIN_FIXTURE = FIXTURE_DIR / "웅진_계열사.md"
ILDONG_FIXTURE = FIXTURE_DIR / "일동제약_수익성분석.md"

WOONGJIN_ELEMENT_TYPES = [
    "PARAGRAPH", "SEPARATOR", "HEADING_2", "PARAGRAPH", "PARAGRAPH",
    "SEPARATOR", "HEADING_2", "PARAGRAPH", "HEADING_3", "TABLE",
    "HEADING_3", "TABLE", "HEADING_3", "TABLE", "HEADING_3", "PARAGRAPH",
    "TABLE", "HEADING_3", "TABLE", "SEPARATOR", "HEADING_2", "CODE_BLOCK",
    "SEPARATOR", "HEADING_2", "PARAGRAPH", "HEADING_3", "PARAGRAPH",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "HEADING_3",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH",
    "HEADING_3", "PARAGRAPH", "PARAGRAPH", "HEADING_3", "PARAGRAPH",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH",
    "HEADING_3", "PARAGRAPH", "SEPARATOR", "HEADING_2", "PARAGRAPH",
    "PARAGRAPH", "HEADING_2", "TABLE", "PARAGRAPH", "SEPARATOR",
    "HEADING_2", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH",
]

ILDONG_ELEMENT_TYPES = [
    "SEPARATOR", "HEADING_2", "PARAGRAPH", "SEPARATOR", "HEADING_2",
    "HEADING_3", "TABLE", "PARAGRAPH", "HEADING_3", "TABLE", "PARAGRAPH",
    "SEPARATOR", "HEADING_2", "HEADING_3", "TABLE", "PARAGRAPH",
    "HEADING_3", "TABLE", "HEADING_3", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "PARAGRAPH",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "SEPARATOR", "HEADING_2",
    "HEADING_3", "TABLE", "PARAGRAPH", "HEADING_3", "TABLE", "PARAGRAPH",
    "SEPARATOR", "HEADING_2", "TABLE", "PARAGRAPH", "PARAGRAPH",
    "SEPARATOR", "HEADING_2", "HEADING_3", "PARAGRAPH", "HEADING_3",
    "PARAGRAPH", "HEADING_3", "PARAGRAPH", "SEPARATOR", "HEADING_2",
    "HEADING_3", "PARAGRAPH", "HEADING_3", "PARAGRAPH", "PARAGRAPH",
    "PARAGRAPH", "PARAGRAPH", "PARAGRAPH", "HEADING_3", "PARAGRAPH",
]


def _semantic_bookmark_names(doc: Document):
    """Collect ib-report semantic bookmark names from a rendered document."""
    bookmark_starts = doc.element.xpath(".//*[local-name()='bookmarkStart']")
    names = []
    for bookmark in bookmark_starts:
        name = bookmark.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name",
            "",
        )
        if name.startswith("_ibrep_"):
            names.append(name)
    return names


def _element_type_names(fixture_path: Path) -> List[str]:
    """Return the ordered element-type names parsed from a Markdown fixture."""
    model = parse_markdown_file(fixture_path)
    return [element.element_type.name for element in model.elements]


def _assert_model_snapshot(
    fixture_path: Path,
    expected_sequence: List[str],
    expected_length: int,
    expected_table_count: int,
    expected_heading_count: int,
) -> None:
    """Assert an exact element snapshot and its embedded count invariants."""
    assert len(expected_sequence) == expected_length
    assert expected_sequence.count("TABLE") == expected_table_count
    assert (
        sum(1 for name in expected_sequence if name in HEADING_ELEMENT_TYPES)
        == expected_heading_count
    )

    sequence = _element_type_names(fixture_path)

    assert sequence == expected_sequence
    assert sum(1 for name in sequence if name == "TABLE") == expected_table_count
    assert (
        sum(1 for name in sequence if name in HEADING_ELEMENT_TYPES)
        == expected_heading_count
    )


def _assert_rendered_docx(
    fixture_path: Path,
    output_path: Path,
    expected_table_count: int,
) -> None:
    """Render a fixture and assert structural properties in the saved DOCX."""
    rendered_path = IBReportConverter(
        md_file_path=str(fixture_path),
        output_path=str(output_path),
    ).convert()
    doc = Document(rendered_path)

    assert len(doc.tables) >= expected_table_count

    bookmark_names = _semantic_bookmark_names(doc)
    assert bookmark_names
    assert any(name.startswith("_ibrep_SEPARATOR_") for name in bookmark_names)
    # The current renderer adds semantic bookmarks for separators, code blocks,
    # and diagrams, but not tables; do not assert an unimplemented TABLE prefix.

    heading_paragraphs = [
        paragraph
        for paragraph in doc.paragraphs
        if paragraph.style is not None
        and paragraph.style.name
        and "Heading" in paragraph.style.name
    ]
    assert len(heading_paragraphs) > 0


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENTMODEL SNAPSHOT GATES
# ═══════════════════════════════════════════════════════════════════════════════


def test_woongjin_document_model_snapshot():
    """Keep the Woongjin fixture's parsed element structure stable."""
    _assert_model_snapshot(
        WOONGJIN_FIXTURE,
        WOONGJIN_ELEMENT_TYPES,
        expected_length=62,
        expected_table_count=6,
        expected_heading_count=17,
    )


def test_ildong_document_model_snapshot():
    """Keep the Ildong fixture's parsed element structure stable."""
    _assert_model_snapshot(
        ILDONG_FIXTURE,
        ILDONG_ELEMENT_TYPES,
        expected_length=63,
        expected_table_count=7,
        expected_heading_count=20,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# RENDERED DOCX XML GATES
# ═══════════════════════════════════════════════════════════════════════════════


def test_woongjin_rendered_docx_structure(tmp_path):
    """Render the Woongjin fixture and preserve its DOCX structural markers."""
    _assert_rendered_docx(
        WOONGJIN_FIXTURE,
        tmp_path / "woongjin-gate.docx",
        expected_table_count=6,
    )


def test_ildong_rendered_docx_structure(tmp_path):
    """Render the Ildong fixture and preserve its DOCX structural markers."""
    _assert_rendered_docx(
        ILDONG_FIXTURE,
        tmp_path / "ildong-gate.docx",
        expected_table_count=7,
    )
