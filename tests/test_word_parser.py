"""Tests for word_parser module."""

import base64
from pathlib import Path

from docx import Document

from md_parser import ElementType
from word_parser import parse_word_file

PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Z8eQAAAAASUVORK5CYII="
)


def _save_doc(doc: Document, output_path: Path) -> Path:
    """Save a temporary docx fixture and return its path."""
    doc.save(str(output_path))
    return output_path


def _write_png(output_path: Path) -> Path:
    """Write a tiny PNG fixture used for image extraction tests."""
    output_path.write_bytes(PNG_BYTES)
    return output_path


def test_parse_preserves_paragraph_table_paragraph_order(tmp_path):
    """Paragraph-table-paragraph order should survive Word parsing."""
    doc = Document()
    doc.add_paragraph("Opening paragraph")

    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Header"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Revenue"
    table.cell(1, 1).text = "100"

    doc.add_paragraph("Closing paragraph")
    docx_path = _save_doc(doc, tmp_path / "ordered.docx")

    model = parse_word_file(str(docx_path), extract_images=False)

    assert [element.element_type for element in model.elements[:3]] == [
        ElementType.PARAGRAPH,
        ElementType.TABLE,
        ElementType.PARAGRAPH,
    ]
    assert model.elements[0].raw_text == "Opening paragraph"
    assert model.elements[2].raw_text == "Closing paragraph"


def test_parse_callout_table_in_document_order(tmp_path):
    """A single-cell callout table should become one blockquote in place."""
    doc = Document()
    doc.add_paragraph("Before callout")

    callout = doc.add_table(rows=1, cols=1)
    callout.cell(0, 0).text = "요약: 핵심 포인트"

    doc.add_paragraph("After callout")
    docx_path = _save_doc(doc, tmp_path / "callout.docx")

    model = parse_word_file(str(docx_path), extract_images=False)

    assert [element.element_type for element in model.elements[:3]] == [
        ElementType.PARAGRAPH,
        ElementType.BLOCKQUOTE,
        ElementType.PARAGRAPH,
    ]
    assert model.elements[1].content.title == "요약"
    assert model.elements[1].content.text == "핵심 포인트"


def test_parse_image_paragraph_with_extraction(tmp_path):
    """Image-only paragraphs should become image elements with extracted files."""
    image_path = _write_png(tmp_path / "tiny.png")
    image_dir = tmp_path / "images"

    doc = Document()
    doc.add_paragraph("Chart below")
    doc.add_picture(str(image_path))
    docx_path = _save_doc(doc, tmp_path / "images.docx")

    model = parse_word_file(
        str(docx_path),
        extract_images=True,
        image_output_dir=str(image_dir),
    )

    image_elements = [element for element in model.elements if element.element_type == ElementType.IMAGE]
    assert len(image_elements) == 1
    assert image_elements[0].content.path.startswith("image_")
    assert list(image_dir.glob("image_*"))
