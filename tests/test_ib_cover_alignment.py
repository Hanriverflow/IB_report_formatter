"""IB cover and disclaimer paragraphs keep their own alignment.

The IB Body style's default alignment may change (it is left-aligned since the
Korean typesetting change); the cover title and subtitle, the cover disclaimer
box, the disclaimer page text (which uses IB Body) and the copyright line set
their alignment directly and must keep it whatever the style default is.
"""

from pathlib import Path

import pytest
import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from document_profiles import RenderOptions
from ib_renderer import CoverRenderer, DisclaimerRenderer
from md_to_word import IBReportConverter

SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "profiles" / "ib-report.md"


@pytest.fixture(scope="module")
def rendered(tmp_path_factory):
    output = tmp_path_factory.mktemp("ib-cover-alignment") / "ib-report.docx"
    IBReportConverter(str(SAMPLE), str(output), RenderOptions(strict=True)).convert()
    return Document(str(output))


def _alignments(paragraphs, text):
    return [paragraph.alignment for paragraph in paragraphs if paragraph.text == text]


def test_cover_title_and_subtitle_are_centred(rendered):
    metadata = yaml.safe_load(SAMPLE.read_text(encoding="utf-8").split("---", 2)[1])
    assert _alignments(rendered.paragraphs, metadata["title"]) == [WD_ALIGN_PARAGRAPH.CENTER]
    assert _alignments(rendered.paragraphs, metadata["subtitle"]) == [WD_ALIGN_PARAGRAPH.CENTER]


def test_cover_disclaimer_box_is_justified(rendered):
    cells = [
        paragraph
        for table in rendered.tables
        for cell in table._cells
        for paragraph in cell.paragraphs
    ]
    assert _alignments(cells, CoverRenderer._COVER_DISCLAIMER_TEXT) == [WD_ALIGN_PARAGRAPH.JUSTIFY]


def test_disclaimer_page_text_is_justified(rendered):
    lines = [
        line
        for _, content in DisclaimerRenderer._SECTIONS
        for line in DisclaimerRenderer._split_content_lines(content)
    ]
    assert lines
    for line in lines:
        matches = [paragraph for paragraph in rendered.paragraphs if paragraph.text == line]
        assert [paragraph.style.name for paragraph in matches] == ["IB Body"], line
        assert [paragraph.alignment for paragraph in matches] == [WD_ALIGN_PARAGRAPH.JUSTIFY], line


def test_copyright_line_is_centred(rendered):
    alignments = [
        paragraph.alignment
        for paragraph in rendered.paragraphs
        if paragraph.text.startswith("(c) ") and paragraph.text.endswith("All rights reserved.")
    ]
    assert alignments == [WD_ALIGN_PARAGRAPH.CENTER]
