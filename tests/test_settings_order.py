"""Settings order regressions through parsing, rendering and DOCX serialization."""

from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import parse_markdown_file

ROOT = Path(__file__).resolve().parents[1]

# CT_Settings order for children in python-docx's default template, independently
# transcribed from python-docx 1.2.0 docx/oxml/settings.py CT_Settings._tag_seq.
# The final two children are the template's ignorable Word 2010 extensions.
EXPECTED_SETTINGS = (
    "w:zoom",
    "w:proofState",
    "w:defaultTabStop",
    "w:characterSpacingControl",
    "w:savePreviewPicture",
    "w:updateFields",
    "w:compat",
    "w:rsids",
    "m:mathPr",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotAutoCompressPictures",
    "w:shapeDefaults",
    "w:decimalSymbol",
    "w:listSeparator",
    "w14:docId",
    "w14:defaultImageDpi",
)


@pytest.mark.parametrize(
    "profile",
    [
        "ib-report",
        "ib-memo",
        "plain",
        "office-letter",
        "business-report",
        "meeting-minutes",
        "term-sheet",
    ],
)
def test_rendered_settings_follow_schema_order_without_other_changes(profile: str) -> None:
    """Keep updateFields at its schema position and preserve every template child."""
    model = parse_markdown_file(str(ROOT / "samples" / "profiles" / f"{profile}.md"))
    document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    payload = BytesIO()
    document.save(payload)
    settings = Document(payload).settings.element

    assert [child.tag for child in settings] == [qn(tag) for tag in EXPECTED_SETTINGS]
    update_fields = settings.findall(qn("w:updateFields"))
    assert len(update_fields) == 1
    assert dict(update_fields[0].attrib) == {qn("w:val"): "true"}
    assert len(update_fields[0]) == 0
    settings.remove(update_fields[0])
    assert etree.tostring(settings) == etree.tostring(Document().settings.element)
