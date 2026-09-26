"""A render-time profile override cannot recover a different parser's losses."""

import pytest

from converters import get_default_registry
from document_model import DocumentModel, Element, ElementType, Paragraph
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

SOURCE = "## References\n\n1. Source material\n\n| Code | Notes |\n|---|---|\n| 001234 | Keep |"


@pytest.mark.parametrize("parsed,requested", [("ib-report", "plain"), ("plain", "ib-memo")])
@pytest.mark.parametrize("strict", [False, True])
@pytest.mark.parametrize("via_registry", [False, True])
def test_profile_change_after_parsing_requires_reparse(parsed, requested, strict, via_registry):
    model = MarkdownParser(profile=parsed).parse(SOURCE)
    options = RenderOptions(profile=requested, strict=strict)
    with pytest.raises(ValueError, match="[Rr]epars"):
        if via_registry:
            get_default_registry().convert(model, output_format="docx", render_options=options)
        else:
            IBDocumentRenderer(options=options).render(model)
    assert model.metadata.profile == parsed


def test_parsing_with_requested_profile_preserves_source_and_columns():
    model = MarkdownParser(profile="plain").parse(SOURCE)
    document = IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model)
    assert model.parsed_profile == "plain"
    assert "Source material" in "\n".join(p.text for p in document.paragraphs)
    assert [[c.text for c in row.cells] for row in document.tables[0].rows] == [
        ["Code", "Notes"], ["001234", "Keep"]
    ]


def test_hand_built_model_can_select_render_profile_without_parser_provenance():
    model = DocumentModel(elements=[Element(ElementType.PARAGRAPH, Paragraph(text="Keep body"))])
    document = IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model)
    assert "Keep body" in "\n".join(p.text for p in document.paragraphs)
    assert model.parsed_profile is None
