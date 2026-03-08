"""Tests for md_renderer module."""

from md_parser import (
    DocumentMetadata,
    DocumentModel,
    Element,
    ElementType,
    Image,
    Paragraph,
    TextRun,
)
from md_renderer import render_to_markdown


def test_render_to_markdown_strips_inline_formatting():
    """Strip mode should preserve text while removing markdown markers."""
    model = DocumentModel(
        metadata=DocumentMetadata(title="Report"),
        elements=[
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(
                    text="Bold text",
                    runs=[TextRun(text="Bold", bold=True), TextRun(text=" text")],
                ),
            )
        ],
    )

    rendered = render_to_markdown(model, include_frontmatter=False, strip_formatting=True)

    assert rendered.strip() == "Bold text"
    assert "**" not in rendered


def test_render_to_markdown_prefixes_image_paths():
    """Image path prefixes should be applied in rendered markdown."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.IMAGE,
                content=Image(alt_text="Chart", path="chart.png"),
            )
        ]
    )

    rendered = render_to_markdown(
        model,
        include_frontmatter=False,
        image_path_prefix="images",
    )

    assert rendered.strip() == "![Chart](images/chart.png)"
