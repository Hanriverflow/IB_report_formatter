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


def test_render_to_markdown_normalizes_windows_image_paths():
    """Windows-style image prefixes should be normalized for Markdown links."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.IMAGE,
                content=Image(alt_text="Chart", path="chart.png"),
            )
        ]
    )

    relative_rendered = render_to_markdown(
        model,
        include_frontmatter=False,
        image_path_prefix=r"images\charts",
    )
    absolute_rendered = render_to_markdown(
        model,
        include_frontmatter=False,
        image_path_prefix=r"C:\tmp\images",
    )

    assert relative_rendered.strip() == "![Chart](images/charts/chart.png)"
    assert absolute_rendered.strip() == "![Chart](C:/tmp/images/chart.png)"


def test_render_to_markdown_appends_endnotes_block():
    """Footnotes should be rendered in parser-compatible endnotes form."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="본문", runs=[TextRun(text="본문"), TextRun(text="1", superscript=True)]),
            )
        ],
        footnotes={1: "Source A"},
    )

    rendered = render_to_markdown(model, include_frontmatter=False)

    assert "^1^" in rendered
    assert "## Citations" in rendered
    assert "1. Source A" in rendered
