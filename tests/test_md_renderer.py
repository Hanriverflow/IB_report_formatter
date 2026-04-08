"""Tests for md_renderer module."""

import yaml

from md_parser import (
    Blockquote,
    CodeBlock,
    Diagram,
    DiagramArrow,
    DiagramBox,
    DocumentMetadata,
    DocumentModel,
    Element,
    ElementType,
    FrontmatterParser,
    Heading,
    Image,
    LaTeXEquation,
    ListItem,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
)
from md_renderer import _normalize_markdown, render_to_markdown


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


def test_render_to_markdown_preserves_subscript_runs():
    """Subscript runs should round-trip to ~text~ markers."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(
                    text="H2O",
                    runs=[
                        TextRun(text="H"),
                        TextRun(text="2", subscript=True),
                        TextRun(text="O"),
                    ],
                ),
            )
        ]
    )

    rendered = render_to_markdown(model, include_frontmatter=False)

    assert rendered.strip() == "H~2~O"


def test_render_to_markdown_uses_table_cell_runs():
    """Table cells should prefer structured runs over plain content."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.TABLE,
                content=Table(
                    rows=[
                        TableRow(cells=[TableCell(content="Metric"), TableCell(content="Value")]),
                        TableRow(
                            cells=[
                                TableCell(
                                    content="Revenue",
                                    runs=[TextRun(text="Revenue", bold=True)],
                                ),
                                TableCell(
                                    content="H2O",
                                    runs=[
                                        TextRun(text="H"),
                                        TextRun(text="2", subscript=True),
                                        TextRun(text="O"),
                                    ],
                                ),
                            ]
                        ),
                    ]
                ),
            )
        ]
    )

    rendered = render_to_markdown(model, include_frontmatter=False)

    assert "| **Revenue** | H~2~O |" in rendered


def test_render_to_markdown_embeds_base64_images():
    """Embedded images should render as data URIs when requested."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.IMAGE,
                content=Image(
                    alt_text="Chart",
                    path="chart.png",
                    base64_data="YWJj",
                    mime_type="image/png",
                ),
            )
        ]
    )

    rendered = render_to_markdown(
        model,
        include_frontmatter=False,
        embed_images_base64=True,
    )

    assert rendered.strip() == "![Chart](data:image/png;base64,YWJj)"


def test_render_to_markdown_preserves_run_colors():
    """Colored runs should render as HTML spans with inline color."""
    model = DocumentModel(
        elements=[
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(
                    text="Loss",
                    runs=[TextRun(text="Loss", bold=True, color_hex="#C00000")],
                ),
            )
        ]
    )

    rendered = render_to_markdown(model, include_frontmatter=False)

    assert rendered.strip() == '<span style="color:#C00000">**Loss**</span>'


def test_render_to_markdown_frontmatter_quotes_are_yaml_safe():
    """Regression: quoted metadata must render as valid YAML frontmatter."""
    model = DocumentModel(
        metadata=DocumentMetadata(
            title='He said "Hi"',
            extra={"summary": 'Quoted "metadata" stays parseable'},
        )
    )

    rendered = render_to_markdown(model, include_frontmatter=True)
    metadata, remaining = FrontmatterParser.parse(rendered.splitlines())

    assert metadata.title == 'He said "Hi"'
    assert metadata.extra["summary"] == 'Quoted "metadata" stays parseable'
    assert remaining == []


def test_render_to_markdown_preserves_pipe_lines_inside_fenced_code_blocks():
    """Regression: table normalization must not rewrite fenced code that contains pipes."""
    model = DocumentModel(
        elements=[
            Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="Before")),
            Element(
                element_type=ElementType.CODE_BLOCK,
                content=CodeBlock(
                    language="text",
                    code="| not | a | table |\n| still | code |",
                ),
            ),
            Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="After")),
        ]
    )

    rendered = render_to_markdown(model, include_frontmatter=False)
    fenced_body = rendered.split("```text\n", 1)[1].split("\n```", 1)[0]

    assert fenced_body == "| not | a | table |\n| still | code |"


def test_render_to_markdown_renders_diagram_elements_as_fenced_yaml():
    """Regression: DIAGRAM elements should render as parser-compatible fenced blocks."""
    diagram = Diagram(
        diagram_type="flow",
        title="Credit Flow",
        boxes=[
            DiagramBox(id="start", label="Start", pos=[0, 0], style="entry"),
            DiagramBox(id="end", label="End", pos=[1, 0], style="exit"),
        ],
        arrows=[DiagramArrow(from_id="start", to_id="end", label="step", style="dashed")],
        notes=["Check covenants"],
    )
    model = DocumentModel(
        elements=[Element(element_type=ElementType.DIAGRAM, content=diagram)]
    )

    rendered = render_to_markdown(model, include_frontmatter=False)
    payload = rendered.split("```diagram:flow\n", 1)[1].split("\n```", 1)[0]
    parsed = yaml.safe_load(payload)

    assert rendered.startswith("```diagram:flow\n")
    assert parsed["title"] == "Credit Flow"
    assert parsed["boxes"][0]["id"] == "start"
    assert parsed["arrows"][0]["from"] == "start"
    assert parsed["notes"] == ["Check covenants"]


# ═══════════════════════════════════════════════════════════════════════════════
# OUTPUT NORMALIZATION TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestNormalizeMarkdown:
    """Tests for _normalize_markdown post-processing."""

    def test_trailing_whitespace_stripped(self):
        result = _normalize_markdown("hello   \nworld\t\n")
        assert "   \n" not in result
        assert "\t\n" not in result
        assert result == "hello\nworld\n"

    def test_excessive_blank_lines_compressed(self):
        result = _normalize_markdown("one\n\n\n\n\ntwo\n")
        assert result == "one\n\ntwo\n"

    def test_single_blank_line_preserved(self):
        result = _normalize_markdown("one\n\ntwo\n")
        assert result == "one\n\ntwo\n"

    def test_math_block_gets_surrounding_blanks(self):
        result = _normalize_markdown("text\n$$\nx^2\n$$\nmore text\n")
        # Should have blank line before and after math block
        assert "\n\n$$\nx^2\n$$\n\n" in result

    def test_table_block_gets_surrounding_blanks(self):
        result = _normalize_markdown("text\n| A | B |\n| --- | --- |\n| 1 | 2 |\nmore\n")
        # Table should be surrounded by blank lines
        lines = result.split("\n")
        table_start = next(i for i, l in enumerate(lines) if l.startswith("| A"))
        # Line before table should be blank
        assert lines[table_start - 1] == ""

    def test_final_newline_ensured(self):
        result = _normalize_markdown("hello")
        assert result.endswith("\n")
        assert not result.endswith("\n\n")

    def test_leading_whitespace_stripped(self):
        result = _normalize_markdown("\n\n\nhello\n")
        assert result == "hello\n"

    def test_fenced_code_block_gets_blanks(self):
        result = _normalize_markdown("text\n```python\nprint('hi')\n```\nmore\n")
        assert "\n\n```python\nprint('hi')\n```\n\n" in result

    def test_no_change_for_clean_input(self):
        """Already-clean markdown should pass through with minimal changes."""
        clean = "# Title\n\nParagraph text.\n\n- Item 1\n- Item 2\n"
        result = _normalize_markdown(clean)
        assert result == clean


class TestNormalizationIntegration:
    """Integration tests: full render pipeline produces normalized output."""

    def test_render_no_trailing_whitespace(self):
        model = DocumentModel(
            elements=[
                Element(
                    element_type=ElementType.HEADING_1,
                    content=Heading(text="Title", level=1),
                ),
                Element(
                    element_type=ElementType.PARAGRAPH,
                    content=Paragraph(text="Body text"),
                ),
            ]
        )
        rendered = render_to_markdown(model, include_frontmatter=False)
        for line in rendered.split("\n"):
            assert line == line.rstrip(), f"Trailing whitespace found: {line!r}"

    def test_render_no_triple_blank_lines(self):
        model = DocumentModel(
            elements=[
                Element(element_type=ElementType.HEADING_1, content=Heading(text="H1", level=1)),
                Element(element_type=ElementType.EMPTY, content=None),
                Element(element_type=ElementType.EMPTY, content=None),
                Element(element_type=ElementType.EMPTY, content=None),
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="After gaps")),
            ]
        )
        rendered = render_to_markdown(model, include_frontmatter=False)
        assert "\n\n\n" not in rendered

    def test_render_table_has_blank_lines(self):
        table = Table(
            rows=[
                TableRow(cells=[TableCell(content="A"), TableCell(content="B")]),
                TableRow(cells=[TableCell(content="1"), TableCell(content="2")]),
            ]
        )
        model = DocumentModel(
            elements=[
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="Before")),
                Element(element_type=ElementType.TABLE, content=table),
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="After")),
            ]
        )
        rendered = render_to_markdown(model, include_frontmatter=False)
        # Table should be set off from surrounding text
        lines = rendered.strip().split("\n")
        table_idx = next(i for i, l in enumerate(lines) if l.startswith("| A"))
        assert lines[table_idx - 1] == "", "Missing blank line before table"

    def test_render_math_block_has_blank_lines(self):
        model = DocumentModel(
            elements=[
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="Before")),
                Element(
                    element_type=ElementType.LATEX_BLOCK,
                    content=LaTeXEquation(expression="E=mc^2", is_block=True),
                ),
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="After")),
            ]
        )
        rendered = render_to_markdown(model, include_frontmatter=False)
        assert "\n\n$$\nE=mc^2\n$$\n\n" in rendered

    def test_strip_formatting_with_normalization(self):
        """Strip mode should still produce normalized output."""
        model = DocumentModel(
            elements=[
                Element(
                    element_type=ElementType.PARAGRAPH,
                    content=Paragraph(
                        text="Bold italic",
                        runs=[
                            TextRun(text="Bold", bold=True),
                            TextRun(text=" "),
                            TextRun(text="italic", italic=True),
                        ],
                    ),
                ),
            ]
        )
        rendered = render_to_markdown(model, include_frontmatter=False, strip_formatting=True)
        assert "**" not in rendered
        assert "*italic*" not in rendered
        assert "Bold italic" in rendered
        assert rendered.endswith("\n")

    def test_strip_formatting_removes_subscript_markers(self):
        """Strip mode should keep subscript text without markers."""
        model = DocumentModel(
            elements=[
                Element(
                    element_type=ElementType.PARAGRAPH,
                    content=Paragraph(
                        text="H2O",
                        runs=[
                            TextRun(text="H"),
                            TextRun(text="2", subscript=True),
                            TextRun(text="O"),
                        ],
                    ),
                ),
            ]
        )

        rendered = render_to_markdown(model, include_frontmatter=False, strip_formatting=True)

        assert rendered == "H2O\n"

    def test_strip_formatting_removes_color_spans(self):
        """Strip mode should keep text while dropping color markup."""
        model = DocumentModel(
            elements=[
                Element(
                    element_type=ElementType.PARAGRAPH,
                    content=Paragraph(
                        text="Loss",
                        runs=[TextRun(text="Loss", color_hex="#C00000")],
                    ),
                ),
            ]
        )

        rendered = render_to_markdown(model, include_frontmatter=False, strip_formatting=True)

        assert rendered == "Loss\n"

    def test_render_complex_document_normalized(self):
        """A complex document with mixed elements should be cleanly normalized."""
        model = DocumentModel(
            metadata=DocumentMetadata(title="Test Report"),
            elements=[
                Element(element_type=ElementType.HEADING_1, content=Heading(text="Summary", level=1)),
                Element(element_type=ElementType.PARAGRAPH, content=Paragraph(text="Introduction.")),
                Element(
                    element_type=ElementType.TABLE,
                    content=Table(
                        rows=[
                            TableRow(cells=[TableCell(content="Key"), TableCell(content="Value")]),
                            TableRow(cells=[TableCell(content="A"), TableCell(content="100")]),
                        ]
                    ),
                ),
                Element(element_type=ElementType.BULLET_LIST, content=ListItem(text="Point one")),
                Element(element_type=ElementType.BULLET_LIST, content=ListItem(text="Point two")),
                Element(
                    element_type=ElementType.BLOCKQUOTE,
                    content=Blockquote(title="NOTE", text="Important info"),
                ),
                Element(element_type=ElementType.SEPARATOR, content=None),
                Element(
                    element_type=ElementType.LATEX_BLOCK,
                    content=LaTeXEquation(expression="\\sum_{i=1}^n x_i", is_block=True),
                ),
            ],
            footnotes={1: "Source document"},
        )
        rendered = render_to_markdown(model, include_frontmatter=True)

        # Basic quality checks
        assert rendered.endswith("\n")
        assert "\n\n\n" not in rendered
        for line in rendered.split("\n"):
            assert line == line.rstrip(), f"Trailing whitespace: {line!r}"
        # All elements present
        assert "# Summary" in rendered
        assert "Introduction." in rendered
        assert "| Key | Value |" in rendered
        assert "- Point one" in rendered
        assert "> **[NOTE]**" in rendered
        assert "---" in rendered
        assert "$$" in rendered
        assert "## Citations" in rendered
