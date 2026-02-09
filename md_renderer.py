"""
Markdown Renderer Module for Word to Markdown Converter
Handles rendering of DocumentModel to Markdown text.

Renders clean, LLM-friendly Markdown output.
"""

import logging
from dataclasses import dataclass
from typing import List, Dict, Optional, Union, Tuple, cast

from md_parser import (
    DocumentModel,
    DocumentMetadata,
    Element,
    ElementType,
    Heading,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    ListItem,
    Blockquote,
    Image,
    TextRun,
    LaTeXEquation,
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class RenderConfig:
    """Configuration for Markdown rendering."""

    include_frontmatter: bool = True
    embed_images_base64: bool = False
    image_path_prefix: str = ""
    table_alignment: bool = True
    bold_marker: str = "**"
    italic_marker: str = "*"
    strip_formatting: bool = False  # For LLM-optimized output


class TextRunRenderer:
    """Renders TextRun objects to Markdown inline formatting."""

    def __init__(self, config: RenderConfig):
        self.config = config

    def render(self, runs: List[TextRun]) -> str:
        """Render text runs with bold/italic formatting."""
        if self.config.strip_formatting:
            return "".join(run.text for run in runs)

        result = []
        for run in runs:
            text = run.text
            if run.bold:
                text = f"{self.config.bold_marker}{text}{self.config.bold_marker}"
            if run.italic:
                text = f"{self.config.italic_marker}{text}{self.config.italic_marker}"
            if run.superscript:
                text = f"^{text}^"
            result.append(text)
        return "".join(result)


class FrontmatterRenderer:
    """Renders YAML frontmatter from document metadata."""

    @staticmethod
    def render(metadata: DocumentMetadata) -> str:
        """Render metadata as YAML frontmatter block."""
        lines = ["---"]

        if metadata.title and metadata.title != "IB Report":
            lines.append(f'title: "{metadata.title}"')
        if metadata.subtitle:
            lines.append(f'subtitle: "{metadata.subtitle}"')
        if metadata.company and metadata.company != "Korea Development Bank":
            lines.append(f'company: "{metadata.company}"')
        if metadata.ticker:
            lines.append(f'ticker: "{metadata.ticker}"')
        if metadata.sector and metadata.sector != "SECTOR":
            lines.append(f'sector: "{metadata.sector}"')
        if metadata.analyst and metadata.analyst != "DCM Team 1":
            lines.append(f'analyst: "{metadata.analyst}"')

        for key, value in metadata.extra.items():
            lines.append(f'{key}: "{value}"')

        lines.append("---")
        lines.append("")

        # Only return if actual content
        if len(lines) > 3:
            return "\n".join(lines)
        return ""


class HeadingRenderer:
    """Renders heading elements to Markdown."""

    @staticmethod
    def render(heading: Heading) -> str:
        """Render heading with appropriate # prefix."""
        prefix = "#" * heading.level
        text = heading.text.replace("**", "").strip()
        return f"{prefix} {text}\n"


class ParagraphRenderer:
    """Renders paragraph elements to Markdown."""

    def __init__(self, config: RenderConfig):
        self.run_renderer = TextRunRenderer(config)

    def render(self, para: Paragraph) -> str:
        """Render paragraph with inline formatting."""
        if para.runs:
            text = self.run_renderer.render(para.runs)
        else:
            text = para.text
        return f"{text}\n"


class TableRenderer:
    """Renders table elements to Markdown."""

    def __init__(self, config: RenderConfig):
        self.config = config

    def render(self, table: Table) -> str:
        """Render a table to Markdown format."""
        if not table.rows:
            return ""

        lines = []

        # Header row
        if table.rows:
            header_row = table.rows[0]
            header_cells = [self._escape_cell(c.content) for c in header_row.cells]
            lines.append("| " + " | ".join(header_cells) + " |")

            # Separator row with alignment
            if self.config.table_alignment and table.alignments:
                sep_cells = []
                for i, align in enumerate(table.alignments):
                    if i < len(table.alignments):
                        if align == "center":
                            sep_cells.append(":---:")
                        elif align == "right":
                            sep_cells.append("---:")
                        else:
                            sep_cells.append("---")
                    else:
                        sep_cells.append("---")
                # Ensure we have enough separators
                while len(sep_cells) < len(header_cells):
                    sep_cells.append("---")
                lines.append("| " + " | ".join(sep_cells) + " |")
            else:
                sep = ["---"] * len(header_cells)
                lines.append("| " + " | ".join(sep) + " |")

        # Data rows
        for row in table.rows[1:]:
            cells = [self._escape_cell(c.content) for c in row.cells]
            lines.append("| " + " | ".join(cells) + " |")

        lines.append("")
        return "\n".join(lines)

    @staticmethod
    def _escape_cell(text: str) -> str:
        """Escape pipe characters and newlines in cell content."""
        return text.replace("|", "\\|").replace("\n", " ").strip()


class ListRenderer:
    """Renders list elements to Markdown."""

    def __init__(self, config: RenderConfig):
        self.run_renderer = TextRunRenderer(config)

    def render_bullet(self, item: ListItem) -> str:
        """Render a bullet list item."""
        if item.runs:
            text = self.run_renderer.render(item.runs)
        else:
            text = item.text

        indent = "  " * item.indent_level
        return f"{indent}- {text}\n"

    def render_numbered(self, number: str, item: ListItem) -> str:
        """Render a numbered list item."""
        if item.runs:
            text = self.run_renderer.render(item.runs)
        else:
            text = item.text

        indent = "  " * item.indent_level
        return f"{indent}{number}. {text}\n"


class BlockquoteRenderer:
    """Renders blockquote/callout elements to Markdown."""

    @staticmethod
    def render(blockquote: Blockquote) -> str:
        """Render a blockquote with title to Markdown."""
        lines = []

        # Title line with bold
        lines.append(f"> **[{blockquote.title}]**")

        # Content lines
        for line in blockquote.text.split("\n"):
            if line.strip():
                lines.append(f"> {line}")
            else:
                lines.append(">")

        lines.append("")
        return "\n".join(lines)


class ImageRenderer:
    """Renders image elements to Markdown."""

    def __init__(self, config: RenderConfig):
        self.config = config

    def render(self, image: Image) -> str:
        """Render an image to Markdown format."""
        alt = image.alt_text or "Image"

        # Handle base64 embedded images
        if image.base64_data and self.config.embed_images_base64:
            return f"![{alt}](data:{image.mime_type};base64,{image.base64_data})\n"

        # Regular path-based images
        path = image.path
        if self.config.image_path_prefix and path:
            path = f"{self.config.image_path_prefix}/{path}"

        if not path:
            return f"<!-- Image: {alt} (path not available) -->\n"

        return f"![{alt}]({path})\n"


class MarkdownRenderer:
    """Main renderer for converting DocumentModel to Markdown."""

    def __init__(self, config: Optional[RenderConfig] = None):
        self.config = config or RenderConfig()

        # Initialize sub-renderers
        self.frontmatter_renderer = FrontmatterRenderer()
        self.heading_renderer = HeadingRenderer()
        self.paragraph_renderer = ParagraphRenderer(self.config)
        self.table_renderer = TableRenderer(self.config)
        self.list_renderer = ListRenderer(self.config)
        self.blockquote_renderer = BlockquoteRenderer()
        self.image_renderer = ImageRenderer(self.config)

    def render(self, model: DocumentModel) -> str:
        """Render DocumentModel to Markdown string."""
        output_parts = []

        # Frontmatter
        if self.config.include_frontmatter:
            frontmatter = self.frontmatter_renderer.render(model.metadata)
            if frontmatter:
                output_parts.append(frontmatter)

        # Elements
        for element in model.elements:
            try:
                rendered = self._render_element(element)
                if rendered:
                    output_parts.append(rendered)
            except Exception as e:
                logger.warning("Failed to render element %s: %s", element.element_type, e)
                output_parts.append(f"<!-- Render error: {element.element_type.name} -->\n")

        return "\n".join(output_parts)

    def _render_element(self, element: Element) -> str:
        """Render a single element to Markdown."""
        content = element.content

        if element.element_type in (
            ElementType.HEADING_1,
            ElementType.HEADING_2,
            ElementType.HEADING_3,
            ElementType.HEADING_4,
            ElementType.NUMBERED_HEADING,
        ) and isinstance(content, Heading):
            return self.heading_renderer.render(content)

        elif element.element_type == ElementType.PARAGRAPH and isinstance(content, Paragraph):
            return self.paragraph_renderer.render(content)

        elif element.element_type == ElementType.BULLET_LIST and isinstance(content, ListItem):
            return self.list_renderer.render_bullet(content)

        elif element.element_type == ElementType.NUMBERED_LIST and isinstance(content, tuple):
            number, item = content
            return self.list_renderer.render_numbered(number, item)

        elif element.element_type == ElementType.TABLE and isinstance(content, Table):
            return self.table_renderer.render(content)

        elif element.element_type == ElementType.BLOCKQUOTE and isinstance(content, Blockquote):
            return self.blockquote_renderer.render(content)

        elif element.element_type == ElementType.IMAGE and isinstance(content, Image):
            return self.image_renderer.render(content)

        elif element.element_type == ElementType.LATEX_BLOCK and isinstance(content, LaTeXEquation):
            return f"$$\n{content.expression}\n$$\n"

        elif element.element_type == ElementType.LATEX_INLINE and isinstance(
            content, LaTeXEquation
        ):
            return f"${content.expression}$\n"

        elif element.element_type == ElementType.SEPARATOR:
            return "---\n"

        elif element.element_type == ElementType.EMPTY:
            return ""

        else:
            logger.warning("Unknown element type or content mismatch: %s", element.element_type)
            return ""


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════


def render_to_markdown(
    model: DocumentModel,
    include_frontmatter: bool = True,
    strip_formatting: bool = False,
) -> str:
    """
    Render a DocumentModel to Markdown string.

    Args:
        model: The DocumentModel to render
        include_frontmatter: Whether to include YAML frontmatter
        strip_formatting: Remove bold/italic for cleaner LLM input

    Returns:
        Markdown string
    """
    config = RenderConfig(
        include_frontmatter=include_frontmatter,
        strip_formatting=strip_formatting,
    )
    renderer = MarkdownRenderer(config)
    return renderer.render(model)
