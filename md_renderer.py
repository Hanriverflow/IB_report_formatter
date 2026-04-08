"""
Markdown Renderer Module for Word to Markdown Converter
Handles rendering of DocumentModel to Markdown text.

Renders clean, LLM-friendly Markdown output with post-processing
normalization: trailing whitespace removal, blank line compression,
consistent spacing around block elements (tables, code, math).
"""

import logging
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

try:
    import yaml
except ImportError:  # pragma: no cover - optional dependency
    yaml = None

from md_parser import (
    Blockquote,
    CodeBlock,
    Diagram,
    DocumentMetadata,
    DocumentModel,
    Element,
    ElementType,
    Heading,
    Image,
    LaTeXEquation,
    ListItem,
    Paragraph,
    Table,
    TextRun,
)

logger = logging.getLogger(__name__)


def _safe_yaml_dump(data: Dict[str, Any]) -> str:
    """Serialize metadata or diagram payloads with YAML-safe escaping."""
    if yaml is not None:
        rendered = yaml.safe_dump(
            data,
            sort_keys=False,
            allow_unicode=True,
            default_flow_style=False,
        )
        return rendered.strip() or "{}"
    return _fallback_yaml_dump(data)


def _fallback_yaml_dump(value: Any) -> str:
    """Minimal YAML serializer used when PyYAML is unavailable."""
    lines = _fallback_yaml_lines(value)
    return "\n".join(lines) if lines else "{}"


def _fallback_yaml_lines(value: Any, indent: int = 0) -> List[str]:
    """Render nested dict/list data as YAML-compatible lines."""
    prefix = "  " * indent

    if isinstance(value, dict):
        if not value:
            return [f"{prefix}{{}}"] if indent else ["{}"]

        lines: List[str] = []
        for key, item in value.items():
            rendered_key = str(key)
            if isinstance(item, dict):
                if item:
                    lines.append(f"{prefix}{rendered_key}:")
                    lines.extend(_fallback_yaml_lines(item, indent + 1))
                else:
                    lines.append(f"{prefix}{rendered_key}: {{}}")
            elif isinstance(item, list):
                if item:
                    lines.append(f"{prefix}{rendered_key}:")
                    lines.extend(_fallback_yaml_lines(item, indent + 1))
                else:
                    lines.append(f"{prefix}{rendered_key}: []")
            else:
                lines.append(f"{prefix}{rendered_key}: {_fallback_yaml_scalar(item)}")
        return lines

    if isinstance(value, list):
        if not value:
            return [f"{prefix}[]"]

        lines = []
        for item in value:
            if isinstance(item, (dict, list)):
                if item:
                    lines.append(f"{prefix}-")
                    lines.extend(_fallback_yaml_lines(item, indent + 1))
                else:
                    empty_value = "{}" if isinstance(item, dict) else "[]"
                    lines.append(f"{prefix}- {empty_value}")
            else:
                lines.append(f"{prefix}- {_fallback_yaml_scalar(item)}")
        return lines

    return [f"{prefix}{_fallback_yaml_scalar(value)}"]


def _fallback_yaml_scalar(value: Any) -> str:
    """Render a scalar with conservative escaping."""
    if value is None:
        return '""'
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)

    text = str(value)
    text = text.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    return f'"{text}"'


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
            elif run.subscript:
                text = f"~{text}~"
            if run.color_hex:
                text = f'<span style="color:{run.color_hex}">{text}</span>'
            result.append(text)
        return "".join(result)


class EndnoteRenderer:
    """Renders parser-compatible endnotes blocks."""

    _TITLE = "## Citations"

    @classmethod
    def render(cls, footnotes: dict) -> str:
        """Render endnotes in a format md_parser can read back."""
        if not footnotes:
            return ""

        lines = [cls._TITLE]
        for number, text in sorted(footnotes.items()):
            lines.append(f"{number}. {text}")
        lines.append("")
        return "\n".join(lines)


class FrontmatterRenderer:
    """Renders YAML frontmatter from document metadata."""

    @classmethod
    def render(cls, metadata: DocumentMetadata) -> str:
        """Render metadata as YAML frontmatter block."""
        payload = cls._build_payload(metadata)
        if not payload:
            return ""

        lines = ["---", _safe_yaml_dump(payload), "---", ""]

        return "\n".join(lines)

    @staticmethod
    def _build_payload(metadata: DocumentMetadata) -> Dict[str, Any]:
        """Return only non-default metadata fields for frontmatter output."""
        payload: Dict[str, Any] = {}
        if metadata.title and metadata.title != "IB Report":
            payload["title"] = metadata.title
        if metadata.subtitle:
            payload["subtitle"] = metadata.subtitle
        if metadata.company and metadata.company != "Korea Development Bank":
            payload["company"] = metadata.company
        if metadata.ticker:
            payload["ticker"] = metadata.ticker
        if metadata.sector and metadata.sector != "SECTOR":
            payload["sector"] = metadata.sector
        if metadata.analyst and metadata.analyst != "DCM Team 1":
            payload["analyst"] = metadata.analyst

        for key, value in metadata.extra.items():
            payload[str(key)] = value

        return payload


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
        self.run_renderer = TextRunRenderer(config)

    def render(self, table: Table) -> str:
        """Render a table to Markdown format."""
        if not table.rows:
            return ""

        lines = []

        # Header row
        if table.rows:
            header_row = table.rows[0]
            header_cells = [self._render_cell(c) for c in header_row.cells]
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
            cells = [self._render_cell(c) for c in row.cells]
            lines.append("| " + " | ".join(cells) + " |")

        lines.append("")
        return "\n".join(lines)

    def _render_cell(self, cell) -> str:
        """Render a table cell, preferring structured runs when available."""
        text = self.run_renderer.render(cell.runs) if cell.runs else cell.content
        return self._escape_cell(text)

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


class DiagramRenderer:
    """Renders diagram elements to fenced YAML blocks."""

    @staticmethod
    def render(diagram: Diagram) -> str:
        """Render diagram payload in parser-compatible fenced syntax."""
        diagram_type = diagram.diagram_type.strip()
        language = f"diagram:{diagram_type}" if diagram_type else "diagram"
        payload: Dict[str, Any] = {}

        if diagram.title:
            payload["title"] = diagram.title
        if diagram.boxes:
            payload["boxes"] = [
                {
                    "id": box.id,
                    "label": box.label,
                    "pos": box.pos,
                    "style": box.style,
                }
                for box in diagram.boxes
            ]
        if diagram.arrows:
            payload["arrows"] = [
                {
                    "from": arrow.from_id,
                    "to": arrow.to_id,
                    "label": arrow.label,
                    "style": arrow.style,
                }
                for arrow in diagram.arrows
            ]
        if diagram.notes:
            payload["notes"] = diagram.notes

        return f"```{language}\n{_safe_yaml_dump(payload)}\n```\n"


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
            path = self._join_markdown_path(self.config.image_path_prefix, path)

        if not path:
            return f"<!-- Image: {alt} (path not available) -->\n"

        return f"![{alt}]({path})\n"

    @staticmethod
    def _join_markdown_path(prefix: str, path: str) -> str:
        """Join image path parts and normalize separators for Markdown links."""
        normalized_prefix = prefix.replace("\\", "/").rstrip("/")
        normalized_path = path.replace("\\", "/").lstrip("/")

        if not normalized_prefix:
            return normalized_path
        if not normalized_path:
            return normalized_prefix
        return f"{normalized_prefix}/{normalized_path}"


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
        self.diagram_renderer = DiagramRenderer()
        self.image_renderer = ImageRenderer(self.config)
        self.endnote_renderer = EndnoteRenderer()

    def render(self, model: DocumentModel) -> str:
        """Render DocumentModel to Markdown string.

        Output is normalized for LLM consumption:
        - Trailing whitespace stripped from every line
        - 3+ consecutive blank lines compressed to 2
        - Consistent blank lines around block elements
        - Clean final newline
        """
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

        if model.footnotes:
            output_parts.append(self.endnote_renderer.render(model.footnotes))

        raw = "\n".join(output_parts)
        return _normalize_markdown(raw)

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

        elif element.element_type == ElementType.CODE_BLOCK and isinstance(content, CodeBlock):
            info = content.language.strip()
            fence = f"```{info}" if info else "```"
            return f"{fence}\n{content.code}\n```\n"

        elif element.element_type == ElementType.DIAGRAM and isinstance(content, Diagram):
            return self.diagram_renderer.render(content)

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
    image_path_prefix: str = "",
    embed_images_base64: bool = False,
) -> str:
    """
    Render a DocumentModel to Markdown string.

    Args:
        model: The DocumentModel to render
        include_frontmatter: Whether to include YAML frontmatter
        strip_formatting: Remove bold/italic for cleaner LLM input
        embed_images_base64: Inline images as data URIs instead of file paths

    Returns:
        Markdown string
    """
    config = RenderConfig(
        include_frontmatter=include_frontmatter,
        image_path_prefix=image_path_prefix,
        strip_formatting=strip_formatting,
        embed_images_base64=embed_images_base64,
    )
    renderer = MarkdownRenderer(config)
    return renderer.render(model)


# ═══════════════════════════════════════════════════════════════════════════════
# OUTPUT NORMALIZATION (LLM-optimized)
# ═══════════════════════════════════════════════════════════════════════════════

# Precompiled patterns for performance
_RE_TRAILING_WHITESPACE = re.compile(r"[ \t]+$", re.MULTILINE)
_RE_EXCESSIVE_BLANKS = re.compile(r"\n{3,}")
_RE_TABLE_BLOCK = re.compile(r"(\|[^\n]+\|\n(?:\|[^\n]+\|\n)*)")
_RE_MATH_BLOCK = re.compile(r"(\$\$\n.*?\n\$\$)", re.DOTALL)
_RE_FENCE_BLOCK = re.compile(r"(```[^\n]*\n.*?\n```)", re.DOTALL)
_RE_PROTECTED_FENCE_TOKEN = re.compile(r"(__MD_FENCE_BLOCK_\d+__)")


def _protect_fenced_code_blocks(text: str) -> Tuple[str, List[str]]:
    """Replace fenced code blocks with tokens during block normalization."""
    protected_blocks: List[str] = []

    def _replace(match) -> str:
        token = f"__MD_FENCE_BLOCK_{len(protected_blocks)}__"
        protected_blocks.append(match.group(1))
        return token

    return _RE_FENCE_BLOCK.sub(_replace, text), protected_blocks


def _restore_fenced_code_blocks(text: str, protected_blocks: List[str]) -> str:
    """Restore fenced code blocks after non-fenced normalization steps."""
    for index, block in enumerate(protected_blocks):
        text = text.replace(f"__MD_FENCE_BLOCK_{index}__", block)
    return text


def _normalize_markdown(text: str) -> str:
    """Normalize Markdown output for clean, consistent LLM consumption.

    Steps:
        1. Strip trailing whitespace from every line
        2. Compress 3+ consecutive blank lines to exactly 2 (one blank line)
        3. Ensure block elements (tables, math, fences) have blank lines around them
        4. Ensure single trailing newline
    """
    # 1. Strip trailing whitespace per line
    text = _RE_TRAILING_WHITESPACE.sub("", text)

    # 2. Protect fenced blocks before table/math normalization so block content
    #    is not mistaken for Markdown tables or math blocks.
    text, protected_fences = _protect_fenced_code_blocks(text)

    # 3. Ensure blank line before/after block elements
    #    (tables, $$ math $$, protected fenced blocks)
    for pattern in (_RE_TABLE_BLOCK, _RE_MATH_BLOCK, _RE_PROTECTED_FENCE_TOKEN):
        text = pattern.sub(r"\n\1\n", text)

    # 4. Compress excessive blank lines (3+ newlines → 2)
    text = _RE_EXCESSIVE_BLANKS.sub("\n\n", text)

    # 5. Restore protected fences and ensure final newline
    text = _restore_fenced_code_blocks(text, protected_fences)
    text = text.strip() + "\n"

    return text
