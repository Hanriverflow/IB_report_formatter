"""
Plugin/converter architecture for IB Report Formatter.

Inspired by Microsoft's markitdown project. Provides BaseConverter
and ConverterRegistry for extensible format support.

Usage:
    from converters import get_default_registry

    registry = get_default_registry()

    # Parse any supported file -> DocumentModel
    model = registry.convert("report.md")
    model = registry.convert("report.docx")

    # Render DocumentModel -> output
    registry.convert(model, output_format="docx", output_path="out.docx")
    md_text = registry.convert(model, output_format="md")

    # Add a new format
    class PdfOutputConverter(OutputConverter):
        name = "pdf"
        output_format = "pdf"
        def convert(self, source, **kwargs): ...
    registry.register(PdfOutputConverter())
"""

import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import IO, Any, BinaryIO, List, Optional, Union

from md_parser import DocumentModel
from stream_utils import detect_format, ensure_seekable, is_stream

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# BASE CLASSES
# ═══════════════════════════════════════════════════════════════════════════════


class BaseConverter(ABC):
    """Abstract base class for all format converters."""

    name: str = ""
    priority: int = 100  # Lower = tried first

    @abstractmethod
    def accepts(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> bool:
        """Return True if this converter can handle the given source."""
        ...

    @abstractmethod
    def convert(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> Any:
        """Execute the conversion."""
        ...


class InputConverter(BaseConverter):
    """Base for converters that parse files into DocumentModel.

    Accepts file paths (str/Path) and binary streams (BinaryIO).
    For streams, uses extension hint or signature-based format detection.
    """

    supported_extensions: List[str] = []
    # Format name for stream detection (e.g. "docx", "md")
    supported_format: str = ""

    def accepts(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> bool:
        if isinstance(source, DocumentModel):
            return False

        # Stream path: use hint or detect format
        if is_stream(source):
            hint = kwargs.get("extension_hint")
            fmt = detect_format(source, hint=hint)
            return fmt == self.supported_format

        # File path
        path = Path(source) if isinstance(source, str) else source
        return path.suffix.lower() in self.supported_extensions


class OutputConverter(BaseConverter):
    """Base for converters that render DocumentModel to output."""

    output_format: str = ""

    def accepts(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> bool:
        if not isinstance(source, DocumentModel):
            return False
        fmt = kwargs.get("output_format", "")
        return fmt.lower() == self.output_format


# ═══════════════════════════════════════════════════════════════════════════════
# REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════


class ConverterRegistry:
    """Registry of available converters with priority-based lookup."""

    def __init__(self) -> None:
        self._converters: List[BaseConverter] = []

    def register(self, converter: BaseConverter) -> None:
        """Register a converter. Lower priority values are tried first."""
        self._converters.append(converter)
        self._converters.sort(key=lambda c: c.priority)

    def find_converter(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> Optional[BaseConverter]:
        """Find the first converter that accepts the source."""
        for converter in self._converters:
            if converter.accepts(source, **kwargs):
                return converter
        return None

    def convert(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> Any:
        """Find a matching converter and execute conversion.

        For BinaryIO streams, pass extension_hint="docx" (or ".docx") to
        help format detection when the stream lacks a file signature.
        """
        # Ensure stream is seekable before converter lookup (accepts may peek)
        if is_stream(source):
            source = ensure_seekable(source)

        converter = self.find_converter(source, **kwargs)
        if converter is None:
            raise ValueError(
                "No converter found for: {!r} with kwargs={}".format(source, kwargs)
            )
        return converter.convert(source, **kwargs)

    @property
    def converters(self) -> List[BaseConverter]:
        """List all registered converters (sorted by priority)."""
        return list(self._converters)


# ═══════════════════════════════════════════════════════════════════════════════
# BUILT-IN CONVERTER WRAPPERS
# ═══════════════════════════════════════════════════════════════════════════════


class MarkdownInputConverter(InputConverter):
    """Parse Markdown files into DocumentModel."""

    name = "markdown-input"
    priority = 100
    supported_extensions = [".md", ".markdown"]
    supported_format = "md"

    def convert(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> DocumentModel:
        from md_parser import parse_markdown_file

        # parse_markdown_file now accepts both str and BinaryIO
        if is_stream(source):
            return parse_markdown_file(source)
        return parse_markdown_file(str(source))


class DocxInputConverter(InputConverter):
    """Parse Word (.docx) files into DocumentModel."""

    name = "docx-input"
    priority = 100
    supported_extensions = [".docx"]
    supported_format = "docx"

    def convert(
        self, source: Union[str, Path, BinaryIO, DocumentModel], **kwargs: Any
    ) -> DocumentModel:
        from word_parser import parse_word_file

        # parse_word_file now accepts both str and BinaryIO
        if is_stream(source):
            return parse_word_file(
                source,
                extract_images=kwargs.get("extract_images", True),
                image_output_dir=kwargs.get("image_output_dir"),
                embed_images_base64=kwargs.get("embed_images_base64", False),
            )
        return parse_word_file(
            str(source),
            extract_images=kwargs.get("extract_images", True),
            image_output_dir=kwargs.get("image_output_dir"),
            embed_images_base64=kwargs.get("embed_images_base64", False),
        )


class DocxOutputConverter(OutputConverter):
    """Render DocumentModel to Word (.docx) format."""

    name = "docx-output"
    priority = 100
    output_format = "docx"

    def convert(self, source: Union[str, Path, DocumentModel], **kwargs: Any) -> Any:
        from ib_renderer import IBDocumentRenderer

        assert isinstance(source, DocumentModel)
        renderer = IBDocumentRenderer()
        doc = renderer.render(source)
        output_path = kwargs.get("output_path")
        if output_path:
            doc.save(str(output_path))
            return str(output_path)
        return doc


class MarkdownOutputConverter(OutputConverter):
    """Render DocumentModel to Markdown string/file."""

    name = "markdown-output"
    priority = 100
    output_format = "md"

    def convert(self, source: Union[str, Path, DocumentModel], **kwargs: Any) -> Any:
        from md_renderer import render_to_markdown

        assert isinstance(source, DocumentModel)
        md_text = render_to_markdown(
            source,
            include_frontmatter=kwargs.get("include_frontmatter", True),
            strip_formatting=kwargs.get("strip_formatting", False),
            image_path_prefix=kwargs.get("image_path_prefix", ""),
            embed_images_base64=kwargs.get("embed_images_base64", False),
        )
        output_path = kwargs.get("output_path")
        if output_path:
            Path(output_path).write_text(md_text, encoding="utf-8")
            return str(output_path)
        return md_text


# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════

_default_registry: Optional[ConverterRegistry] = None


def get_default_registry() -> ConverterRegistry:
    """Get (and lazily initialize) the default converter registry."""
    global _default_registry
    if _default_registry is None:
        _default_registry = ConverterRegistry()
        _register_builtin_converters(_default_registry)
    return _default_registry


def _register_builtin_converters(registry: ConverterRegistry) -> None:
    """Register the built-in converters."""
    registry.register(MarkdownInputConverter())
    registry.register(DocxInputConverter())
    registry.register(DocxOutputConverter())
    registry.register(MarkdownOutputConverter())
