"""
Plugin/converter architecture for IB Report Formatter.

Inspired by Microsoft's markitdown project. Provides BaseConverter
and ConverterRegistry for extensible format support.

Changelog (hardening):
    - Publish the default registry only after locked, complete initialization.

Changelog (render warnings):
    - The DOCX output converter adds the renderer's warnings to a caller's
      `render_warnings` list.

Usage:
    from converters import get_default_registry

    registry = get_default_registry()

    # Parse any supported file -> DocumentModel
    model = registry.convert("report.md")

    # Render DocumentModel -> output
    registry.convert(model, output_format="docx", output_path="out.docx")

    # Add a new format
    class PdfOutputConverter(OutputConverter):
        name = "pdf"
        output_format = "pdf"
        def convert(self, source, **kwargs): ...
    registry.register(PdfOutputConverter())
"""

import logging
import threading
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, BinaryIO, cast

from input_safety import DEFAULT_LIMITS
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
    def accepts(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> bool:
        """Return True if this converter can handle the given source."""
        ...

    @abstractmethod
    def convert(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> Any:
        """Execute the conversion."""
        ...


class InputConverter(BaseConverter):
    """Base for converters that parse files into DocumentModel.

    Accepts file paths (str/Path) and binary streams (BinaryIO).
    For streams, uses extension hint or signature-based format detection.
    """

    supported_extensions: list[str] = []
    # Format name for stream detection (e.g. "docx", "md")
    supported_format: str = ""

    def accepts(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> bool:
        if isinstance(source, DocumentModel):
            return False

        # Stream path: use hint or detect format
        if is_stream(source):
            hint = kwargs.get("extension_hint")
            options = kwargs.get("render_options")
            limits = kwargs.get("limits") or getattr(options, "limits", None) or DEFAULT_LIMITS
            fmt = detect_format(cast(BinaryIO, source), hint=hint, limits=limits)
            return fmt == self.supported_format

        # File path
        if not isinstance(source, (str, Path)):
            return False
        path = Path(source)
        return path.suffix.lower() in self.supported_extensions


class OutputConverter(BaseConverter):
    """Base for converters that render DocumentModel to output."""

    output_format: str = ""

    def accepts(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> bool:
        if not isinstance(source, DocumentModel):
            return False
        fmt = kwargs.get("output_format", "")
        return isinstance(fmt, str) and fmt.lower() == self.output_format


# ═══════════════════════════════════════════════════════════════════════════════
# REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════


class ConverterRegistry:
    """Registry of available converters with priority-based lookup."""

    def __init__(self) -> None:
        self._converters: list[BaseConverter] = []

    def register(self, converter: BaseConverter) -> None:
        """Register a converter. Lower priority values are tried first."""
        self._converters.append(converter)
        self._converters.sort(key=lambda c: c.priority)

    def find_converter(
        self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any
    ) -> BaseConverter | None:
        """Find the first converter that accepts the source."""
        for converter in self._converters:
            if converter.accepts(source, **kwargs):
                return converter
        return None

    def convert(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> Any:
        """Find a matching converter and execute conversion.

        For Markdown BinaryIO streams, pass extension_hint="md" (or ".md") to
        help format detection when the stream lacks a file signature.
        """
        # Ensure stream is seekable before converter lookup (accepts may peek)
        if is_stream(source):
            options = kwargs.get("render_options")
            limits = kwargs.get("limits") or getattr(options, "limits", None) or DEFAULT_LIMITS
            source = ensure_seekable(cast(BinaryIO, source), limits=limits)

        converter = self.find_converter(source, **kwargs)
        if converter is None:
            raise ValueError(f"No converter found for: {source!r} with kwargs={kwargs}")
        return converter.convert(source, **kwargs)

    @property
    def converters(self) -> list[BaseConverter]:
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
        self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any
    ) -> DocumentModel:
        from document_profiles import RenderOptions
        from md_parser import parse_markdown_file

        options = kwargs.get("render_options")
        roots = kwargs.get("asset_roots", ())
        profile = kwargs.get("profile")
        limits = kwargs.get("limits") or getattr(options, "limits", None) or DEFAULT_LIMITS
        if isinstance(options, RenderOptions):
            roots = (*options.asset_roots, *roots)

        # parse_markdown_file now accepts both str and BinaryIO
        if is_stream(source):
            return parse_markdown_file(
                cast(BinaryIO, source), profile=profile, asset_roots=roots, limits=limits
            )
        return parse_markdown_file(str(source), profile=profile, asset_roots=roots, limits=limits)


class DocxOutputConverter(OutputConverter):
    """Render DocumentModel to Word (.docx) format."""

    name = "docx-output"
    priority = 100
    output_format = "docx"

    def convert(self, source: str | Path | BinaryIO | DocumentModel, **kwargs: Any) -> Any:
        from ib_renderer import IBDocumentRenderer, rebase_local_links

        assert isinstance(source, DocumentModel)
        from document_profiles import RenderOptions

        options = kwargs.get("render_options")
        if options is None:
            keys = {
                "include_cover",
                "include_toc",
                "include_disclaimer",
                "separator_mode",
                "profile",
                "theme",
                "strict",
                "confidential",
                "charts",
                "preset",
                "house",
                "term_tags",
            }
            options = RenderOptions(**{key: value for key, value in kwargs.items() if key in keys})
        if not isinstance(options, RenderOptions):
            raise TypeError("render_options must be RenderOptions")
        renderer = IBDocumentRenderer(options=options)
        doc = renderer.render(
            source, asset_roots=kwargs.get("asset_roots", ()), limits=kwargs.get("limits")
        )
        # A caller's list receives the warnings that never reject (IBDocumentRenderer.warnings).
        sink = kwargs.get("render_warnings")
        if sink is not None:
            sink.extend(renderer.warnings)
        output_path = kwargs.get("output_path")
        if output_path:
            from cli_utils import safe_save

            rebase_local_links(doc, source.source_dir, Path(output_path))
            saved = safe_save(
                Path(output_path),
                lambda path: doc.save(str(path)),
                logger,
                "%s is locked; saving with timestamp suffix",
            )
            return str(saved)
        return doc


# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT REGISTRY
# ═══════════════════════════════════════════════════════════════════════════════

_default_registry: ConverterRegistry | None = None
_default_registry_lock = threading.Lock()


def get_default_registry() -> ConverterRegistry:
    """Get the default registry, initializing its built-ins once across threads.

    Returns:
        The shared registry with all built-in converters registered.
    """
    global _default_registry
    if _default_registry is None:
        with _default_registry_lock:
            if _default_registry is None:
                registry = ConverterRegistry()
                _register_builtin_converters(registry)
                _default_registry = registry
    return _default_registry


def _register_builtin_converters(registry: ConverterRegistry) -> None:
    """Register the built-in converters."""
    registry.register(MarkdownInputConverter())
    registry.register(DocxOutputConverter())
