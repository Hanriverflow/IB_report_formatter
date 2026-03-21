"""Tests for the plugin/converter architecture."""

import textwrap
from pathlib import Path

import pytest

from converters import (
    BaseConverter,
    ConverterRegistry,
    DocxInputConverter,
    DocxOutputConverter,
    InputConverter,
    MarkdownInputConverter,
    MarkdownOutputConverter,
    OutputConverter,
    get_default_registry,
)
from md_parser import DocumentMetadata, DocumentModel, Element, ElementType, Heading


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════


def _minimal_model() -> DocumentModel:
    """Create a minimal DocumentModel for testing."""
    return DocumentModel(
        metadata=DocumentMetadata(title="Test"),
        elements=[
            Element(
                element_type=ElementType.HEADING_1,
                content=Heading(level=1, text="Hello"),
                raw_text="Hello",
            ),
        ],
        footnotes={},
    )


def _write_md(tmp_path: Path, content: str) -> Path:
    """Write a markdown file and return its path."""
    md_file = tmp_path / "test.md"
    md_file.write_text(textwrap.dedent(content).strip(), encoding="utf-8")
    return md_file


# ═══════════════════════════════════════════════════════════════════════════════
# REGISTRY TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestConverterRegistry:
    """Test ConverterRegistry behavior."""

    def test_register_and_find(self):
        """Register a converter and find it."""
        registry = ConverterRegistry()
        converter = MarkdownInputConverter()
        registry.register(converter)
        found = registry.find_converter("test.md")
        assert found is converter

    def test_find_returns_none_when_empty(self):
        """Empty registry returns None."""
        registry = ConverterRegistry()
        assert registry.find_converter("test.md") is None

    def test_convert_raises_when_no_match(self):
        """ValueError when no converter matches."""
        registry = ConverterRegistry()
        with pytest.raises(ValueError, match="No converter found"):
            registry.convert("unknown.xyz")

    def test_priority_ordering(self):
        """Lower priority value wins when multiple converters match."""

        class HighPriority(InputConverter):
            name = "high"
            priority = 10
            supported_extensions = [".md"]

            def convert(self, source, **kwargs):
                return "high"

        class LowPriority(InputConverter):
            name = "low"
            priority = 200
            supported_extensions = [".md"]

            def convert(self, source, **kwargs):
                return "low"

        registry = ConverterRegistry()
        registry.register(LowPriority())
        registry.register(HighPriority())

        found = registry.find_converter("test.md")
        assert found is not None
        assert found.name == "high"

    def test_converters_property(self):
        """converters property returns a copy of the list."""
        registry = ConverterRegistry()
        registry.register(MarkdownInputConverter())
        registry.register(DocxInputConverter())
        converters = registry.converters
        assert len(converters) == 2
        # Mutation doesn't affect internal list
        converters.clear()
        assert len(registry.converters) == 2


# ═══════════════════════════════════════════════════════════════════════════════
# INPUT CONVERTER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestInputConverter:
    """Test InputConverter.accepts() logic."""

    def test_accepts_md_extension(self):
        converter = MarkdownInputConverter()
        assert converter.accepts("report.md") is True
        assert converter.accepts("report.markdown") is True

    def test_rejects_wrong_extension(self):
        converter = MarkdownInputConverter()
        assert converter.accepts("report.docx") is False
        assert converter.accepts("report.pdf") is False

    def test_rejects_document_model(self):
        """InputConverter should reject DocumentModel instances."""
        converter = MarkdownInputConverter()
        model = _minimal_model()
        assert converter.accepts(model) is False

    def test_accepts_path_object(self):
        converter = MarkdownInputConverter()
        assert converter.accepts(Path("test.md")) is True

    def test_docx_accepts(self):
        converter = DocxInputConverter()
        assert converter.accepts("report.docx") is True
        assert converter.accepts("report.md") is False

    def test_case_insensitive_extension(self):
        converter = MarkdownInputConverter()
        assert converter.accepts("REPORT.MD") is True


# ═══════════════════════════════════════════════════════════════════════════════
# OUTPUT CONVERTER TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestOutputConverter:
    """Test OutputConverter.accepts() logic."""

    def test_accepts_document_model_with_format(self):
        converter = DocxOutputConverter()
        model = _minimal_model()
        assert converter.accepts(model, output_format="docx") is True

    def test_rejects_wrong_format(self):
        converter = DocxOutputConverter()
        model = _minimal_model()
        assert converter.accepts(model, output_format="pdf") is False

    def test_rejects_file_path(self):
        converter = DocxOutputConverter()
        assert converter.accepts("file.docx", output_format="docx") is False

    def test_markdown_output_accepts(self):
        converter = MarkdownOutputConverter()
        model = _minimal_model()
        assert converter.accepts(model, output_format="md") is True

    def test_rejects_without_format_kwarg(self):
        converter = DocxOutputConverter()
        model = _minimal_model()
        assert converter.accepts(model) is False


# ═══════════════════════════════════════════════════════════════════════════════
# CONVERSION TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestMarkdownInputConversion:
    """Test actual Markdown file parsing via converter."""

    def test_parse_md_file(self, tmp_path):
        md_file = _write_md(tmp_path, """
            ---
            title: Test Report
            ---

            # Introduction

            Hello world.
        """)
        converter = MarkdownInputConverter()
        model = converter.convert(str(md_file))
        assert isinstance(model, DocumentModel)
        assert model.metadata.title == "Test Report"
        assert len(model.elements) > 0


class TestMarkdownOutputConversion:
    """Test DocumentModel rendering to Markdown."""

    def test_render_to_string(self):
        model = _minimal_model()
        converter = MarkdownOutputConverter()
        result = converter.convert(model, output_format="md")
        assert isinstance(result, str)
        assert "# Hello" in result

    def test_render_to_file(self, tmp_path):
        model = _minimal_model()
        converter = MarkdownOutputConverter()
        out_path = tmp_path / "output.md"
        result = converter.convert(model, output_format="md", output_path=str(out_path))
        assert result == str(out_path)
        assert out_path.exists()
        content = out_path.read_text(encoding="utf-8")
        assert "# Hello" in content

    def test_strip_formatting_option(self):
        model = _minimal_model()
        converter = MarkdownOutputConverter()
        result = converter.convert(model, output_format="md", strip_formatting=True)
        assert isinstance(result, str)


class TestDocxOutputConversion:
    """Test DocumentModel rendering to DOCX."""

    def test_render_to_docx_object(self):
        model = _minimal_model()
        converter = DocxOutputConverter()
        doc = converter.convert(model, output_format="docx")
        # Should return a python-docx Document object
        from docx.document import Document as DocxDocument

        assert isinstance(doc, DocxDocument)

    def test_render_to_file(self, tmp_path):
        model = _minimal_model()
        converter = DocxOutputConverter()
        out_path = tmp_path / "output.docx"
        result = converter.convert(
            model, output_format="docx", output_path=str(out_path)
        )
        assert result == str(out_path)
        assert out_path.exists()


# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT REGISTRY TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestDefaultRegistry:
    """Test the default registry with all built-in converters."""

    def test_has_four_builtins(self):
        registry = get_default_registry()
        assert len(registry.converters) == 4

    def test_finds_md_input(self):
        registry = get_default_registry()
        converter = registry.find_converter("test.md")
        assert converter is not None
        assert converter.name == "markdown-input"

    def test_finds_docx_input(self):
        registry = get_default_registry()
        converter = registry.find_converter("test.docx")
        assert converter is not None
        assert converter.name == "docx-input"

    def test_finds_docx_output(self):
        registry = get_default_registry()
        model = _minimal_model()
        converter = registry.find_converter(model, output_format="docx")
        assert converter is not None
        assert converter.name == "docx-output"

    def test_finds_md_output(self):
        registry = get_default_registry()
        model = _minimal_model()
        converter = registry.find_converter(model, output_format="md")
        assert converter is not None
        assert converter.name == "markdown-output"

    def test_end_to_end_md_roundtrip(self, tmp_path):
        """MD file -> DocumentModel -> MD string."""
        md_file = _write_md(tmp_path, """
            ---
            title: Round Trip
            ---

            # Section One

            This is a test.
        """)
        registry = get_default_registry()
        model = registry.convert(str(md_file))
        assert isinstance(model, DocumentModel)

        md_text = registry.convert(model, output_format="md")
        assert isinstance(md_text, str)
        assert "# Section One" in md_text


# ═══════════════════════════════════════════════════════════════════════════════
# BACKWARD COMPATIBILITY
# ═══════════════════════════════════════════════════════════════════════════════


class TestBackwardCompatibility:
    """Verify existing APIs still work directly (not through registry)."""

    def test_parse_markdown_file_direct(self, tmp_path):
        md_file = _write_md(tmp_path, """
            # Direct Call

            Works fine.
        """)
        from md_parser import parse_markdown_file

        model = parse_markdown_file(str(md_file))
        assert isinstance(model, DocumentModel)

    def test_render_to_markdown_direct(self):
        from md_renderer import render_to_markdown

        model = _minimal_model()
        result = render_to_markdown(model)
        assert isinstance(result, str)
        assert "# Hello" in result

    def test_ib_renderer_direct(self):
        from ib_renderer import IBDocumentRenderer

        model = _minimal_model()
        renderer = IBDocumentRenderer()
        doc = renderer.render(model)
        from docx.document import Document as DocxDocument

        assert isinstance(doc, DocxDocument)
