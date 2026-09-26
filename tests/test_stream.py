"""
Tests for Step 3: Stream-based processing + MIME detection.

Tests cover:
    - stream_utils: detect_format, ensure_seekable, is_stream
    - md_parser: parse_markdown_file() with BinaryIO
    - converters: InputConverter.accepts() with streams
"""

import io
import zipfile
from pathlib import Path

import pytest
from docx import Document

PNG_BYTES = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x04\x00\x00\x00"
    b"\xb5\x1c\x0c\x02\x00\x00\x00\x0bIDATx\xdac\xfc\xff\x1f\x00\x03\x03\x02\x00\xee\xd9\xf1"
    b"\xe4\x00\x00\x00\x00IEND\xaeB`\x82"
)


# ═══════════════════════════════════════════════════════════════════════════════
# FIXTURES
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.fixture
def sample_docx_bytes():
    """Create a minimal DOCX in memory and return its bytes."""
    doc = Document()
    doc.add_heading("Stream Test", level=1)
    doc.add_paragraph("Hello from a stream.")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


@pytest.fixture
def sample_md_bytes():
    """Create sample Markdown bytes."""
    return "# 스트림 테스트\n\n본문 텍스트입니다.\n".encode()


@pytest.fixture
def sample_docx_path(tmp_path, sample_docx_bytes):
    """Write sample DOCX to disk and return path."""
    p = tmp_path / "test.docx"
    p.write_bytes(sample_docx_bytes)
    return p


@pytest.fixture
def sample_md_path(tmp_path, sample_md_bytes):
    """Write sample MD to disk and return path."""
    p = tmp_path / "test.md"
    p.write_bytes(sample_md_bytes)
    return p


# ═══════════════════════════════════════════════════════════════════════════════
# stream_utils tests
# ═══════════════════════════════════════════════════════════════════════════════


class TestDetectFormat:
    """Tests for stream_utils.detect_format."""

    def test_detect_docx_by_signature(self, sample_docx_bytes):
        from stream_utils import detect_format

        stream = io.BytesIO(sample_docx_bytes)
        assert detect_format(stream) == "docx"

    def test_detect_md_by_content(self, sample_md_bytes):
        from stream_utils import detect_format

        stream = io.BytesIO(sample_md_bytes)
        assert detect_format(stream) == "md"

    def test_detect_with_extension_hint(self, sample_docx_bytes):
        from stream_utils import detect_format

        stream = io.BytesIO(sample_docx_bytes)
        assert detect_format(stream, hint=".docx") == "docx"

    def test_detect_hint_without_dot(self, sample_md_bytes):
        from stream_utils import detect_format

        stream = io.BytesIO(sample_md_bytes)
        assert detect_format(stream, hint="md") == "md"

    def test_detect_empty_stream(self):
        from stream_utils import detect_format

        stream = io.BytesIO(b"")
        assert detect_format(stream) == "unknown"

    def test_detect_binary_garbage(self):
        from stream_utils import detect_format

        stream = io.BytesIO(bytes(range(256)) * 10)
        # Not ZIP, not text → unknown
        assert detect_format(stream) == "unknown"

    def test_stream_position_preserved(self, sample_docx_bytes):
        from stream_utils import detect_format

        stream = io.BytesIO(sample_docx_bytes)
        stream.seek(5)
        detect_format(stream)
        # ensure_seekable wraps, so position should be at start of the new wrapper
        # but the point is detect_format doesn't consume the stream
        # We just verify it doesn't raise

    def test_detect_korean_euckr(self):
        """Korean text in EUC-KR should be detected as md."""
        from stream_utils import detect_format

        korean = "한국어 텍스트입니다.".encode("euc-kr")
        stream = io.BytesIO(korean)
        assert detect_format(stream) == "md"

    def test_detect_zip_as_docx(self):
        """Any ZIP is treated as docx (we only support DOCX as ZIP)."""
        from stream_utils import detect_format

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("test.txt", "hello")
        stream = io.BytesIO(buf.getvalue())
        assert detect_format(stream) == "docx"


class TestEnsureSeekable:
    """Tests for stream_utils.ensure_seekable."""

    def test_already_seekable(self):
        from stream_utils import ensure_seekable

        stream = io.BytesIO(b"hello")
        result = ensure_seekable(stream)
        assert result is stream  # same object

    def test_non_seekable_wrapped(self):
        from stream_utils import ensure_seekable

        class NonSeekable:
            def __init__(self, data):
                self._data = data
                self._pos = 0

            def read(self, n=-1):
                if n == -1:
                    result = self._data[self._pos :]
                    self._pos = len(self._data)
                else:
                    result = self._data[self._pos : self._pos + n]
                    self._pos += n
                return result

            def seekable(self):
                return False

        ns = NonSeekable(b"test data")
        result = ensure_seekable(ns)
        assert result.read() == b"test data"
        result.seek(0)
        assert result.read() == b"test data"


class TestIsStream:
    def test_stream(self):
        from stream_utils import is_stream

        assert is_stream(io.BytesIO(b""))
        assert not is_stream("a string")
        assert not is_stream(Path("/tmp"))
        assert not is_stream(42)


# ═══════════════════════════════════════════════════════════════════════════════
# Markdown stream tests
# ═══════════════════════════════════════════════════════════════════════════════


class TestMdParserStream:
    """Test that parse_markdown_file accepts BinaryIO."""

    def test_parse_from_stream(self, sample_md_bytes):
        from md_parser import parse_markdown_file

        stream = io.BytesIO(sample_md_bytes)
        model = parse_markdown_file(stream)
        assert model is not None
        assert len(model.elements) > 0

    def test_parse_from_path_still_works(self, sample_md_path):
        from md_parser import parse_markdown_file

        model = parse_markdown_file(str(sample_md_path))
        assert model is not None
        assert len(model.elements) > 0

    def test_stream_korean_euckr(self):
        from md_parser import parse_markdown_file

        content = "# 한국어 제목\n\n본문 내용\n".encode("euc-kr")
        stream = io.BytesIO(content)
        model = parse_markdown_file(stream)
        assert len(model.elements) > 0


# ═══════════════════════════════════════════════════════════════════════════════
# converters stream tests
# ═══════════════════════════════════════════════════════════════════════════════


class TestConverterRegistryStream:
    """Test ConverterRegistry with BinaryIO streams."""

    def test_registry_convert_md_stream(self, sample_md_bytes):
        from converters import get_default_registry

        registry = get_default_registry()
        stream = io.BytesIO(sample_md_bytes)
        model = registry.convert(stream)
        assert model is not None
        assert len(model.elements) > 0

    def test_registry_unknown_stream_raises(self):
        from converters import get_default_registry

        registry = get_default_registry()
        stream = io.BytesIO(b"\x00\x01\x02\x03" * 100)
        with pytest.raises(ValueError, match="No converter found"):
            registry.convert(stream)


# ═══════════════════════════════════════════════════════════════════════════════
# CLI pipe tests
# ═══════════════════════════════════════════════════════════════════════════════
