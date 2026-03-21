"""
Stream utilities and lightweight MIME/format detection.

Provides BinaryIO-based format detection without heavy dependencies
like magika. Uses file signatures (magic bytes) and optional extension hints.

Usage:
    from stream_utils import detect_format, ensure_seekable

    fmt = detect_format(stream)               # "docx", "md", or "unknown"
    fmt = detect_format(stream, hint=".docx")  # hint takes priority

    stream = ensure_seekable(stream)          # wraps non-seekable in BytesIO
"""

import io
import logging
from typing import BinaryIO, Optional, Union

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════════
# FORMAT SIGNATURES
# ═══════════════════════════════════════════════════════════════════════════════

# PK\x03\x04 = ZIP (DOCX, XLSX, PPTX are ZIP-based)
_ZIP_SIGNATURE = b"PK\x03\x04"

# Known extension → format mappings
_EXT_TO_FORMAT = {
    ".docx": "docx",
    ".md": "md",
    ".markdown": "md",
    ".txt": "md",  # treat plain text as markdown
}


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════


def detect_format(
    source: BinaryIO,
    *,
    hint: Optional[str] = None,
) -> str:
    """Detect the format of a binary stream.

    Detection priority:
        1. Extension hint (if provided and recognized)
        2. File signature (magic bytes)
        3. Heuristic (UTF-8 text → markdown)

    The stream position is preserved (reset to original after peeking).

    Args:
        source: A readable binary stream.
        hint: Optional file extension hint (e.g. ".docx", "docx").

    Returns:
        Format string: "docx", "md", or "unknown".
    """
    # 1. Extension hint
    if hint:
        normalized = hint.lower() if hint.startswith(".") else f".{hint.lower()}"
        fmt = _EXT_TO_FORMAT.get(normalized)
        if fmt:
            return fmt

    # 2. Peek at first bytes for signature detection
    source = ensure_seekable(source)
    pos = source.tell()
    header = source.read(4)
    source.seek(pos)

    if not header:
        return "unknown"

    # ZIP signature → DOCX (the only ZIP format we support)
    if header[:4] == _ZIP_SIGNATURE:
        return "docx"

    # 3. Heuristic: if it looks like valid UTF-8 text, treat as markdown
    source.seek(pos)
    sample = source.read(8192)
    source.seek(pos)

    if _looks_like_text(sample):
        return "md"

    return "unknown"


def ensure_seekable(stream: BinaryIO) -> BinaryIO:
    """Wrap a non-seekable stream in BytesIO to make it seekable.

    If the stream is already seekable, returns it unchanged.
    Otherwise, reads all content into a BytesIO wrapper.
    """
    if hasattr(stream, "seekable") and stream.seekable():
        return stream

    data = stream.read()
    return io.BytesIO(data)


def normalize_source(
    source: Union[str, "pathlib.Path", BinaryIO],
) -> Union[str, BinaryIO]:
    """Normalize source to either a file path string or a BinaryIO stream.

    - str / Path → str (file path)
    - BinaryIO → BinaryIO (stream)
    """
    import pathlib

    if isinstance(source, pathlib.Path):
        return str(source)
    return source


def is_stream(source: object) -> bool:
    """Check if source is a binary stream (has read method)."""
    return hasattr(source, "read") and callable(getattr(source, "read"))


# ═══════════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ═══════════════════════════════════════════════════════════════════════════════


def _looks_like_text(data: bytes) -> bool:
    """Heuristic check: does this data look like UTF-8 text?

    Returns True if the data can be decoded as UTF-8 and doesn't
    contain excessive control characters (excluding newlines/tabs).
    """
    if not data:
        return False

    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        # Try other common encodings for Korean text
        for enc in ("euc-kr", "cp949"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            return False

    # Count control characters (excluding whitespace)
    control_count = sum(
        1
        for ch in text
        if ord(ch) < 32 and ch not in ("\n", "\r", "\t")
    )
    # If more than 5% are control chars, probably binary
    return control_count / max(len(text), 1) < 0.05
