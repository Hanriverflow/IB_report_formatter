"""Request-scoped asset access, input limits and conversion rejections.

Changelog (asset boundary):
    - Resolve local assets against operator-supplied roots before checking existence.
    - Preserve issue prefixes and expose codes on strict validation failures.
    - Bound Markdown reads, table grids, nesting and image headers per request.
"""

import base64
import logging
from collections.abc import Iterable
from dataclasses import dataclass, fields
from enum import StrEnum
from io import BytesIO
from pathlib import Path
from typing import BinaryIO
from urllib.parse import unquote

from docx.image.image import _ImageHeaderFactory

logger = logging.getLogger(__name__)


class RejectionCode(StrEnum):
    """Stable codes for the engine's documented rejections."""

    ASSET_OUTSIDE_ROOT = "asset_outside_root"
    LIMIT_INPUT_BYTES = "limit_input_bytes"
    LIMIT_TABLE_CELLS = "limit_table_cells"
    LIMIT_LIST_DEPTH = "limit_list_depth"
    LIMIT_IMAGE_BYTES = "limit_image_bytes"
    LIMIT_IMAGE_PIXELS = "limit_image_pixels"


def coded_issue(code: RejectionCode, message: str) -> str:
    """Prefix one issue without changing the existing string diagnostic API.

    Args:
        code: Machine-readable rejection code.
        message: Individual issue sentence.

    Returns:
        The issue with its bracketed code and one separating space.
    """
    return f"[{code}] {message}"


class ConversionRejection(ValueError):
    """Base for documented strict-mode and input-limit rejections."""

    def __init__(self, message: str, codes: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.codes = list(dict.fromkeys(codes))


class InputLimitError(ConversionRejection):
    """A limit failure in either mode, retaining its code and measured quantity."""

    def __init__(self, code: RejectionCode, limit: int, measured: int) -> None:
        self.code = code
        self.limit = limit
        self.measured = measured
        super().__init__(coded_issue(code, f"Limit {limit} exceeded; measured {measured}"), (code,))


@dataclass(frozen=True)
class InputLimits:
    """Operator-only limits; document metadata never participates in resolution."""

    max_input_bytes: int = 16 * 1024**2
    max_table_cells: int = 50_000
    max_list_depth: int = 32
    max_image_bytes: int = 20 * 1024**2
    max_image_pixels: int = 50_000_000

    def __post_init__(self) -> None:
        for field in fields(self):
            value = getattr(self, field.name)
            if type(value) is not int or value <= 0:
                raise ValueError(f"{field.name} must be a positive integer")

    def check(self, code: RejectionCode, measured: int) -> None:
        """Reject a measured quantity above its corresponding limit.

        Args:
            code: One of the five limit codes.
            measured: Actual bytes, cells, nesting depth or pixels.
        """
        limit = getattr(self, code.value.replace("limit_", "max_", 1))
        if measured > limit:
            raise InputLimitError(code, limit, measured)


DEFAULT_LIMITS = InputLimits()


def check_input_file(path: Path, limits: InputLimits = DEFAULT_LIMITS) -> None:
    """Reject oversized Markdown before opening it.

    Args:
        path: Original or derived Markdown file.
        limits: Caller-supplied request limits.
    """
    limits.check(RejectionCode.LIMIT_INPUT_BYTES, path.stat().st_size)


def read_input_stream(stream: BinaryIO, limits: InputLimits = DEFAULT_LIMITS) -> bytes:
    """Read at most the limit plus one byte, stopping at the first excess byte.

    Args:
        stream: Binary Markdown stream, read from its current position.
        limits: Caller-supplied request limits.

    Returns:
        The original bytes, without decoding or newline translation.
    """
    result = bytearray()
    while True:
        chunk = stream.read(min(64 * 1024, limits.max_input_bytes + 1 - len(result)))
        if not chunk:
            return bytes(result)
        result.extend(chunk)
        limits.check(RejectionCode.LIMIT_INPUT_BYTES, len(result))


def read_input_file(path: Path, limits: InputLimits = DEFAULT_LIMITS) -> bytes:
    """Check file size first and keep the read bounded if the file grows.

    Args:
        path: Markdown file to read.
        limits: Caller-supplied request limits.

    Returns:
        Original Markdown bytes.
    """
    check_input_file(path, limits)
    with path.open("rb") as stream:
        return read_input_stream(stream, limits)


def check_image_source(source: str | BytesIO, limits: InputLimits = DEFAULT_LIMITS) -> None:
    """Check image bytes and header dimensions without decoding pixel data.

    Args:
        source: Already authorized file path or decoded in-memory image.
        limits: Caller-supplied request limits.
    """
    if isinstance(source, str):
        try:
            limits.check(RejectionCode.LIMIT_IMAGE_BYTES, Path(source).stat().st_size)
            with Path(source).open("rb") as stream:
                _check_image_header(stream, limits)
        except OSError as error:
            # Let the existing insertion path retain its exact missing/unreadable warning.
            logger.debug("Image source could not be inspected: %s", error)
    else:
        limits.check(RejectionCode.LIMIT_IMAGE_BYTES, source.getbuffer().nbytes)
        position = source.tell()
        try:
            _check_image_header(source, limits)
        finally:
            source.seek(position)


def _check_image_header(stream: BinaryIO, limits: InputLimits) -> None:
    """Use python-docx's header reader; leave invalid-image diagnostics to insertion."""
    try:
        header = _ImageHeaderFactory(stream)
    except Exception as error:
        logger.debug("Image header could not be inspected: %s", error)
        return
    limits.check(RejectionCode.LIMIT_IMAGE_PIXELS, header.px_width * header.px_height)


def decode_image(data: str, limits: InputLimits = DEFAULT_LIMITS) -> BytesIO:
    """Reject oversized Base64 before decoding and check the resulting header.

    Args:
        data: Base64 payload, with the same ignored characters as b64decode.
        limits: Caller-supplied request limits.

    Returns:
        A checked decoded image stream ready for insertion.
    """
    # Six bits per alphabet character, excluding padding and ignored whitespace.
    # This also bounds malformed payloads before the decoder allocates anything.
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    measured = sum(character in alphabet for character in data) * 3 // 4
    limits.check(RejectionCode.LIMIT_IMAGE_BYTES, measured)
    source = BytesIO(base64.b64decode(data))
    check_image_source(source, limits)
    return source


class StrictValidationError(ConversionRejection):
    """Strict-mode input or document rejection, with distinct codes in issue order."""

    def __init__(self, message: str, issues: Iterable[str]) -> None:
        issues = tuple(issues)
        codes = (
            code.value
            for issue in issues
            for code in RejectionCode
            if issue.startswith(f"[{code}] ")
        )
        super().__init__(message + "; ".join(issues), codes)


class AssetOutsideRootError(ValueError):
    """Internal asset failure; renderers turn it into a warning or strict rejection."""

    def __init__(self, path: str) -> None:
        super().__init__(
            coded_issue(
                RejectionCode.ASSET_OUTSIDE_ROOT,
                f"Image or logo is outside allowed folders: {path}",
            )
        )


def resolve_asset_roots(roots: Iterable[str | Path]) -> tuple[Path, ...]:
    """Resolve and deduplicate operator-supplied folders, following links.

    Args:
        roots: Allowed folders, including any source-file default.

    Returns:
        Immutable final folder locations for this request.
    """
    return tuple(dict.fromkeys(Path(root).resolve() for root in roots))


def local_asset_path(
    path: str, base_dir: Path | None = None, asset_roots: Iterable[str | Path] = ()
) -> Path | None:
    """Find an allowed file, trying the written destination before percent decoding.

    Args:
        path: Local destination as written in the document.
        base_dir: Relative-path base; None keeps paths cwd-relative without granting access.
        asset_roots: Allowed folders supplied by the caller, never document metadata.

    Returns:
        The resolved existing file, or None for an unreadable/missing allowed asset.

    Raises:
        AssetOutsideRootError: No candidate can be loaded and a candidate escapes the roots.
    """
    roots = resolve_asset_roots(asset_roots)
    outside = False
    for candidate in dict.fromkeys((path, unquote(path))):
        location = Path(candidate)
        if base_dir is not None and not location.is_absolute():
            location = base_dir / location
        try:
            location = location.resolve()
        except (OSError, ValueError, RuntimeError):
            # An unresolved final location cannot be proven to stay within a root.
            outside = True
            continue
        if not any(location.is_relative_to(root) for root in roots):
            outside = True
            continue
        try:
            if location.is_file():
                return location
        except (OSError, ValueError):
            continue
    if outside:
        raise AssetOutsideRootError(path)
    return None
