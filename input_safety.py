"""Request-scoped asset access, input limits and conversion rejections.

Changelog (asset boundary):
    - Resolve local assets against operator-supplied roots before checking existence.
    - Preserve issue prefixes and expose codes on strict validation failures.
    - Bound Markdown reads, table grids, nesting and image headers per request.
"""

import base64
import logging
import re
from collections.abc import Iterable, Iterator, Mapping
from contextlib import contextmanager
from dataclasses import dataclass, fields, is_dataclass
from enum import StrEnum
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO
from urllib.parse import unquote

from docx.image.image import _ImageHeaderFactory
from yaml import SafeLoader
from yaml.nodes import MappingNode, Node, ScalarNode, SequenceNode

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
        self.issues = issues
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


@contextmanager
def document_validation(
    document: bool = True, *, errors: tuple[type[Exception], ...] = (ValueError,)
) -> Iterator[None]:
    """Translate a document validation boundary without relabeling caller errors.

    Args:
        document: False when the validated value came from an operator argument.
        errors: Specific validation failures at this boundary, including file
            reading or YAML syntax errors only for document-named configuration.

    Yields:
        Control to the validator, preserving its message and original cause.
    """
    try:
        yield
    except ConversionRejection:
        raise
    except errors as error:
        if document:
            raise StrictValidationError("", (str(error),)) from error
        raise


# Values and characters YAML may expand to through aliases, however short it is.
_YAML_EXPANSION_FLOOR = 100_000
# Merge keys one mapping may hold. YAML needs one (its value may list several
# mappings); the loader removes each from the mapping in turn, so thousands in
# one mapping take quadratic time with no alias involved.
_YAML_MERGE_KEYS = 100
_YAML_MERGE_TAG = "tag:yaml.org,2002:merge"
# Integer keys of this many characters or more that one mapping may hold, counting
# the keys it merges in. Python hashes an integer by its remainder modulo
# 2**61 - 1, so integers at least that large (eighteen characters in the
# shortest notation) can be chosen to share one hash value, and a dict takes
# quadratic time to hold thousands of them.
_YAML_LONG_INTEGER_KEYS = 100
_YAML_LONG_INTEGER_CHARACTERS = 18
_YAML_INTEGER_TAG = "tag:yaml.org,2002:int"
# Colons one integer may hold. YAML 1.1 reads `1:30:00` as a base-60 integer,
# and the loader builds it part by part with ever larger numbers, so a scalar
# of thousands of parts takes quadratic time. A time of day has two.
_YAML_INTEGER_COLONS = 20


def _yaml_children(node: Node) -> list[Node]:
    if isinstance(node, MappingNode):
        return [child for pair in node.value for child in pair]
    return list(node.value)


def _check_yaml_expansion(root: Node, allowance: int) -> None:
    """Refuse parsed YAML that contains itself or expands past the allowance.

    Size counts one for each node and each character of a scalar, and counts an
    aliased node again wherever it is used, which is what a merge key copies and
    what anything that prints or walks the loaded value pays for. Without
    aliases a node and a character each take at least a character of YAML, so
    YAML that uses none stays near its own length (about one and a half times
    in the densest forms), well under an allowance of four times.

    Each collection is opened once, however often it is aliased, so the check
    itself takes time in proportion to the parsed YAML.

    Args:
        root: The composed document, before any value is constructed.
        allowance: The largest size accepted.

    Raises:
        ValueError: The graph is recursive or larger than the allowance, one
            mapping holds more merge keys than `_YAML_MERGE_KEYS` or more very
            long integer keys than `_YAML_LONG_INTEGER_KEYS`, or an integer
            holds more colons than `_YAML_INTEGER_COLONS`.
    """
    sizes: dict[int, int] = {}
    # Very long integer keys a mapping ends up with, or a list's mappings hold together.
    long_keys: dict[int, int] = {}
    active: set[int] = set()
    pending: list[tuple[Node, bool]] = [(root, False)]
    while pending:
        node, finished = pending.pop()
        if not isinstance(node, (MappingNode, SequenceNode)):
            if node.tag == _YAML_INTEGER_TAG and node.value.count(":") > _YAML_INTEGER_COLONS:
                raise ValueError("YAML integer has too many parts separated by ':'")
            continue
        identity = id(node)
        if finished:
            active.discard(identity)
            sizes[identity] = 1 + sum(
                max(1, len(child.value)) if isinstance(child, ScalarNode) else sizes[id(child)]
                for child in _yaml_children(node)
            )
            if sizes[identity] > allowance:
                raise ValueError("YAML aliases expand to too much content")
            if isinstance(node, SequenceNode):
                # What a merge key takes from this list when the list is its value.
                long_keys[identity] = sum(long_keys.get(id(child), 0) for child in node.value)
                continue
            long_keys[identity] = sum(
                long_keys.get(id(value), 0)
                if key.tag == _YAML_MERGE_TAG
                else isinstance(key, ScalarNode)
                and key.tag == _YAML_INTEGER_TAG
                and len(key.value) >= _YAML_LONG_INTEGER_CHARACTERS
                for key, value in node.value
            )
            if long_keys[identity] > _YAML_LONG_INTEGER_KEYS:
                raise ValueError(
                    "YAML mapping has too many very long integer keys"
                    " (quote the keys to keep them as text)"
                )
        elif identity in active:
            raise ValueError("YAML structure refers to itself")
        elif identity not in sizes:
            if isinstance(node, MappingNode):
                merges = sum(1 for key, _ in node.value if key.tag == _YAML_MERGE_TAG)
                if merges > _YAML_MERGE_KEYS:
                    raise ValueError("YAML mapping repeats the merge key too many times")
            active.add(identity)
            pending.append((node, True))
            pending.extend((child, False) for child in _yaml_children(node))


def _check_yaml_integers(data: object) -> None:
    """Refuse an integer that loads but is too long for Python to write out again."""
    seen: set[int] = set()
    pending: list[object] = [data]
    while pending:
        item = pending.pop()
        if isinstance(item, int):
            try:
                str(item)
            except ValueError as error:
                raise ValueError("YAML integer is too long to use") from error
        elif isinstance(item, Mapping):
            if id(item) not in seen:
                seen.add(id(item))
                pending.extend(item.keys())
                pending.extend(item.values())
        elif isinstance(item, (list, tuple, set, frozenset)) and id(item) not in seen:
            seen.add(id(item))
            pending.extend(item)


def load_yaml(text: str) -> Any:
    """Load YAML with the safe loader; a value it cannot construct is a ValueError.

    PyYAML's scalar constructors raise plain AttributeError, KeyError, IndexError
    or TypeError for malformed tagged scalars (`!!timestamp x`, `!!bool x`,
    `!!int ""`), and deep nesting ends in RecursionError. Callers already treat
    yaml.YAMLError and ValueError as invalid YAML, so those are left as they are.
    An integer too long for Python to convert back to text is reported the same way.

    Aliases let a few hundred characters describe a structure that refers to
    itself, that expands to billions of values, or that repeats one long text
    or one large mapping thousands of times. The loader shares aliased values,
    so most of these load at once and then stall whatever prints or walks them,
    and a merge key copies its mapping while loading. They are refused on the
    parsed YAML, before any value is constructed.

    Args:
        text: YAML written in a document or in a file a document or operator named.

    Returns:
        The constructed value, or None for empty YAML.

    Raises:
        yaml.YAMLError: Syntax or structure the loader rejects.
        ValueError: A value the loader could not construct, a structure that is
            recursive or expands too far through aliases, or a mapping that
            repeats the merge key too many times.
    """
    # The steps of yaml.safe_load, with the alias check between parsing and construction.
    try:
        loader = SafeLoader(text)
        try:
            node = loader.get_single_node()
            if node is None:
                return None
            _check_yaml_expansion(node, max(_YAML_EXPANSION_FLOOR, 4 * len(text)))
            data = loader.construct_document(node)
        finally:
            loader.dispose()
    except (AttributeError, KeyError, IndexError, TypeError, OverflowError) as error:
        raise ValueError(f"YAML value cannot be constructed: {error!r}") from error
    except RecursionError as error:
        raise ValueError("YAML is nested too deeply") from error
    _check_yaml_integers(data)
    return data


class _XMLText:
    INVALID = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def validate_xml_strings(value: object) -> None:
    """Check decoded document strings before they reach Word's XML writer.

    Args:
        value: Metadata or decoded configuration, including nested YAML values.

    Raises:
        ValueError: A string contains an XML-incompatible character.
    """
    pending = [value]
    visited: set[int] = set()
    while pending:
        item = pending.pop()
        if isinstance(item, str):
            invalid = _XMLText.INVALID.search(item)
            if invalid is not None:
                raise ValueError(
                    f"XML-incompatible character U+{ord(invalid.group()):04X} in Markdown input"
                )
        elif id(item) not in visited:
            visited.add(id(item))
            if is_dataclass(item) and not isinstance(item, type):
                pending.extend(getattr(item, field.name) for field in fields(item))
            elif isinstance(item, Mapping):
                pending.extend(item.keys())
                pending.extend(item.values())
            elif isinstance(item, (list, tuple, set)):
                pending.extend(item)


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
