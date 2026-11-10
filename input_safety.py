"""Request-scoped asset access and machine-readable conversion rejections.

Changelog (asset boundary):
    - Resolve local assets against operator-supplied roots before checking existence.
    - Preserve issue prefixes and expose codes on strict validation failures.
"""

from collections.abc import Iterable
from enum import StrEnum
from pathlib import Path
from urllib.parse import unquote


class RejectionCode(StrEnum):
    """Codes implemented by the engine; future input limits extend this enum."""

    ASSET_OUTSIDE_ROOT = "asset_outside_root"


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
    """Base for documented rejections, including future input-limit exceptions."""

    def __init__(self, message: str, codes: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.codes = list(dict.fromkeys(codes))


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
