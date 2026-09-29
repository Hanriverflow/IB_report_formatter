"""Term-sheet profile composition: house boilerplate, opening block and layout helpers.

Called from `IBDocumentRenderer.render` (the single composition path) when the
resolved profile is `term-sheet`. Design: docs/term-sheet-design-20260929.md.
This module depends only on the model, styles, YAML and python-docx primitives;
it must not import `ib_renderer` (the renderer injects run-rendering callbacks).

Changelog (A1 foundation):
    - Validate immutable house boilerplate with presence-based frontmatter precedence.

Changelog (A2 rendering):
    - Schema-ordered single cell fills shared with generic merged-table emission.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

import yaml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tc

from document_model import DocumentMetadata

ROW_SPLIT_THRESHOLD = 12
# Elements that follow w:shd inside w:tcPr (ECMA-376 CT_TcPr sequence).
_TCPR_AFTER_SHD = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
    "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
)
_HOUSE_KEYS = frozenset({"prepared_by", "disclaimer", "confidential_label", "confirmation"})
_CONFIRMATION_KEYS = frozenset({"intro", "items", "signature"})


@dataclass(frozen=True)
class ConfirmationText:
    """Validated confirmation wording without mutable item collections."""

    intro: str = ""
    items: Tuple[str, ...] = ()
    signature: str = ""


@dataclass(frozen=True)
class TermSheetTexts:
    """Resolved boilerplate for one render; empty labels intentionally remain empty."""

    prepared_by: str
    disclaimer: str
    confidential_label: str = "Strictly Confidential"
    confirmation: Optional[ConfirmationText] = None


def _confirmation_text(value: Any) -> ConfirmationText:
    """Convert a supported confirmation mapping into an immutable payload."""
    if not isinstance(value, dict) or not value:
        raise ValueError("confirmation must be a mapping with intro, items or signature")
    unknown = set(value) - _CONFIRMATION_KEYS
    if unknown:
        raise ValueError("Unknown confirmation settings: " + ", ".join(sorted(map(str, unknown))))
    for key in ("intro", "signature"):
        if key in value and not isinstance(value[key], str):
            raise ValueError(f"confirmation.{key} must be a string")
    if "items" in value:
        items = value["items"]
        if (
            not isinstance(items, list)
            or not items
            or any(not isinstance(item, str) or not item.strip() for item in items)
        ):
            raise ValueError("confirmation.items must be a non-empty list of non-empty strings")
    return ConfirmationText(
        intro=value.get("intro", ""),
        items=tuple(value.get("items", [])),
        signature=value.get("signature", ""),
    )


def _validate_house_fields(data: Dict[str, Any]) -> None:
    """Reject unsupported keys and malformed values without requiring all fields."""
    unknown = set(data) - _HOUSE_KEYS
    if unknown:
        raise ValueError("Unknown house settings: " + ", ".join(sorted(map(str, unknown))))
    for key in ("prepared_by", "disclaimer", "confidential_label"):
        if key in data and not isinstance(data[key], str):
            raise ValueError(f"{key} must be a string")
    if "confirmation" in data:
        _confirmation_text(data["confirmation"])


def load_house(path: Union[str, Path]) -> Dict[str, Any]:
    """Load and validate the fields supplied by a house YAML file.

    Args:
        path: House file to read; callers resolve its source-relative location.

    Returns:
        Validated supplied fields, retaining which optional keys were present.

    Raises:
        ValueError: YAML syntax, mapping shape or a supplied field is invalid.
        OSError: The selected file cannot be read.
    """
    try:
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8-sig"))
    except yaml.YAMLError as exc:
        raise ValueError(f"Invalid house YAML: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("house must be a YAML mapping")
    _validate_house_fields(data)
    return data


def resolve_term_sheet_texts(
    metadata: DocumentMetadata, house_path: Optional[str] = None
) -> TermSheetTexts:
    """Resolve boilerplate before rendering, preserving explicit empty values.

    Args:
        metadata: Term-sheet metadata containing optional boilerplate overrides.
        house_path: Effective house path after caller/frontmatter precedence.
            Source-relative paths must already be absolute; absent paths use no file.

    Returns:
        Immutable texts for the renderer's current request.

    Raises:
        ValueError: A path is relative, a field is malformed, or required text is absent.
        OSError: The effective house file cannot be read.
    """
    house: Dict[str, Any] = {}
    if house_path is not None:
        if not isinstance(house_path, str) or not house_path.strip():
            raise ValueError("house must be a YAML file path")
        if not Path(house_path).is_absolute():
            raise ValueError("house path must be absolute for input without a source file")
        house = load_house(house_path)
    values = {
        key: metadata.extra[key] if key in metadata.extra else house[key]
        for key in _HOUSE_KEYS
        if key in metadata.extra or key in house
    }
    _validate_house_fields(values)
    for key in ("prepared_by", "disclaimer"):
        if key not in values or not values[key].strip():
            raise ValueError(f"term-sheet requires non-empty {key}")
    return TermSheetTexts(
        prepared_by=values["prepared_by"],
        disclaimer=values["disclaimer"],
        confidential_label=values.get("confidential_label", "Strictly Confidential"),
        confirmation=_confirmation_text(values["confirmation"])
        if "confirmation" in values else None,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE CELL PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════════


def set_cell_fill(tc: CT_Tc, hex_color: str) -> None:
    """Give one `w:tc` exactly one solid fill, placed in schema order.

    Args:
        tc: The cell element; a covered vertical-merge cell is a distinct `w:tc`.
        hex_color: Six-digit RGB fill without a leading `#`.
    """
    tc_pr = tc.get_or_add_tcPr()
    for shading in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(shading)
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), hex_color)
    tc_pr.insert_element_before(shading, *_TCPR_AFTER_SHD)
