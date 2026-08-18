"""Load external YAML profiles into the IB report style class and singleton.

Theme values are converted according to the current value type on ``IBStyle``.
This keeps serialized point, inch, and color values unambiguous while allowing
plain YAML scalar values to retain their native types.
"""

import logging
import re
from pathlib import Path
from typing import Dict

import yaml
from docx.shared import Inches, Pt, RGBColor

import ib_renderer

logger = logging.getLogger(__name__)

THEMES_DIR = Path(__file__).resolve().parent / "themes"

_SENTINEL = object()
_HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _public_style_attributes() -> Dict[str, object]:
    """Return public style names with values from the live singleton.

    Returns:
        Mapping of public attribute names to their current values.
    """
    attributes = {}
    for name in dir(ib_renderer.IBStyle):
        if name.startswith("_"):
            continue

        class_value = getattr(ib_renderer.IBStyle, name)
        if callable(class_value):
            continue

        attributes[name] = getattr(ib_renderer.STYLE, name)

    return attributes


def resolve_theme_path(name_or_path: str) -> Path:
    """Resolve an existing path or a bundled theme name.

    Args:
        name_or_path: Existing YAML path or a theme name without an extension.

    Returns:
        Resolved input path, or the matching path under ``THEMES_DIR``.

    Raises:
        FileNotFoundError: If no candidate path exists.
    """
    supplied_path = Path(name_or_path)
    theme_path = THEMES_DIR / (name_or_path + ".yaml")

    if supplied_path.exists():
        return supplied_path
    if theme_path.exists():
        return theme_path

    tried = [str(supplied_path), str(theme_path)]
    raise FileNotFoundError(
        "Theme {!r} was not found. Tried: {}".format(
            name_or_path,
            ", ".join(tried),
        )
    )


def _convert_value(name: str, value: object, current: object) -> object:
    """Convert one serialized theme value to its target attribute type.

    Args:
        name: Style attribute name, used in validation errors.
        value: Value parsed from YAML.
        current: Current value on ``IBStyle`` that determines target type.

    Returns:
        Converted value suitable for assignment to ``IBStyle``.

    Raises:
        ValueError: If a color or measurement value is malformed.
    """
    if isinstance(current, RGBColor):
        if not isinstance(value, str) or not _HEX_COLOR_RE.fullmatch(value):
            raise ValueError(
                f"Theme field {name} must be a 6-character hexadecimal color"
            )
        return RGBColor.from_string(value.upper())

    if isinstance(current, Pt):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"Theme field {name} must be numeric")
        return Pt(value)

    if isinstance(current, Inches):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"Theme field {name} must be numeric")
        return Inches(value)

    if isinstance(current, bool):
        if not isinstance(value, bool):
            raise ValueError(f"Theme field {name} must be a boolean")
        return value

    if isinstance(current, int):
        if isinstance(value, bool) or not isinstance(value, int):
            raise ValueError(f"Theme field {name} must be an integer")
        return value

    if isinstance(current, float):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"Theme field {name} must be numeric")
        return float(value)

    if isinstance(current, str):
        if not isinstance(value, str):
            raise ValueError(f"Theme field {name} must be a string")
        return value

    if not isinstance(value, type(current)):
        raise ValueError(
            f"Theme field {name} must have type {type(current).__name__}"
        )
    return value


def load_theme(name_or_path: str) -> Dict[str, object]:
    """Load a YAML theme and apply it to the style class and singleton.

    Args:
        name_or_path: Existing YAML path or bundled theme name.

    Returns:
        Mapping of attribute names to the converted values that were applied.

    Raises:
        FileNotFoundError: If the theme cannot be resolved.
        ValueError: If the YAML root is not a mapping or a key is not public.
    """
    theme_path = resolve_theme_path(name_or_path)
    with theme_path.open("r", encoding="utf-8") as stream:
        raw_theme = yaml.safe_load(stream)

    if not isinstance(raw_theme, dict):
        raise ValueError("Theme YAML must contain a top-level mapping")

    public_attributes = _public_style_attributes()
    converted_values = {}

    for key, value in raw_theme.items():
        current = (
            getattr(ib_renderer.STYLE, key, _SENTINEL)
            if isinstance(key, str)
            else _SENTINEL
        )
        if current is _SENTINEL or key not in public_attributes:
            raise ValueError(f"Unknown IBStyle theme field: {key!r}")

        converted_values[key] = _convert_value(key, value, current)

    # Validate the complete document before mutating global renderer state. This
    # keeps a malformed profile from leaking a partially applied theme into later
    # conversions performed by the same process.
    for key, converted in converted_values.items():
        setattr(ib_renderer.IBStyle, key, converted)
        object.__setattr__(ib_renderer.STYLE, key, converted)

    logger.info("Loaded %d style values from %s", len(converted_values), theme_path)
    return converted_values


def snapshot_style() -> Dict[str, object]:
    """Capture every public style value from the live singleton."""
    return _public_style_attributes()


def restore_style(snapshot: Dict[str, object]) -> None:
    """Restore style class and singleton attributes from a prior snapshot.

    Args:
        snapshot: Mapping previously returned by :func:`snapshot_style`.
    """
    for key, value in snapshot.items():
        setattr(ib_renderer.IBStyle, key, value)
        object.__setattr__(ib_renderer.STYLE, key, value)
