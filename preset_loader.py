"""Load document-type presets from YAML configuration files.

Presets provide optional overrides for document structure and presentation.
Missing fields and explicit YAML null values leave the corresponding default
unchanged for later CLI or renderer handling.
"""

import logging
from dataclasses import dataclass, fields
from pathlib import Path
from typing import Dict, Optional

import yaml

logger = logging.getLogger(__name__)

PRESETS_DIR = Path(__file__).resolve().parent / "presets"

_SEPARATOR_MODES = {"auto", "rule", "page-break"}
_BOOLEAN_FIELDS = {"include_cover", "include_toc", "include_disclaimer"}


@dataclass
class PresetConfig:
    """Optional document settings supplied by a preset."""

    theme: Optional[str] = None
    include_cover: Optional[bool] = None
    include_toc: Optional[bool] = None
    include_disclaimer: Optional[bool] = None
    separator_mode: Optional[str] = None


def resolve_preset_path(name_or_path: str) -> Path:
    """Resolve an existing path or a bundled preset name.

    Args:
        name_or_path: Existing YAML path or a preset name without an extension.

    Returns:
        Resolved input path, or the matching path under ``PRESETS_DIR``.

    Raises:
        FileNotFoundError: If no candidate path exists.
    """
    supplied_path = Path(name_or_path)
    preset_path = PRESETS_DIR / (name_or_path + ".yaml")

    if supplied_path.exists():
        return supplied_path
    if preset_path.exists():
        return preset_path

    tried = [str(supplied_path), str(preset_path)]
    raise FileNotFoundError(
        "Preset {!r} was not found. Tried: {}".format(
            name_or_path,
            ", ".join(tried),
        )
    )


def _validate_value(name: str, value: object) -> object:
    """Validate one non-null preset value.

    Args:
        name: Preset field name, used in validation errors.
        value: Value parsed from YAML.

    Returns:
        The validated value unchanged.

    Raises:
        ValueError: If the value does not satisfy the field schema.
    """
    if name == "theme":
        if not isinstance(value, str):
            raise ValueError(f"Preset field {name} must be a string")
        return value

    if name in _BOOLEAN_FIELDS:
        if not isinstance(value, bool):
            raise ValueError(f"Preset field {name} must be a boolean")
        return value

    if name == "separator_mode":
        if not isinstance(value, str) or value not in _SEPARATOR_MODES:
            valid = ", ".join(sorted(_SEPARATOR_MODES))
            raise ValueError(
                f"Preset field {name} has invalid value {value!r}; "
                f"valid options are: {valid}"
            )
        return value

    return value


def load_preset(name_or_path: str) -> PresetConfig:
    """Load and validate a YAML document preset.

    Args:
        name_or_path: Existing YAML path or bundled preset name.

    Returns:
        Fresh preset configuration containing only declared overrides.

    Raises:
        FileNotFoundError: If the preset cannot be resolved.
        ValueError: If the YAML root, keys, or values are invalid.
    """
    preset_path = resolve_preset_path(name_or_path)
    with preset_path.open("r", encoding="utf-8") as stream:
        raw_preset = yaml.safe_load(stream)

    if raw_preset is None:
        raw_preset = {}
    if not isinstance(raw_preset, dict):
        raise ValueError("Preset YAML must contain a top-level mapping")

    valid_fields = {field.name for field in fields(PresetConfig)}
    unknown_fields = [key for key in raw_preset if key not in valid_fields]
    if unknown_fields:
        unknown = ", ".join(repr(key) for key in unknown_fields)
        raise ValueError(f"Unknown preset field: {unknown}")

    validated_kwargs: Dict[str, object] = {}
    for key, value in raw_preset.items():
        if value is not None:
            validated_kwargs[key] = _validate_value(key, value)

    config = PresetConfig(**validated_kwargs)
    overridden = ", ".join(validated_kwargs) or "none"
    logger.info("Loaded preset from %s with overrides: %s", preset_path, overridden)
    return config
