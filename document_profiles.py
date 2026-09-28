"""Document profiles, validated metadata and request-local rendering options.

Changelog (feature port):
    - Immutable section presets and opt-in charts resolved in the shared path.
    - Typed presentation themes, including PR #5's uppercase field names.
"""

import math
import re
from dataclasses import dataclass, fields, replace
from pathlib import Path
from types import MappingProxyType
from typing import Any, Dict, List, Mapping, Optional

import yaml
from docx.shared import Inches, Pt, RGBColor

from document_model import DocumentMetadata, DocumentModel, ElementType, Table, TableType
from render_styles import IBStyle


@dataclass(frozen=True)
class DocumentProfile:
    """Defaults for a document type; explicit options may override composition."""

    name: str
    is_ib: bool = False
    cover: bool = False
    toc: bool = False
    disclaimer: bool = False
    confidential: bool = False
    a4: bool = True


PROFILES = {
    "ib-report": DocumentProfile("ib-report", True, True, True, True, True, False),
    "ib-memo": DocumentProfile("ib-memo", is_ib=True, confidential=True),
    "plain": DocumentProfile("plain"),
    "office-letter": DocumentProfile("office-letter"),
    "business-report": DocumentProfile("business-report"),
    "meeting-minutes": DocumentProfile("meeting-minutes"),
}


def get_profile(name: str) -> DocumentProfile:
    """Resolve a supported profile and reject spelling mistakes."""
    if name not in PROFILES:
        raise ValueError("Unknown profile {!r}; choose {}".format(name, ", ".join(PROFILES)))
    return PROFILES[name]


def default_metadata(profile: str = "ib-report") -> DocumentMetadata:
    """Return neutral metadata for office documents, legacy defaults for IB."""
    if get_profile(profile).is_ib:
        return DocumentMetadata(profile=profile)
    return DocumentMetadata(title="Document", company="", sector="", analyst="", profile=profile)


@dataclass(frozen=True)
class RenderOptions:
    """Explicit caller overrides; None inherits frontmatter/profile defaults."""

    include_cover: Optional[bool] = None
    include_toc: Optional[bool] = None
    include_disclaimer: Optional[bool] = None
    separator_mode: Optional[str] = None
    profile: Optional[str] = None
    theme: Optional[str] = None
    strict: Optional[bool] = None
    confidential: Optional[bool] = None
    charts: Optional[bool] = None
    preset: Optional[str] = None


PRESETS: Mapping[str, RenderOptions] = MappingProxyType({
    "ib-report": RenderOptions(),
    "termsheet": RenderOptions(include_cover=False, include_toc=False, include_disclaimer=False),
    "legal-memo": RenderOptions(include_cover=False, include_toc=True, include_disclaimer=False),
    "lecture-note": RenderOptions(include_cover=True, include_toc=True, include_disclaimer=False),
})


def get_preset(name: Optional[str]) -> RenderOptions:
    """Resolve an immutable section bundle, with None supplying no overrides.

    Args:
        name: Built-in preset name, independent of the document profile.

    Returns:
        An immutable bundle of optional section toggles.
    """
    if name is None:
        return RenderOptions()
    if not isinstance(name, str) or name not in PRESETS:
        raise ValueError("Unknown preset {!r}; choose {}".format(name, ", ".join(PRESETS)))
    return PRESETS[name]


@dataclass(frozen=True)
class ResolvedOptions:
    """Fully resolved, immutable settings for one conversion."""

    profile: DocumentProfile
    cover: bool
    toc: bool
    disclaimer: bool
    confidential: bool
    separator_mode: str
    theme: Optional[str]
    strict: bool
    charts: bool


def resolve_options(metadata: DocumentMetadata, options: RenderOptions) -> ResolvedOptions:
    """Resolve CLI/API > frontmatter > profile, validating known settings."""
    profile = get_profile(options.profile or metadata.profile)
    layout = metadata.extra.get("layout", {})
    if not isinstance(layout, dict):
        raise ValueError("layout must be a YAML mapping")
    allowed = {"cover", "toc", "disclaimer", "confidential", "separator_mode", "strict"}
    unknown = set(layout) - allowed
    if unknown:
        raise ValueError("Unknown layout settings: {}".format(", ".join(sorted(unknown))))

    caller_preset = get_preset(options.preset)
    yaml_preset = get_preset(metadata.extra.get("preset"))

    def flag(key: str, override: Optional[bool], default: bool) -> bool:
        attribute = "include_" + key
        caller_value = getattr(caller_preset, attribute, None)
        yaml_value = getattr(yaml_preset, attribute, None)
        value = default if yaml_value is None else yaml_value
        value = layout.get(key, value)
        if caller_value is not None:
            value = caller_value
        if override is not None:
            value = override
        if not isinstance(value, bool):
            raise ValueError(f"layout.{key} must be true or false")
        return value

    separator = options.separator_mode or layout.get("separator_mode", "auto")
    if separator not in {"auto", "rule", "page-break"}:
        raise ValueError("separator_mode must be auto, rule or page-break")
    theme = options.theme if options.theme is not None else metadata.extra.get("theme")
    if theme is not None and not isinstance(theme, str):
        raise ValueError("theme must be a name or YAML file path")
    charts = options.charts if options.charts is not None else metadata.extra.get("charts", False)
    if not isinstance(charts, bool):
        raise ValueError("charts must be true or false")
    return ResolvedOptions(
        profile=profile,
        cover=flag("cover", options.include_cover, profile.cover),
        toc=flag("toc", options.include_toc, profile.toc),
        disclaimer=flag("disclaimer", options.include_disclaimer, profile.disclaimer),
        confidential=flag("confidential", options.confidential, profile.confidential),
        separator_mode=separator,
        theme=theme,
        strict=flag("strict", options.strict, False),
        charts=charts,
    )


def load_style(profile: DocumentProfile, theme: Optional[str] = None) -> IBStyle:
    """Build an isolated style from a profile and a validated presentation theme.

    Args:
        profile: Profile defaults, including its numbering policy.
        theme: default, mono, or a YAML path with typed presentation fields.

    Returns:
        A new immutable style; no global values are changed.
    """
    style = IBStyle()
    if profile.name == "ib-memo":
        style = replace(
            style,
            BODY_SPACE_AFTER=Pt(6),
            H2_SPACE_BEFORE=Pt(10),
            TOP_MARGIN=Inches(0.79),
            BOTTOM_MARGIN=Inches(0.79),
            LEFT_MARGIN=Inches(0.85),
            RIGHT_MARGIN=Inches(0.85),
        )
    if not profile.is_ib or theme == "mono":
        style = replace(
            style,
            NAVY=RGBColor(32, 32, 32),
            NAVY_HEX="202020",
            HEADING_FONT="Malgun Gothic",
            BODY_FONT="Malgun Gothic",
            H1_SIZE=Pt(16),
            BODY_SIZE=Pt(11),
            BODY_LINE_SPACING=1.2,
            TOP_MARGIN=Inches(0.79),
            BOTTOM_MARGIN=Inches(0.79),
            LEFT_MARGIN=Inches(0.79),
            RIGHT_MARGIN=Inches(0.79),
            TABLE_HEADER_BG="EEEEEE",
            TABLE_HEADER_COLOR=RGBColor(32, 32, 32),
            TABLE_ZEBRA=False,
            BODY_JUSTIFY=False,
            HEADING_BORDER=False,
            NATIVE_NUMBERING=True,
            TOC_TITLE="목차",
            PAGE_LABEL="",
            PAGE_OF_LABEL=" / ",
            CHART_NEGATIVE_COLOR=RGBColor(64, 64, 64),
        )
    if profile.name == "office-letter":
        style = replace(style, BODY_LINE_SPACING=1.45, BODY_SPACE_AFTER=Pt(8))
    if not theme or theme in {"default", "mono"}:
        return style
    path = Path(theme)
    data = yaml.safe_load(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict):
        raise ValueError("Theme must be a YAML mapping")
    aliases = {
        "body_font",
        "heading_font",
        "korean_font",
        "body_size",
        "primary_color",
        "margin_mm",
    }
    attributes = {field.name for field in fields(IBStyle)
                  if not field.name.startswith("STYLE_") and field.name != "NATIVE_NUMBERING"}
    if set(data) - (aliases | attributes):
        raise ValueError(
            "Unknown theme settings: {}".format(", ".join(sorted(map(str, set(data) - (aliases | attributes)))))
        )
    values: Dict[str, Any] = {}
    for key, attribute in [
        ("body_font", "BODY_FONT"),
        ("heading_font", "HEADING_FONT"),
        ("korean_font", "KOREAN_FONT"),
    ]:
        if key in data:
            if not isinstance(data[key], str) or not data[key].strip():
                raise ValueError(f"{key} must be a non-empty font name")
            values[attribute] = data[key]
    if "korean_font" in data:
        values.update(COVER_FONT=data["korean_font"], TOC_FONT=data["korean_font"])
    if "body_size" in data:
        size = data["body_size"]
        if isinstance(size, bool) or not isinstance(size, (int, float)) or not 6 <= size <= 30:
            raise ValueError("body_size must be between 6 and 30 points")
        values["BODY_SIZE"] = Pt(size)
    if "primary_color" in data:
        color = _ThemeSchema.color("primary_color", data["primary_color"])
        values.update(NAVY=RGBColor.from_string(color.upper()), NAVY_HEX=color.upper())
        if profile.is_ib:
            values["TABLE_HEADER_BG"] = color.upper()
    if "margin_mm" in data:
        margin = data["margin_mm"]
        if (
            isinstance(margin, bool)
            or not isinstance(margin, (int, float))
            or not 5 <= margin <= 60
        ):
            raise ValueError("margin_mm must be between 5 and 60")
        for key in ["TOP_MARGIN", "BOTTOM_MARGIN", "LEFT_MARGIN", "RIGHT_MARGIN"]:
            values[key] = Inches(margin / 25.4)
    explicit = {}
    for key in data:
        if key in attributes:
            explicit[key] = _ThemeSchema.convert(key, data[key], getattr(style, key))
    conflicts = set(values) & set(explicit)
    if any(values[key] != explicit[key] for key in conflicts):
        raise ValueError("Conflicting theme aliases: " + ", ".join(sorted(conflicts)))
    values.update(explicit)
    # Keep RGB/OOXML representations aligned unless both are explicitly supplied.
    for rgb, hex_key in (("NAVY", "NAVY_HEX"), ("LIGHT_GRAY", "LIGHT_GRAY_HEX"),
                         ("ACCENT_BLUE", "ACCENT_BLUE_HEX")):
        if rgb in values and hex_key not in values:
            values[hex_key] = str(values[rgb])
        elif hex_key in values and rgb not in values:
            values[rgb] = RGBColor.from_string(values[hex_key])
    if profile.is_ib and "NAVY" in values and "TABLE_HEADER_BG" not in values:
        values["TABLE_HEADER_BG"] = str(values["NAVY"])
    if "RED" in values and "CHART_NEGATIVE_COLOR" not in values:
        values["CHART_NEGATIVE_COLOR"] = values["RED"]
    return replace(style, **values)


class _ThemeSchema:
    """Strict scalar conversion for presentation fields, without shared mutation."""

    COLOR = re.compile(r"^#?([0-9a-fA-F]{6})$")

    @classmethod
    def color(cls, key: str, value: Any) -> str:
        if not isinstance(value, str) or not cls.COLOR.fullmatch(value):
            raise ValueError(f"{key} must be a six-digit hexadecimal color string")
        return value.lstrip("#").upper()

    @classmethod
    def convert(cls, key: str, value: Any, current: Any) -> Any:
        if isinstance(current, RGBColor):
            return RGBColor.from_string(cls.color(key, value))
        if key.endswith("_HEX") or key == "TABLE_HEADER_BG":
            return cls.color(key, value)
        if isinstance(current, bool):
            if not isinstance(value, bool):
                raise ValueError(f"{key} must be a boolean")
            return value
        if isinstance(current, (Pt, Inches, int, float)):
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ValueError(f"{key} must be numeric")
            try:
                finite = math.isfinite(value)
            except OverflowError:
                finite = False
            if not finite:
                raise ValueError(f"{key} must be finite")
            if isinstance(current, Pt):
                minimum = 0 if "SPACE" in key else 1
                if not minimum <= value <= 144:
                    raise ValueError(f"{key} must be between {minimum} and 144 points")
                return Pt(value)
            if isinstance(current, Inches):
                if not 0 <= value <= 3:
                    raise ValueError(f"{key} must be between 0 and 3 inches")
                return Inches(value)
            if isinstance(current, int):
                if not isinstance(value, int) or not 0 <= value <= 9:
                    raise ValueError(f"{key} must be an integer between 0 and 9")
                return value
            if not 0 < value <= 5:
                raise ValueError(f"{key} must be greater than 0 and at most 5")
            return float(value)
        if not isinstance(value, str) or (key.endswith("_FONT") and not value.strip()):
            raise ValueError(f"{key} must be a string (font names must be non-empty)")
        return value


def string_list(extra: Dict[str, Any], key: str) -> List[str]:
    """Read a scalar or string list without stringifying mappings silently."""
    value = extra.get(key, [])
    if isinstance(value, str):
        return [value] if value.strip() else []
    if not isinstance(value, list) or not all(isinstance(v, str) and v.strip() for v in value):
        raise ValueError(f"{key} must be a string or a list of non-empty strings")
    return value


def validate_office_metadata(metadata: DocumentMetadata) -> None:
    """Validate structured fields used by office composers."""
    for key in ("recipients", "cc", "attachments", "attendees"):
        if key in metadata.extra:
            string_list(metadata.extra, key)
    sender = metadata.extra.get("sender", {})
    if not isinstance(sender, dict) or not all(isinstance(v, str) for v in sender.values()):
        raise ValueError("sender must be a mapping of text fields")
    if metadata.profile == "office-letter":
        if not string_list(metadata.extra, "recipients"):
            raise ValueError("office-letter requires recipients")
        if not sender.get("organization", "").strip():
            raise ValueError("office-letter requires sender.organization")
        letter = metadata.extra.get("letter", {})
        if not isinstance(letter, dict):
            raise ValueError("letter must be a YAML mapping")
        if set(letter) - {"appendix_heading", "appendix_label"}:
            raise ValueError("Unknown letter settings")
        if any(not isinstance(v, str) or not v.strip() for v in letter.values()):
            raise ValueError("letter settings must be non-empty text")
        if "appendix_label" in letter and "appendix_heading" not in letter:
            raise ValueError("letter.appendix_label requires appendix_heading")


def apply_table_specs(model: DocumentModel) -> None:
    """Apply explicit table semantics in document order, never infer a base case."""
    tables = [
        e.content
        for e in model.elements
        if e.element_type == ElementType.TABLE and isinstance(e.content, Table)
    ]
    specs = model.metadata.extra.get("tables", [])
    if not isinstance(specs, list) or len(specs) > len(tables):
        raise ValueError("tables must be a list with at most one specification per body table")
    kinds = {
        "generic": TableType.GENERIC,
        "financial": TableType.FINANCIAL,
        "sensitivity": TableType.BEP_SENSITIVITY,
        "risk": TableType.RISK_MATRIX,
    }
    roles = {"text", "code", "date", "number", "money", "percent", "bps", "multiple"}
    for index, spec in enumerate(specs):
        if not isinstance(spec, dict):
            raise ValueError(f"tables[{index}] must be a mapping")
        unknown = set(spec) - {
            "type",
            "columns",
            "caption",
            "unit",
            "source",
            "as_of",
            "landscape",
            "base_case",
        }
        if unknown:
            raise ValueError("Unknown table settings: {}".format(", ".join(sorted(unknown))))
        table = tables[index]
        if "type" in spec:
            if spec["type"] not in kinds:
                raise ValueError("Unknown table type: {}".format(spec["type"]))
            table.table_type = kinds[spec["type"]]
        columns = spec.get("columns", [])
        if not isinstance(columns, list) or (columns and len(columns) != table.col_count):
            raise ValueError("Table columns must match the table width")
        if any(not isinstance(c, str) or c not in roles for c in columns):
            raise ValueError("Unknown table column role")
        table.column_types = columns
        for key in ("caption", "unit", "source", "as_of"):
            value = spec.get(key, "")
            if isinstance(value, (dict, list)):
                raise ValueError(f"Table {key} must be text")
            setattr(table, key, str(value))
        landscape = spec.get("landscape", False)
        if not isinstance(landscape, bool):
            raise ValueError("Table landscape must be true or false")
        table.landscape = landscape
        for row in table.rows:
            for cell in row.cells:
                cell.is_base_case = False
                cell.risk_level = None
        if table.table_type == TableType.RISK_MATRIX:
            for body_row in table.rows[1:]:
                for column_index, cell in enumerate(body_row.cells):
                    header = table.rows[0].cells[column_index].content.lower()
                    if any(
                        key in header
                        for key in ("impact", "probability", "영향", "확률", "등급", "level")
                    ):
                        value = "".join(run.text for run in cell.runs).lower()
                        for level, labels in [
                            ("high", ("high", "높음")),
                            ("medium", ("medium", "moderate", "중간", "보통")),
                            ("low", ("low", "낮음")),
                        ]:
                            if value in labels:
                                cell.risk_level = level
        if "base_case" in spec:
            base = spec["base_case"]
            if not isinstance(base, dict) or set(base) != {"row", "column"}:
                raise ValueError("base_case requires one-based row and column")
            row, column = base["row"], base["column"]
            if (
                type(row) is not int
                or type(column) is not int
                or not (1 <= row < len(table.rows) and 1 <= column <= table.col_count)
            ):
                raise ValueError("base_case is outside the table body")
            if table.table_type != TableType.BEP_SENSITIVITY:
                raise ValueError("base_case requires a sensitivity table")
            table.rows[row].cells[column - 1].is_base_case = True
