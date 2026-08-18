"""Tests for YAML preset resolution, defaults, and field validation."""

import pytest

from preset_loader import PresetConfig, load_preset

# ═══════════════════════════════════════════════════════════════════════════════
# Bundled presets
# ═══════════════════════════════════════════════════════════════════════════════


def test_all_bundled_presets_load():
    """Every bundled document-type preset loads without error."""
    for name in ("ib-report", "termsheet", "legal-memo", "lecture-note"):
        assert isinstance(load_preset(name), PresetConfig)


def test_ib_report_preset_has_no_overrides():
    """The default IB report preset leaves every setting unchanged."""
    preset = load_preset("ib-report")

    assert preset.theme is None
    assert preset.include_cover is None
    assert preset.include_toc is None
    assert preset.include_disclaimer is None
    assert preset.separator_mode is None


def test_termsheet_preset_disables_optional_sections():
    """The term sheet preset disables sections and defaults omitted fields."""
    preset = load_preset("termsheet")

    assert preset.include_cover is False
    assert preset.include_toc is False
    assert preset.include_disclaimer is False
    assert preset.theme is None
    assert preset.separator_mode is None


# ═══════════════════════════════════════════════════════════════════════════════
# Validation and resolution errors
# ═══════════════════════════════════════════════════════════════════════════════


def test_unknown_preset_key_raises_value_error(tmp_path):
    """Unknown keys are rejected as preset-schema drift."""
    preset_path = tmp_path / "unknown.yaml"
    preset_path.write_text("not_a_real_field: true\n", encoding="utf-8")

    with pytest.raises(ValueError, match="not_a_real_field"):
        load_preset(str(preset_path))


def test_bad_separator_mode_raises_value_error(tmp_path):
    """Separator modes outside the fixed option set are rejected."""
    preset_path = tmp_path / "bad-separator.yaml"
    preset_path.write_text(
        'separator_mode: "not-a-real-mode"\n', encoding="utf-8"
    )

    with pytest.raises(ValueError, match="not-a-real-mode"):
        load_preset(str(preset_path))


@pytest.mark.parametrize("value", ['"yes"', "1"])
def test_non_bool_include_cover_raises_value_error(tmp_path, value):
    """String and integer values are not accepted as booleans."""
    preset_path = tmp_path / "bad-cover.yaml"
    preset_path.write_text(f"include_cover: {value}\n", encoding="utf-8")

    with pytest.raises(ValueError, match="include_cover"):
        load_preset(str(preset_path))


def test_missing_preset_raises_file_not_found_error():
    """A missing path and bundled name produce a clear resolution error."""
    with pytest.raises(FileNotFoundError, match="this-preset-does-not-exist"):
        load_preset("this-preset-does-not-exist")
