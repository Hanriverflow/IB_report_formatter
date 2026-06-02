"""Tests for the style profile system (classic / ib-pro global swap)."""

import pytest

import ib_renderer
import style_profiles


def test_default_profile_is_classic():
    # conftest autouse fixture resets to classic before each test
    assert style_profiles.get_active_profile() == "classic"


def test_switch_to_ib_pro_rebinds_global_style():
    style_profiles.set_active_profile("ib-pro")
    assert ib_renderer.STYLE.PROFILE == "ib-pro"
    assert style_profiles.get_active_profile() == "ib-pro"


def test_set_active_profile_returns_instance():
    inst = style_profiles.set_active_profile("ib-pro")
    assert inst.PROFILE == "ib-pro"


def test_ib_pro_preserves_classic_values_in_phase1():
    classic = style_profiles.set_active_profile("classic")
    ibpro = style_profiles.set_active_profile("ib-pro")
    # Phase 1: ib-pro is identical to classic apart from its profile label.
    assert ibpro.NAVY == classic.NAVY
    assert ibpro.H1_SIZE == classic.H1_SIZE
    assert ibpro.BODY_FONT == classic.BODY_FONT
    assert ibpro.TOP_MARGIN == classic.TOP_MARGIN


def test_unknown_profile_raises():
    with pytest.raises(ValueError):
        style_profiles.set_active_profile("bogus")


def test_valid_profiles_constant():
    assert style_profiles.VALID_PROFILES == ("classic", "ib-pro")
