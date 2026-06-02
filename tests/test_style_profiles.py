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


def test_ib_pro_sets_improvement_toggles():
    ibpro = style_profiles.set_active_profile("ib-pro")
    assert ibpro.BODY_JUSTIFY is False
    assert ibpro.TABLE_BORDER_STYLE == "horizontal"


def test_classic_keeps_legacy_toggles():
    classic = style_profiles.set_active_profile("classic")
    assert classic.BODY_JUSTIFY is True
    assert classic.TABLE_BORDER_STYLE == "grid"


def test_ib_pro_table_borders_drop_vertical_rules():
    from docx import Document
    from docx.oxml.ns import qn

    from ib_renderer import TableStyler

    style_profiles.set_active_profile("ib-pro")
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    TableStyler.set_table_borders(table)
    borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    assert borders.find(qn("w:insideV")).get(qn("w:val")) == "none"
    assert borders.find(qn("w:insideH")).get(qn("w:val")) == "single"


def test_classic_table_borders_keep_vertical_rules():
    from docx import Document
    from docx.oxml.ns import qn

    from ib_renderer import TableStyler

    style_profiles.set_active_profile("classic")
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    TableStyler.set_table_borders(table)
    borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    assert borders.find(qn("w:insideV")).get(qn("w:val")) == "dotted"


def test_ib_pro_header_has_thick_navy_rule():
    from docx import Document
    from docx.oxml.ns import qn

    from ib_renderer import STYLE, TableStyler

    style_profiles.set_active_profile("ib-pro")
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    TableStyler.set_table_borders(table)
    cell = table.rows[0].cells[0]
    tc_borders = cell._tc.get_or_add_tcPr().find(qn("w:tcBorders"))
    bottom = tc_borders.find(qn("w:bottom"))
    assert bottom.get(qn("w:sz")) == "18"
    assert bottom.get(qn("w:color")) == STYLE.NAVY_HEX


def test_classic_header_has_no_per_cell_rule():
    from docx import Document
    from docx.oxml.ns import qn

    from ib_renderer import TableStyler

    style_profiles.set_active_profile("classic")
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    TableStyler.set_table_borders(table)
    tcPr = table.rows[0].cells[0]._tc.tcPr
    assert tcPr is None or tcPr.find(qn("w:tcBorders")) is None
