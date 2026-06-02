"""Tests for the IB style audit tool (tools/ib_style_audit.py)."""

import sys
from pathlib import Path

_TOOLS = Path(__file__).resolve().parent.parent / "tools"
if str(_TOOLS) not in sys.path:
    sys.path.insert(0, str(_TOOLS))

import ib_style_audit  # noqa: E402


# ── pure-function rubric / diff tests (no docx needed) ──────────────────────


def test_diff_identical_returns_empty():
    a = {"x": 1, "y": {"z": 2}, "list": [1, 2]}
    b = {"x": 1, "y": {"z": 2}, "list": [1, 2]}
    assert ib_style_audit.diff_features(a, b) == []


def test_diff_detects_value_change():
    diffs = ib_style_audit.diff_features({"size": 10}, {"size": 12})
    assert any("10" in d and "12" in d for d in diffs)


def test_diff_detects_added_and_removed_keys():
    diffs = ib_style_audit.diff_features({"only_a": 1}, {"only_b": 2})
    joined = "\n".join(diffs)
    assert "only_a" in joined and "only_b" in joined


def test_rubric_type_hierarchy_pass():
    features = {
        "styles": {
            "Heading 1": {"size_pt": 14},
            "Heading 2": {"size_pt": 12},
            "Heading 3": {"size_pt": 11},
        },
        "colors": [],
        "tables": [],
        "alignment": {},
    }
    results = ib_style_audit.evaluate_rubric(features)
    th = next(r for r in results if r["item"] == "type_hierarchy")
    assert th["verdict"] == "pass"


def test_rubric_type_hierarchy_fail_when_inverted():
    features = {
        "styles": {
            "Heading 1": {"size_pt": 10},
            "Heading 2": {"size_pt": 12},
            "Heading 3": {"size_pt": 14},
        },
        "colors": [],
        "tables": [],
        "alignment": {},
    }
    results = ib_style_audit.evaluate_rubric(features)
    th = next(r for r in results if r["item"] == "type_hierarchy")
    assert th["verdict"] == "fail"


def test_rubric_flags_justified_body():
    features = {"styles": {}, "colors": [], "tables": [], "alignment": {"JUSTIFY": 3}}
    results = ib_style_audit.evaluate_rubric(features)
    ba = next(r for r in results if r["item"] == "body_alignment")
    assert ba["verdict"] == "warn"


def test_rubric_color_restraint_levels():
    few = {"styles": {}, "colors": ["A", "B"], "tables": [], "alignment": {}}
    many = {"styles": {}, "colors": [f"C{i}" for i in range(12)], "tables": [], "alignment": {}}
    few_v = next(r for r in ib_style_audit.evaluate_rubric(few) if r["item"] == "color_restraint")
    many_v = next(r for r in ib_style_audit.evaluate_rubric(many) if r["item"] == "color_restraint")
    assert few_v["verdict"] == "pass"
    assert many_v["verdict"] == "fail"


# ── integration: render + extract + diff (slower, exercises full path) ───────


def test_audit_markdown_phase1_diff_is_identical(tmp_path):
    md = tmp_path / "sample.md"
    md.write_text(
        "# Title\n\nBody paragraph.\n\n## Section\n\n| A | B |\n|---|---|\n| 1 | 2 |\n",
        encoding="utf-8",
    )
    result = ib_style_audit.audit_markdown(md)
    # Phase 1: ib-pro must reproduce classic exactly.
    assert result["diff"] == []
    assert result["classic"]["rubric"]
    assert result["ib_pro"]["rubric"]
