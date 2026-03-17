"""Tests for roundtrip_audit helper."""

from roundtrip_audit import build_audit_report


def test_roundtrip_audit_reports_markdown_summary(tmp_path):
    """Audit helper should summarize markdown round-trip results."""
    markdown = """---
title: "Audit Test"
---

# Heading

본문 문장입니다.
"""
    input_path = tmp_path / "audit.md"
    input_path.write_text(markdown, encoding="utf-8")

    report = build_audit_report(input_path)

    assert report.input_type == "markdown"
    assert report.source.title == "Audit Test"
    assert "HEADING_1" in report.source.element_counts
    assert report.roundtrip is not None
    assert report.diff is not None
