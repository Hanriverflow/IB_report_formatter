"""Tests for the Reports MD to Word CLI script."""

import argparse
import sys
from pathlib import Path
from types import SimpleNamespace
import pytest

from reports_md_to_word import (
    get_reports_dir,
    scan_markdown_reports,
    run_batch_conversion,
    build_parser,
)


def test_get_reports_dir():
    """Should correctly locate the reports/ directory."""
    reports_dir = get_reports_dir()
    assert reports_dir.name == "reports"
    assert reports_dir.parent == Path(__file__).resolve().parent.parent


def test_scan_markdown_reports_filters_out_formatted_and_cleaned(tmp_path, monkeypatch):
    """Should only list original .md files, filtering out _formatted.md and _cleaned.md."""
    # Mock get_reports_dir to return our temporary path
    monkeypatch.setattr("reports_md_to_word.get_reports_dir", lambda: tmp_path)

    # Create dummy files
    (tmp_path / "report_1.md").write_text("# Report 1", encoding="utf-8")
    (tmp_path / "report_1_formatted.md").write_text("# Formatted 1", encoding="utf-8")
    (tmp_path / "report_2.md").write_text("# Report 2", encoding="utf-8")
    (tmp_path / "report_2_cleaned.md").write_text("# Cleaned 2", encoding="utf-8")
    (tmp_path / "doc.docx").write_bytes(b"dummy docx")

    md_files, filter_count = scan_markdown_reports()

    assert len(md_files) == 2
    assert sorted([f.name for f in md_files]) == ["report_1.md", "report_2.md"]
    assert filter_count == 2


def test_run_batch_conversion(tmp_path, monkeypatch):
    """Batch conversion should scan all markdown files and execute conversion."""
    reports_mock_dir = tmp_path / "reports"
    reports_mock_dir.mkdir()
    monkeypatch.setattr("reports_md_to_word.get_reports_dir", lambda: reports_mock_dir)

    (reports_mock_dir / "report_a.md").write_text("# A", encoding="utf-8")
    (reports_mock_dir / "report_b.md").write_text("# B", encoding="utf-8")

    # Mock md_to_word.run_conversion to just return success
    conversions_called = []

    def mock_run_conversion(resolved_path, args):
        conversions_called.append(resolved_path)
        return 0

    monkeypatch.setattr("reports_md_to_word.run_conversion", mock_run_conversion)

    args = SimpleNamespace(
        output_file=None,
        no_cover=True,
        no_toc=True,
        no_disclaimer=True,
        separator_mode="auto",
        format=False,
        deepresearch_cleaner="off",
        cite_mode="footnote",
        drop_unknown_markers=False,
        cleaner_report=False,
        verbose=False,
        batch=True,
    )

    exit_code = run_batch_conversion(args)

    assert exit_code == 0
    assert len(conversions_called) == 2
    assert sorted([Path(p).name for p in conversions_called]) == ["report_a.md", "report_b.md"]


def test_build_parser_defaults():
    """Verify built parser arguments."""
    parser = build_parser()
    
    # Test blank args (interactive defaults)
    args = parser.parse_args([])
    assert args.input_file is None
    assert args.output_file is None
    assert args.interactive is False
    assert args.batch is False
    assert args.format is False
    assert args.deepresearch_cleaner == "off"
    assert args.cite_mode == "footnote"
    assert args.no_cover is False
    assert args.no_toc is False
    assert args.no_disclaimer is False
    assert args.separator_mode == "auto"

    # Test parser custom inputs
    args = parser.parse_args(["my_report.md", "out.docx", "--format", "--no-cover"])
    assert args.input_file == "my_report.md"
    assert args.output_file == "out.docx"
    assert args.format is True
    assert args.no_cover is True
