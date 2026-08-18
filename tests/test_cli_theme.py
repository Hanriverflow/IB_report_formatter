"""Theme-option smoke tests for the Markdown-to-Word CLI."""

import logging
import sys
from types import SimpleNamespace

import pytest
from docx import Document
from docx.shared import RGBColor

import md_to_word
import theme_loader


@pytest.fixture(autouse=True)
def preserve_ib_style():
    """Restore all shared IB style values after every test."""
    snapshot = theme_loader.snapshot_style()
    yield
    theme_loader.restore_style(snapshot)


def _conversion_args(output_file, theme=None):
    """Build the CLI argument fields consumed by single-file conversion."""
    return SimpleNamespace(
        output_file=str(output_file),
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
        batch=False,
        theme=theme,
    )


def test_parser_theme_default_and_explicit_value():
    """The parser should default to no theme and accept a name or path value."""
    parser = md_to_word.build_parser()

    default_args = parser.parse_args(["file.md"])
    themed_args = parser.parse_args(["file.md", "--theme", "navy"])

    assert default_args.theme is None
    assert themed_args.theme == "navy"


def test_no_theme_does_not_import_loader(monkeypatch, tmp_path):
    """The default path should convert successfully without importing theme_loader."""
    monkeypatch.setitem(sys.modules, "theme_loader", None)
    input_path = tmp_path / "plain.md"
    output_path = tmp_path / "plain.docx"
    input_path.write_text("# Title\n\nPlain paragraph.\n", encoding="utf-8")

    assert md_to_word.apply_theme(None) == 0
    assert md_to_word.run_conversion(
        input_path,
        _conversion_args(output_path),
    ) == 0


def test_custom_theme_changes_rendered_bullet_color(tmp_path):
    """A custom theme should affect bullet text in the generated document."""
    theme_path = tmp_path / "custom.yaml"
    input_path = tmp_path / "themed.md"
    output_path = tmp_path / "themed.docx"
    theme_path.write_text('NAVY: "1A2B3C"\n', encoding="utf-8")
    input_path.write_text("# Title\n\n- Bullet item\n", encoding="utf-8")

    assert md_to_word.apply_theme(str(theme_path)) == 0
    assert md_to_word.run_conversion(
        input_path,
        _conversion_args(output_path, theme=str(theme_path)),
    ) == 0

    document = Document(str(output_path))
    bullet = next(p for p in document.paragraphs if "Bullet item" in p.text)

    assert any(
        run.font.color.rgb == RGBColor(0x1A, 0x2B, 0x3C)
        for run in bullet.runs
    )


def test_apply_theme_returns_error_for_missing_theme(caplog):
    """A missing theme should produce a non-zero CLI helper result."""
    with caplog.at_level(logging.ERROR, logger="md_to_word"):
        exit_code = md_to_word.apply_theme("does_not_exist")

    assert exit_code == 1
    assert "Theme load failed" in caplog.text
