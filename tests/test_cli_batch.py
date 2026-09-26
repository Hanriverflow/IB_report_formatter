"""Batch-mode smoke tests for the CLI entry points."""

from types import SimpleNamespace

from md_to_word import run_batch_conversion as run_md_batch_conversion


def test_md_to_word_batch_conversion_creates_outputs(tmp_path):
    """Batch mode should convert every markdown file in a directory."""
    input_dir = tmp_path / "md_inputs"
    output_dir = tmp_path / "docx_outputs"
    input_dir.mkdir()

    (input_dir / "one.md").write_text("# One\n\n본문입니다.\n", encoding="utf-8")
    (input_dir / "two.md").write_text("# Two\n\n둘째 문서입니다.\n", encoding="utf-8")

    args = SimpleNamespace(
        output_file=str(output_dir),
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

    exit_code = run_md_batch_conversion(input_dir, args)

    assert exit_code == 0
    assert len(list(output_dir.glob("*.docx"))) == 2
