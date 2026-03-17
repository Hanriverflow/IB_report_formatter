"""Tests for shared CLI helpers."""

from cli_utils import resolve_input_path


def test_resolve_input_path_prefers_nested_cwd_path(tmp_path, monkeypatch):
    """Nested relative paths should resolve from cwd before parent basename fallback."""
    project_parent = tmp_path / "project_parent"
    cwd = tmp_path / "workspace"
    script_dir = cwd / "tools"
    nested_dir = cwd / "reports" / "q1"

    project_parent.mkdir()
    nested_dir.mkdir(parents=True)
    script_dir.mkdir(parents=True)

    wrong_file = project_parent / "report.docx"
    right_file = nested_dir / "report.docx"
    wrong_file.write_text("wrong", encoding="utf-8")
    right_file.write_text("right", encoding="utf-8")

    monkeypatch.chdir(cwd)

    resolved = resolve_input_path(
        "reports/q1/report.docx",
        parent_dir=project_parent,
        script_path=script_dir / "word_to_md.py",
    )

    assert resolved == right_file
