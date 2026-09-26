"""Word QA preflight safety tests; do not launch Word or require Office."""

import os
import shutil
import subprocess
from pathlib import Path

import pytest

POWERSHELL = shutil.which("powershell.exe")
pytestmark = pytest.mark.skipif(os.name != "nt" or not POWERSHELL, reason="Windows PowerShell required")
SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "word_visual_qa.ps1"


def run_qa(source, output, *args):
    return subprocess.run(
        [POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT),
         "-InputDirectory", str(source), "-OutputDirectory", str(output), *args],
        capture_output=True, timeout=30,
    )


def test_qa_rejects_reused_output_before_opening_word(tmp_path):
    source, output = tmp_path / "source", tmp_path / "qa"
    source.mkdir()
    output.mkdir()
    evidence = output / ".prior-evidence"
    evidence.write_text("Keep existing evidence", encoding="utf-8")
    result = run_qa(source, output)
    assert result.returncode != 0
    assert b"OutputDirectory must be new or empty" in result.stderr
    assert evidence.read_text(encoding="utf-8") == "Keep existing evidence"


def test_qa_rejects_negative_expected_page_count_before_opening_word(tmp_path):
    result = run_qa(tmp_path, tmp_path / "qa", "-ExpectedPages", "-1")
    assert result.returncode != 0
    assert b"ExpectedPages must be zero" in result.stderr
    assert not (tmp_path / "qa").exists()
