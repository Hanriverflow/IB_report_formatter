"""Evidence and save-safety checks for the Codex-native rendering wrapper."""

import base64
import hashlib
import json
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest
from docx import Document

import cli_utils
from scripts.agent_render import render_document

ROOT = Path(__file__).resolve().parents[1]
SOURCE = """---
profile: plain
title: 검토안
terms:
  amount: "300억원"
  company: "가상회사"
---
# 검토안

{{company}} 조달금액은 {{amount}}입니다.
"""


def make_source(tmp_path: Path, text: str = SOURCE) -> Path:
    source = tmp_path / "한글 문서.md"
    source.write_text(text, encoding="utf-8")
    return source


def test_real_render_reports_hashes_audit_and_utf8_manifest(tmp_path):
    source = make_source(tmp_path)
    expected = tmp_path / "확정 조건.json"
    expected.write_text(json.dumps({"amount": "300억원"}), encoding="utf-8")
    result = render_document(str(source), str(tmp_path / "출력 폴더"), expected_terms=str(expected))
    assert result["ok"] and result["status"] == "ready_for_review"
    assert result["profile"] == "plain"
    assert result["expected_terms"]["status"] == "matched"
    assert result["expected_terms"]["checked"] == 1
    assert result["visual_review"] == "not_performed"
    assert result["audit"]["issues"] == []
    assert "warnings" in result["audit"]
    output = Path(result["output"]["path"])
    assert output.exists()
    assert result["output"]["sha256"] == hashlib.sha256(output.read_bytes()).hexdigest()
    assert result["input"]["sha256"] == hashlib.sha256(source.read_bytes()).hexdigest()
    assert json.loads((output.parent / "result.json").read_text(encoding="utf-8")) == result
    assert "300억원" in "".join(Document(output).element.itertext())


@pytest.mark.parametrize("expected", [{"amount": "301억원"}, {"absent": "300억원"}])
def test_expected_terms_mismatch_never_renders(tmp_path, expected):
    source = make_source(tmp_path)
    terms = tmp_path / "terms.json"
    terms.write_text(json.dumps(expected), encoding="utf-8")
    output = tmp_path / "out"
    result = render_document(str(source), str(output), expected_terms=str(terms))
    assert not result["ok"] and result["stage"] == "expected_terms"
    assert result["expected_terms"]["status"] == "mismatch"
    assert not list(output.glob("*.docx"))
    assert (output / "result.json").exists()


@pytest.mark.parametrize("payload", ['{"amount":300}', '[]', '{"x":"a","x":"b"}'])
def test_expected_terms_rejects_invalid_maps(tmp_path, payload):
    source = make_source(tmp_path)
    terms = tmp_path / "terms.json"
    terms.write_text(payload, encoding="utf-8")
    result = render_document(str(source), str(tmp_path / "out"), expected_terms=str(terms))
    assert not result["ok"] and result["stage"] == "expected_terms"
    assert result["output"]["path"] is None


def test_empty_comparison_does_not_claim_terms_verified(tmp_path):
    terms = tmp_path / "terms.json"
    terms.write_text("{}", encoding="utf-8")
    result = render_document(str(make_source(tmp_path)), str(tmp_path / "out"), expected_terms=str(terms))
    assert result["ok"]
    assert result["expected_terms"]["status"] == "empty"
    assert result["expected_terms"]["checked"] == 0


def test_raw_term_spacing_is_compared_exactly(tmp_path):
    source = make_source(tmp_path, SOURCE.replace('"300억원"', '" 300  억원 "'))
    terms = tmp_path / "terms.json"
    terms.write_text(json.dumps({"amount": " 300  억원 "}), encoding="utf-8")
    result = render_document(str(source), str(tmp_path / "out"), expected_terms=str(terms))
    assert result["ok"] and result["expected_terms"]["status"] == "matched"


@pytest.mark.parametrize("source_text", [SOURCE + "\n{{unknown}}\n", SOURCE + "\n![missing](missing.png)\n"])
def test_strict_failures_leave_no_docx(tmp_path, source_text):
    source = make_source(tmp_path, source_text)
    output = tmp_path / "out"
    result = render_document(str(source), str(output))
    assert not result["ok"] and result["diagnostics"]
    assert result["stage"] in {"parse", "render"}
    assert not list(output.glob("*.docx"))


def test_existing_directory_and_source_containment_are_safe(tmp_path):
    source = make_source(tmp_path)
    output = tmp_path / "out"
    output.mkdir()
    sentinel = output / "result.json"
    sentinel.write_bytes(b"existing content")
    result = render_document(str(source), str(output))
    assert not result["ok"] and result["stage"] == "setup"
    assert sentinel.read_bytes() == b"existing content"
    before = source.read_bytes()
    result = render_document(str(source), str(tmp_path))
    assert not result["ok"] and result["stage"] == "setup"
    assert source.read_bytes() == before


def test_concurrent_reservation_allows_only_one_job(tmp_path):
    source = make_source(tmp_path)
    with ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(
            lambda _: render_document(str(source), str(tmp_path / "shared")), range(2),
        ))
    assert sum(result["ok"] for result in results) == 1
    assert next(result for result in results if not result["ok"])["stage"] == "setup"


def test_profile_override_reaches_parser_and_renderer(tmp_path):
    source = make_source(tmp_path, "---\nprofile: ib-report\n---\n1. 짧은 항목\n")
    result = render_document(str(source), str(tmp_path / "out"), profile="plain")
    assert result["ok"] and result["profile"] == "plain"
    document = Document(result["output"]["path"])
    assert document.element.xpath(".//w:pPr/w:numPr")
    assert result["audit"]["headings"] == 0


def test_sample_house_resolves_from_original_source_directory(tmp_path):
    result = render_document(str(ROOT / "samples/profiles/term-sheet.md"), str(tmp_path / "out"))
    assert result["ok"], result
    assert result["profile"] == "term-sheet"
    assert result["audit"]["terms"]["failed_checks"] == []


def test_relative_image_resolves_from_original_source_directory(tmp_path):
    source_dir = tmp_path / "원본 자료"
    source_dir.mkdir()
    (source_dir / "그림.png").write_bytes(base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0xsAAAAASUVORK5CYII="
    ))
    source = make_source(source_dir, SOURCE + "\n![그림](그림.png)\n")
    result = render_document(str(source), str(tmp_path / "out"))
    assert result["ok"], result
    assert result["audit"]["images"] == 1


def test_audit_warnings_remain_distinct_from_structural_issues(tmp_path):
    source = make_source(tmp_path, SOURCE + "\n[Image: literal quoted example]\n")
    result = render_document(str(source), str(tmp_path / "out"))
    assert result["ok"]
    assert result["audit"]["issues"] == []
    assert any("placeholder" in warning for warning in result["audit"]["warnings"])
    assert result["diagnostics"] == []


def test_locked_save_reports_actual_fallback_path(tmp_path, monkeypatch):
    atomic_save = cli_utils._atomic_save

    def locked_first(path, save_action):
        if path.name == "document.docx":
            raise PermissionError("Simulated locked destination")
        atomic_save(path, save_action)

    monkeypatch.setattr(cli_utils, "_atomic_save", locked_first)
    result = render_document(str(make_source(tmp_path)), str(tmp_path / "out"))
    assert result["ok"]
    assert Path(result["output"]["path"]).name != "document.docx"
    assert Path(result["output"]["path"]).exists()


def test_cli_unicode_and_structured_argument_errors(tmp_path):
    source = make_source(tmp_path)
    command = [sys.executable, str(ROOT / "scripts/agent_render.py")]
    success = subprocess.run(
        command + [str(source), "--output-dir", str(tmp_path / "out")],
        cwd=tmp_path, capture_output=True, check=False,
    )
    assert success.returncode == 0, success.stderr.decode("utf-8", errors="replace")
    assert json.loads(success.stdout)["input"]["path"] == str(source.resolve())
    failure = subprocess.run(command, cwd=tmp_path, capture_output=True, check=False)
    assert failure.returncode == 2
    assert json.loads(failure.stdout)["stage"] == "arguments"
