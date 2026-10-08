"""Source bundle confidentiality, integrity and extracted-engine checks."""

import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("build_codex_bundle", ROOT / "scripts/build_codex_bundle.py")
assert SPEC is not None and SPEC.loader is not None
builder = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(builder)


@pytest.fixture
def public_checkout(tmp_path):
    checkout = tmp_path / "source"
    _, paths = builder.source_files(ROOT)
    for relative in paths:
        target = checkout / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, target)
    return checkout


def test_bundle_has_exact_allowlist_and_verifiable_hashes(public_checkout, tmp_path):
    secret_paths = [
        ".env", "draft.md", "private/real-deal.md", "docs/internal-draft.md",
        "samples/harness/client-secret.yaml", ".agents/skills/ib-document/private.md",
        "dist/old.exe", ".venv/secret.txt", "output/private.docx",
    ]
    for relative in secret_paths:
        target = public_checkout / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("DO_NOT_DISTRIBUTE", encoding="utf-8")
    archive = builder.build_bundle(public_checkout, tmp_path / "release")
    version, paths = builder.source_files(public_checkout)
    prefix = f"IB_report_formatter_Codex_{version}/"
    with zipfile.ZipFile(archive) as bundle:
        assert {
            prefix + name for name in (
                "docs/distribution-build.md", "docs/distribution/시작하기.html",
                "scripts/build_portable.py", "tools/__init__.py", "tools/portable_launcher.py",
                "samples/distribution/minimal-term-sheet.md",
                "samples/distribution/minimal-house.yaml",
            )
        }.issubset(bundle.namelist())
        assert not any(name.endswith((".exe", ".dll", ".pyd", ".docx")) for name in bundle.namelist())
        assert set(bundle.namelist()) == {
            prefix + name for name in [*paths, "시작하기.html", "bundle-manifest.json"]
        }
        manifest = json.loads(bundle.read(prefix + "bundle-manifest.json"))
        assert {item["path"] for item in manifest["files"]} == set(paths) | {"시작하기.html"}
        for item in manifest["files"]:
            data = bundle.read(prefix + item["path"])
            assert len(data) == item["bytes"]
            assert hashlib.sha256(data).hexdigest() == item["sha256"]
            assert b"DO_NOT_DISTRIBUTE" not in data
        assert bundle.read(prefix + "시작하기.html") == bundle.read(prefix + "docs/codex-start.html")
        assert manifest["source"]["base_commit"] is None
        assert manifest["source"]["dirty"] is None
    assert archive.with_suffix(".zip.sha256").read_text().split()[0] == hashlib.sha256(archive.read_bytes()).hexdigest()


def test_repeated_build_preserves_previous_release(public_checkout, tmp_path):
    first = builder.build_bundle(public_checkout, tmp_path / "release")
    original = first.read_bytes()
    second = builder.build_bundle(public_checkout, tmp_path / "release")
    assert first != second
    assert first.read_bytes() == original


def test_missing_required_file_fails_before_creating_release(public_checkout, tmp_path):
    (public_checkout / "samples/harness/house.yaml").unlink()
    with pytest.raises(FileNotFoundError, match="samples/harness/house.yaml"):
        builder.build_bundle(public_checkout, tmp_path / "release")
    assert not (tmp_path / "release").exists()


def test_module_path_escape_is_rejected(public_checkout, tmp_path):
    config = public_checkout / "pyproject.toml"
    config.write_text(config.read_text(encoding="utf-8").replace('"chart_renderer.py"', '"../private.py"'), encoding="utf-8")
    with pytest.raises(ValueError, match="top-level Python modules"):
        builder.build_bundle(public_checkout, tmp_path / "release")


def test_symlink_source_escape_is_rejected(public_checkout, tmp_path):
    target = public_checkout / "samples/harness/house.yaml"
    outside = tmp_path / "outside.yaml"
    outside.write_text("private material", encoding="utf-8")
    target.unlink()
    try:
        target.symlink_to(outside)
    except OSError:
        pytest.skip("Creating symbolic links is unavailable for this account")
    with pytest.raises(ValueError, match="leaves the project"):
        builder.build_bundle(public_checkout, tmp_path / "release")


def test_extracted_source_renders_without_original_checkout(public_checkout, tmp_path):
    archive = builder.build_bundle(public_checkout, tmp_path / "release")
    extracted = tmp_path / "압축 해제"
    with zipfile.ZipFile(archive) as bundle:
        bundle.extractall(extracted)
    project = next(extracted.iterdir())
    # The interpreter supplies installed third-party dependencies. All engine
    # code and example input come from the extracted ZIP, with a foreign cwd.
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    # -I also suppresses cwd and PYTHONPATH. This entry point explicitly adds
    # its own extracted source root, so imports cannot fall back to checkout.
    result = subprocess.run(
        [sys.executable, "-I", str(project / "scripts/agent_render.py"),
         str(project / "samples/profiles/plain.md"), "--output-dir", str(tmp_path / "run")],
        cwd=tmp_path, env=environment, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    assert result.returncode == 0, result.stdout + result.stderr
    evidence = json.loads(result.stdout)
    assert evidence["ok"] is True
    rendered = Path(evidence["output"]["path"])
    with zipfile.ZipFile(rendered) as docx:
        assert "word/document.xml" in docx.namelist()
