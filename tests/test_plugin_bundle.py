"""Validate root-layout plugin archives, compatibility metadata and source isolation."""

import hashlib
import importlib.util
import json
import shutil
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("build_plugin_bundle", ROOT / "scripts/build_plugin_bundle.py")
assert SPEC is not None and SPEC.loader is not None
builder = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(builder)


@pytest.fixture
def plugin_checkout(tmp_path):
    checkout = tmp_path / "plugin-source"
    _, _, paths = builder.public_sources(ROOT)
    for relative in paths:
        target = checkout / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / relative, target)
    return checkout


def test_plugin_archive_layout_hashes_and_no_private_inputs(plugin_checkout, tmp_path):
    for name in (".env", "private/real.md", "skills/setup/secret.md", "dist/program.exe",
                 ".git/config", ".venv/private.py", "output/real.docx", "docs/internal.md"):
        path = plugin_checkout / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("PRIVATE_MUST_NOT_SHIP", encoding="utf-8")
    archive = builder.build_bundle(plugin_checkout, tmp_path / "release")
    assert archive.read_bytes() == archive.with_suffix(".plugin").read_bytes()
    plugin_version, engine_version, paths = builder.public_sources(plugin_checkout)
    with zipfile.ZipFile(archive) as bundle:
        assert set(bundle.namelist()) == set(paths) | {"bundle-manifest.json"}
        assert "plugin.json" in bundle.namelist()
        assert "skills/setup/SKILL.md" in bundle.namelist()
        assert not any(name.startswith((".git/", ".venv/", "private/", ".agents/skills/"))
                       for name in bundle.namelist())
        manifest = json.loads(bundle.read("bundle-manifest.json"))
        assert manifest["plugin_version"] == plugin_version == "0.1.0"
        assert manifest["engine_version"] == engine_version == "2.0.0"
        assert {item["path"] for item in manifest["files"]} == set(paths)
        for item in manifest["files"]:
            data = bundle.read(item["path"])
            assert len(data) == item["bytes"]
            assert hashlib.sha256(data).hexdigest() == item["sha256"]
            assert b"PRIVATE_MUST_NOT_SHIP" not in data
    for suffix in (".zip", ".plugin"):
        artifact = archive.with_suffix(suffix)
        assert artifact.with_suffix(suffix + ".sha256").read_text().split()[0] == hashlib.sha256(artifact.read_bytes()).hexdigest()


def test_manifests_use_same_identity_and_local_root(plugin_checkout):
    manifests = [json.loads((plugin_checkout / path).read_text(encoding="utf-8"))
                 for path in builder.MANIFESTS]
    assert {(item["name"], item["version"]) for item in manifests} == {("ib-report-formatter", "0.1.0")}
    onboarding = manifests[0]["extensions"]["com.openai"]["onboardingSkill"]
    assert onboarding == "./skills/setup/SKILL.md"
    assert (plugin_checkout / onboarding).is_file()
    codex = json.loads((plugin_checkout / ".agents/plugins/marketplace.json").read_text())
    claude = json.loads((plugin_checkout / ".claude-plugin/marketplace.json").read_text())
    assert codex["plugins"][0]["source"] == {"source": "local", "path": "./"}
    assert claude["plugins"][0]["source"] == "./"
    assert codex["name"] == claude["name"] == "ib-report-formatter-local"


def test_missing_skill_prevents_partial_release(plugin_checkout, tmp_path):
    (plugin_checkout / "skills/setup/SKILL.md").unlink()
    with pytest.raises(FileNotFoundError, match="skills/setup/SKILL.md"):
        builder.build_bundle(plugin_checkout, tmp_path / "release")
    assert not (tmp_path / "release").exists()


def test_mismatched_plugin_versions_fail(plugin_checkout, tmp_path):
    path = plugin_checkout / ".claude-plugin/plugin.json"
    data = json.loads(path.read_text())
    data["version"] = "9.9.9"
    path.write_text(json.dumps(data))
    with pytest.raises(ValueError, match="identities must agree"):
        builder.build_bundle(plugin_checkout, tmp_path / "release")


def test_module_traversal_rejected(plugin_checkout, tmp_path):
    path = plugin_checkout / "pyproject.toml"
    path.write_text(path.read_text(encoding="utf-8").replace('"chart_renderer.py"', '"../secret.py"'), encoding="utf-8")
    with pytest.raises(ValueError, match="top-level Python modules"):
        builder.build_bundle(plugin_checkout, tmp_path / "release")


def test_source_link_escape_rejected(plugin_checkout, tmp_path):
    external = tmp_path / "secret.md"
    external.write_text("secret")
    path = plugin_checkout / "skills/setup/SKILL.md"
    path.unlink()
    try:
        path.symlink_to(external)
    except OSError:
        pytest.skip("Account cannot create symbolic links")
    with pytest.raises(ValueError, match="leaves the project"):
        builder.build_bundle(plugin_checkout, tmp_path / "release")


def test_repeated_plugin_build_preserves_artifacts(plugin_checkout, tmp_path):
    first = builder.build_bundle(plugin_checkout, tmp_path / "release")
    data = first.read_bytes()
    second = builder.build_bundle(plugin_checkout, tmp_path / "release")
    assert first != second
    assert first.read_bytes() == data
