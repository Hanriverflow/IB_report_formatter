"""Package the source-based document plugin for local Codex and Claude installs.

Archives have plugin.json and skills/ at their root. The .plugin file is the
same ZIP bytes under a manual-import extension, not proof of host acceptance.
No dependency runtime, executable, generated document or private input is bundled.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import subprocess
import tomllib
import uuid
import zipfile
from datetime import UTC, datetime
from pathlib import Path

logger = logging.getLogger(__name__)
PUBLIC_FILES = (
    "docs/start-here.md",
    "plugin.json", ".codex-plugin/plugin.json", ".claude-plugin/plugin.json",
    ".agents/plugins/marketplace.json", ".claude-plugin/marketplace.json",
    "skills/ib-document/SKILL.md", "skills/ib-document/references/term-sheet.md",
    "skills/setup/SKILL.md", "docs/plugin-guide.md", "docs/codex-start.html",
    "pyproject.toml", "uv.lock", ".python-version", "LICENSE", "README.md", "README.ko.md",
    "scripts/agent_render.py", "scripts/plugin_runtime.py", "scripts/build_plugin_bundle.py",
    "scripts/word_visual_qa.ps1",
    "samples/profiles/business-report.md", "samples/profiles/company-theme.yaml",
    "samples/profiles/ib-memo.md", "samples/profiles/ib-report.md",
    "samples/profiles/meeting-minutes.md", "samples/profiles/office-letter.md",
    "samples/profiles/plain.md", "samples/profiles/term-sheet-house.yaml",
    "samples/profiles/term-sheet.md",
    "samples/harness/brief.md", "samples/harness/request.txt",
    "samples/harness/house.yaml", "samples/harness/expected-terms.json",
)
MANIFESTS = ("plugin.json", ".codex-plugin/plugin.json", ".claude-plugin/plugin.json")


def public_sources(source_root: Path) -> tuple[str, str, list[str]]:
    """Validate public sources and return plugin version, engine version and paths.

    Args:
        source_root: Plugin source directory, also the engine project root.

    Returns:
        Plugin version, engine version and the explicit file allowlist.
    """
    root = source_root.resolve()
    config = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))
    modules = config["tool"]["hatch"]["build"]["targets"]["wheel"]["only-include"]
    if not modules or any(not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*\.py", name)
                          for name in modules):
        raise ValueError("Wheel allowlist must contain only top-level Python modules")
    paths = sorted(set(PUBLIC_FILES).union(modules))
    for relative in paths:
        path = root / relative
        if not path.is_file():
            raise FileNotFoundError(f"Missing public plugin source: {relative}")
        if not path.resolve().is_relative_to(root):
            raise ValueError(f"Public plugin source leaves the project: {relative}")
        if any(part.is_symlink() for part in (path, *path.parents) if part != root):
            raise ValueError(f"Public plugin source contains a symbolic link: {relative}")
    manifests = [json.loads((root / name).read_text(encoding="utf-8")) for name in MANIFESTS]
    identity = {(manifest["name"], manifest["version"]) for manifest in manifests}
    if len(identity) != 1:
        raise ValueError("Portable, Codex and Claude plugin identities must agree")
    name, version = identity.pop()
    if name != "ib-report-formatter" or not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise ValueError("Plugin name or version is invalid")
    return version, config["project"]["version"], paths


def _provenance(root: Path, paths: list[str]) -> dict:
    def git(*args: str) -> str:
        return subprocess.run(
            ["git", "-c", "core.quotepath=false", *args], cwd=root, check=True,
            capture_output=True, text=True, encoding="utf-8",
        ).stdout.strip()

    try:
        return {
            "base_commit": git("rev-parse", "HEAD"),
            "dirty": bool(git("status", "--porcelain", "--untracked-files=normal")),
            "modified_files": git("diff", "--name-only", "HEAD", "--", *paths).splitlines(),
            "new_files": git("ls-files", "--others", "--exclude-standard", "--", *paths).splitlines(),
        }
    except (OSError, subprocess.CalledProcessError):
        return {"base_commit": None, "dirty": None, "modified_files": None, "new_files": None}


def build_bundle(source_root: Path, output_dir: Path) -> Path:
    """Build a unique plugin ZIP and a byte-identical .plugin archive.

    Args:
        source_root: Plugin and engine checkout.
        output_dir: Parent for a new release directory; existing files are preserved.

    Returns:
        Absolute path of the new ZIP. Its .plugin sibling has the same contents.
    """
    root = source_root.resolve()
    version, engine_version, paths = public_sources(root)
    payloads = {name: (root / name).read_bytes() for name in paths}
    manifest = {
        "format_version": 1, "distribution": "document-plugin",
        "plugin_version": version, "engine_version": engine_version,
        "created_at_utc": datetime.now(UTC).isoformat(),
        "source": _provenance(root, paths),
        "provenance_note": (
            "base_commit does not identify uncommitted changes. File hashes identify the "
            "actual package bytes; the manifest itself is excluded from its file list."
        ),
        "files": [
            {"path": name, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
            for name, data in sorted(payloads.items())
        ],
    }
    release = output_dir.resolve() / (
        "plugin-" + datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8]
    )
    release.mkdir(parents=True, exist_ok=False)
    archive = release / f"ib-report-formatter-plugin-{version}.zip"
    with zipfile.ZipFile(archive, "x", compression=zipfile.ZIP_DEFLATED) as bundle:
        for name, data in sorted(payloads.items()):
            bundle.writestr(name, data)
        bundle.writestr("bundle-manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    data = archive.read_bytes()
    plugin = archive.with_suffix(".plugin")
    with plugin.open("xb") as stream:
        stream.write(data)
    checksum = hashlib.sha256(data).hexdigest()
    for artifact in (archive, plugin):
        artifact.with_suffix(artifact.suffix + ".sha256").write_text(
            f"{checksum}  {artifact.name}\n", encoding="utf-8"
        )
    return archive


def main() -> int:
    """Build a local plugin package without installing or enabling it."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output-dir", type=Path, default=Path("dist"))
    arguments = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    try:
        archive = build_bundle(arguments.source_root, arguments.output_dir)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        logger.error("Plugin bundle failed: %s", exc)
        return 1
    logger.info("Plugin ZIP: %s", archive)
    logger.info("Manual-import archive: %s", archive.with_suffix(".plugin"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
