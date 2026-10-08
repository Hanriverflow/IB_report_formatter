"""Build a source-only ZIP for opening this project in Codex.

No Python runtime, executable, private inputs or generated documents are shipped.
Public build sources for both distributions are included; Codex remains the
default starting workflow in this source bundle.
Every source path is explicitly approved here or is an engine module listed in
the wheel configuration. New files are never discovered by recursive globbing.
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
    "pyproject.toml", "uv.lock", ".python-version", ".gitignore", "LICENSE",
    "AGENTS.md", "README.md", "README.ko.md", "CHANGELOG.md",
    ".agents/skills/ib-document/SKILL.md",
    ".agents/skills/ib-document/references/term-sheet.md",
    "docs/codex-start.html",
    "docs/distribution-build.md", "docs/distribution/시작하기.html",
    "docs/converted-input-cleanup-plan-20260929.md",
    "docs/format-preprocessing-guide.md",
    "docs/implementation-plan-20260914.md",
    "docs/improvement-plan-20260915.md",
    "docs/improvement-roadmap-20260929.md",
    "docs/pr-consolidation-20260926.md",
    "docs/term-sheet-design-20260929.md",
    "docs/term-sheet-plan-20260929.md",
    "docs/verification-20260914.md",
    "docs/verification-20260915.md",
    "docs/verification-20260929.md",
    "docs/verification-feature-port-20260929.md",
    "docs/verification-term-sheet-20260929.md",
    "docs/verification-title-block-20260929.md",
    "scripts/agent_render.py", "scripts/build_codex_bundle.py",
    "scripts/build_portable.py", "tools/portable_launcher.py",
    "scripts/word_visual_qa.ps1",
    "samples/profiles/business-report.md", "samples/profiles/company-theme.yaml",
    "samples/profiles/ib-memo.md", "samples/profiles/ib-report.md",
    "samples/profiles/meeting-minutes.md", "samples/profiles/office-letter.md",
    "samples/profiles/plain.md", "samples/profiles/term-sheet-house.yaml",
    "samples/profiles/term-sheet.md",
    "samples/harness/brief.md", "samples/harness/request.txt",
    "samples/harness/house.yaml", "samples/harness/expected-terms.json",
    "samples/distribution/minimal-term-sheet.md", "samples/distribution/minimal-house.yaml",
)


def source_files(source_root: Path) -> tuple[str, list[str]]:
    """Return the version and validated explicit source allowlist.

    Args:
        source_root: Project directory containing pyproject.toml.

    Returns:
        Project version and sorted public file paths.
    """
    config = tomllib.loads((source_root / "pyproject.toml").read_text(encoding="utf-8"))
    version = config["project"]["version"]
    if not re.fullmatch(r"[0-9A-Za-z][0-9A-Za-z._-]*", version):
        raise ValueError("Project version is not safe for a bundle filename")
    modules = config["tool"]["hatch"]["build"]["targets"]["wheel"]["only-include"]
    if not modules or any(not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*\.py", name)
                          for name in modules):
        raise ValueError("Wheel allowlist must contain only top-level Python modules")
    paths = sorted(set(PUBLIC_FILES).union(modules))
    root = source_root.resolve()
    for relative in paths:
        path = root / relative
        if not path.is_file():
            raise FileNotFoundError(f"Required public source file is missing: {relative}")
        if not path.resolve().is_relative_to(root):
            raise ValueError(f"Public source path leaves the project: {relative}")
        # Reject links even when their current target happens to be inside root.
        if any(part.is_symlink() for part in (path, *path.parents) if part != root):
            raise ValueError(f"Public source path contains a symbolic link: {relative}")
    return version, paths


def git_provenance(source_root: Path, paths: list[str]) -> dict:
    """Read base commit and worktree state without exposing private filenames."""
    try:
        commit = subprocess.run(
            ["git", "rev-parse", "HEAD"], cwd=source_root, check=True,
            capture_output=True, text=True,
        ).stdout.strip()
        status = subprocess.run(
            ["git", "status", "--porcelain", "--untracked-files=normal"],
            cwd=source_root, check=True, capture_output=True, text=True,
        ).stdout
        modified = subprocess.run(
            ["git", "-c", "core.quotepath=false", "diff", "--name-only", "HEAD", "--", *paths],
            cwd=source_root, check=True, capture_output=True, text=True, encoding="utf-8",
        ).stdout.splitlines()
        new = subprocess.run(
            ["git", "-c", "core.quotepath=false", "ls-files", "--others", "--exclude-standard",
             "--", *paths], cwd=source_root, check=True, capture_output=True, text=True,
            encoding="utf-8",
        ).stdout.splitlines()
    except (OSError, subprocess.CalledProcessError):
        return {"base_commit": None, "dirty": None, "modified_files": None, "new_files": None}
    return {"base_commit": commit, "dirty": bool(status.strip()),
            "modified_files": modified, "new_files": new}


def build_bundle(source_root: Path, output_dir: Path) -> Path:
    """Build a unique ZIP with a manifest hashing every bundled source file.

    Args:
        source_root: Source checkout to package.
        output_dir: Parent directory for a newly created release directory.

    Returns:
        Absolute path to the new ZIP. Existing releases are never overwritten.
    """
    source_root = source_root.resolve()
    version, paths = source_files(source_root)
    # Read once: hashes always describe the exact bytes written to the archive.
    payloads = {relative: (source_root / relative).read_bytes() for relative in paths}
    payloads["시작하기.html"] = payloads["docs/codex-start.html"]
    manifest = {
        "format_version": 1,
        "distribution": "codex-source",
        "version": version,
        "created_at_utc": datetime.now(UTC).isoformat(),
        "source": git_provenance(source_root, paths),
        "provenance_note": (
            "base_commit identifies the starting revision, not uncommitted changes. "
            "File hashes identify the actual bundled source. "
            "The manifest itself is not included in its file hashes."
        ),
        "files": [
            {"path": name, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
            for name, data in sorted(payloads.items())
        ],
    }
    root_name = f"IB_report_formatter_Codex_{version}"
    release = output_dir.resolve() / (
        "codex-source-" + datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
        + "-" + uuid.uuid4().hex[:8]
    )
    release.mkdir(parents=True, exist_ok=False)
    archive = release / f"{root_name}.zip"
    with zipfile.ZipFile(archive, "x", compression=zipfile.ZIP_DEFLATED) as bundle:
        for name, data in sorted(payloads.items()):
            bundle.writestr(f"{root_name}/{name}", data)
        bundle.writestr(
            f"{root_name}/bundle-manifest.json",
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        )
    checksum = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(
        f"{checksum}  {archive.name}\n", encoding="utf-8"
    )
    return archive


def main() -> int:
    """Build the source distribution from this checkout."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output-dir", type=Path, default=Path("dist"))
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    try:
        archive = build_bundle(args.source_root, args.output_dir)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        logger.error("Source bundle failed: %s", exc)
        return 1
    logger.info("Source ZIP: %s", archive)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
