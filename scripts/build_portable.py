"""Build an allowlisted Windows portable bundle in a new output directory.

Run using an isolated Python 3.12 environment with the project's locked runtime
dependencies, charset-normalizer and PyInstaller installed. No source reports,
working documents, or repository-wide directory trees are copied.
"""

import argparse
import hashlib
import importlib.metadata
import json
import logging
import platform
import shutil
import subprocess
import sys
import tomllib
import urllib.request
from datetime import UTC, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP_NAME = "문서변환기"
PROFILE_FILES = (
    "business-report.md", "company-theme.yaml", "ib-memo.md", "ib-report.md",
    "meeting-minutes.md", "office-letter.md", "plain.md", "term-sheet-house.yaml",
    "term-sheet.md",
)
INTRO_FILES = ("minimal-term-sheet.md", "minimal-house.yaml")
RUNTIME_DISTRIBUTIONS = (
    "python-docx", "PyYAML", "matplotlib", "charset-normalizer", "lxml", "numpy",
    "pillow", "contourpy", "cycler", "fonttools", "kiwisolver", "packaging",
    "pyparsing", "python-dateutil", "six", "typing_extensions", "pyinstaller",
)


def package_inputs() -> list[tuple[Path, Path]]:
    """Return the explicit public input-file allowlist and package destinations."""
    files = [(ROOT / "docs/distribution/시작하기.html", Path("시작하기.html")),
             (ROOT / "LICENSE", Path("라이선스/LICENSE.txt"))]
    files.extend((ROOT / "samples/distribution" / name, Path("예제/빠른시작") / name)
                 for name in INTRO_FILES)
    files.extend((ROOT / "samples/profiles" / name, Path("예제/전체예제") / name)
                 for name in PROFILE_FILES)
    return files


def collect_notices(destination: Path) -> dict[str, str]:
    """Collect installed dependency licenses and exact Python/Tcl/Tk notices."""
    versions = {}
    for name in RUNTIME_DISTRIBUTIONS:
        distribution = importlib.metadata.distribution(name)
        versions[name] = distribution.version
        folder = destination / name
        folder.mkdir(parents=True)
        found = []
        for entry in distribution.files or []:
            if any(part.lower() in {"licenses", "license"} for part in entry.parts) or entry.name.lower().startswith(("license", "copying", "notice")):
                source = Path(distribution.locate_file(entry))
                if source.is_file():
                    # Preserve nested license paths and avoid path traversal.
                    relative = Path(*[part for part in entry.parts if part not in {"..", "."}])
                    target = folder / relative
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, target)
                    found.append(str(relative))
        metadata = distribution.read_text("METADATA") or ""
        (folder / "METADATA.txt").write_text(metadata, encoding="utf-8")
        if not found:
            raise RuntimeError(f"No license file found for installed dependency: {name}")
    python_license = Path(sys.base_prefix) / "LICENSE.txt"
    shutil.copy2(python_license, destination / "Python-LICENSE.txt")
    import tkinter
    interpreter = tkinter.Tcl()
    tcl_library = Path(interpreter.eval("info library"))
    patchlevel = interpreter.eval("info patchlevel")
    versions["Python"] = platform.python_version()
    versions["Tcl/Tk"] = patchlevel
    for component, directory in (("Tcl", tcl_library), ("Tk", tcl_library.parent / "tk8.6")):
        license_file = directory / "license.terms"
        target = destination / f"{component}-license.terms"
        if license_file.is_file():
            shutil.copy2(license_file, target)
        else:
            # Some Python distributions omit Tcl's notice. Retrieve the exact
            # upstream release notice, rather than silently leaving it out.
            tag = "core-" + patchlevel.replace(".", "-")
            url = f"https://raw.githubusercontent.com/tcltk/{component.lower()}/{tag}/license.terms"
            with urllib.request.urlopen(url, timeout=30) as response:
                content = response.read()
            if b"copyright" not in content.lower():
                raise RuntimeError(f"Unexpected upstream license content: {url}")
            target.write_bytes(content)
            (destination / f"{component}-source.txt").write_text(url + "\n", encoding="utf-8")
    return versions


def build(output: Path, app_dir: Path | None = None) -> Path:
    """Create a fresh staging folder, sample results, manifest and distributable ZIP."""
    if sys.platform != "win32":
        raise RuntimeError("Build the Windows package on Windows.")
    inputs = package_inputs()
    for source, _ in inputs:
        if not source.is_file():
            raise FileNotFoundError(source)
    output = output.resolve()
    if output.exists():
        raise FileExistsError(f"Use a new/empty output path that does not yet exist: {output}")
    output.mkdir(parents=True)
    if app_dir is None:
        subprocess.run([
            sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir",
            "--windowed", "--name", APP_NAME, "--paths", str(ROOT),
            "--distpath", str(output / "frozen"), "--workpath", str(output / "work"),
            "--specpath", str(output), "--collect-data", "matplotlib",
            "--hidden-import", "matplotlib.backends.backend_agg",
            "--hidden-import", "charset_normalizer",
            "--exclude-module", "pytest", "--exclude-module", "IPython",
            str(ROOT / "tools/portable_launcher.py"),
        ], check=True, cwd=ROOT)
        app_dir = output / "frozen" / APP_NAME
    if not (app_dir / f"{APP_NAME}.exe").is_file():
        raise FileNotFoundError(f"Missing frozen executable in {app_dir}")
    with (ROOT / "pyproject.toml").open("rb") as handle:
        version = tomllib.load(handle)["project"]["version"]
    label = f"IB문서변환기_{version}_Windows_x64"
    package = output / label
    shutil.copytree(app_dir, package)
    for source, relative in inputs:
        target = package / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    for name in ("입력", "출력"):
        (package / name).mkdir()
    versions = collect_notices(package / "라이선스")
    sys.path.insert(0, str(ROOT))
    from tools.portable_launcher import convert_document
    sample_records = []
    for source, relative in inputs:
        if source.suffix == ".md":
            packaged_source = package / relative
            saved = convert_document(str(packaged_source), str(packaged_source.parent))
            sample_records.append({"source": relative.as_posix(), "output": saved.relative_to(package).as_posix(), "strict": True, "visual_review": "pending"})
    (package / "버전정보.txt").write_text(
        f"IB 문서변환기 {version}\nWindows x64 · Python 설치 불필요\n"
        "ZIP 전체 압축을 해제한 뒤 문서변환기.exe를 실행하세요.\n"
        "시작하기.html에 텀시트 작성법과 예제가 있습니다.\n"
        "예제는 가상 자료입니다. 출력 내용·숫자·페이지 배치는 Word에서 확인하세요.\n"
        f"Build UTC: {datetime.now(UTC).isoformat()}\n\n"
        + "\n".join(f"{key}: {value}" for key, value in versions.items()) + "\n",
        encoding="utf-8",
    )
    manifest = {"version": version, "platform": platform.platform(), "dependencies": versions,
                "samples": sample_records, "files": {}}
    for item in sorted(package.rglob("*")):
        if item.is_file():
            manifest["files"][item.relative_to(package).as_posix()] = hashlib.sha256(item.read_bytes()).hexdigest()
    (package / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    archive = Path(shutil.make_archive(str(output / label), "zip", output, label))
    logging.info("Created %s", archive)
    return archive


def main() -> None:
    """Read explicit build paths from the command line."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--app-dir", type=Path, help="Reuse an already-built frozen app folder")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    build(args.output_dir, args.app_dir)


if __name__ == "__main__":
    main()
