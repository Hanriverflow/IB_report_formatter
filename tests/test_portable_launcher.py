"""Portable entry point checks against the actual parsing/rendering/save path."""

import json
import zipfile
from pathlib import Path

import pytest
from docx import Document

from input_safety import InputLimitError, RejectionCode, StrictValidationError
from scripts import build_portable
from scripts.build_portable import package_inputs
from tools import portable_launcher


@pytest.mark.parametrize("wrapped", [False, True])
@pytest.mark.parametrize(
    "code, cause, remedy",
    [
        (RejectionCode.ASSET_OUTSIDE_ROOT, "폴더 밖", "상대 경로"),
        (RejectionCode.LIMIT_INPUT_BYTES, "Markdown 파일의 바이트 수", "여러 Markdown 파일"),
        (RejectionCode.LIMIT_TABLE_CELLS, "표 하나의 칸 수", "여러 표 또는 문서"),
        (RejectionCode.LIMIT_LIST_DEPTH, "목록의 중첩 단계 수", "들여쓰기 단계를 줄이거나"),
        (RejectionCode.LIMIT_IMAGE_BYTES, "이미지 또는 로고 하나의 바이트 수", "이미지를 압축"),
        (RejectionCode.LIMIT_IMAGE_PIXELS, "픽셀 수(가로 × 세로)", "해상도를 낮추거나"),
    ],
)
def test_korean_rejection_causes_and_remedies(code, cause, remedy, wrapped):
    rejection = (
        StrictValidationError("Document validation failed: ", [f"[{code}] 예시 경로"])
        if code == RejectionCode.ASSET_OUTSIDE_ROOT
        else InputLimitError(code, 1000, 2000)
    )
    error = rejection
    if wrapped:
        error = RuntimeError(f"Rendering failed: {rejection}")
        error.__cause__ = rejection
    message = portable_launcher.conversion_error_message(error)
    assert f"[{code}] " in message
    assert cause in message and remedy in message
    assert message.endswith(str(error))
    if code != RejectionCode.ASSET_OUTSIDE_ROOT:
        assert "한도: 1,000, 측정값: 2,000." in message
    else:
        assert "같은 폴더나 그 하위 폴더로 복사" in message


def test_portable_outside_image_rejection_has_korean_guidance(tmp_path):
    job = tmp_path / "job"
    job.mkdir()
    source = job / "예시.md"
    source.write_text("---\nprofile: plain\n---\n![예시](../outside.png)\n", encoding="utf-8")
    with pytest.raises(RuntimeError) as caught:
        portable_launcher.convert_document(str(source), str(tmp_path / "out"))
    message = portable_launcher.conversion_error_message(caught.value)
    assert "[asset_outside_root]" in message
    assert "폴더 밖" in message and "상대 경로" in message
    assert not list(tmp_path.rglob("*.docx"))


def test_uncoded_portable_errors_keep_original_message():
    error = ValueError("기존 입력 오류")
    assert portable_launcher.conversion_error_message(error) == str(error)


@pytest.mark.parametrize(
    "option",
    [
        "--asset-root",
        "--max-input-bytes",
        "--max-table-cells",
        "--max-list-depth",
        "--max-image-bytes",
        "--max-image-pixels",
    ],
)
def test_portable_rejects_safety_override_arguments(option, monkeypatch):
    monkeypatch.setattr(portable_launcher.sys, "argv", ["portable_launcher", option, "1"])
    with pytest.raises(SystemExit) as caught:
        portable_launcher.main()
    assert caught.value.code == 2


def test_conversion_preserves_yaml_and_korean_spaced_paths(tmp_path):
    source_dir = tmp_path / "한글 입력 폴더"
    source_dir.mkdir()
    (source_dir / "house.yaml").write_text(
        (Path(__file__).resolve().parents[1] / "samples/profiles/term-sheet-house.yaml").read_text(
            encoding="utf-8"
        ),
        encoding="utf-8",
    )
    source = source_dir / "첫 텀시트.md"
    source.write_text(
        '---\nprofile: term-sheet\nhouse: house.yaml\ntitle: "{{borrower}} 조건표"\n'
        "terms:\n  borrower: 가상기업\n---\n# 주요 조건\n\n"
        "| 항목 | 내용 |\n|---|---|\n| 차주 | {{borrower}} |\n",
        encoding="utf-8",
    )
    result = portable_launcher.convert_document(str(source), str(tmp_path / "한글 출력 폴더"))
    assert result.is_file()
    document = Document(result)
    xml = document._element.xml
    assert "가상기업" in xml
    assert "{{borrower}}" not in xml
    assert "주요 조건" in xml


def test_conversion_does_not_overwrite_existing_result(tmp_path):
    source = tmp_path / "sample.md"
    source.write_text("---\nprofile: plain\n---\n새 결과\n", encoding="utf-8")
    existing = tmp_path / "sample.docx"
    existing.write_bytes(b"existing document")
    saved = portable_launcher.convert_document(str(source), str(tmp_path))
    assert saved != existing
    assert existing.read_bytes() == b"existing document"
    assert Document(saved).paragraphs[0].text == "새 결과"


def test_strict_failure_does_not_save_partial_docx(tmp_path):
    source = tmp_path / "누락 이미지.md"
    source.write_text("---\nprofile: plain\n---\n![누락](absent.png)\n", encoding="utf-8")
    with pytest.raises(RuntimeError, match="Rendering failed"):
        portable_launcher.convert_document(str(source), str(tmp_path / "출력"))
    assert not list(tmp_path.rglob("*.docx"))


@pytest.mark.parametrize(
    "source, output, message",
    [
        ("", ".", "선택"),
        ("a.md", "", "폴더"),
        ("a.docx", ".", "Markdown"),
    ],
)
def test_actionable_input_errors(source, output, message):
    with pytest.raises(ValueError, match=message):
        portable_launcher.convert_document(source, output)


def test_engine_saved_path_is_returned_without_reconstruction(tmp_path, monkeypatch):
    source = tmp_path / "source.md"
    source.write_text("---\nprofile: plain\n---\n본문", encoding="utf-8")
    actual = tmp_path / "lock fallback.docx"

    def save_fallback(doc, requested):
        doc.save(actual)
        return actual

    monkeypatch.setattr("md_to_word.safe_save", save_fallback)
    assert portable_launcher.convert_document(str(source), str(tmp_path)) == actual


def test_frozen_root_follows_executable_not_cwd(tmp_path, monkeypatch):
    monkeypatch.setattr(portable_launcher.sys, "frozen", True, raising=False)
    monkeypatch.setattr(portable_launcher.sys, "executable", str(tmp_path / "문서변환기.exe"))
    assert portable_launcher.application_root() == tmp_path


def test_package_allowlist_only_includes_public_guide_license_and_samples():
    inputs = package_inputs()
    assert len(inputs) == 13
    assert {destination.parts[0] for _, destination in inputs} == {
        "시작하기.html",
        "라이선스",
        "예제",
    }
    assert all(
        ".." not in destination.parts and not destination.is_absolute() for _, destination in inputs
    )
    assert not any(source.suffix == ".docx" for source, _ in inputs)
    assert (Path("예제/빠른시작/minimal-house.yaml")) in {dest for _, dest in inputs}


def test_portable_archive_filename_and_matching_root(tmp_path, monkeypatch):
    source_root = tmp_path / "source"
    source_root.mkdir()
    (source_root / "pyproject.toml").write_text('[project]\nversion = "9.8.7"\n', encoding="utf-8")
    license_file = source_root / "LICENSE"
    license_file.write_text("test license", encoding="utf-8")
    app_dir = tmp_path / "frozen-app"
    app_dir.mkdir()
    executable = f"{build_portable.APP_NAME}.exe"
    (app_dir / executable).write_bytes(b"test frozen executable")
    monkeypatch.setattr(build_portable, "ROOT", source_root)
    monkeypatch.setattr(build_portable.sys, "platform", "win32")
    monkeypatch.setattr(
        build_portable,
        "package_inputs",
        lambda: [(license_file, Path("라이선스/LICENSE.txt"))],
    )
    monkeypatch.setattr(build_portable, "collect_notices", lambda destination: {})
    monkeypatch.setattr(build_portable.sys, "path", list(build_portable.sys.path))

    monkeypatch.setattr(build_portable, "collect_runtime_notices", lambda destination, source: {})
    archive = build_portable.build(tmp_path / "release", app_dir, tmp_path / "runtime-source")

    assert archive.name == "ib-report-formatter-windows-x64-9.8.7.zip"
    prefix = "ib-report-formatter-windows-x64-9.8.7/"
    with zipfile.ZipFile(archive) as bundle:
        assert all(name.startswith(prefix) for name in bundle.namelist())
        assert bundle.read(prefix + executable) == b"test frozen executable"
        assert bundle.read(prefix + "라이선스/LICENSE.txt") == b"test license"
        assert prefix + "입력/" in bundle.namelist()
        assert prefix + "출력/" in bundle.namelist()
        assert json.loads(bundle.read(prefix + "manifest.json"))["version"] == "9.8.7"


def _runtime_notice_fixture(tmp_path, monkeypatch):
    monkeypatch.setattr(build_portable.platform, "python_version", lambda: "3.12.12")
    source = tmp_path / "runtime-source"
    base = tmp_path / "active-runtime"
    (source / "install" / "DLLs").mkdir(parents=True)
    (source / "licenses").mkdir()
    (base / "DLLs").mkdir(parents=True)
    for name in ("python.exe", "python312.dll", "DLLs/libffi-8.dll"):
        (source / "install" / name).write_bytes(b"verified-runtime")
        (base / name).write_bytes(b"verified-runtime")
    (source / "PYTHON.json").write_text(
        json.dumps(
            {
                "python_version": build_portable.platform.python_version(),
                "target_triple": "x86_64-pc-windows-msvc",
            }
        ),
        encoding="utf-8",
    )
    for name in (
        "cpython",
        "expat",
        "libffi",
        "liblzma",
        "openssl-3",
        "zlib",
        "bzip2",
        "mpdecimal",
    ):
        (source / "licenses" / f"LICENSE.{name}.txt").write_text(
            "upstream notice", encoding="utf-8"
        )
    monkeypatch.setattr(build_portable.sys, "base_prefix", str(base))
    return source, base


def test_runtime_notices_match_and_preserve_upstream_metadata(tmp_path, monkeypatch):
    source, base = _runtime_notice_fixture(tmp_path, monkeypatch)
    destination = tmp_path / "notices"
    build_portable.collect_runtime_notices(destination, source)
    saved = destination / "Python-standalone"
    assert (saved / "LICENSE.libffi.txt").read_text() == "upstream notice"
    assert (saved / "PYTHON.json").read_bytes() == (source / "PYTHON.json").read_bytes()
    proof = json.loads((saved / "provenance.json").read_text())
    assert len(proof["runtime_binaries"]) == 3


def test_runtime_notices_reject_different_binary(tmp_path, monkeypatch):
    source, base = _runtime_notice_fixture(tmp_path, monkeypatch)
    (base / "DLLs/libffi-8.dll").write_bytes(b"different-runtime")
    with pytest.raises(RuntimeError, match="does not match"):
        build_portable.collect_runtime_notices(tmp_path / "notices", source)
    assert not (tmp_path / "notices/Python-standalone").exists()


def test_runtime_notices_reject_missing_license(tmp_path, monkeypatch):
    source, base = _runtime_notice_fixture(tmp_path, monkeypatch)
    (source / "licenses/LICENSE.expat.txt").unlink()
    with pytest.raises(RuntimeError, match="Incomplete"):
        build_portable.collect_runtime_notices(tmp_path / "notices", source)


def test_runtime_notices_reject_wrong_platform(tmp_path, monkeypatch):
    source, base = _runtime_notice_fixture(tmp_path, monkeypatch)
    (source / "PYTHON.json").write_text(
        json.dumps({"python_version": "0.0", "target_triple": "wrong"})
    )
    with pytest.raises(RuntimeError, match="platform/version mismatch"):
        build_portable.collect_runtime_notices(tmp_path / "notices", source)
