"""Schema decisions use fixed validator data; no executable is downloaded or run."""

import copy
import hashlib
import json
import subprocess
from pathlib import Path
from zipfile import ZipFile

import pytest
import yaml

from scripts import sample_corpus
from scripts import validate_word_schema as schema

ERROR = {
    "Id": "Sch_Example",
    "Description": "Example unexpected child element",
    "ErrorType": "Schema",
    "Path": {
        "Namespaces": {"w": "urn:example"},
        "PartUri": "/word/settings.xml",
        "XPath": "/w:settings[1]",
    },
}


def _payload(path: Path, errors: list, *, encoded: bool = False) -> str:
    return json.dumps(
        [{"FilePath": str(path), "ValidationErrors": json.dumps(errors) if encoded else errors}]
    )


def _exception_entry() -> dict:
    return {
        "id": ERROR["Id"],
        "part_uri": ERROR["Path"]["PartUri"],
        "xpath": ERROR["Path"]["XPath"],
        "reason": "Synthetic example of a Word-saved reproduction",
        "evidence": {
            "kind": "word_saved_reproduction",
            "procedure": "예시: create this construction in Word, save as DOCX, validate with Microsoft365",
            "word_version": "예시 Word version and build (synthetic fixture)",
            "validator_output": json.loads(_payload(Path("example-word-saved.docx"), [ERROR])),
        },
    }


def _load_entries(tmp_path: Path, entries: list) -> tuple[schema.SchemaException, ...]:
    path = tmp_path / "exceptions.json"
    path.write_text(json.dumps(entries), encoding="utf-8")
    return schema.load_exceptions(path)


@pytest.mark.parametrize("encoded", [False, True])
@pytest.mark.parametrize("bom", ["", "\ufeff"])
def test_validator_output_shapes_and_bom(tmp_path, encoded, bom):
    path = tmp_path / "example.docx"
    result = schema.evaluate_output(bom + _payload(path, [ERROR], encoded=encoded), [path], [])
    assert result["ok"] is False
    assert result["error_count"] == result["uncovered_error_count"] == 1
    assert result["documents"] == [
        {
            "file_path": str(path.resolve()),
            "errors": [
                {
                    "id": ERROR["Id"],
                    "part_uri": ERROR["Path"]["PartUri"],
                    "xpath": ERROR["Path"]["XPath"],
                    "description": ERROR["Description"],
                    "error_type": "Schema",
                    "covered_by_exception": False,
                    "exception_index": None,
                }
            ],
        }
    ]


def test_exception_covers_same_identity_across_documents(tmp_path):
    exceptions = _load_entries(tmp_path, [_exception_entry()])
    paths = [tmp_path / "one.docx", tmp_path / "two.docx"]
    records = [json.loads(_payload(path, [ERROR]))[0] for path in paths]
    result = schema.evaluate_output(json.dumps(records), paths, exceptions)
    assert result["ok"] is True
    assert result["error_count"] == 2
    assert result["uncovered_error_count"] == 0
    assert result["unused_exceptions"] == []
    assert all(doc["errors"][0]["covered_by_exception"] for doc in result["documents"])
    assert result["exceptions"][0]["reason"] == exceptions[0].reason


@pytest.mark.parametrize("field", ["Id", "PartUri", "XPath"])
def test_exception_identity_requires_all_three_exact_fields(tmp_path, field):
    exceptions = _load_entries(tmp_path, [_exception_entry()])
    different = copy.deepcopy(ERROR)
    target = different if field == "Id" else different["Path"]
    target[field] += "-different"
    path = tmp_path / "example.docx"
    result = schema.evaluate_output(_payload(path, [different]), [path], exceptions)
    assert result["ok"] is False
    assert result["uncovered_error_count"] == 1
    assert len(result["unused_exceptions"]) == 1


def test_unused_exception_fails_even_without_errors(tmp_path):
    exceptions = _load_entries(tmp_path, [_exception_entry()])
    path = tmp_path / "example.docx"
    result = schema.evaluate_output(_payload(path, []), [path], exceptions)
    assert result["ok"] is False
    assert result["error_count"] == 0
    assert result["unused_exceptions"] == [
        {"id": ERROR["Id"], "part_uri": "/word/settings.xml", "xpath": "/w:settings[1]"}
    ]


def test_covered_error_does_not_hide_an_uncovered_error(tmp_path):
    exceptions = _load_entries(tmp_path, [_exception_entry()])
    different = copy.deepcopy(ERROR)
    different["Id"] = "Sch_AnotherExample"
    path = tmp_path / "example.docx"
    result = schema.evaluate_output(_payload(path, [ERROR, different]), [path], exceptions)
    assert result["ok"] is False
    assert result["error_count"] == 2
    assert result["uncovered_error_count"] == 1
    assert result["unused_exceptions"] == []


def test_clean_documents_and_empty_exceptions_pass(tmp_path):
    path = tmp_path / "example.docx"
    result = schema.evaluate_output(_payload(path, []), [path], [])
    assert result["ok"] is True
    assert result["documents"][0]["errors"] == []
    assert schema.load_exceptions(schema.EXCEPTIONS_FILE) == ()


@pytest.mark.parametrize(
    "change",
    [
        "missing_reason",
        "missing_procedure",
        "missing_word_version",
        "wrong_kind",
        "wrong_error",
        "missing_output",
    ],
)
def test_exception_requires_word_saved_reproduction(tmp_path, change):
    entry = _exception_entry()
    if change == "missing_reason":
        entry["reason"] = " "
    elif change == "wrong_kind":
        entry["evidence"]["kind"] = "standards_citation"
    elif change == "wrong_error":
        entry["evidence"]["validator_output"][0]["ValidationErrors"][0]["Id"] = "Other"
    else:
        field = {
            "missing_procedure": "procedure",
            "missing_word_version": "word_version",
            "missing_output": "validator_output",
        }[change]
        entry["evidence"].pop(field)
    with pytest.raises(ValueError):
        _load_entries(tmp_path, [entry])


def test_duplicate_exceptions_are_rejected(tmp_path):
    with pytest.raises(ValueError, match="Duplicate exception"):
        _load_entries(tmp_path, [_exception_entry(), _exception_entry()])


def test_word_evidence_accepts_string_encoded_errors(tmp_path):
    entry = _exception_entry()
    record = entry["evidence"]["validator_output"][0]
    record["ValidationErrors"] = json.dumps(record["ValidationErrors"])
    assert len(_load_entries(tmp_path, [entry])) == 1


@pytest.mark.parametrize(
    "raw",
    [
        "not json",
        "{}",
        '[{"FilePath": "example.docx"}]',
        '[{"FilePath": "example.docx", "ValidationErrors": "null"}]',
    ],
)
def test_malformed_validator_output_is_rejected(raw):
    with pytest.raises(ValueError):
        schema.parse_validator_output(raw)


@pytest.mark.parametrize("field", ["Id", "Description", "ErrorType", "Path"])
def test_incomplete_error_is_rejected(tmp_path, field):
    error = copy.deepcopy(ERROR)
    error.pop(field)
    with pytest.raises(ValueError):
        schema.parse_validator_output(_payload(tmp_path / "example.docx", [error]))


@pytest.mark.parametrize("kind", ["missing", "unexpected", "duplicate", "empty_corpus"])
def test_incomplete_or_ambiguous_corpus_results_fail(tmp_path, kind):
    path = tmp_path / "example.docx"
    records = json.loads(_payload(path, []))
    expected = [path]
    if kind == "missing":
        records = []
    elif kind == "unexpected":
        records += json.loads(_payload(tmp_path / "extra.docx", []))
    elif kind == "duplicate":
        records += records
    else:
        expected = []
    with pytest.raises(ValueError):
        schema.evaluate_output(json.dumps(records), expected, [])


def test_asset_hash_match_and_mismatch(tmp_path):
    asset = tmp_path / "example.zip"
    asset.write_bytes(b"example archive")
    manifest = tmp_path / "assets.json"
    manifest.write_text(
        json.dumps(
            {"assets": {asset.name: {"sha256": hashlib.sha256(asset.read_bytes()).hexdigest()}}}
        ),
        encoding="utf-8",
    )
    schema.verify_asset(asset, manifest)
    asset.write_bytes(b"changed archive")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        schema.verify_asset(asset, manifest)
    with pytest.raises(ValueError, match="Pinned asset"):
        schema.verify_asset(tmp_path / "unknown.zip", manifest)


def test_hash_command_fails_for_wrong_release_asset(tmp_path):
    asset = tmp_path / "linux-x64.zip"
    asset.write_bytes(b"tiny fixed data, not a validator")
    assert schema.main(["verify-asset", str(asset)]) == 1


@pytest.mark.parametrize(
    "outcome, exit_code",
    [("clean", 0), ("errors", 1), ("bad_json", 1), ("process_failure", 1), ("timeout", 1)],
)
def test_cli_writes_report_and_uses_content_not_process_success(
    tmp_path, monkeypatch, outcome, exit_code
):
    output_dir = tmp_path / "rendered"
    path = output_dir / "example.docx"
    monkeypatch.setattr(schema, "render_corpus", lambda destination: (path,))
    validator = tmp_path / "not-executed-validator"

    def fake_run(command, **kwargs):
        assert command == [
            str(validator.resolve()),
            str(output_dir.resolve()),
            "Microsoft365",
            "--recursive",
            "--all",
        ]
        assert kwargs["encoding"] == "utf-8-sig"
        assert kwargs["check"] is False
        if outcome == "timeout":
            raise subprocess.TimeoutExpired(command, kwargs["timeout"])
        output = (
            "invalid json"
            if outcome == "bad_json"
            else _payload(path, [ERROR] if outcome == "errors" else [])
        )
        return subprocess.CompletedProcess(
            command, 2 if outcome == "process_failure" else 0, output, "example stderr"
        )

    monkeypatch.setattr(schema.subprocess, "run", fake_run)
    report_path = tmp_path / "report.json"
    result = schema.main(
        ["validate", str(validator), "--output-dir", str(output_dir), "--report", str(report_path)]
    )
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert result == exit_code
    assert report["ok"] is (exit_code == 0)
    if outcome == "errors":
        assert report["validator_exit_code"] == 0
        assert report["uncovered_error_count"] == 1
    elif outcome != "clean":
        assert report["failure"]


def test_render_failure_still_writes_report(tmp_path, monkeypatch):
    def fail_render(destination):
        raise ValueError("Example strict rendering failure")

    monkeypatch.setattr(schema, "render_corpus", fail_render)
    report_path = tmp_path / "report.json"
    assert (
        schema.main(
            [
                "validate",
                "unused",
                "--output-dir",
                str(tmp_path / "new"),
                "--report",
                str(report_path),
            ]
        )
        == 1
    )
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["stage"] == "render"
    assert "strict rendering failure" in report["failure"]


@pytest.mark.parametrize("has_error", [False, True])
def test_cli_applies_exception_file_and_office_override(tmp_path, monkeypatch, has_error):
    _load_entries(tmp_path, [_exception_entry()])
    path = tmp_path / "example.docx"
    monkeypatch.setattr(schema, "render_corpus", lambda destination: (path,))

    def fake_run(command, **kwargs):
        assert command[2] == "Office2021"
        return subprocess.CompletedProcess(
            command, 0, _payload(path, [ERROR] if has_error else []), ""
        )

    monkeypatch.setattr(schema.subprocess, "run", fake_run)
    report_path = tmp_path / "report.json"
    result = schema.main(
        [
            "validate",
            "unused",
            "--output-dir",
            str(tmp_path / "new"),
            "--report",
            str(report_path),
            "--exceptions",
            str(tmp_path / "exceptions.json"),
            "--office-version",
            "Office2021",
        ]
    )
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert result == (0 if has_error else 1)
    assert report["ok"] is has_error
    assert report["office_version"] == "Office2021"
    assert bool(report["unused_exceptions"]) is not has_error


def test_corpus_recursively_discovers_sources_and_keeps_variants_distinct(tmp_path):
    for directory in sample_corpus.SAMPLE_DIRECTORIES:
        path = tmp_path / "samples" / directory / "nested" / "example.md"
        path.parent.mkdir(parents=True)
        path.write_text("# 예시 문서\n", encoding="utf-8")
    chart = tmp_path / sample_corpus.CHARTS_SAMPLE
    chart.write_text("# 예시 차트\n", encoding="utf-8")
    cases = sample_corpus.sample_cases(tmp_path)
    assert len(cases) == 5
    assert len({case.output_relative for case in cases}) == 5
    assert cases[-1] == sample_corpus.SampleCase(sample_corpus.CHARTS_SAMPLE, charts=True)
    assert all(case.options.strict is True and case.options.profile is None for case in cases)
    assert all(case.options.charts is None for case in cases[:-1])
    assert cases[-1].options.charts is True


def test_corpus_requires_samples_and_charts(tmp_path):
    with pytest.raises(ValueError, match="No Markdown"):
        sample_corpus.sample_cases(tmp_path)
    for directory in sample_corpus.SAMPLE_DIRECTORIES:
        path = tmp_path / "samples" / directory / "example.md"
        path.parent.mkdir(parents=True)
        path.write_text("# 예시 문서\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Missing charts"):
        sample_corpus.sample_cases(tmp_path)


def test_shared_corpus_really_renders_all_samples_strictly(tmp_path):
    cases = sample_corpus.sample_cases()
    discovered = {
        path.relative_to(sample_corpus.PROJECT_ROOT)
        for directory in sample_corpus.SAMPLE_DIRECTORIES
        for path in (sample_corpus.PROJECT_ROOT / "samples" / directory).rglob("*.md")
    }
    assert {case.source for case in cases[:-1]} == discovered
    assert len(cases) == len(discovered) + 1
    outputs = sample_corpus.render_corpus(tmp_path / "corpus")
    assert len(outputs) == len(cases)
    for case, output in zip(cases, outputs, strict=True):
        assert output == tmp_path / "corpus" / case.output_relative
        with ZipFile(output) as archive:
            assert "word/document.xml" in archive.namelist()
            if case.charts:
                assert any(name.startswith("word/media/") for name in archive.namelist())
    with pytest.raises(FileExistsError):
        sample_corpus.render_corpus(tmp_path / "corpus")


def test_ci_schema_job_is_independent_nonblocking_and_always_reports():
    root = sample_corpus.PROJECT_ROOT
    workflow = yaml.safe_load((root / ".github/workflows/ci.yml").read_text(encoding="utf-8"))
    job = workflow["jobs"]["word-schema"]
    assert job["runs-on"] == "ubuntu-latest"
    assert job["continue-on-error"] is True
    assert "needs" not in job and "needs" not in workflow["jobs"]["quality"]
    steps = job["steps"]
    download = next(step for step in steps if step.get("name") == "Download and verify validator")
    command = download["run"]
    assert "word_schema_validator_assets.json" in command
    assert command.index("curl --fail") < command.index("verify-asset") < command.index("unzip")
    validation = next(
        step for step in steps if step.get("name") == "Render corpus and validate Word schema"
    )
    assert steps.index(download) < steps.index(validation)
    assert "--report word-schema-report.json" in validation["run"]
    assert "continue-on-error" not in validation
    fallback = next(
        step for step in steps if step.get("name") == "Ensure a report exists after setup failures"
    )
    upload = next(
        step for step in steps if step.get("uses", "").startswith("actions/upload-artifact@")
    )
    assert fallback["if"] == upload["if"] == "always()"
    assert steps.index(validation) < steps.index(fallback) < steps.index(upload)
    assert upload["with"]["path"] == "word-schema-report.json"
    assert upload["with"]["if-no-files-found"] == "error"


def test_pinned_release_assets():
    manifest = json.loads(schema.ASSET_MANIFEST.read_text(encoding="utf-8"))
    assert manifest["version"] == "v2.1.6"
    expected = {
        "linux-x64.zip": "e1dcebba46b937a68831ec194f5a67e7312a300902678254c3b463dcff5cf499",
        "win-x64.zip": "0f22c0da723d589c20168a82f7e8057c177698c621635e6924c4fdaa19ef1acb",
    }
    for name, digest in expected.items():
        assert manifest["assets"][name] == {
            "url": f"https://github.com/mikeebowen/OOXML-Validator/releases/download/v2.1.6/{name}",
            "sha256": digest,
        }
