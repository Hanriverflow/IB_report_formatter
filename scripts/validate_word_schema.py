"""Report OOXML-Validator findings without changing generated engine output.

Changelog (Word schema validation):
    - Verify pinned archives, validate the shared corpus, and enforce evidenced
      exceptions. The CI job that runs this script is a blocking gate.
"""

import argparse
import hashlib
import json
import logging
import shutil
import stat
import subprocess
import sys
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from zipfile import ZipFile

# Support both direct invocation and imports from tests in the source checkout.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.sample_corpus import render_corpus  # noqa: E402

LOGGER = logging.getLogger(__name__)
SCRIPT_DIR = Path(__file__).resolve().parent
ASSET_MANIFEST = SCRIPT_DIR / "word_schema_validator_assets.json"
EXCEPTIONS_FILE = SCRIPT_DIR / "word_schema_exceptions.json"


@dataclass(frozen=True)
class ErrorIdentity:
    """Exact validator identity; no descriptions or wildcard matching."""

    id: str
    part_uri: str
    xpath: str


@dataclass(frozen=True)
class ValidationError:
    """An observed error with its complete reportable location."""

    identity: ErrorIdentity
    description: str
    error_type: str


@dataclass(frozen=True)
class SchemaException:
    """An exception backed by a Word-saved reproduction."""

    identity: ErrorIdentity
    reason: str
    evidence: dict[str, Any]


def _object(value: object, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be a JSON object")
    return value


def _text(data: dict[str, Any], key: str) -> str:
    value = data.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{key} must be a nonempty string")
    return value


def parse_validator_output(output: str) -> dict[Path, tuple[ValidationError, ...]]:
    """Decode BOM-tolerant output and either ValidationErrors representation.

    Args:
        output: Validator stdout, including files without errors via --all.

    Returns:
        Absolute file paths mapped to validated, typed error records.
    """
    records = json.loads(output.lstrip("\ufeff"))
    if not isinstance(records, list):
        raise ValueError("Validator output must be a JSON array")
    documents: dict[Path, tuple[ValidationError, ...]] = {}
    for record in records:
        document = _object(record, "Validator file record")
        path = Path(_text(document, "FilePath")).resolve()
        if path in documents:
            raise ValueError(f"Duplicate validator result: {path}")
        raw_errors = document.get("ValidationErrors")
        if isinstance(raw_errors, str):
            raw_errors = json.loads(raw_errors.lstrip("\ufeff"))
        if not isinstance(raw_errors, list):
            raise ValueError("ValidationErrors must be an array or a JSON-encoded array")
        errors: list[ValidationError] = []
        for raw in raw_errors:
            error = _object(raw, "Validation error")
            location = _object(error.get("Path"), "Error Path")
            errors.append(
                ValidationError(
                    ErrorIdentity(
                        _text(error, "Id"),
                        _text(location, "PartUri"),
                        _text(location, "XPath"),
                    ),
                    _text(error, "Description"),
                    _text(error, "ErrorType"),
                )
            )
        documents[path] = tuple(errors)
    return documents


def load_exceptions(path: Path) -> tuple[SchemaException, ...]:
    """Require exact identities, reasons, and matching Word reproduction output.

    Args:
        path: JSON list following word_schema_validation.md.

    Returns:
        Unique, structurally evidenced exception entries.
    """
    entries = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(entries, list):
        raise ValueError("Exceptions must be a JSON array")
    exceptions: list[SchemaException] = []
    identities: set[ErrorIdentity] = set()
    for entry in entries:
        data = _object(entry, "Exception")
        identity = ErrorIdentity(*(_text(data, key) for key in ("id", "part_uri", "xpath")))
        if identity in identities:
            raise ValueError(f"Duplicate exception identity: {identity}")
        evidence = _object(data.get("evidence"), "Exception evidence")
        if evidence.get("kind") != "word_saved_reproduction":
            raise ValueError("Only word_saved_reproduction evidence is accepted")
        _text(evidence, "procedure")
        _text(evidence, "word_version")
        observed = parse_validator_output(json.dumps(evidence.get("validator_output")))
        if not any(error.identity == identity for errors in observed.values() for error in errors):
            raise ValueError(f"Word reproduction does not contain exception identity: {identity}")
        exceptions.append(SchemaException(identity, _text(data, "reason"), evidence))
        identities.add(identity)
    return tuple(exceptions)


def evaluate_output(
    output: str, expected_files: Sequence[Path], exceptions: Sequence[SchemaException]
) -> dict[str, Any]:
    """Apply exact exceptions and fail closed on incomplete document coverage.

    Args:
        output: Validator stdout.
        expected_files: Every DOCX produced by the current corpus run.
        exceptions: Validated exception entries.

    Returns:
        JSON-ready report; ok requires no uncovered errors or unused exceptions.
    """
    documents = parse_validator_output(output)
    expected = {path.resolve() for path in expected_files}
    if not expected or len(expected) != len(expected_files):
        raise ValueError("Expected corpus must contain unique document paths and not be empty")
    if set(documents) != expected:
        missing = sorted(str(path) for path in expected - documents.keys())
        unexpected = sorted(str(path) for path in documents.keys() - expected)
        raise ValueError(
            f"Validator document coverage differs: missing={missing}, unexpected={unexpected}"
        )
    if len({entry.identity for entry in exceptions}) != len(exceptions):
        raise ValueError("Duplicate exception identities")
    lookup = {
        error: index
        for index, entry in enumerate(exceptions)
        for errors in parse_validator_output(
            json.dumps(entry.evidence["validator_output"])
        ).values()
        for error in errors
        if error.identity == entry.identity
    }
    matched: set[int] = set()
    reported_documents: list[dict[str, Any]] = []
    uncovered = 0
    total = 0
    for path in expected_files:
        reported_errors: list[dict[str, Any]] = []
        for error in documents[path.resolve()]:
            index = lookup.get(error)
            if index is None:
                uncovered += 1
            else:
                matched.add(index)
            total += 1
            reported_errors.append(
                {
                    **asdict(error.identity),
                    "description": error.description,
                    "error_type": error.error_type,
                    "covered_by_exception": index is not None,
                    "exception_index": index,
                }
            )
        reported_documents.append({"file_path": str(path.resolve()), "errors": reported_errors})
    unused = [asdict(entry.identity) for i, entry in enumerate(exceptions) if i not in matched]
    return {
        "ok": uncovered == 0 and not unused,
        "documents": reported_documents,
        "error_count": total,
        "uncovered_error_count": uncovered,
        "exceptions": [asdict(entry) for entry in exceptions],
        "unused_exceptions": unused,
    }


def verify_asset(path: Path, manifest_path: Path = ASSET_MANIFEST) -> None:
    """Reject an archive whose name or SHA-256 differs from the pinned manifest.

    Args:
        path: Downloaded release archive, before extraction or execution.
        manifest_path: Repository manifest of release URLs and hashes.
    """
    manifest = _object(json.loads(manifest_path.read_text(encoding="utf-8-sig")), "Manifest")
    assets = _object(manifest.get("assets"), "Manifest assets")
    asset = _object(assets.get(path.name), f"Pinned asset {path.name}")
    expected = _text(asset, "sha256")
    if len(expected) != 64 or any(char not in "0123456789abcdef" for char in expected):
        raise ValueError(f"Invalid pinned SHA-256 for {path.name}")
    with path.open("rb") as stream:
        actual = hashlib.file_digest(stream, "sha256").hexdigest()
    if actual != expected:
        raise ValueError(f"SHA-256 mismatch for {path.name}: expected {expected}, got {actual}")
    LOGGER.info("Verified SHA-256 for %s", path.name)


@contextmanager
def _verified_validator(asset: Path) -> Iterator[Path]:
    """Extract only a private copy verified during this invocation.

    Args:
        asset: Downloaded archive whose name and hash must match the pinned manifest.

    Yields:
        Executable from a fresh directory, removed after the validator finishes.
    """
    with TemporaryDirectory(prefix="word-schema-validator-") as directory:
        root = Path(directory)
        copied = root / asset.name
        shutil.copyfile(asset, copied)
        verify_asset(copied, ASSET_MANIFEST)
        extracted = root / "extracted"
        with ZipFile(copied) as archive:
            archive.extractall(extracted)
        executable = extracted / (
            "OOXMLValidatorCLI.exe" if sys.platform == "win32" else "OOXMLValidatorCLI"
        )
        executable.chmod(executable.stat().st_mode | stat.S_IXUSR)
        yield executable


def main(argv: Sequence[str] | None = None) -> int:
    """Run archive verification or render and validate the sample corpus.

    Args:
        argv: Command arguments; defaults to the process arguments.

    Returns:
        Zero only for a matching archive or a fully passing schema report.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    verify = commands.add_parser("verify-asset", help="Verify an archive before extraction")
    verify.add_argument("asset", type=Path)
    validate = commands.add_parser("validate", help="Render and validate all sample cases")
    validate.add_argument("asset", type=Path, help="Pinned release archive to verify and extract")
    validate.add_argument("--output-dir", type=Path, required=True)
    validate.add_argument("--report", type=Path, required=True)
    validate.add_argument("--exceptions", type=Path, default=EXCEPTIONS_FILE)
    validate.add_argument("--office-version", default="Microsoft365")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    if args.command == "verify-asset":
        try:
            verify_asset(args.asset)
        except (OSError, ValueError) as exc:
            LOGGER.error("%s", exc)
            return 1
        return 0

    report: dict[str, Any] = {
        "ok": False,
        "stage": "exceptions",
        "office_version": args.office_version,
        "asset": str(args.asset.resolve()),
        "documents": [],
    }
    try:
        exceptions = load_exceptions(args.exceptions)
        report["stage"] = "verify_asset"
        with _verified_validator(args.asset) as validator:
            report["validator"] = str(validator)
            report["stage"] = "render"
            outputs = render_corpus(args.output_dir)
            report["stage"] = "validator"
            result = subprocess.run(
                [
                    str(validator),
                    str(args.output_dir.resolve()),
                    args.office_version,
                    "--recursive",
                    "--all",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8-sig",
                timeout=300,
                check=False,
            )
        report["validator_exit_code"] = result.returncode
        report["validator_stderr"] = result.stderr
        if result.returncode:
            raise RuntimeError(f"Validator process failed with exit code {result.returncode}")
        report["stage"] = "evaluate"
        report.update(evaluate_output(result.stdout, outputs, exceptions))
        report["stage"] = "complete"
    except Exception as exc:
        report["failure"] = f"{type(exc).__name__}: {exc}"
        LOGGER.error("Schema validation failed at %s: %s", report["stage"], exc)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    LOGGER.info("Schema report: %s (ok=%s)", args.report, report["ok"])
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
