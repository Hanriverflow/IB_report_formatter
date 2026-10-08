"""Strict, JSON-reporting entry point for a Codex document authoring workflow.

Changelog (Codex workflow):
    - Reuse the converter registry, preserve source-relative assets, and report
      exact term comparisons separately from structural and visual review.
"""

import argparse
import hashlib
import json
import sys
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Never, Optional, Sequence

# This source-checkout tool must also work when invoked from another directory.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def _sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for key, value in pairs:
        if key in values:
            raise ValueError(f"Duplicate expected-term key: {key}")
        values[key] = value
    return values


def render_document(
    source: str,
    output_dir: str,
    profile: Optional[str] = None,
    expected_terms: Optional[str] = None,
) -> dict[str, Any]:
    """Render one source in an exclusively reserved output directory.

    Args:
        source: Original Markdown path, retained for relative asset resolution.
        output_dir: A directory that must not already exist.
        profile: Explicit parser and renderer profile override, if requested.
        expected_terms: Optional JSON object of exact string term values.

    Returns:
        JSON-compatible evidence; ``ok`` covers strict rendering and structural
        issues only, not factual completeness, financial validity or page layout.
        Audit failures retain generated files for diagnosis; input, term and
        strict-render failures produce no DOCX.
    """
    result: dict[str, Any] = {
        "ok": False, "status": "failed", "stage": "setup", "diagnostics": [],
        "input": {"path": str(source), "sha256": None},
        "output": {"directory": str(output_dir), "path": None, "sha256": None},
        "profile": profile, "audit": None, "visual_review": "not_performed",
        "expected_terms": {
            "status": "not_requested", "checked": 0,
            "scope": "Exact comparison with parsed terms only; not source truth or prose completeness.",
        },
    }
    owned_directory: Optional[Path] = None
    try:
        source_path = Path(source).resolve()
        destination = Path(output_dir).resolve()
        result["input"]["path"] = str(source_path)
        result["output"]["directory"] = str(destination)
        expected_path = Path(expected_terms).resolve() if expected_terms is not None else None
        for input_path in (source_path, expected_path):
            if input_path is not None and input_path.is_relative_to(destination):
                raise ValueError("Output directory must not contain an input file")
        # mkdir is the reservation: no pre-check race and no reuse of even empty directories.
        destination.mkdir(parents=True, exist_ok=False)
        owned_directory = destination
        result["stage"] = "input"
        result["input"]["sha256"] = _sha256(source_path)
        with redirect_stdout(sys.stderr):
            from converters import get_default_registry

            registry = get_default_registry()
            result["stage"] = "parse"
            model = registry.convert(source_path, profile=profile)
            result["profile"] = model.parsed_profile
            if model.warnings:
                result["diagnostics"].extend(
                    {"code": "input_validation", "message": message}
                    for message in model.warnings
                )
                raise ValueError("Strict input validation failed")
            if expected_path is not None:
                result["stage"] = "expected_terms"
                comparison = result["expected_terms"]
                comparison.update({"path": str(expected_path), "status": "invalid"})
                comparison["sha256"] = _sha256(expected_path)
                expected = json.loads(
                    expected_path.read_text(encoding="utf-8-sig"),
                    object_pairs_hook=_reject_duplicate_keys,
                )
                if not isinstance(expected, dict) or any(
                    not isinstance(key, str) or not isinstance(value, str)
                    for key, value in expected.items()
                ):
                    raise ValueError("Expected terms must be a JSON object of string keys and values")
                actual = model.metadata.extra.get("terms", {})
                mismatches = [key for key, value in expected.items() if actual.get(key) != value]
                comparison.update({
                    "checked": len(expected),
                    "status": "mismatch" if mismatches else ("matched" if expected else "empty"),
                    "mismatched_keys": mismatches,
                })
                if mismatches:
                    raise ValueError("Expected terms absent or different: " + ", ".join(mismatches))
            result["stage"] = "render"
            saved = Path(registry.convert(
                model, output_format="docx", output_path=destination / "document.docx",
                strict=True, profile=profile,
            )).resolve()
            result["output"].update({"path": str(saved), "sha256": _sha256(saved)})
            result["stage"] = "audit"
            from docx import Document

            from docx_audit import audit_to_dict, inspect_document, inspect_terms

            document = Document(str(saved))
            audit = inspect_document(document)
            result["audit"] = audit_to_dict(audit, inspect_terms(document))
            if audit.issues:
                result["diagnostics"].extend(
                    {"code": "structural_issue", "message": message} for message in audit.issues
                )
                raise ValueError("Structural audit failed; generated DOCX retained for diagnosis")
        result.update({"ok": True, "status": "ready_for_review", "stage": "complete"})
    except Exception as exc:
        result["diagnostics"].append({"code": type(exc).__name__, "message": str(exc)})
    if owned_directory is not None:
        try:
            with (owned_directory / "result.json").open("x", encoding="utf-8") as stream:
                json.dump(result, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
        except Exception as exc:
            result.update({"ok": False, "status": "failed", "stage": "report"})
            result["diagnostics"].append({"code": type(exc).__name__, "message": str(exc)})
    return result


class _JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> Never:
        raise ValueError(message)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Print one JSON result and return zero only for a successful strict render.

    Args:
        argv: Optional CLI arguments for embedded invocation and tests.

    Returns:
        Zero on success, one on conversion failure, two on invalid CLI arguments.
    """
    parser = _JsonArgumentParser(description=__doc__)
    parser.add_argument("source")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--profile")
    parser.add_argument("--expected-terms")
    try:
        arguments = parser.parse_args(argv)
    except ValueError as exc:
        result = {
            "ok": False, "status": "failed", "stage": "arguments",
            "diagnostics": [{"code": "invalid_arguments", "message": str(exc)}],
            "audit": None, "visual_review": "not_performed",
        }
        print(json.dumps(result, ensure_ascii=True))
        return 2
    result = render_document(
        arguments.source, arguments.output_dir, arguments.profile, arguments.expected_terms,
    )
    # ASCII-escaped JSON is lossless, even when a Windows caller captures a legacy code page.
    print(json.dumps(result, ensure_ascii=True))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
