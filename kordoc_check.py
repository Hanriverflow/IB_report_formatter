"""Cross-check the local Markdown-to-DOCX pipeline with kordoc.

This is a developer smoke tool, not a CI gate. A drift result is a signal for a
human to investigate and does not by itself identify a bug in this script.
"""

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from md_parser import ElementType, parse_markdown_file
from md_to_word import IBReportConverter

EXIT_OK = 0
EXIT_DRIFT = 2
EXIT_KORDOC_MISSING = 3
EXIT_KORDOC_FAILURE = 4
EXIT_PIPELINE_FAILURE = 5

HEADING_TYPES = {
    ElementType.HEADING_1,
    ElementType.HEADING_2,
    ElementType.HEADING_3,
    ElementType.HEADING_4,
    ElementType.NUMBERED_HEADING,
}

def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Cross-check md-to-docx output with the kordoc parser."
    )
    parser.add_argument("input_md", help="Path to the source Markdown file")
    return parser


def find_kordoc() -> Optional[str]:
    return shutil.which("kordoc") or shutil.which("kordoc.cmd")


def kordoc_uses_parse_subcommand(kordoc_bin: str) -> bool:
    result = subprocess.run(
        [kordoc_bin, "--help"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        shell=False,
    )
    help_text = result.stdout + result.stderr
    return "parse" in help_text.lower()


def extract_with_kordoc(kordoc_bin: str, docx_path: Path) -> str:
    """Extract Markdown text, raising RuntimeError on non-zero exit."""
    has_parse = kordoc_uses_parse_subcommand(kordoc_bin)
    command = [kordoc_bin, str(docx_path)]
    if has_parse:
        command = [kordoc_bin, "parse", str(docx_path)]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        shell=False,
    )
    if result.returncode != 0:
        stderr = result.stderr.strip() or "(kordoc produced no stderr output)"
        raise RuntimeError(stderr)
    return result.stdout


def count_source_elements(input_md_path: Path) -> Tuple[int, int]:
    model = parse_markdown_file(str(input_md_path))
    heading_count = sum(
        element.element_type in HEADING_TYPES for element in model.elements
    )
    table_count = sum(
        element.element_type is ElementType.TABLE for element in model.elements
    )
    return heading_count, table_count


def count_kordoc_elements(kordoc_text: str) -> Tuple[int, int]:
    heading_count = len(re.findall(r"^#{1,4} ", kordoc_text, re.MULTILINE))
    table_count = 0
    inside_table = False

    for line in kordoc_text.splitlines():
        starts_table_line = line.startswith("|")
        if starts_table_line and not inside_table:
            table_count += 1
        inside_table = starts_table_line

    return heading_count, table_count


def format_report(rows: Sequence[Sequence[str]]) -> str:
    headers = ("metric", "source model", "kordoc-extracted", "verdict")
    all_rows: List[Sequence[str]] = [headers]
    all_rows.extend(rows)
    widths = [
        max(len(row[index]) for row in all_rows) for index in range(len(headers))
    ]

    def format_row(row: Sequence[str]) -> str:
        return " | ".join(
            value.ljust(widths[index]) for index, value in enumerate(row)
        )

    separator = "-|-".join("-" * width for width in widths)
    return "\n".join([format_row(headers), separator] + [format_row(row) for row in rows])


def main() -> int:
    """Run the kordoc smoke comparison and return its process exit code."""
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_argument_parser().parse_args()

    kordoc_bin = find_kordoc()
    if kordoc_bin is None:
        print(
            "kordoc was not found on PATH. Install it with: "
            "npm install -g @clazic/kordoc"
        )
        return EXIT_KORDOC_MISSING

    input_md_path = Path(args.input_md)
    if not input_md_path.is_file():
        print(f"Input Markdown file does not exist: {input_md_path}")
        return EXIT_PIPELINE_FAILURE

    try:
        source_counts = count_source_elements(input_md_path)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_docx_path = Path(temp_dir) / (input_md_path.stem + ".docx")
            converter = IBReportConverter(
                md_file_path=str(input_md_path),
                output_path=str(temp_docx_path),
            )
            saved_path_str = converter.convert()
            try:
                kordoc_text = extract_with_kordoc(
                    kordoc_bin, Path(saved_path_str)
                )
            except (OSError, UnicodeError, ValueError, RuntimeError) as exc:
                print(f"kordoc failed: {exc}")
                return EXIT_KORDOC_FAILURE
    except (OSError, UnicodeError, ValueError) as exc:
        print(f"Smoke-check pipeline failed: {exc}")
        return EXIT_PIPELINE_FAILURE

    kordoc_counts = count_kordoc_elements(kordoc_text)
    heading_ok = source_counts[0] == kordoc_counts[0]
    table_ok = source_counts[1] == kordoc_counts[1]
    rows = [
        (
            "Headings",
            str(source_counts[0]),
            str(kordoc_counts[0]),
            "OK" if heading_ok else "DRIFT",
        ),
        (
            "Tables",
            str(source_counts[1]),
            str(kordoc_counts[1]),
            "OK" if table_ok else "DRIFT",
        ),
    ]
    print(format_report(rows))
    return EXIT_OK if heading_ok and table_ok else EXIT_DRIFT


if __name__ == "__main__":
    sys.exit(main())
