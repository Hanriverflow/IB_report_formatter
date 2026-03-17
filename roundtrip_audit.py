"""
Round-trip audit helper for Markdown <-> Word conversion.
"""

import argparse
import json
import logging
import sys
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Dict, List

from cli_utils import setup_logging as configure_logging
from md_parser import DocumentModel, ElementType, Heading, parse_markdown_file
from md_to_word import IBReportConverter
from word_parser import parse_word_file
from word_to_md import WordToMarkdownConverter

logger = logging.getLogger("roundtrip_audit")


@dataclass
class ModelSummary:
    """Compact semantic summary used for round-trip comparisons."""

    title: str
    subtitle: str
    company: str
    element_counts: Dict[str, int]
    headings: List[str]
    footnote_count: int
    image_count: int
    table_count: int


@dataclass
class AuditDiff:
    """Count deltas between source and round-tripped content."""

    element_counts: Dict[str, int]
    footnote_count: int
    image_count: int
    table_count: int


@dataclass
class AuditReport:
    """Round-trip audit result."""

    input: str
    input_type: str
    source: ModelSummary
    roundtrip: ModelSummary
    diff: AuditDiff


def summarize_model(model: DocumentModel) -> ModelSummary:
    """Summarize a parsed document model for round-trip comparison."""
    element_counts: Dict[str, int] = {}
    for element in model.elements:
        key = element.element_type.name
        element_counts[key] = element_counts.get(key, 0) + 1

    headings = [
        element.content.text
        for element in model.elements
        if element.element_type
        in (
            ElementType.HEADING_1,
            ElementType.HEADING_2,
            ElementType.HEADING_3,
            ElementType.HEADING_4,
            ElementType.NUMBERED_HEADING,
        )
        and isinstance(element.content, Heading)
    ]

    return ModelSummary(
        title=model.metadata.title,
        subtitle=model.metadata.subtitle,
        company=model.metadata.company,
        element_counts=element_counts,
        headings=headings,
        footnote_count=len(model.footnotes),
        image_count=element_counts.get(ElementType.IMAGE.name, 0),
        table_count=element_counts.get(ElementType.TABLE.name, 0),
    )


def build_audit_report(input_path: Path) -> AuditReport:
    """Run a round-trip audit and return a summary report."""
    with tempfile.TemporaryDirectory() as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)

        if input_path.suffix.lower() == ".md":
            source_model = parse_markdown_file(str(input_path))
            docx_path = tmp_dir / f"{input_path.stem}_audit.docx"
            markdown_path = tmp_dir / f"{input_path.stem}_audit.md"

            IBReportConverter(str(input_path), output_path=str(docx_path)).convert()
            WordToMarkdownConverter(
                docx_file_path=str(docx_path),
                output_path=str(markdown_path),
            ).convert()
            roundtrip_model = parse_markdown_file(str(markdown_path))
            source_kind = "markdown"
        elif input_path.suffix.lower() == ".docx":
            source_model = parse_word_file(str(input_path), extract_images=False)
            markdown_path = tmp_dir / f"{input_path.stem}_audit.md"

            WordToMarkdownConverter(
                docx_file_path=str(input_path),
                output_path=str(markdown_path),
            ).convert()
            roundtrip_model = parse_markdown_file(str(markdown_path))
            source_kind = "word"
        else:
            raise ValueError(f"Unsupported input type: {input_path}")

        source_summary = summarize_model(source_model)
        roundtrip_summary = summarize_model(roundtrip_model)

        return AuditReport(
            input=str(input_path),
            input_type=source_kind,
            source=source_summary,
            roundtrip=roundtrip_summary,
            diff=AuditDiff(
                element_counts={
                    key: roundtrip_summary.element_counts.get(key, 0)
                    - source_summary.element_counts.get(key, 0)
                    for key in set(source_summary.element_counts)
                    | set(roundtrip_summary.element_counts)
                },
                footnote_count=roundtrip_summary.footnote_count - source_summary.footnote_count,
                image_count=roundtrip_summary.image_count - source_summary.image_count,
                table_count=roundtrip_summary.table_count - source_summary.table_count,
            ),
        )


def build_parser() -> argparse.ArgumentParser:
    """Build the round-trip audit CLI parser."""
    parser = argparse.ArgumentParser(
        description="Audit round-trip preservation for Markdown and Word documents",
    )
    parser.add_argument("input_file", help="Input .md or .docx file to audit")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the audit report as JSON",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    return parser


def main() -> None:
    """Run the round-trip audit CLI."""
    parser = build_parser()
    args = parser.parse_args()
    configure_logging(verbose=args.verbose)

    input_path = Path(args.input_file).resolve()
    if not input_path.exists():
        logger.error("File not found: %s", input_path)
        sys.exit(1)

    report = build_audit_report(input_path)

    if args.json:
        print(json.dumps(asdict(report), ensure_ascii=False, indent=2))
        return

    print(f"Input: {report.input}")
    print(f"Type: {report.input_type}")
    print(f"Title: {report.source.title} -> {report.roundtrip.title}")
    print(f"Elements: {report.source.element_counts} -> {report.roundtrip.element_counts}")
    print(f"Footnotes: {report.source.footnote_count} -> {report.roundtrip.footnote_count}")
    print(f"Images: {report.source.image_count} -> {report.roundtrip.image_count}")
    print(f"Tables: {report.source.table_count} -> {report.roundtrip.table_count}")
    print(f"Diff: {report.diff}")


if __name__ == "__main__":
    main()
