"""
Round-trip audit helper for Markdown <-> Word conversion.
"""

import argparse
import difflib
import json
import logging
import re
import sys
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from cli_utils import setup_logging as configure_logging
from md_parser import (
    Blockquote,
    CodeBlock,
    Diagram,
    DocumentModel,
    Element,
    ElementType,
    Heading,
    Image,
    LaTeXEquation,
    ListItem,
    Paragraph,
    Table,
    TextRun,
    parse_markdown_file,
)
from md_to_word import IBReportConverter
from word_parser import parse_word_file
from word_to_md import WordToMarkdownConverter

logger = logging.getLogger("roundtrip_audit")
_INLINE_WHITESPACE_RE = re.compile(r"[ \t]+")


@dataclass
class ModelSummary:
    """Compact semantic summary used for round-trip comparisons."""

    title: str
    subtitle: str
    company: str
    metadata_extra: Dict[str, str]
    element_counts: Dict[str, int]
    headings: List[str]
    footnote_count: int
    image_count: int
    table_count: int
    elements: List["ElementSummary"] = field(default_factory=list)


@dataclass
class ElementSummary:
    """Canonical representation of a single document element."""

    type: str
    text: str = ""
    level: Optional[int] = None
    language: str = ""
    row_count: Optional[int] = None
    column_count: Optional[int] = None
    alt_text: str = ""
    list_depth: Optional[int] = None
    numbering_format: str = ""
    bold_run_count: int = 0
    italic_run_count: int = 0
    diagram_type: str = ""
    box_count: int = 0
    arrow_count: int = 0


@dataclass
class ElementChange:
    """A specific element-level change detected in the round-trip."""

    change_type: str
    source_index: Optional[int]
    roundtrip_index: Optional[int]
    changed_fields: List[str]
    source: Optional[ElementSummary]
    roundtrip: Optional[ElementSummary]


@dataclass
class AuditDiff:
    """Semantic deltas between source and round-tripped content."""

    metadata_fields: Dict[str, Dict[str, Optional[str]]]
    metadata_extra: Dict[str, Dict[str, Optional[str]]]
    element_counts: Dict[str, int]
    footnote_count: int
    image_count: int
    table_count: int
    changed_elements: List[ElementChange]


@dataclass
class AuditReport:
    """Round-trip audit result."""

    input: str
    input_type: str
    source: ModelSummary
    roundtrip: ModelSummary
    diff: AuditDiff


def _normalize_text(text: str) -> str:
    """Collapse inline whitespace for stable text comparisons."""
    return _INLINE_WHITESPACE_RE.sub(" ", text.strip())


def _normalize_block_text(text: str) -> str:
    """Normalize block text while preserving line boundaries."""
    lines = text.splitlines()
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(line.rstrip() for line in lines)


def _summarize_runs(runs: List[TextRun]) -> Tuple[int, int]:
    """Count bold and italic runs for compact formatting diffs."""
    bold_runs = sum(1 for run in runs if getattr(run, "bold", False))
    italic_runs = sum(1 for run in runs if getattr(run, "italic", False))
    return bold_runs, italic_runs


def _infer_numbering_format(number: str) -> str:
    """Infer a compact numbering format label."""
    stripped = str(number).strip()
    if stripped.isdigit():
        return "decimal"
    if stripped:
        return "custom"
    return ""


def _summarize_element_runs(content: object) -> Tuple[int, int]:
    """Count bold/italic runs for the given element content."""
    if isinstance(content, (Paragraph, ListItem)):
        return _summarize_runs(content.runs)
    if isinstance(content, tuple) and len(content) == 2 and isinstance(content[1], ListItem):
        return _summarize_runs(content[1].runs)
    if isinstance(content, Table):
        bold_runs = 0
        italic_runs = 0
        for row in content.rows:
            for cell in row.cells:
                cell_bold, cell_italic = _summarize_runs(cell.runs)
                bold_runs += cell_bold
                italic_runs += cell_italic
        return bold_runs, italic_runs
    return 0, 0


def summarize_element(element: Element) -> ElementSummary:
    """Build a canonical summary for a single element."""
    content = element.content
    bold_runs, italic_runs = _summarize_element_runs(content)
    summary = ElementSummary(
        type=element.element_type.name,
        bold_run_count=bold_runs,
        italic_run_count=italic_runs,
    )

    if isinstance(content, Heading):
        summary.level = content.level
        summary.text = _normalize_text(content.text)
    elif isinstance(content, Paragraph):
        summary.text = _normalize_text(content.text)
    elif isinstance(content, tuple) and len(content) == 2 and isinstance(content[1], ListItem):
        number, item = content
        summary.text = _normalize_text(item.text)
        summary.list_depth = item.indent_level
        summary.numbering_format = _infer_numbering_format(str(number))
    elif isinstance(content, ListItem):
        summary.text = _normalize_text(content.text)
        summary.list_depth = content.indent_level
        if element.element_type == ElementType.BULLET_LIST:
            summary.numbering_format = "bullet"
    elif isinstance(content, Table):
        summary.row_count = len(content.rows)
        summary.column_count = content.col_count or max(
            (len(row.cells) for row in content.rows),
            default=0,
        )
    elif isinstance(content, Blockquote):
        title = _normalize_text(content.title)
        body = _normalize_text(content.text)
        summary.text = f"{title}: {body}".strip(": ")
    elif isinstance(content, Image):
        summary.alt_text = _normalize_text(content.alt_text)
    elif isinstance(content, CodeBlock):
        summary.language = _normalize_text(content.language)
        summary.text = _normalize_block_text(content.code)
    elif isinstance(content, Diagram):
        summary.diagram_type = _normalize_text(content.diagram_type)
        summary.text = _normalize_text(content.title)
        summary.box_count = len(content.boxes)
        summary.arrow_count = len(content.arrows)
    elif isinstance(content, LaTeXEquation):
        summary.text = _normalize_block_text(content.expression)
    elif element.element_type == ElementType.SEPARATOR:
        summary.text = "---"
    elif element.raw_text:
        summary.text = _normalize_block_text(element.raw_text)

    return summary


def summarize_model(model: DocumentModel) -> ModelSummary:
    """Summarize a parsed document model for round-trip comparison."""
    element_counts: Dict[str, int] = {}
    elements = []
    for element in model.elements:
        key = element.element_type.name
        element_counts[key] = element_counts.get(key, 0) + 1
        elements.append(summarize_element(element))

    headings = [
        element.text
        for element in elements
        if element.level is not None
    ]

    return ModelSummary(
        title=model.metadata.title,
        subtitle=model.metadata.subtitle,
        company=model.metadata.company,
        metadata_extra=dict(sorted(model.metadata.extra.items())),
        element_counts=element_counts,
        headings=headings,
        footnote_count=len(model.footnotes),
        image_count=element_counts.get(ElementType.IMAGE.name, 0),
        table_count=element_counts.get(ElementType.TABLE.name, 0),
        elements=elements,
    )


def _diff_mapping(
    source: Dict[str, str],
    roundtrip: Dict[str, str],
) -> Dict[str, Dict[str, Optional[str]]]:
    """Return changed key/value pairs between two dictionaries."""
    differences: Dict[str, Dict[str, Optional[str]]] = {}
    for key in sorted(set(source) | set(roundtrip)):
        source_value = source.get(key)
        roundtrip_value = roundtrip.get(key)
        if source_value != roundtrip_value:
            differences[key] = {
                "source": source_value,
                "roundtrip": roundtrip_value,
            }
    return differences


def _element_key(summary: ElementSummary) -> str:
    """Return a stable comparison key for element alignment."""
    return json.dumps(asdict(summary), ensure_ascii=False, sort_keys=True)


def _changed_element_fields(
    source: ElementSummary,
    roundtrip: ElementSummary,
) -> List[str]:
    """Return the names of element fields that changed."""
    source_dict = asdict(source)
    roundtrip_dict = asdict(roundtrip)
    return [
        key
        for key in source_dict
        if source_dict[key] != roundtrip_dict[key]
    ]


def _diff_elements(
    source: List[ElementSummary],
    roundtrip: List[ElementSummary],
) -> List[ElementChange]:
    """Diff canonical element snapshots without cascading position noise."""
    matcher = difflib.SequenceMatcher(
        a=[_element_key(element) for element in source],
        b=[_element_key(element) for element in roundtrip],
        autojunk=False,
    )
    changes: List[ElementChange] = []

    for tag, source_start, source_end, roundtrip_start, roundtrip_end in matcher.get_opcodes():
        if tag == "equal":
            continue

        if tag == "replace":
            overlap = min(source_end - source_start, roundtrip_end - roundtrip_start)
            for offset in range(overlap):
                source_element = source[source_start + offset]
                roundtrip_element = roundtrip[roundtrip_start + offset]
                changes.append(
                    ElementChange(
                        change_type="changed",
                        source_index=source_start + offset + 1,
                        roundtrip_index=roundtrip_start + offset + 1,
                        changed_fields=_changed_element_fields(source_element, roundtrip_element),
                        source=source_element,
                        roundtrip=roundtrip_element,
                    )
                )

            for source_index in range(source_start + overlap, source_end):
                changes.append(
                    ElementChange(
                        change_type="removed",
                        source_index=source_index + 1,
                        roundtrip_index=None,
                        changed_fields=[],
                        source=source[source_index],
                        roundtrip=None,
                    )
                )

            for roundtrip_index in range(roundtrip_start + overlap, roundtrip_end):
                changes.append(
                    ElementChange(
                        change_type="added",
                        source_index=None,
                        roundtrip_index=roundtrip_index + 1,
                        changed_fields=[],
                        source=None,
                        roundtrip=roundtrip[roundtrip_index],
                    )
                )
            continue

        if tag == "delete":
            for source_index in range(source_start, source_end):
                changes.append(
                    ElementChange(
                        change_type="removed",
                        source_index=source_index + 1,
                        roundtrip_index=None,
                        changed_fields=[],
                        source=source[source_index],
                        roundtrip=None,
                    )
                )
            continue

        if tag == "insert":
            for roundtrip_index in range(roundtrip_start, roundtrip_end):
                changes.append(
                    ElementChange(
                        change_type="added",
                        source_index=None,
                        roundtrip_index=roundtrip_index + 1,
                        changed_fields=[],
                        source=None,
                        roundtrip=roundtrip[roundtrip_index],
                    )
                )

    return changes


def build_audit_diff(source: ModelSummary, roundtrip: ModelSummary) -> AuditDiff:
    """Build a structured semantic diff between two model summaries."""
    metadata_fields = _diff_mapping(
        {
            "title": source.title,
            "subtitle": source.subtitle,
            "company": source.company,
        },
        {
            "title": roundtrip.title,
            "subtitle": roundtrip.subtitle,
            "company": roundtrip.company,
        },
    )
    element_counts = {
        key: roundtrip.element_counts.get(key, 0) - source.element_counts.get(key, 0)
        for key in sorted(set(source.element_counts) | set(roundtrip.element_counts))
        if roundtrip.element_counts.get(key, 0) != source.element_counts.get(key, 0)
    }

    return AuditDiff(
        metadata_fields=metadata_fields,
        metadata_extra=_diff_mapping(source.metadata_extra, roundtrip.metadata_extra),
        element_counts=element_counts,
        footnote_count=roundtrip.footnote_count - source.footnote_count,
        image_count=roundtrip.image_count - source.image_count,
        table_count=roundtrip.table_count - source.table_count,
        changed_elements=_diff_elements(source.elements, roundtrip.elements),
    )


def build_report_from_models(
    input_path: Path,
    input_type: str,
    source_model: DocumentModel,
    roundtrip_model: DocumentModel,
) -> AuditReport:
    """Build an audit report from parsed source and round-tripped models."""
    source_summary = summarize_model(source_model)
    roundtrip_summary = summarize_model(roundtrip_model)
    return AuditReport(
        input=str(input_path),
        input_type=input_type,
        source=source_summary,
        roundtrip=roundtrip_summary,
        diff=build_audit_diff(source_summary, roundtrip_summary),
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

        return build_report_from_models(
            input_path=input_path,
            input_type=source_kind,
            source_model=source_model,
            roundtrip_model=roundtrip_model,
        )


def _format_value(value: Optional[str]) -> str:
    """Render optional values for human-readable output."""
    if value is None or value == "":
        return "<missing>"
    return value


def _preview_text(text: str, limit: int = 80) -> str:
    """Create a short single-line preview for human-readable diffs."""
    preview = text.replace("\n", "\\n")
    if len(preview) <= limit:
        return preview
    return f"{preview[: limit - 3]}..."


def has_differences(report: AuditReport) -> bool:
    """Return True when the report contains a semantic regression."""
    diff = report.diff
    return any(
        [
            diff.metadata_fields,
            diff.metadata_extra,
            diff.element_counts,
            diff.footnote_count,
            diff.image_count,
            diff.table_count,
            diff.changed_elements,
        ]
    )


def format_audit_report(report: AuditReport) -> str:
    """Render a human-readable summary of the audit report."""
    lines = [
        f"Input: {report.input}",
        f"Type: {report.input_type}",
        f"Status: {'FAIL' if has_differences(report) else 'PASS'}",
    ]

    if not has_differences(report):
        lines.append("No semantic differences detected.")
        return "\n".join(lines)

    if report.diff.metadata_fields:
        lines.append("Metadata fields:")
        for key, values in report.diff.metadata_fields.items():
            lines.append(
                "  - {key}: {source} -> {roundtrip}".format(
                    key=key,
                    source=_format_value(values["source"]),
                    roundtrip=_format_value(values["roundtrip"]),
                )
            )

    if report.diff.metadata_extra:
        lines.append("Metadata extra:")
        for key, values in report.diff.metadata_extra.items():
            lines.append(
                "  - {key}: {source} -> {roundtrip}".format(
                    key=key,
                    source=_format_value(values["source"]),
                    roundtrip=_format_value(values["roundtrip"]),
                )
            )

    if report.diff.element_counts:
        lines.append(f"Element count deltas: {report.diff.element_counts}")

    if report.diff.footnote_count:
        lines.append(f"Footnotes delta: {report.diff.footnote_count:+d}")
    if report.diff.image_count:
        lines.append(f"Images delta: {report.diff.image_count:+d}")
    if report.diff.table_count:
        lines.append(f"Tables delta: {report.diff.table_count:+d}")

    if report.diff.changed_elements:
        lines.append("Changed elements:")
        for change in report.diff.changed_elements:
            source_label = (
                f"src#{change.source_index}:{change.source.type}"
                if change.source_index is not None and change.source is not None
                else "src#<missing>"
            )
            roundtrip_label = (
                f"rt#{change.roundtrip_index}:{change.roundtrip.type}"
                if change.roundtrip_index is not None and change.roundtrip is not None
                else "rt#<missing>"
            )
            fields = f" fields={','.join(change.changed_fields)}" if change.changed_fields else ""
            preview_source = _preview_text(change.source.text) if change.source else "<missing>"
            preview_roundtrip = (
                _preview_text(change.roundtrip.text) if change.roundtrip else "<missing>"
            )
            lines.append(
                f"  - {change.change_type}: {source_label} -> {roundtrip_label}{fields}"
            )
            lines.append(
                f"    text: {preview_source} -> {preview_roundtrip}"
            )

    return "\n".join(lines)


def audit(input_path: Path) -> str:
    """Return a human-readable round-trip audit diff."""
    return format_audit_report(build_audit_report(input_path))


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

    input_path = Path(args.input_file).resolve()
    if not input_path.exists():
        logger.error("File not found: %s", input_path)
        sys.exit(1)

    if args.json:
        logging.disable(logging.CRITICAL)
        try:
            report = build_audit_report(input_path)
        finally:
            logging.disable(logging.NOTSET)
        print(json.dumps(asdict(report), ensure_ascii=False, indent=2, sort_keys=True))
        return

    configure_logging(verbose=args.verbose)
    print(audit(input_path))


if __name__ == "__main__":
    main()
