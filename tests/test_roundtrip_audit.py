"""Tests for roundtrip_audit helper."""

import json
import subprocess
import sys
from pathlib import Path

from md_parser import CodeBlock, DocumentMetadata, DocumentModel, Element, ElementType, Paragraph
from roundtrip_audit import (
    build_audit_diff,
    build_audit_report,
    build_report_from_models,
    format_audit_report,
    summarize_model,
)


def _make_model(elements, extra=None):
    """Build a minimal document model for audit tests."""
    metadata = DocumentMetadata(title="Audit Test", extra=extra or {})
    return DocumentModel(metadata=metadata, elements=elements, footnotes={})


def test_roundtrip_audit_reports_markdown_summary(tmp_path):
    """Audit helper should summarize metadata extras and canonical elements."""
    markdown = """---
title: "Audit Test"
analyst: "Alice Kim"
recipient: "Board"
---

# Heading

본문 문장입니다.
"""
    input_path = tmp_path / "audit.md"
    input_path.write_text(markdown, encoding="utf-8")

    report = build_audit_report(input_path)

    assert report.input_type == "markdown"
    assert report.source.title == "Audit Test"
    assert report.source.metadata_extra == {"recipient": "Board"}
    assert report.source.elements[0].type == "HEADING_1"
    assert report.source.elements[0].level == 1
    assert report.source.elements[1].type == "PARAGRAPH"
    assert report.roundtrip is not None
    assert report.diff is not None


def test_roundtrip_audit_detects_code_block_loss_with_same_counts():
    """Element-level diff should catch a replaced code block even when counts match."""
    source_model = _make_model(
        [
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="Intro"),
            ),
            Element(
                element_type=ElementType.CODE_BLOCK,
                content=CodeBlock(code="print('alpha')", language="python"),
            ),
            Element(
                element_type=ElementType.CODE_BLOCK,
                content=CodeBlock(code="SELECT * FROM deals", language="sql"),
            ),
        ]
    )
    roundtrip_model = _make_model(
        [
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="Intro"),
            ),
            Element(
                element_type=ElementType.CODE_BLOCK,
                content=CodeBlock(code="print('alpha')", language="python"),
            ),
            Element(
                element_type=ElementType.CODE_BLOCK,
                content=CodeBlock(code="print('alpha')", language="python"),
            ),
        ]
    )

    diff = build_audit_diff(summarize_model(source_model), summarize_model(roundtrip_model))

    assert diff.element_counts == {}
    assert len(diff.changed_elements) == 1
    assert diff.changed_elements[0].change_type == "changed"
    assert diff.changed_elements[0].source_index == 3
    assert diff.changed_elements[0].source.type == "CODE_BLOCK"
    assert diff.changed_elements[0].roundtrip.type == "CODE_BLOCK"
    assert "language" in diff.changed_elements[0].changed_fields
    assert diff.changed_elements[0].source.text == "SELECT * FROM deals"


def test_roundtrip_audit_detects_separator_loss_with_same_counts():
    """Element-level diff should catch a moved separator when counts still match."""
    source_model = _make_model(
        [
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="Before"),
            ),
            Element(
                element_type=ElementType.SEPARATOR,
                content=None,
                raw_text="---",
            ),
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="After"),
            ),
        ]
    )
    roundtrip_model = _make_model(
        [
            Element(
                element_type=ElementType.SEPARATOR,
                content=None,
                raw_text="---",
            ),
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="Before"),
            ),
            Element(
                element_type=ElementType.PARAGRAPH,
                content=Paragraph(text="After"),
            ),
        ]
    )

    report = build_report_from_models(
        input_path=Path("audit.md"),
        input_type="markdown",
        source_model=source_model,
        roundtrip_model=roundtrip_model,
    )
    rendered = format_audit_report(report)

    assert report.diff.element_counts == {}
    assert report.diff.changed_elements
    assert any(
        change.source is not None and change.source.type == "SEPARATOR"
        for change in report.diff.changed_elements
    )
    assert "Changed elements:" in rendered
    assert "SEPARATOR" in rendered


def test_roundtrip_audit_cli_json_outputs_expected_keys(tmp_path):
    """CLI JSON mode should emit structured summaries and element diffs."""
    markdown = """---
title: "Audit Test"
analyst: "Alice Kim"
---

# Heading

```python
print("hello")
```
"""
    input_path = tmp_path / "audit.md"
    input_path.write_text(markdown, encoding="utf-8")

    result = subprocess.run(
        [
            sys.executable,
            "roundtrip_audit.py",
            str(input_path),
            "--json",
        ],
        capture_output=True,
        cwd=str(Path(__file__).resolve().parent.parent),
        text=True,
        timeout=60,
    )

    payload = json.loads(result.stdout)

    assert result.returncode == 0
    assert "source" in payload
    assert "roundtrip" in payload
    assert "diff" in payload
    assert "metadata_extra" in payload["source"]
    assert "elements" in payload["source"]
    assert payload["source"]["elements"][0]["type"] == "HEADING_1"
    assert "changed_elements" in payload["diff"]
