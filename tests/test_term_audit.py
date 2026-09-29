"""docx-audit term checks on generated documents edited as Word would edit them."""

import json
import logging
from dataclasses import asdict
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pytest
from docx import Document
from docx.document import Document as DocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

import docx_audit
from document_profiles import RenderOptions
from docx_audit import DocumentAudit, TermAudit, audit_to_dict, inspect_document, inspect_terms
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
VALUES = {"amount": "500억원", "tenor": "3년", "rate": "CD + [1.10]%", "blank": ""}
SOURCE = (
    "Deal {{amount}} for {{tenor}}.\n\n"
    "- Amount {{amount}} at {{rate}}{{blank}}\n\n"
    "| Item | Value |\n|---|---|\n| Amount | {{amount}} |\n"
)
OLD_KEYS = [
    "paragraphs", "tables", "headings", "images", "numbered_paragraphs", "issues",
    "warnings", "pagination_marked_paragraphs", "visual_review",
]


def xp(element, path: str) -> List:
    """XPath with explicit namespaces; python-docx has no class for `w:sdt`."""
    return etree.XPath(path, namespaces=NS)(element)


def front(values: Dict[str, str]) -> str:
    """Frontmatter with exact YAML string values (JSON is valid YAML)."""
    lines = "".join(
        f"  {key}: {json.dumps(value, ensure_ascii=False)}\n" for key, value in values.items()
    )
    return "---\nterms:\n" + lines + "---\n"


def reopen(doc: DocxDocument) -> DocxDocument:
    """Save and reopen, as after editing in Word."""
    payload = BytesIO()
    doc.save(payload)
    payload.seek(0)
    return Document(payload)


def generate(markdown: str = SOURCE, values: Optional[Dict[str, str]] = None) -> DocxDocument:
    """Render a tagged document strictly and reopen it."""
    model = MarkdownParser(profile="plain").parse(front(VALUES if values is None else values) + markdown)
    return reopen(IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model))


def audit(doc: DocxDocument) -> Tuple[DocumentAudit, Optional[TermAudit]]:
    """Audit a saved copy of an edited document."""
    saved = reopen(doc)
    return inspect_document(saved), inspect_terms(saved)


def controls(doc: DocxDocument, key: str) -> List:
    """Term controls of one key in document order."""
    return xp(doc.element.body, f".//w:sdt[w:sdtPr/w:tag/@w:val='ibrep:term:{key}']")


def type_value(control, text: str) -> None:
    """Replace the value, as typing inside the control does."""
    [node] = xp(control, "./w:sdtContent/w:r/w:t")
    node.text = text


def element(tag: str, text: Optional[str] = None, **attributes: str):
    """Build a WordprocessingML element."""
    node = OxmlElement(tag)
    for name, value in attributes.items():
        node.set(qn("w:" + name), value)
    if text is not None:
        node.text = text
    return node


def run(text: str, text_tag: str = "w:t"):
    """A run holding one text node."""
    node = element("w:r")
    node.append(element(text_tag, text))
    return node


def term_warnings(result: DocumentAudit) -> List[str]:
    """Warnings added by the term checks."""
    return [warning for warning in result.warnings if warning.startswith("Term ")]


# ─────────────────────────────────────────────────────────────────────────────
# Consistency results
# ─────────────────────────────────────────────────────────────────────────────


def test_unedited_document_is_consistent_and_lists_indicative_terms() -> None:
    result, terms = audit(generate())
    assert terms == TermAudit(mismatched={}, changed={}, missing=[], indicative=["rate"])
    assert result.issues == [] and term_warnings(result) == []
    assert audit_to_dict(result, terms)["terms"] == {
        "mismatched": {}, "changed": {}, "missing": [], "indicative": ["rate"],
    }


def test_editing_one_occurrence_reports_mismatch_and_change() -> None:
    doc = generate()
    type_value(controls(doc, "amount")[1], "600억원")
    result, terms = audit(doc)
    assert terms is not None
    assert terms.mismatched == {"amount": ["500억원", "600억원"]}
    assert terms.changed == {"amount": {"generated": "500억원", "current": "600억원"}}
    assert terms.missing == []
    assert term_warnings(result) == [
        "Term 'amount' has inconsistent values in the document: '500억원', '600억원'",
    ]
    assert result.issues == []


def test_editing_every_occurrence_is_a_consistent_change() -> None:
    doc = generate()
    for control in controls(doc, "amount"):
        type_value(control, "600억원")
    result, terms = audit(doc)
    assert terms is not None
    assert terms.mismatched == {}
    assert terms.changed == {"amount": {"generated": "500억원", "current": "600억원"}}
    assert term_warnings(result) == []


def test_deleting_the_last_control_of_a_key_reports_missing() -> None:
    doc = generate()
    [tenor] = controls(doc, "tenor")
    tenor.getparent().remove(tenor)
    first_amount = controls(doc, "amount")[0]
    first_amount.getparent().remove(first_amount)
    result, terms = audit(doc)
    assert terms is not None
    assert terms.missing == ["tenor"]
    assert terms.mismatched == {} and terms.changed == {}
    assert term_warnings(result) == [
        "Term 'tenor' was generated but none of its content controls remain in the document",
    ]


def test_removing_the_control_but_keeping_its_text_reports_missing() -> None:
    doc = generate()
    [tenor] = controls(doc, "tenor")
    for content_run in xp(tenor, "./w:sdtContent/w:r"):
        tenor.addprevious(content_run)
    tenor.getparent().remove(tenor)
    _, terms = audit(doc)
    assert terms is not None and terms.missing == ["tenor"]
    assert "3년" in "".join(xp(doc.element.body, ".//w:t/text()"))


def test_every_control_removed_still_reports_the_snapshot() -> None:
    doc = generate()
    for sdt in xp(doc.element.body, ".//w:sdt"):
        sdt.getparent().remove(sdt)
    result, terms = audit(doc)
    assert terms == TermAudit(
        mismatched={}, changed={}, missing=["amount", "blank", "rate", "tenor"], indicative=[],
    )
    assert len(term_warnings(result)) == 4


# ─────────────────────────────────────────────────────────────────────────────
# Current text
# ─────────────────────────────────────────────────────────────────────────────


def test_tracked_changes_count_inserted_and_moved_to_text_only() -> None:
    doc = generate()
    first, second, third = controls(doc, "amount")
    content = first.find(qn("w:sdtContent"))
    [original] = xp(content, "./w:r")
    deleted = element("w:del", id="901", author="Reviewer")
    content.remove(original)
    deleted.append(run("500억원", "w:delText"))
    content.append(deleted)
    inserted = element("w:ins", id="902", author="Reviewer")
    inserted.append(run("700억원"))
    content.append(inserted)
    # Nonstandard w:t under a deletion is still deleted text.
    second_content = second.find(qn("w:sdtContent"))
    odd_delete = element("w:del", id="903", author="Reviewer")
    odd_delete.append(run("삭제"))
    second_content.insert(0, odd_delete)
    third_content = third.find(qn("w:sdtContent"))
    [third_run] = xp(third_content, "./w:r")
    third_content.remove(third_run)
    move_from = element("w:moveFrom", id="904", author="Reviewer")
    move_from.append(run("500억원"))
    move_to = element("w:moveTo", id="905", author="Reviewer")
    move_to.append(run("500억원"))
    third_content.extend([move_from, move_to])
    _, terms = audit(doc)
    assert terms is not None
    assert terms.mismatched == {"amount": ["700억원", "500억원"]}
    assert terms.changed == {"amount": {"generated": "500억원", "current": "700억원"}}


def test_a_control_deleted_with_tracking_counts_as_removed() -> None:
    doc = generate()
    [tenor] = controls(doc, "tenor")
    deleted = element("w:del", id="910", author="Reviewer")
    tenor.addprevious(deleted)
    deleted.append(tenor)
    _, terms = audit(doc)
    assert terms is not None and terms.missing == ["tenor"]


def test_tabs_and_breaks_are_part_of_the_current_value() -> None:
    doc = generate()
    [tenor] = controls(doc, "tenor")
    [value_run] = xp(tenor, "./w:sdtContent/w:r")
    for node in xp(value_run, "./w:t"):
        value_run.remove(node)
    value_run.extend([
        element("w:t", "3"), element("w:tab"), element("w:t", "년"), element("w:br"),
        element("w:t", "만기"),
    ])
    _, terms = audit(doc)
    assert terms is not None
    assert terms.changed == {"tenor": {"generated": "3년", "current": "3\t년\n만기"}}


def test_empty_values_compare_as_text() -> None:
    doc = generate()
    _, terms = audit(doc)
    assert terms is not None and "blank" not in terms.changed and "blank" not in terms.indicative
    [blank] = controls(doc, "blank")
    blank.find(qn("w:sdtContent")).append(run("x"))
    _, terms = audit(doc)
    assert terms is not None
    assert terms.changed == {"blank": {"generated": "", "current": "x"}}


def test_brackets_mark_indicative_values_until_they_are_removed() -> None:
    doc = generate()
    [rate] = controls(doc, "rate")
    type_value(rate, "CD + 1.10%")
    _, terms = audit(doc)
    assert terms is not None
    assert terms.indicative == []
    assert terms.changed == {"rate": {"generated": "CD + [1.10]%", "current": "CD + 1.10%"}}


def test_controls_without_a_snapshot_are_still_compared() -> None:
    doc = generate()
    part = doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
    root = etree.fromstring(part.blob)
    for prop in list(root):
        if prop.get("name").startswith("ibrep.term."):
            root.remove(prop)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    type_value(controls(doc, "amount")[0], "600억원")
    result, terms = audit(doc)
    assert terms is not None
    assert terms.mismatched == {"amount": ["600억원", "500억원"]}
    assert terms.changed == {} and terms.missing == []
    assert len(term_warnings(result)) == 1


def test_other_content_controls_are_ignored() -> None:
    doc = generate()
    [tenor] = controls(doc, "tenor")
    tenor.find(qn("w:sdtPr")).find(qn("w:tag")).set(qn("w:val"), "company:field")
    _, terms = audit(doc)
    assert terms is not None and terms.missing == ["tenor"]


# ─────────────────────────────────────────────────────────────────────────────
# Output contract
# ─────────────────────────────────────────────────────────────────────────────


def test_documents_without_terms_keep_the_previous_json(tmp_path: Path, caplog, monkeypatch) -> None:
    model = MarkdownParser(profile="plain").parse("# Title\n\nBody {{literal}}.\n\n| A |\n|---|\n| 1 |")
    doc = reopen(IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model))
    result = inspect_document(doc)
    assert inspect_terms(doc) is None
    assert audit_to_dict(result, None) == asdict(result)
    assert list(audit_to_dict(result, None)) == OLD_KEYS
    output = tmp_path / "plain.docx"
    doc.save(output)
    payload = run_cli(output, caplog, monkeypatch, expected_code=0)
    assert list(payload) == OLD_KEYS


def test_cli_adds_terms_without_changing_the_exit_code(tmp_path: Path, caplog, monkeypatch) -> None:
    doc = generate()
    type_value(controls(doc, "amount")[1], "600억원")
    output = tmp_path / "edited.docx"
    doc.save(output)
    payload = run_cli(output, caplog, monkeypatch, expected_code=0)
    assert list(payload) == OLD_KEYS + ["terms"]
    assert payload["issues"] == []
    assert payload["terms"]["mismatched"] == {"amount": ["500억원", "600억원"]}
    assert any("inconsistent" in warning for warning in payload["warnings"])


def run_cli(path: Path, caplog, monkeypatch, expected_code: int) -> Dict:
    """Run `docx-audit` and return its JSON report."""
    monkeypatch.setattr("sys.argv", ["docx-audit", str(path)])
    with caplog.at_level(logging.INFO), pytest.raises(SystemExit) as exit_info:
        docx_audit.main()
    assert exit_info.value.code == expected_code
    [report] = [record.getMessage() for record in caplog.records if record.getMessage().startswith("{")]
    caplog.clear()
    return json.loads(report)
