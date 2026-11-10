"""Audit untrusted XML parts in fictional, engine-generated DOCX packages."""

import json
import logging
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

import docx_audit
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser
from numeric_checks import CHECKS_PROPERTY

SOURCE = (
    "---\nterms:\n  amount: '500'\nchecks:\n  - amount = 500\n---\n"
    "# 가상 감사 문서\n\n금액 {{amount}}.[^1]\n\n"
    "1. 예시 항목\n\n[^1]: 가상 각주.\n"
)
MARKER = "가상_ENTITY_REPLACEMENT_MUST_NOT_BE_READ"
REFERENCE = "AUDIT_ENTITY_REFERENCE"


def _render_package() -> bytes:
    """Parse and strictly render the fictional fixture into a DOCX package."""
    model = MarkdownParser(profile="plain").parse(SOURCE)
    document = IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model)
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def _inject_entity(package: bytes, target: str, entity_kind: str, tmp_path: Path) -> Path:
    """Rewrite one generated XML part with an internal or file-backed entity.

    Args:
        package: Original DOCX bytes from the engine.
        target: Audit reader to exercise.
        entity_kind: Internal declaration or external file reference.
        tmp_path: Isolated test directory for the package and external payload.

    Returns:
        Path to the rewritten DOCX.
    """
    part_name = "word/footnotes.xml" if target == "footnotes" else "docProps/custom.xml"
    with ZipFile(BytesIO(package)) as archive:
        root = etree.fromstring(archive.read(part_name))
    replacement = MARKER
    if target == "snapshot-name":
        prop = next(node for node in root if node.get("name") == "ibrep.term.amount")
        prop.set("name", "ibrep.term." + REFERENCE)
    elif target in {"snapshot", "checks"}:
        property_name = "ibrep.term.amount" if target == "snapshot" else CHECKS_PROPERTY
        prop = next(node for node in root if node.get("name") == property_name)
        prop[0].text = REFERENCE
        if target == "checks":
            replacement = json.dumps(
                {"checks": ["left = right"], "values": {"left": MARKER, "right": "1"}},
                ensure_ascii=False,
            )
    elif target == "footnotes":
        note = next(node for node in root if node.get(qn("w:id")) == "1")
        note.find(".//" + qn("w:t")).text = MARKER
        replacement = etree.tostring(note, encoding="unicode")
        root.remove(note)
        root[-1].tail = REFERENCE
    else:
        replacement = (
            '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:numPr><w:numId w:val="{MARKER}"/></w:numPr></w:pPr>'
        )
        root[-1].tail = REFERENCE
    if entity_kind == "external":
        external = tmp_path / "entity-payload.txt"
        external.write_text(replacement, encoding="utf-8")
        declaration = f'<!ENTITY audit_entity SYSTEM "{external.as_uri()}">'
    else:
        declaration = f"<!ENTITY audit_entity '{replacement}'>"
    root_name = (root.prefix + ":" if root.prefix else "") + etree.QName(root).localname
    xml = (
        f"<!DOCTYPE {root_name} [{declaration}]>"
        + etree.tostring(root, encoding="unicode").replace(REFERENCE, "&audit_entity;")
    ).encode("utf-8")
    output = tmp_path / "entity-audit.docx"
    with ZipFile(BytesIO(package)) as source, ZipFile(output, "w") as destination:
        for item in source.infolist():
            destination.writestr(item, xml if item.filename == part_name else source.read(item))
    return output


@pytest.mark.parametrize(
    "target", ["snapshot", "snapshot-name", "checks", "footnotes", "numbering"]
)
@pytest.mark.parametrize("entity_kind", ["internal", "external"])
def test_audit_does_not_expand_entities(
    target: str,
    entity_kind: str,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """All four raw XML readers reject entity payloads without crashing or leaking."""
    output = _inject_entity(_render_package(), target, entity_kind, tmp_path)
    document = Document(output)
    result = docx_audit.audit_file(str(output))
    terms = docx_audit.inspect_terms(document)
    payload = docx_audit.audit_to_dict(result, terms)
    assert MARKER not in json.dumps(payload, ensure_ascii=False)
    assert result.issues
    monkeypatch.setattr("sys.argv", ["docx-audit", str(output)])
    with caplog.at_level(logging.INFO), pytest.raises(SystemExit) as exc:
        docx_audit.main()
    assert exc.value.code == 1
    reports = [record.getMessage() for record in caplog.records if record.name == "docx_audit"]
    assert MARKER not in "".join(reports)
    assert json.loads(reports[-1]) == payload


def test_normal_audit_result_and_json_are_unchanged(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Preserve the pre-hardening result and serialized CLI report for normal XML."""
    package = _render_package()
    output = tmp_path / "normal-audit.docx"
    output.write_bytes(package)
    document = Document(BytesIO(package))
    result = docx_audit.inspect_document(document)
    payload = docx_audit.audit_to_dict(result, docx_audit.inspect_terms(document))
    expected = {
        "paragraphs": 3,
        "tables": 0,
        "headings": 1,
        "images": 0,
        "numbered_paragraphs": 1,
        "issues": [],
        "warnings": [],
        "pagination_marked_paragraphs": 1,
        "visual_review": "not_performed",
        "terms": {
            "mismatched": {},
            "changed": {},
            "missing": [],
            "indicative": [],
            "failed_checks": [],
        },
    }
    assert payload == expected
    monkeypatch.setattr("sys.argv", ["docx-audit", str(output)])
    with caplog.at_level(logging.INFO), pytest.raises(SystemExit) as exc:
        docx_audit.main()
    assert exc.value.code == 0
    [report] = [record.getMessage() for record in caplog.records if record.name == "docx_audit"]
    assert report == json.dumps(expected, ensure_ascii=False, indent=2)
