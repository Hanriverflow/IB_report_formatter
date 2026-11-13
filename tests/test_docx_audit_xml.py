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
from scripts.sample_corpus import render_corpus

SOURCE = (
    "---\nterms:\n  amount: '500'\nchecks:\n  - amount = 500\n---\n"
    "# 가상 감사 문서\n\n금액 {{amount}}.[^1]\n\n"
    "1. 예시 항목\n\n[^1]: 가상 각주.\n"
)
MARKER = "가상_ENTITY_REPLACEMENT_MUST_NOT_BE_READ"
REFERENCE = "AUDIT_ENTITY_REFERENCE"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _write_parts(path, parts):
    with ZipFile(path, "w") as archive:
        for name, blob in parts.items():
            archive.writestr(name, blob)


def _package_parts():
    with ZipFile(BytesIO(_render_package())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


@pytest.mark.parametrize("name", ["document.bin", "document.XML"])
@pytest.mark.parametrize("mapping", ["override", "default"])
def test_renamed_xml_main_part_is_checked(tmp_path, monkeypatch, name, mapping):
    parts = _package_parts()
    root = etree.fromstring(parts.pop("word/document.xml"))
    root.set("audit_probe", REFERENCE)
    parts["word/" + name] = (
        f'<!DOCTYPE probe [<!ENTITY probe "{MARKER}">]>'
        + etree.tostring(root, encoding="unicode").replace(REFERENCE, "&probe;")
    ).encode()
    parts["_rels/.rels"] = parts["_rels/.rels"].replace(b"document.xml", name.encode())
    parts[f"word/_rels/{name}.rels"] = parts.pop("word/_rels/document.xml.rels")
    types = etree.fromstring(parts["[Content_Types].xml"])
    main = next(node for node in types if node.get("PartName") == "/word/document.xml")
    if mapping == "override":
        main.set("PartName", "/WORD/" + name.upper())
    else:
        extension = name.rsplit(".", 1)[1].upper()
        etree.SubElement(
            types, f"{{{CT_NS}}}Default", Extension=extension, ContentType=main.get("ContentType")
        )
        types.remove(main)
    parts["[Content_Types].xml"] = etree.tostring(types)
    path = tmp_path / "example-renamed.docx"
    _write_parts(path, parts)
    # Demonstrate that the installed python-docx parses the renamed part and expands it.
    assert Document(path).element.get("audit_probe") == MARKER
    monkeypatch.setattr(docx_audit, "Document", lambda *a: pytest.fail("unsafe package opened"))
    assert docx_audit.audit_file(str(path)).issues == [
        f"XML part cannot be read safely: /word/{name}"
    ]


def test_binary_attachment_named_xml_is_not_xml(tmp_path):
    parts = _package_parts()
    types = etree.fromstring(parts["[Content_Types].xml"])
    etree.SubElement(
        types,
        f"{{{CT_NS}}}Override",
        PartName="/WORD/EXAMPLE.XML",
        ContentType="application/octet-stream",
    )
    parts["[Content_Types].xml"] = etree.tostring(types)
    parts["word/example.xml"] = b"\x00\xffexample binary"
    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    etree.SubElement(
        rels,
        "{" + etree.QName(rels).namespace + "}Relationship",
        Id="rIdExample",
        Type="urn:example:attachment",
        Target="example.xml",
    )
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels)
    path = tmp_path / "example-binary.docx"
    _write_parts(path, parts)
    expected = docx_audit.inspect_document(Document(path))
    assert not expected.issues
    assert docx_audit.audit_file(str(path)) == expected


@pytest.mark.parametrize("failure", ["missing", "malformed", "dtd"])
def test_content_types_checked_before_open(tmp_path, monkeypatch, failure):
    parts = _package_parts()
    if failure == "missing":
        del parts["[Content_Types].xml"]
    elif failure == "malformed":
        parts["[Content_Types].xml"] = b"not XML"
    else:
        parts["[Content_Types].xml"] = (
            b"<!DOCTYPE Types>" + parts["[Content_Types].xml"].split(b"?>", 1)[1]
        )
    path = tmp_path / "example-types.docx"
    _write_parts(path, parts)
    monkeypatch.setattr(docx_audit, "Document", lambda *a: pytest.fail("unsafe package opened"))
    assert docx_audit.audit_file(str(path)).issues == [
        "XML part cannot be read safely: /[Content_Types].xml"
    ]


@pytest.mark.parametrize("content_type", ["application/xml", "text/xml", "application/example+xml"])
@pytest.mark.parametrize("name", ["example.BIN", "example.RELS"])
def test_xml_types_and_uppercase_relationships_checked(tmp_path, monkeypatch, content_type, name):
    parts = _package_parts()
    types = etree.fromstring(parts["[Content_Types].xml"])
    etree.SubElement(
        types,
        f"{{{CT_NS}}}Default",
        Extension=name.rsplit(".", 1)[1].lower(),
        ContentType=content_type if name.endswith("BIN") else "application/octet-stream",
    )
    parts["[Content_Types].xml"] = etree.tostring(types)
    parts[name] = b"<!DOCTYPE example><example/>"
    path = tmp_path / "example-extra.docx"
    _write_parts(path, parts)
    monkeypatch.setattr(docx_audit, "Document", lambda *a: pytest.fail("unsafe package opened"))
    assert docx_audit.audit_file(str(path)).issues == [f"XML part cannot be read safely: /{name}"]


def test_all_sample_audits_and_json_unchanged(tmp_path):
    for path in render_corpus(tmp_path / "corpus"):
        document = Document(path)
        expected = docx_audit.inspect_document(document)
        expected_terms = docx_audit.inspect_terms(document)
        actual, actual_terms = docx_audit._audit_package(str(path))
        assert actual == expected, path.name
        assert json.dumps(
            docx_audit.audit_to_dict(actual, actual_terms), ensure_ascii=False
        ) == json.dumps(docx_audit.audit_to_dict(expected, expected_terms), ensure_ascii=False), (
            path.name
        )


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
    # File audits now stop before opening any part of a rejected package.
    assert MARKER not in json.dumps(docx_audit.audit_to_dict(result, terms), ensure_ascii=False)
    payload = docx_audit.audit_to_dict(result)
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


@pytest.mark.parametrize(
    "part_name",
    [
        "word/document.xml",
        "word/numbering.xml",
        "word/styles.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
        "unused.xml",
    ],
)
@pytest.mark.parametrize("entity_kind", ["internal", "external"])
@pytest.mark.parametrize("entry_point", ["api", "cli"])
def test_package_preflight_rejects_dtd_before_python_docx(
    tmp_path, monkeypatch, caplog, part_name, entity_kind, entry_point
):
    package = _render_package()
    with ZipFile(BytesIO(package)) as archive:
        blob = archive.read(part_name) if part_name != "unused.xml" else b"<unused/>"
    root = etree.fromstring(blob)
    root.set("audit_probe", REFERENCE)
    declaration = f'<!ENTITY probe "{MARKER}">'
    if entity_kind == "external":
        external = tmp_path / "example-entity.txt"
        external.write_text(MARKER, encoding="utf-8")
        declaration = f'<!ENTITY probe SYSTEM "{external.as_uri()}">'
    xml = (
        f"<!DOCTYPE probe [{declaration}]>"
        + etree.tostring(root, encoding="unicode").replace(REFERENCE, "&probe;")
    ).encode("utf-8")
    path = tmp_path / "example-dtd.docx"
    with ZipFile(BytesIO(package)) as source, ZipFile(path, "w") as destination:
        for item in source.infolist():
            if item.filename != part_name:
                destination.writestr(item, source.read(item))
        destination.writestr(part_name, xml)

    def forbidden_open(*args, **kwargs):
        pytest.fail("python-docx consumed a DTD-bearing package")

    monkeypatch.setattr(docx_audit, "Document", forbidden_open)
    if entry_point == "api":
        payload = docx_audit.audit_to_dict(docx_audit.audit_file(str(path)))
    else:
        monkeypatch.setattr("sys.argv", ["docx-audit", str(path)])
        with caplog.at_level(logging.INFO), pytest.raises(SystemExit) as caught:
            docx_audit.main()
        assert caught.value.code == 1
        payload = json.loads(caplog.records[-1].getMessage())
    assert payload["issues"] == [f"XML part cannot be read safely: /{part_name}"]
    assert MARKER not in json.dumps(payload, ensure_ascii=False)


@pytest.mark.parametrize("name", ["document", ".document"])
def test_extensionless_xml_main_part_is_checked(tmp_path, monkeypatch, name):
    """python-docx gives a name without an extension, and a dotfile, the extension ""."""
    parts = _package_parts()
    root = etree.fromstring(parts.pop("word/document.xml"))
    root.set("audit_probe", REFERENCE)
    parts["word/" + name] = (
        f'<!DOCTYPE probe [<!ENTITY probe "{MARKER}">]>'
        + etree.tostring(root, encoding="unicode").replace(REFERENCE, "&probe;")
    ).encode()
    parts["_rels/.rels"] = parts["_rels/.rels"].replace(b"document.xml", name.encode())
    parts[f"word/_rels/{name}.rels"] = parts.pop("word/_rels/document.xml.rels")
    types = etree.fromstring(parts["[Content_Types].xml"])
    main = next(node for node in types if node.get("PartName") == "/word/document.xml")
    etree.SubElement(
        types, f"{{{CT_NS}}}Default", Extension="", ContentType=main.get("ContentType")
    )
    types.remove(main)
    parts["[Content_Types].xml"] = etree.tostring(types)
    path = tmp_path / "example-extensionless.docx"
    _write_parts(path, parts)
    # Demonstrate that the installed python-docx parses the part and expands the entity.
    assert Document(path).element.get("audit_probe") == MARKER
    monkeypatch.setattr(docx_audit, "Document", lambda *a: pytest.fail("unsafe package opened"))
    assert docx_audit.audit_file(str(path)).issues == [
        f"XML part cannot be read safely: /word/{name}"
    ]
