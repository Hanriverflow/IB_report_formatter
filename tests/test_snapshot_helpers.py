"""Guard snapshot coverage and the narrow boundary of each normalization."""

import base64
import hashlib
import json
from dataclasses import fields
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

import pytest
from lxml import etree

import ib_renderer
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file
from tests.snapshot_helpers import ModelSerializer, XMLNormalizer, model_json, xml_parts


def test_model_preserves_fields_order_paths_enums_and_binary(tmp_path: Path) -> None:
    payload = b"example embedded image bytes"
    encoded = base64.b64encode(payload).decode("ascii")
    source = tmp_path / "example.md"
    source.write_text(
        f"---\nprofile: plain\ntitle: 예시 문서\n---\n# 예시 제목\n\n![예시 이미지](data:image/png;base64,{encoded})\n",
        encoding="utf-8",
    )
    model = parse_markdown_file(str(source))
    model.metadata.extra.update(
        {
            "z_first": {"z": "first", "a": "second"},
            "a_second": [3, 1, 2],
            "house": str(tmp_path / "nested" / "example.yaml"),
            "bytes": payload,
            "source_date": date(2026, 11, 10),
        }
    )
    text = model_json(model, source.parent)
    result = json.loads(text)
    assert list(result) == [field.name for field in fields(model)]
    assert list(result["metadata"]) == [field.name for field in fields(model.metadata)]
    assert list(result["metadata"]["extra"]) == list(model.metadata.extra)
    assert list(result["metadata"]["extra"]["z_first"]) == ["z", "a"]
    assert result["metadata"]["extra"]["a_second"] == [3, 1, 2]
    assert result["metadata"]["extra"]["house"] == "nested/example.yaml"
    assert result["metadata"]["extra"]["source_date"] == "2026-11-10"
    assert result["source_dir"] == "."
    assert result["asset_roots"] == ["."]
    assert str(tmp_path) not in text and tmp_path.as_posix() not in text
    assert [element["element_type"] for element in result["elements"]] == [
        element.element_type.name for element in model.elements
    ]
    assert len(result["sections"]) == len(model.sections)
    image = next(e for e in result["elements"] if e["element_type"] == "IMAGE")
    expected = {"length": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}
    assert image["content"]["base64_data"] == result["metadata"]["extra"]["bytes"] == expected
    assert encoded not in text
    assert "![예시 이미지](data:image/png;base64,[length=" in image["raw_text"]


def test_model_rejects_unknown_types_instead_of_repr(tmp_path: Path) -> None:
    with pytest.raises(TypeError, match="Unsupported snapshot model type"):
        ModelSerializer(tmp_path).convert(object())


@pytest.mark.parametrize("text", [r"C:\example\title.txt", "/example/title.txt"])
def test_absolute_looking_title_and_paragraph_are_literal_text(tmp_path, text):
    source = "---\nprofile: plain\ntitle: '" + text + "'\n---\n" + text
    model = MarkdownParser(profile="plain").parse(source)
    model.metadata.extra["note"] = text
    # Even a path-named key inside arbitrary prose metadata is not a file field.
    model.metadata.extra["example"] = {"house": text, "logo": text, "source": text}
    serialized = json.loads(model_json(model, tmp_path))
    assert serialized["metadata"]["title"] == text
    assert serialized["metadata"]["extra"]["note"] == text
    assert serialized["metadata"]["extra"]["example"] == model.metadata.extra["example"]
    assert serialized["elements"][0]["content"]["text"] == text
    assert serialized["elements"][0]["content"]["runs"][0]["text"] == text


def test_known_path_fields_are_normalized_without_rewriting_labels(tmp_path):
    path = str(tmp_path / "example.png")
    model = MarkdownParser(profile="plain").parse(f"![예시]({path})")
    model.metadata.extra.update({"house": path, "theme": path, "style": {"logo": path}})
    serialized = json.loads(model_json(model, tmp_path))
    extra = serialized["metadata"]["extra"]
    assert extra["house"] == extra["theme"] == extra["style"]["logo"] == "example.png"
    assert serialized["elements"][0]["content"]["path"] == "example.png"


def test_xml_normalization_preserves_content_ids_and_semantics() -> None:
    word = XMLNormalizer.WORD
    data = f'''<w:document xmlns:w="{word}" xmlns:x="urn:example">
<w:body><w:p w:rsidR="A123" x:rsidR="keep" z="last" a="first">
<w:bookmarkStart w:id="42" w:name="_ibrep_CODE_BLOCK_python_a1b2"/>
<w:bookmarkStart w:id="43" w:name="user_a1b2"/>
<w:bookmarkStart w:id="44" w:name="_ibrep_CHART_nothex"/>
<w:r><w:t xml:space="preserve">  예시 2026 text  </w:t></w:r>
</w:p><w:rsids><w:rsid w:val="A123"/><w:rsidRoot w:val="B456"/></w:rsids>
</w:body></w:document>'''.encode()
    normalized = XMLNormalizer.normalize(data, "word/document.xml")
    root = etree.fromstring(normalized.encode())
    ns = {"w": word}
    paragraph = root.find("w:body/w:p", ns)
    assert paragraph is not None
    assert paragraph.attrib[f"{{{word}}}rsidR"] == "REVISION"
    assert paragraph.attrib["{urn:example}rsidR"] == "keep"
    assert list(paragraph.attrib) == sorted(paragraph.attrib)
    bookmarks = paragraph.findall("w:bookmarkStart", ns)
    assert [b.get(f"{{{word}}}name") for b in bookmarks] == [
        "_ibrep_CODE_BLOCK_python_RANDOM",
        "user_a1b2",
        "_ibrep_CHART_nothex",
    ]
    assert [b.get(f"{{{word}}}id") for b in bookmarks] == ["42", "43", "44"]
    assert paragraph.find("w:r/w:t", ns).text == "  예시 2026 text  "
    assert "A123" not in normalized and "B456" not in normalized
    assert XMLNormalizer.normalize(normalized.encode(), "word/document.xml") == normalized


def test_only_core_creation_and_modification_timestamps_are_normalized() -> None:
    data = f'''<properties xmlns:dcterms="{XMLNormalizer.DCTERMS}">
<dcterms:created>2026-01-01T00:00:00Z</dcterms:created>
<dcterms:modified>2027-01-01T00:00:00Z</dcterms:modified>
<dcterms:date>2028-01-01</dcterms:date><revision>7</revision></properties>'''.encode()
    core = XMLNormalizer.normalize(data, "docProps/core.xml")
    assert core.count("TIMESTAMP") == 2
    assert "2028-01-01" in core and "<revision>7</revision>" in core
    custom = XMLNormalizer.normalize(data, "docProps/custom.xml")
    assert "2026-01-01T00:00:00Z" in custom and "2027-01-01T00:00:00Z" in custom


def test_xml_declaration_is_not_discarded() -> None:
    declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    normalized = XMLNormalizer.normalize(declaration + b"<example/>", "example.xml")
    assert normalized == declaration.decode() + "\n<example/>\n"
    changed = XMLNormalizer.normalize(
        declaration.replace(b'"yes"', b'"no"') + b"<example/>", "example.xml"
    )
    assert normalized != changed
    assert XMLNormalizer.normalize(normalized.encode(), "example.xml") == normalized


def test_generated_copyright_year_is_volatile(monkeypatch: pytest.MonkeyPatch) -> None:
    """Prove the extra normalization against actual output from two run years."""
    model = MarkdownParser(profile="ib-report").parse(
        '---\ncompany: 예시금융\ndate: "2026-11-10"\n---\n'
        "# 예시 보고서\n\n(c) 1999 예시금융. All rights reserved.\n"
    )
    outputs = []
    for year in (2035, 2036):
        # Replace this module's clock reference only; do not patch global time.
        monkeypatch.setattr(
            ib_renderer, "time", SimpleNamespace(strftime=lambda fmt, y=year: str(y))
        )
        document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
        raw = document.part.blob
        assert f"(c) {year} ".encode() in raw
        outputs.append(
            XMLNormalizer.normalize(raw, "word/document.xml", copyright_company="예시금융")
        )
    assert outputs[0] == outputs[1]
    assert "(c) YEAR 예시금융. All rights reserved." in outputs[0]
    assert "(c) 1999 예시금융. All rights reserved." in outputs[0]
    assert "2026-11-10" in outputs[0]


def test_copyright_like_source_text_without_disclaimer_is_preserved() -> None:
    model = MarkdownParser(profile="plain").parse("(c) 1999 예시금융. All rights reserved.")
    document = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
    normalized = XMLNormalizer.normalize(document.part.blob, "word/document.xml")
    assert "(c) 1999 예시금융. All rights reserved." in normalized
    assert "YEAR" not in normalized


def test_all_xml_parts_are_included_and_binary_is_excluded(tmp_path: Path) -> None:
    document = tmp_path / "example.docx"
    names = [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/custom.xml",
        "word/footnotes.xml",
        "word/header1.xml",
        "word/footer1.xml",
        "word/settings.xml",
        "word/_rels/header1.xml.rels",
        "customXml/item99.xml",
    ]
    with ZipFile(document, "w") as archive:
        for name in names:
            archive.writestr(name, b'<example z="2" a="1"/>')
        archive.writestr("word/media/image1.png", b"example binary")
    parts = xml_parts(document)
    assert list(parts) == sorted(names)
    assert all(xml == '<example a="1" z="2"/>\n' for xml in parts.values())
    with ZipFile(document, "a") as archive:
        archive.writestr("unanticipated/new-part.xml", b"<new/>")
    assert list(xml_parts(document)) != list(parts)
