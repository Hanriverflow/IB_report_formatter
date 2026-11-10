"""Readable, narrowly normalized representations for the sample snapshots.

No font, image extent, relationship target, content, or whole XML part is masked.
Every normalization below has a concrete source of nondeterminism documented next
to it. Unknown model types fail instead of falling back to machine-specific repr.
"""

import base64
import hashlib
import json
import os
import re
from collections.abc import Mapping
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from lxml import etree
from syrupy.extensions.single_file import SingleFileSnapshotExtension, WriteMode

from document_model import DocumentModel, Image


class JSONTextSnapshotExtension(SingleFileSnapshotExtension):
    """Store already formatted JSON without reordering model metadata.

    Syrupy's text reader uses universal newlines, so a checkout's CRLF versus LF
    cannot change comparison. The snapshot directory also has Git eol=lf.
    """

    file_extension = "json"
    _write_mode = WriteMode.TEXT


class XMLTextSnapshotExtension(SingleFileSnapshotExtension):
    """Store XML as readable text with the same newline policy as model JSON."""

    file_extension = "xml"
    _write_mode = WriteMode.TEXT


class ModelSerializer:
    """Retain dataclass field, mapping insertion, and sequence order."""

    _DATA_URI_RE = re.compile(r"(data:image/[a-zA-Z0-9+.-]+;base64,)([A-Za-z0-9+/=\r\n\t ]+)")

    def __init__(self, sample_dir: Path):
        self.sample_dir = sample_dir.resolve()

    @staticmethod
    def binary_summary(value: bytes) -> dict[str, int | str]:
        """Retain binary identity without dumping image bytes into a text diff."""
        return {"length": len(value), "sha256": hashlib.sha256(value).hexdigest()}

    def relative_path(self, value: Path) -> str:
        """Remove checkout/user roots; POSIX separators are platform independent."""
        if value.is_absolute():
            return Path(os.path.relpath(value, self.sample_dir)).as_posix()
        return value.as_posix()

    def summarize_data_uri(self, match: re.Match[str]) -> str:
        """Summarize duplicated image bytes in raw Markdown and metadata logos."""
        data = base64.b64decode("".join(match[2].split()), validate=True)
        summary = self.binary_summary(data)
        return f"{match[1]}[length={summary['length']};sha256={summary['sha256']}]"

    def convert(self, value: Any) -> Any:
        """Convert every model field, rejecting unsupported values.

        Args:
            value: Model, nested dataclass, collection, or scalar.

        Returns:
            JSON-compatible data in the original order.
        """
        if isinstance(value, Enum):
            return value.name
        if isinstance(value, bytes | bytearray | memoryview):
            return self.binary_summary(bytes(value))
        if isinstance(value, Path):
            return self.relative_path(value)
        if is_dataclass(value) and not isinstance(value, type):
            converted = {}
            for field in fields(value):
                item = getattr(value, field.name)
                # Image.base64_data represents bytes even though the parser stores
                # it as text. Hash decoded bytes, not an encoding/line-wrap choice.
                if isinstance(value, Image) and field.name == "base64_data" and item is not None:
                    item = base64.b64decode("".join(item.split()), validate=True)
                converted[field.name] = self.convert(item)
            return converted
        if isinstance(value, Mapping):
            return {str(key): self.convert(item) for key, item in value.items()}
        if isinstance(value, list | tuple):
            return [self.convert(item) for item in value]
        if isinstance(value, date | datetime):
            # These are source YAML dates, not the run clock; keep their values.
            return value.isoformat()
        if isinstance(value, str):
            # Raw Markdown can retain the same embedded bytes as Image.base64_data;
            # retain the MIME type and surrounding source, replacing only payloads.
            value = self._DATA_URI_RE.sub(self.summarize_data_uri, value)
            # parse_markdown_file resolves house/theme/logo/image paths to absolute
            # strings, in addition to the Path-valued source_dir and asset_roots.
            if "\n" not in value and Path(value).is_absolute():
                return self.relative_path(Path(value))
            return value
        if value is None or isinstance(value, bool | int | float):
            return value
        raise TypeError(f"Unsupported snapshot model type: {type(value).__name__}")


def model_json(model: DocumentModel, sample_dir: Path) -> str:
    """Serialize a complete parsed model before rendering can mutate it.

    Args:
        model: Parsed document, including metadata, sections, and elements.
        sample_dir: Original Markdown folder, the base for all absolute paths.

    Returns:
        Indented UTF-8-compatible JSON with stable source/field order.
    """
    return (
        json.dumps(
            ModelSerializer(sample_dir).convert(model),
            ensure_ascii=False,
            indent=2,
            allow_nan=False,
        )
        + "\n"
    )


class XMLNormalizer:
    """Replace only identified volatile values, keeping their nodes/attributes."""

    WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    DCTERMS = "http://purl.org/dc/terms/"
    _XML_DECLARATION_RE = re.compile(rb"\A(?:\xef\xbb\xbf)?(<\?xml\s+[^?]*\?>)")
    _SEMANTIC_BOOKMARK_RE = re.compile(r"^(_ibrep_[A-Z][A-Za-z0-9_]*_)[0-9a-f]{4}$")
    _COPYRIGHT_YEAR_RE = re.compile(r"^\(c\) [0-9]{4}( .+\. All rights reserved\.)$")

    @classmethod
    def normalize(cls, data: bytes, part_name: str, *, copyright_company: str | None = None) -> str:
        """Sort attributes and pretty-print without stripping text whitespace.

        Args:
            data: Original XML bytes from the DOCX archive.
            part_name: Exact archive member name, including its folder.
            copyright_company: Company only when resolved options enable the
                generated disclaimer; None leaves all body text untouched.

        Returns:
            XML text with only the documented volatile values replaced.
        """
        parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True)
        tree = etree.fromstring(data, parser=parser).getroottree()
        root = tree.getroot()
        for element in root.iter():
            if not isinstance(element.tag, str):
                continue
            for key in list(element.attrib):
                name = etree.QName(key)
                # Word revision-session identifiers do not describe document
                # content. Keep each attribute's presence and its location.
                if name.namespace == cls.WORD and name.localname.startswith("rsid"):
                    element.set(key, "REVISION")
            # settings.xml also stores revision IDs as w:val on these nodes.
            if element.tag in (f"{{{cls.WORD}}}rsid", f"{{{cls.WORD}}}rsidRoot"):
                key = f"{{{cls.WORD}}}val"
                if key in element.attrib:
                    element.set(key, "REVISION")
            # Core property timestamps describe creation/saving, not source dates.
            if part_name == "docProps/core.xml" and element.tag in (
                f"{{{cls.DCTERMS}}}created",
                f"{{{cls.DCTERMS}}}modified",
            ):
                if element.text is not None:
                    element.text = "TIMESTAMP"
            if element.tag == f"{{{cls.WORD}}}bookmarkStart":
                key = f"{{{cls.WORD}}}name"
                name = element.get(key, "")
                # _build_semantic_bookmark_name appends uuid4().hex[:4]. Preserve
                # the kind, optional semantic slug, and numeric bookmark ID.
                match = cls._SEMANTIC_BOOKMARK_RE.fullmatch(name)
                if match:
                    element.set(key, match[1] + "RANDOM")
            attributes = sorted(element.attrib.items())
            element.attrib.clear()
            element.attrib.update(attributes)

        # DisclaimerRenderer appends time.strftime('%Y') to the final body
        # paragraph. Only replace that year in that exact generated sentence;
        # dates, years in the body, company, wording, and formatting remain visible.
        # The corpus's ib-report exercises this branch; two different simulated
        # years are checked in test_generated_copyright_year_is_volatile.
        if part_name == "word/document.xml" and copyright_company is not None:
            paragraphs = root.findall(f"{{{cls.WORD}}}body/{{{cls.WORD}}}p")
            if paragraphs:
                texts = paragraphs[-1].findall(f"{{{cls.WORD}}}r/{{{cls.WORD}}}t")
                if len(texts) == 1 and texts[0].text:
                    match = cls._COPYRIGHT_YEAR_RE.fullmatch(texts[0].text)
                    if match and match[1] == f" {copyright_company}. All rights reserved.":
                        texts[0].text = "(c) YEAR" + match[1]
        rendered = etree.tostring(tree, encoding="unicode", pretty_print=True)
        # lxml's Unicode serialization omits the declaration. Preserve it from
        # the original part so version/encoding/standalone changes remain visible.
        declaration = cls._XML_DECLARATION_RE.match(data)
        if declaration:
            rendered = declaration[1].decode(tree.docinfo.encoding) + "\n" + rendered
        return rendered


def xml_parts(document: Path, *, copyright_company: str | None = None) -> dict[str, str]:
    """Read every XML/relationship part, with no part allowlist.

    Args:
        document: Rendered DOCX archive.
        copyright_company: Company when a generated disclaimer was requested.

    Returns:
        Sorted archive member names mapped to readable normalized XML. The keys
        are snapshotted too, so removing or renaming a part cannot go unnoticed.
    """
    with ZipFile(document) as archive:
        return {
            name: XMLNormalizer.normalize(
                archive.read(name), name, copyright_company=copyright_company
            )
            for name in sorted(archive.namelist())
            if name.endswith((".xml", ".rels"))
        }
