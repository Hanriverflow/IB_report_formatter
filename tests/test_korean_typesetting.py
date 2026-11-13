"""Korean typesetting defaults in every profile, read from real conversions.

Word line breaking itself cannot run in CI; these tests pin the document
properties whose effect was measured in Word (docs/verification-korean-typesetting-20261114.md).
"""

import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from document_profiles import RenderOptions
from md_to_word import IBReportConverter
from office_layout import setup_korean_language

ROOT = Path(__file__).resolve().parents[1]
PROFILES = (
    "ib-report",
    "ib-memo",
    "plain",
    "business-report",
    "meeting-minutes",
    "office-letter",
    "term-sheet",
)
W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@pytest.fixture(scope="module")
def converted(tmp_path_factory) -> dict[str, Path]:
    """Convert each profile's repository sample through the normal strict path."""
    folder = tmp_path_factory.mktemp("korean-typesetting")
    outputs = {}
    for profile in PROFILES:
        output = folder / f"{profile}.docx"
        source = ROOT / "samples" / "profiles" / f"{profile}.md"
        IBReportConverter(str(source), str(output), RenderOptions(strict=True)).convert()
        outputs[profile] = output
    return outputs


def _styles(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as package:
        return etree.fromstring(package.read("word/styles.xml"))


@pytest.mark.parametrize("profile", PROFILES)
def test_every_profile_declares_korean_east_asian_language(converted, profile):
    languages = _styles(converted[profile]).findall("w:docDefaults/w:rPrDefault/w:rPr/w:lang", W)
    assert len(languages) == 1
    assert languages[0].get(qn("w:eastAsia")) == "ko-KR"
    # Only the East Asian language changes; Latin text keeps its language.
    assert languages[0].get(qn("w:val")) == "en-US"


@pytest.mark.parametrize("profile", PROFILES)
def test_no_profile_breaks_korean_per_syllable(converted, profile):
    # wordWrap=0 makes Word break Korean words between syllables.
    with zipfile.ZipFile(converted[profile]) as package:
        for name in package.namelist():
            if not name.endswith(".xml"):
                continue
            root = etree.fromstring(package.read(name))
            for element in root.iter(qn("w:wordWrap")):
                assert element.get(qn("w:val")) not in {"0", "false", "off"}, name


@pytest.mark.parametrize("profile", PROFILES)
def test_every_profile_prints_korean_next_to_latin_and_digits_tight(converted, profile):
    # Word's automatic gap would print "2025 년", "300 억원", "SPC 는".
    paragraph_defaults = _styles(converted[profile]).find("w:docDefaults/w:pPrDefault/w:pPr", W)
    for tag in ("autoSpaceDE", "autoSpaceDN"):
        values = [element.get(qn("w:val")) for element in paragraph_defaults.findall(f"w:{tag}", W)]
        assert values == ["0"], tag


def test_korean_language_is_added_when_the_template_has_none():
    doc = Document()
    run_properties = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault"))[0]
    for language in run_properties.findall(qn("w:lang")):
        run_properties.remove(language)
    setup_korean_language(doc)
    languages = run_properties.findall(qn("w:lang"))
    assert [language.get(qn("w:eastAsia")) for language in languages] == ["ko-KR"]
    # Schema order: w:lang follows w:szCs.
    names = [etree.QName(child).localname for child in run_properties]
    assert names.index("lang") > names.index("szCs")
