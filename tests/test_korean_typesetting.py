"""Korean typesetting defaults in every profile, read from real conversions.

Word line breaking itself cannot run in CI; these tests pin the document
properties whose effect was measured in Word (docs/verification-korean-typesetting-20261114.md).
"""

import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
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


def _body_alignment(path: Path) -> list[str | None]:
    for style in _styles(path).findall("w:style", W):
        name = style.find("w:name", W)
        if name is not None and name.get(qn("w:val")) == "IB Body":
            return [jc.get(qn("w:val")) for jc in style.findall("w:pPr/w:jc", W)]
    raise AssertionError("IB Body style missing")


@pytest.mark.parametrize("profile", ["ib-report", "ib-memo"])
def test_ib_body_is_left_aligned(converted, profile):
    # Justified Korean stretched word gaps to up to 4.6 times a space.
    assert _body_alignment(converted[profile]) == ["left"]


def test_ib_theme_can_still_justify_the_body(tmp_path):
    theme = tmp_path / "justified.yaml"
    theme.write_text("BODY_JUSTIFY: true\n", encoding="utf-8")
    output = tmp_path / "justified.docx"
    source = ROOT / "samples" / "profiles" / "ib-memo.md"
    IBReportConverter(
        str(source), str(output), RenderOptions(strict=True, theme=str(theme))
    ).convert()
    assert _body_alignment(output) == ["both"]


def test_ib_cover_disclaimer_stays_justified(converted):
    # Only the body style changes; the disclaimer keeps its direct justification.
    document = Document(str(converted["ib-report"]))
    justified = [
        paragraph
        for table in document.tables
        for cell in table._cells
        for paragraph in cell.paragraphs
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    ]
    assert justified
