"""Tests for OMML → LaTeX conversion module."""

import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from omml_latex import (
    OMML_NS,
    OmmlToLatex,
    W_NS,
    _convert_omml_in_xml,
    count_equations_in_docx,
    extract_latex_from_paragraph,
    pre_process_docx_math,
)


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════


def _make_omath_xml(inner: str) -> ET.Element:
    """Wrap OMML XML fragments in a proper root with namespaces."""
    xml = (
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        f"{inner}"
        "</m:oMath>"
    )
    return ET.fromstring(xml)


def _make_docx_bytes(document_xml: str) -> BytesIO:
    """Create a minimal .docx (ZIP) with the given word/document.xml content."""
    buf = BytesIO()
    full_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        "<w:body>"
        f"{document_xml}"
        "</w:body>"
        "</w:document>"
    )
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("word/document.xml", full_xml)
        zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>')
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# OmmlToLatex UNIT TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestOmmlToLatex:
    """Test core OMML → LaTeX element conversion."""

    def test_simple_variable(self):
        """A single <m:r> with text 'x' should return 'x'."""
        omath = _make_omath_xml(
            '<m:r><m:t>x</m:t></m:r>'
        )
        result = OmmlToLatex(omath)
        assert result.latex == "x"

    def test_fraction(self):
        """<m:f> should produce \\frac{a}{b}."""
        omath = _make_omath_xml(
            "<m:f>"
            "<m:fPr><m:type m:val='bar'/></m:fPr>"
            "<m:num><m:r><m:t>a</m:t></m:r></m:num>"
            "<m:den><m:r><m:t>b</m:t></m:r></m:den>"
            "</m:f>"
        )
        result = OmmlToLatex(omath)
        assert "\\frac{a}{b}" in result.latex

    def test_superscript(self):
        """<m:sSup> should produce x^{2}."""
        omath = _make_omath_xml(
            "<m:sSup>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "<m:sup><m:r><m:t>2</m:t></m:r></m:sup>"
            "</m:sSup>"
        )
        result = OmmlToLatex(omath)
        assert "x" in result.latex
        assert "^{2}" in result.latex

    def test_subscript(self):
        """<m:sSub> should produce x_{i}."""
        omath = _make_omath_xml(
            "<m:sSub>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "<m:sub><m:r><m:t>i</m:t></m:r></m:sub>"
            "</m:sSub>"
        )
        result = OmmlToLatex(omath)
        assert "x" in result.latex
        assert "_{i}" in result.latex

    def test_square_root(self):
        """<m:rad> without degree should produce \\sqrt{x}."""
        omath = _make_omath_xml(
            "<m:rad>"
            "<m:radPr/>"
            "<m:deg/>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "</m:rad>"
        )
        result = OmmlToLatex(omath)
        assert "\\sqrt{x}" in result.latex

    def test_nth_root(self):
        """<m:rad> with degree should produce \\sqrt[3]{x}."""
        omath = _make_omath_xml(
            "<m:rad>"
            "<m:radPr/>"
            "<m:deg><m:r><m:t>3</m:t></m:r></m:deg>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "</m:rad>"
        )
        result = OmmlToLatex(omath)
        assert "\\sqrt[3]{x}" in result.latex

    def test_parentheses_delimiter(self):
        """<m:d> with default delimiters should produce \\left(...)\\right)."""
        omath = _make_omath_xml(
            "<m:d>"
            "<m:dPr/>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "</m:d>"
        )
        result = OmmlToLatex(omath)
        assert "\\left" in result.latex
        assert "\\right" in result.latex

    def test_summation_nary(self):
        """<m:nary> with sum symbol should produce \\sum."""
        omath = _make_omath_xml(
            "<m:nary>"
            "<m:naryPr><m:chr m:val='\u2211'/></m:naryPr>"
            "<m:sub><m:r><m:t>i=0</m:t></m:r></m:sub>"
            "<m:sup><m:r><m:t>n</m:t></m:r></m:sup>"
            "<m:e><m:r><m:t>i</m:t></m:r></m:e>"
            "</m:nary>"
        )
        result = OmmlToLatex(omath)
        assert "\\sum" in result.latex

    def test_matrix(self):
        """<m:m> should produce \\begin{matrix}...\\end{matrix}."""
        omath = _make_omath_xml(
            "<m:m>"
            "<m:mPr/>"
            "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
            "<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr>"
            "</m:m>"
        )
        result = OmmlToLatex(omath)
        assert "\\begin{matrix}" in result.latex
        assert "\\end{matrix}" in result.latex

    def test_str_method(self):
        """__str__ should return the same as .latex."""
        omath = _make_omath_xml('<m:r><m:t>y</m:t></m:r>')
        conv = OmmlToLatex(omath)
        assert str(conv) == conv.latex


# ═══════════════════════════════════════════════════════════════════════════════
# PARAGRAPH-LEVEL EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════


class TestExtractLatexFromParagraph:
    """Test extraction of LaTeX from paragraph XML elements."""

    def test_block_equation(self):
        """oMathPara should be detected as block."""
        xml = (
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:oMathPara><m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath></m:oMathPara>"
            "</w:p>"
        )
        para = ET.fromstring(xml)
        results = extract_latex_from_paragraph(para)
        assert len(results) >= 1
        latex, is_block = results[0]
        assert is_block is True
        assert "E" in latex

    def test_inline_equation(self):
        """Standalone oMath should be detected as inline."""
        xml = (
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"
            "</w:p>"
        )
        para = ET.fromstring(xml)
        results = extract_latex_from_paragraph(para)
        assert len(results) >= 1
        _, is_block = results[0]
        assert is_block is False

    def test_no_equations(self):
        """Paragraph without math should return empty list."""
        xml = (
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:r><w:t>Hello world</w:t></w:r>"
            "</w:p>"
        )
        para = ET.fromstring(xml)
        results = extract_latex_from_paragraph(para)
        assert results == []


# ═══════════════════════════════════════════════════════════════════════════════
# DOCX PRE-PROCESSOR
# ═══════════════════════════════════════════════════════════════════════════════


class TestPreProcessDocxMath:
    """Test ZIP-level DOCX pre-processing."""

    def test_block_equation_replaced(self):
        """Block equation should be converted to $$...$$ in the XML."""
        doc_xml = (
            "<w:p>"
            "<m:oMathPara><m:oMath>"
            '<m:r><m:t xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">x</m:t></m:r>'
            "</m:oMath></m:oMathPara>"
            "</w:p>"
        )
        docx_stream = _make_docx_bytes(doc_xml)
        result_stream = pre_process_docx_math(docx_stream)

        # Read back the processed XML
        with zipfile.ZipFile(result_stream) as zf:
            processed_xml = zf.read("word/document.xml").decode("utf-8")

        assert "$$" in processed_xml

    def test_inline_equation_replaced(self):
        """Inline equation should be converted to $...$ in the XML."""
        doc_xml = (
            "<w:p>"
            "<w:r><w:t>The value is </w:t></w:r>"
            "<m:oMath>"
            '<m:r><m:t xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">y</m:t></m:r>'
            "</m:oMath>"
            "<w:r><w:t> here.</w:t></w:r>"
            "</w:p>"
        )
        docx_stream = _make_docx_bytes(doc_xml)
        result_stream = pre_process_docx_math(docx_stream)

        with zipfile.ZipFile(result_stream) as zf:
            processed_xml = zf.read("word/document.xml").decode("utf-8")

        # Should have single $ delimiters for inline
        assert "$y$" in processed_xml

    def test_no_equations_passthrough(self):
        """DOCX without equations should pass through unchanged."""
        doc_xml = "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>"
        docx_stream = _make_docx_bytes(doc_xml)
        result_stream = pre_process_docx_math(docx_stream)

        with zipfile.ZipFile(result_stream) as zf:
            processed_xml = zf.read("word/document.xml").decode("utf-8")

        assert "Hello world" in processed_xml

    def test_non_math_files_preserved(self):
        """Non-math XML files in the DOCX should be preserved."""
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>')
            zf.writestr("word/styles.xml", "<styles>custom</styles>")
            zf.writestr("[Content_Types].xml", "<Types/>")
        buf.seek(0)

        result = pre_process_docx_math(buf)
        with zipfile.ZipFile(result) as zf:
            assert zf.read("word/styles.xml") == b"<styles>custom</styles>"


# ═══════════════════════════════════════════════════════════════════════════════
# XML CONVERSION TESTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestConvertOmmlInXml:
    """Test the XML-level OMML conversion."""

    def test_fraction_conversion(self):
        """A fraction in OMML should become \\frac in the output."""
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<w:body><w:p>"
            "<m:oMathPara><m:oMath>"
            "<m:f><m:fPr><m:type m:val='bar'/></m:fPr>"
            "<m:num><m:r><m:t>a</m:t></m:r></m:num>"
            "<m:den><m:r><m:t>b</m:t></m:r></m:den>"
            "</m:f>"
            "</m:oMath></m:oMathPara>"
            "</w:p></w:body></w:document>"
        )
        result = _convert_omml_in_xml(xml.encode("utf-8"))
        text = result.decode("utf-8")
        assert "\\frac{a}{b}" in text
        assert "$$" in text


# ═══════════════════════════════════════════════════════════════════════════════
# EQUATION COUNTER
# ═══════════════════════════════════════════════════════════════════════════════


class TestCountEquations:
    """Test equation counting in DOCX files."""

    def test_count_block_and_inline(self, tmp_path):
        """Count block and inline equations."""
        doc_xml = (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<w:body>"
            "<w:p><m:oMathPara><m:oMath><m:r><m:t>block</m:t></m:r></m:oMath></m:oMathPara></w:p>"
            "<w:p><m:oMath><m:r><m:t>inline</m:t></m:r></m:oMath></w:p>"
            "</w:body></w:document>"
        )
        docx_path = tmp_path / "test.docx"
        with zipfile.ZipFile(docx_path, "w") as zf:
            zf.writestr("word/document.xml", doc_xml)

        counts = count_equations_in_docx(str(docx_path))
        assert counts["block"] == 1
        assert counts["inline"] == 1

    def test_no_equations(self, tmp_path):
        """DOCX without equations should return zeros."""
        doc_xml = (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>"
        )
        docx_path = tmp_path / "test.docx"
        with zipfile.ZipFile(docx_path, "w") as zf:
            zf.writestr("word/document.xml", doc_xml)

        counts = count_equations_in_docx(str(docx_path))
        assert counts["block"] == 0
        assert counts["inline"] == 0
