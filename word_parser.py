"""
Word Parser Module for Word to Markdown Converter
Handles parsing of .docx files into DocumentModel.

Dependencies:
    Required: python-docx
    Optional: Pillow (for image processing)
"""

import re
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Union

from docx import Document
from docx.document import Document as DocxDocument

from md_parser import (
    DocumentModel,
    DocumentMetadata,
    Element,
    ElementType,
    Heading,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    ListItem,
    Blockquote,
    Image,
    TextRun,
)

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# METADATA EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class MetadataExtractor:
    """Extracts document metadata from Word core properties."""

    @staticmethod
    def extract(doc: DocxDocument) -> DocumentMetadata:
        """Extract metadata from document properties.

        Args:
            doc: python-docx Document object

        Returns:
            DocumentMetadata with title, author, dates
        """
        props = doc.core_properties
        metadata = DocumentMetadata()

        if props.title:
            metadata.title = props.title
        if props.author:
            metadata.analyst = props.author
        if props.subject:
            metadata.subtitle = props.subject
        if props.created:
            metadata.extra["date"] = props.created.strftime("%Y-%m-%d")
        if props.category:
            metadata.sector = props.category

        return metadata


# ═══════════════════════════════════════════════════════════════════════════════
# STYLE DETECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class StyleDetector:
    """Detects paragraph styles and formatting from Word documents."""

    # Heading style name patterns (flexible matching)
    HEADING_PATTERNS = [
        (re.compile(r"^Heading\s*1$", re.I), 1),
        (re.compile(r"^Heading\s*2$", re.I), 2),
        (re.compile(r"^Heading\s*3$", re.I), 3),
        (re.compile(r"^Heading\s*4$", re.I), 4),
        (re.compile(r"^제목\s*1$", re.I), 1),  # Korean
        (re.compile(r"^제목\s*2$", re.I), 2),
        (re.compile(r"^제목\s*3$", re.I), 3),
    ]

    LIST_BULLET_PATTERNS = [
        re.compile(r"^List\s*Bullet", re.I),
        re.compile(r"^IB\s*Bullet", re.I),
    ]

    LIST_NUMBER_PATTERNS = [
        re.compile(r"^List\s*Number", re.I),
        re.compile(r"^List\s*Paragraph", re.I),
    ]

    @classmethod
    def detect_heading_level(cls, para) -> Optional[int]:
        """Detect if paragraph is a heading and return level (1-4)."""
        style_name = para.style.name
        for pattern, level in cls.HEADING_PATTERNS:
            if pattern.match(style_name):
                return level

        # Fallback to heuristic
        if cls._is_heading_by_formatting(para):
            return 2  # Default heuristic heading to level 2

        return None

    @classmethod
    def detect_list_type(cls, para) -> Optional[str]:
        """Detect if paragraph is a list. Returns 'bullet' or 'number'."""
        style_name = para.style.name

        for pattern in cls.LIST_BULLET_PATTERNS:
            if pattern.match(style_name):
                return "bullet"

        for pattern in cls.LIST_NUMBER_PATTERNS:
            if pattern.match(style_name):
                return "number"

        return None

    @classmethod
    def _is_heading_by_formatting(cls, para) -> bool:
        """Heuristic: all-bold short text might be a heading."""
        text = para.text.strip()
        if not text or len(text) > 80:
            return False

        if not para.runs:
            return False

        # Check if all runs are bold
        all_bold = True
        for run in para.runs:
            if run.text.strip() and not run.bold:
                all_bold = False
                break

        if all_bold and not text.endswith("."):
            return True

        return False


# ═══════════════════════════════════════════════════════════════════════════════
# RUN EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class RunExtractor:
    """Extracts TextRun objects from Word paragraph runs."""

    @staticmethod
    def extract_runs(para) -> List[TextRun]:
        """Extract text runs with formatting.

        Returns list of TextRun with bold, italic, superscript flags.
        """
        runs: List[TextRun] = []
        for r in para.runs:
            if not r.text:
                continue

            runs.append(
                TextRun(
                    text=r.text,
                    bold=bool(r.bold),
                    italic=bool(r.italic),
                    superscript=bool(r.font.superscript),
                )
            )
        return runs


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class TableExtractor:
    """Extracts Table objects from Word tables."""

    # IB-style Navy color for header detection (hex)
    NAVY_HEX = "003366"

    @classmethod
    def extract(cls, word_table) -> Table:
        """Convert a Word table to our Table model."""
        table = Table()
        table.col_count = len(word_table.columns)

        for row_idx, word_row in enumerate(word_table.rows):
            is_header = row_idx == 0
            table_row = TableRow(is_header=is_header)

            for cell in word_row.cells:
                cell_text = "\n".join(p.text for p in cell.paragraphs).strip()

                table_cell = TableCell(
                    content=cell_text,
                    is_header=is_header,
                    runs=RunExtractor.extract_runs(cell.paragraphs[0]) if cell.paragraphs else [],
                )

                # Detect numeric content
                table_cell.is_numeric = any(c.isdigit() for c in cell_text)

                # Detect negative numbers
                table_cell.is_negative = (
                    cell_text.startswith("(")
                    and cell_text.endswith(")")
                    and any(c.isdigit() for c in cell_text)
                ) or (cell_text.startswith("-") and any(c.isdigit() for c in cell_text))

                table_row.cells.append(table_cell)

            table.rows.append(table_row)

        return table

    @classmethod
    def is_header_row_navy(cls, word_table) -> bool:
        """Check if first row has Navy background (IB-style header)."""
        try:
            if not word_table.rows:
                return False
            first_row = word_table.rows[0]
            if not first_row.cells:
                return False
            cell = first_row.cells[0]
            # Check shading
            shd = cell._tc.get_or_add_tcPr().find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"
            )
            if shd is not None:
                fill = shd.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", ""
                )
                return fill.upper() == cls.NAVY_HEX.upper()
        except Exception:
            pass
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# CALLOUT DETECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class CalloutDetector:
    """Detects callout boxes (rendered as single-cell tables with styling)."""

    # Background color to callout title mapping
    CALLOUT_COLORS = {
        "003366": "요약",  # Navy - Executive Summary
        "E6F0FA": "시사점",  # Accent Blue - Key Insight
        "FFF3CD": "주의",  # Light Yellow - Warning
        "F5F5F5": "참고",  # Light Gray - Note
    }

    @classmethod
    def detect_callout(cls, word_table) -> Optional[Blockquote]:
        """Check if a 1x1 table is actually a callout box."""
        if len(word_table.rows) != 1 or len(word_table.columns) != 1:
            return None

        cell = word_table.rows[0].cells[0]
        text = "\n".join(p.text for p in cell.paragraphs).strip()

        if not text:
            return None

        # Try to detect by background color
        title = cls._detect_callout_by_color(cell)

        # Fallback: detect by content keywords
        if not title:
            title = cls._detect_callout_by_content(text)

        if title:
            # Remove title prefix from content if present
            content = text
            for keyword in [
                "요약",
                "시사점",
                "주의",
                "참고",
                "SUMMARY",
                "KEY INSIGHT",
                "WARNING",
                "NOTE",
            ]:
                if content.upper().startswith(keyword.upper()):
                    content = content[len(keyword) :].lstrip(":").lstrip()
                    break

            return Blockquote(title=title, text=content)

        return None

    @classmethod
    def _detect_callout_by_color(cls, cell) -> Optional[str]:
        """Detect callout type by cell background color."""
        try:
            shd = cell._tc.get_or_add_tcPr().find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"
            )
            if shd is not None:
                fill = shd.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", ""
                )
                return cls.CALLOUT_COLORS.get(fill.upper())
        except Exception:
            pass
        return None

    @classmethod
    def _detect_callout_by_content(cls, text: str) -> Optional[str]:
        """Detect callout type by content keywords."""
        text_upper = text.upper()
        if any(kw in text_upper for kw in ["EXECUTIVE SUMMARY", "요약", "핵심"]):
            return "요약"
        if any(kw in text_upper for kw in ["KEY INSIGHT", "시사점", "결론"]):
            return "시사점"
        if any(kw in text_upper for kw in ["WARNING", "주의", "RISK"]):
            return "주의"
        if any(kw in text_upper for kw in ["NOTE", "참고"]):
            return "참고"
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# IMAGE EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class ImageExtractor:
    """Extracts images from Word documents."""

    def __init__(self, output_dir: Optional[Path] = None, embed_base64: bool = False):
        self.output_dir = output_dir
        self.embed_base64 = embed_base64
        self._image_counter = 0
        self._extracted_images: Dict[str, Path] = {}

    def extract_all(self, doc: DocxDocument) -> Dict[str, Path]:
        """Extract all images from document to output directory."""
        if not self.output_dir:
            return {}

        self.output_dir.mkdir(parents=True, exist_ok=True)

        import base64

        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                try:
                    image_part = rel.target_part
                    image_bytes = image_part.blob

                    # Generate filename
                    self._image_counter += 1
                    ext = Path(rel.target_ref).suffix or ".png"
                    filename = f"image_{self._image_counter}{ext}"
                    output_path = self.output_dir / filename

                    # Write to file
                    output_path.write_bytes(image_bytes)
                    self._extracted_images[rel.target_ref] = output_path

                    logger.debug("Extracted image: %s", filename)
                except Exception as e:
                    logger.warning("Failed to extract image %s: %s", rel.target_ref, e)

        return self._extracted_images

    def get_image_for_rel(self, rel_id: str) -> Optional[Image]:
        """Get Image object for a relationship ID."""
        if rel_id in self._extracted_images:
            path = self._extracted_images[rel_id]
            return Image(
                alt_text=f"Figure {self._image_counter}",
                path=str(path.name),
            )
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# WORD PARSER
# ═══════════════════════════════════════════════════════════════════════════════


class WordParser:
    """Main parser for Word documents."""

    def __init__(
        self,
        extract_images: bool = True,
        image_output_dir: Optional[Path] = None,
    ):
        self.extract_images = extract_images
        self.image_output_dir = image_output_dir
        self.image_extractor: Optional[ImageExtractor] = None

    def parse(self, file_path: str) -> DocumentModel:
        """Parse a Word document into DocumentModel."""
        doc = Document(file_path)

        # Extract metadata
        metadata = MetadataExtractor.extract(doc)

        # If no title from properties, try first heading
        if metadata.title == "IB Report":
            first_heading = self._find_first_heading(doc)
            if first_heading:
                metadata.title = first_heading

        # Extract images if requested
        if self.extract_images and self.image_output_dir:
            self.image_extractor = ImageExtractor(self.image_output_dir)
            self.image_extractor.extract_all(doc)

        # Parse elements
        elements = self._parse_elements(doc)

        return DocumentModel(
            metadata=metadata,
            elements=elements,
        )

    def _find_first_heading(self, doc: DocxDocument) -> Optional[str]:
        """Find the first heading in the document."""
        for para in doc.paragraphs:
            level = StyleDetector.detect_heading_level(para)
            if level == 1:
                return para.text.strip()
        return None

    def _parse_elements(self, doc: DocxDocument) -> List[Element]:
        """Parse all document elements."""
        elements: List[Element] = []

        # Track which tables we've processed (for callout detection)
        processed_tables = set()

        for para in doc.paragraphs:
            text = para.text.strip()

            if not text:
                continue

            element = self._parse_paragraph(para)
            if element:
                elements.append(element)

        # Process tables
        for table in doc.tables:
            element = self._parse_table(table)
            if element:
                elements.append(element)

        return elements

    def _parse_paragraph(self, para) -> Optional[Element]:
        """Parse a single paragraph."""
        text = para.text.strip()

        if not text:
            return None

        # Check for heading
        level = StyleDetector.detect_heading_level(para)
        if level:
            return Element(
                element_type=getattr(ElementType, f"HEADING_{level}", ElementType.HEADING_2),
                content=Heading(level=level, text=text),
                raw_text=text,
            )

        # Check for list
        list_type = StyleDetector.detect_list_type(para)
        if list_type == "bullet":
            return Element(
                element_type=ElementType.BULLET_LIST,
                content=ListItem(text=text, runs=RunExtractor.extract_runs(para)),
                raw_text=text,
            )
        if list_type == "number":
            import re

            match = re.match(r"^(\d+)\.\s*(.+)$", text)
            if match:
                number, content = match.groups()
                return Element(
                    element_type=ElementType.NUMBERED_LIST,
                    content=(number, ListItem(text=content, runs=RunExtractor.extract_runs(para))),
                    raw_text=text,
                )
            return Element(
                element_type=ElementType.NUMBERED_LIST,
                content=("1", ListItem(text=text, runs=RunExtractor.extract_runs(para))),
                raw_text=text,
            )

        # Default: paragraph
        return Element(
            element_type=ElementType.PARAGRAPH,
            content=Paragraph(text=text, runs=RunExtractor.extract_runs(para)),
            raw_text=text,
        )

    def _parse_table(self, word_table) -> Optional[Element]:
        """Parse a Word table."""
        # Check if it's a callout box first
        callout = CalloutDetector.detect_callout(word_table)
        if callout:
            return Element(
                element_type=ElementType.BLOCKQUOTE,
                content=callout,
                raw_text=callout.text,
            )

        # Regular table
        table = TableExtractor.extract(word_table)
        return Element(
            element_type=ElementType.TABLE,
            content=table,
            raw_text="[TABLE]",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════


def parse_word_file(
    file_path: str,
    extract_images: bool = True,
    image_output_dir: Optional[str] = None,
) -> DocumentModel:
    """
    Parse a Word document into a DocumentModel.

    Args:
        file_path: Path to the .docx file
        extract_images: Whether to extract embedded images
        image_output_dir: Directory to save extracted images

    Returns:
        DocumentModel containing parsed elements
    """
    parser = WordParser(
        extract_images=extract_images,
        image_output_dir=Path(image_output_dir) if image_output_dir else None,
    )
    return parser.parse(file_path)
