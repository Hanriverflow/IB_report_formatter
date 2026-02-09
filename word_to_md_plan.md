# Word to Markdown Converter - Implementation Plan

> Word 문서를 LLM 활용 가능한 Markdown으로 변환

## 1. 개요

### 1.1 목적
- Word 문서(.docx)를 구조화된 Markdown으로 변환
- LLM에서 쉽게 활용할 수 있는 clean한 텍스트 출력
- 기존 `md_to_word.py` 파이프라인의 역방향 구현

### 1.2 주요 Use Cases
1. **LLM 컨텍스트 입력**: Word 보고서를 LLM에 제공하기 위한 텍스트 변환
2. **문서 분석**: Word 문서 내용을 프로그래매틱하게 처리
3. **Round-trip 변환**: Word → MD → (편집) → Word

---

## 2. 아키텍처

### 2.1 Pipeline 구조

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         word_to_md.py (CLI)                             │
│                     Main entry point, orchestration                      │
└────────────────────────────────┬────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                        word_parser.py                                    │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────────┐   │
│  │ DocumentReader   │  │ ElementExtractor │  │ TableExtractor       │   │
│  │ - Open .docx     │  │ - Headings       │  │ - Parse Word tables  │   │
│  │ - Read styles    │  │ - Paragraphs     │  │ - Detect alignment   │   │
│  │ - Encoding       │  │ - Lists          │  │ - Extract cells      │   │
│  └──────────────────┘  └──────────────────┘  └──────────────────────┘   │
│                                                                          │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────────┐   │
│  │ ImageExtractor   │  │ StyleAnalyzer    │  │ MetadataExtractor    │   │
│  │ - Embedded imgs  │  │ - Detect bold    │  │ - Title from styles  │   │
│  │ - Save as files  │  │ - Detect italic  │  │ - Author, date       │   │
│  │ - or Base64      │  │ - Font colors    │  │ - Custom properties  │   │
│  └──────────────────┘  └──────────────────┘  └──────────────────────┘   │
└────────────────────────────────┬────────────────────────────────────────┘
                                 │
                                 ▼
                      ┌──────────────────────┐
                      │   DocumentModel      │
                      │   (from md_parser)   │
                      │   - Reuse existing   │
                      │     data structures  │
                      └──────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                        md_renderer.py                                    │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────────┐   │
│  │ HeadingRenderer  │  │ TableRenderer    │  │ ImageRenderer        │   │
│  │ - # / ## / ###   │  │ - | col | col |  │  │ - ![alt](path)       │   │
│  │ - Numbered       │  │ - Alignment      │  │ - Base64 embed       │   │
│  └──────────────────┘  └──────────────────┘  └──────────────────────┘   │
│                                                                          │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────────┐   │
│  │ ListRenderer     │  │ CalloutRenderer  │  │ FrontmatterRenderer  │   │
│  │ - Bullet: -      │  │ - > [TYPE]       │  │ - YAML metadata      │   │
│  │ - Numbered: 1.   │  │ - > content      │  │ - --- block          │   │
│  └──────────────────┘  └──────────────────┘  └──────────────────────┘   │
└────────────────────────────────┬────────────────────────────────────────┘
                                 │
                                 ▼
                        ┌────────────────┐
                        │  output.md     │
                        │  (Markdown)    │
                        └────────────────┘
```

### 2.2 모듈 구조

```
IB_report_formatter/
├── word_to_md.py      # CLI entry point (NEW)
├── word_parser.py     # Word document parsing (NEW)
├── md_renderer.py     # Markdown rendering (NEW)
├── md_parser.py       # Existing - reuse data models
├── md_to_word.py      # Existing
├── ib_renderer.py     # Existing
└── md_formatter.py    # Existing
```

---

## 3. 상세 설계

### 3.1 word_parser.py

```python
"""
Word Parser Module for Word to Markdown Converter
Handles parsing of .docx files into DocumentModel.

Dependencies:
    Required: python-docx
    Optional: Pillow (for image processing)
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Union
from pathlib import Path
import re
import logging

from docx import Document
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable

# Reuse existing data models
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
# STYLE DETECTION
# ═══════════════════════════════════════════════════════════════════════════════


class StyleDetector:
    """Detects paragraph styles and formatting from Word documents."""
    
    # Heading style name patterns
    HEADING_PATTERNS = [
        (re.compile(r"^Heading\s*1$", re.I), 1),
        (re.compile(r"^Heading\s*2$", re.I), 2),
        (re.compile(r"^Heading\s*3$", re.I), 3),
        (re.compile(r"^Heading\s*4$", re.I), 4),
        (re.compile(r"^제목\s*1$", re.I), 1),  # Korean
        (re.compile(r"^제목\s*2$", re.I), 2),
        (re.compile(r"^제목\s*3$", re.I), 3),
    ]
    
    # List style patterns
    LIST_BULLET_PATTERNS = [
        re.compile(r"^List\s*Bullet", re.I),
        re.compile(r"^IB\s*Bullet", re.I),
    ]
    
    LIST_NUMBER_PATTERNS = [
        re.compile(r"^List\s*Number", re.I),
        re.compile(r"^List\s*Paragraph", re.I),
    ]
    
    @classmethod
    def detect_heading_level(cls, para: DocxParagraph) -> Optional[int]:
        """Detect if paragraph is a heading and return its level."""
        style_name = para.style.name if para.style else ""
        
        for pattern, level in cls.HEADING_PATTERNS:
            if pattern.match(style_name):
                return level
        
        # Fallback: check for bold + large font as heading indicator
        if cls._is_heading_by_formatting(para):
            return 2  # Default to level 2 for formatted headings
        
        return None
    
    @classmethod
    def _is_heading_by_formatting(cls, para: DocxParagraph) -> bool:
        """Heuristic: bold text with larger font might be a heading."""
        if not para.runs:
            return False
        
        all_bold = all(run.bold for run in para.runs if run.text.strip())
        text = para.text.strip()
        
        # Short, all-bold text is likely a heading
        return all_bold and 0 < len(text) <= 100
    
    @classmethod
    def detect_list_type(cls, para: DocxParagraph) -> Optional[str]:
        """Detect if paragraph is a list item. Returns 'bullet' or 'number'."""
        style_name = para.style.name if para.style else ""
        
        for pattern in cls.LIST_BULLET_PATTERNS:
            if pattern.match(style_name):
                return "bullet"
        
        for pattern in cls.LIST_NUMBER_PATTERNS:
            if pattern.match(style_name):
                return "number"
        
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# RUN EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class RunExtractor:
    """Extracts TextRun objects from Word paragraph runs."""
    
    @staticmethod
    def extract_runs(para: DocxParagraph) -> List[TextRun]:
        """Extract text runs with formatting from a paragraph."""
        runs = []
        
        for run in para.runs:
            text = run.text
            if not text:
                continue
            
            runs.append(TextRun(
                text=text,
                bold=run.bold or False,
                italic=run.italic or False,
                superscript=run.font.superscript or False,
            ))
        
        return runs


# ═══════════════════════════════════════════════════════════════════════════════
# TABLE EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class TableExtractor:
    """Extracts Table objects from Word tables."""
    
    @staticmethod
    def extract(word_table: DocxTable) -> Table:
        """Convert a Word table to our Table model."""
        table = Table()
        table.col_count = len(word_table.columns)
        
        for row_idx, word_row in enumerate(word_table.rows):
            is_header = row_idx == 0
            table_row = TableRow(is_header=is_header)
            
            for cell in word_row.cells:
                # Combine all paragraphs in cell
                cell_text = "\n".join(p.text for p in cell.paragraphs).strip()
                
                table_cell = TableCell(
                    content=cell_text,
                    is_header=is_header,
                    runs=RunExtractor.extract_runs(cell.paragraphs[0]) if cell.paragraphs else [],
                )
                
                # Detect numeric content
                table_cell.is_numeric = any(c.isdigit() for c in cell_text)
                
                table_row.cells.append(table_cell)
            
            table.rows.append(table_row)
        
        return table


# ═══════════════════════════════════════════════════════════════════════════════
# IMAGE EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class ImageExtractor:
    """Extracts images from Word documents."""
    
    def __init__(self, output_dir: Optional[Path] = None, embed_base64: bool = False):
        self.output_dir = output_dir
        self.embed_base64 = embed_base64
        self._image_counter = 0
    
    def extract_from_paragraph(self, para: DocxParagraph) -> Optional[Image]:
        """Extract image from a paragraph if it contains one."""
        # Check for inline shapes (embedded images)
        for run in para.runs:
            if hasattr(run, '_r') and run._r.xml:
                if 'drawing' in run._r.xml or 'pict' in run._r.xml:
                    return self._extract_image(run)
        return None
    
    def _extract_image(self, run) -> Optional[Image]:
        """Extract image data from a run."""
        # Implementation depends on python-docx internals
        # This is a simplified version
        self._image_counter += 1
        
        return Image(
            alt_text=f"Figure {self._image_counter}",
            path=f"images/figure_{self._image_counter}.png",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# CALLOUT DETECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class CalloutDetector:
    """Detects callout boxes (rendered as single-cell tables with styling)."""
    
    # Keywords that indicate callout types
    CALLOUT_KEYWORDS = {
        "EXECUTIVE SUMMARY": "EXECUTIVE SUMMARY",
        "요약": "요약",
        "KEY INSIGHT": "KEY INSIGHT",
        "시사점": "시사점",
        "WARNING": "WARNING",
        "주의": "주의",
        "NOTE": "NOTE",
        "참고": "참고",
    }
    
    @classmethod
    def detect_callout(cls, word_table: DocxTable) -> Optional[Blockquote]:
        """Check if a 1x1 table is actually a callout box."""
        if len(word_table.rows) != 1 or len(word_table.columns) != 1:
            return None
        
        cell = word_table.rows[0].cells[0]
        text = "\n".join(p.text for p in cell.paragraphs).strip()
        
        # Check for callout keywords
        for keyword, title in cls.CALLOUT_KEYWORDS.items():
            if text.upper().startswith(keyword) or keyword in text.upper():
                # Extract content after the keyword
                content = text
                for kw in cls.CALLOUT_KEYWORDS.keys():
                    content = re.sub(rf"^{re.escape(kw)}[:\s]*", "", content, flags=re.I)
                
                return Blockquote(
                    title=title,
                    text=content.strip(),
                )
        
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# METADATA EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class MetadataExtractor:
    """Extracts document metadata from Word core properties."""
    
    @staticmethod
    def extract(doc: DocxDocument) -> DocumentMetadata:
        """Extract metadata from document properties."""
        props = doc.core_properties
        
        metadata = DocumentMetadata()
        
        if props.title:
            metadata.title = props.title
        if props.author:
            metadata.analyst = props.author
        if props.subject:
            metadata.subtitle = props.subject
        
        # Store additional properties in extra
        if props.created:
            metadata.extra["date"] = props.created.strftime("%Y-%m-%d")
        if props.category:
            metadata.sector = props.category
        
        return metadata


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN PARSER
# ═══════════════════════════════════════════════════════════════════════════════


class WordParser:
    """Main parser for Word documents."""
    
    def __init__(
        self,
        extract_images: bool = True,
        image_output_dir: Optional[Path] = None,
        embed_images_base64: bool = False,
    ):
        self.extract_images = extract_images
        self.image_output_dir = image_output_dir
        self.embed_images_base64 = embed_images_base64
    
    def parse(self, file_path: str) -> DocumentModel:
        """Parse a Word document into DocumentModel."""
        doc = Document(file_path)
        
        # Extract metadata
        metadata = MetadataExtractor.extract(doc)
        
        # If no title from properties, try to extract from first heading
        if metadata.title == "IB Report":
            first_heading = self._find_first_heading(doc)
            if first_heading:
                metadata.title = first_heading
        
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
        elements = []
        
        # Iterate through document body
        for block in doc.element.body:
            if block.tag.endswith('p'):  # Paragraph
                para = self._find_paragraph_by_element(doc, block)
                if para:
                    element = self._parse_paragraph(para)
                    if element:
                        elements.append(element)
            
            elif block.tag.endswith('tbl'):  # Table
                table = self._find_table_by_element(doc, block)
                if table:
                    element = self._parse_table(table)
                    if element:
                        elements.append(element)
        
        return elements
    
    def _find_paragraph_by_element(self, doc: DocxDocument, element) -> Optional[DocxParagraph]:
        """Find paragraph object by XML element."""
        for para in doc.paragraphs:
            if para._element is element:
                return para
        return None
    
    def _find_table_by_element(self, doc: DocxDocument, element) -> Optional[DocxTable]:
        """Find table object by XML element."""
        for table in doc.tables:
            if table._element is element:
                return table
        return None
    
    def _parse_paragraph(self, para: DocxParagraph) -> Optional[Element]:
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
            # Extract number if present
            match = re.match(r"^(\d+)\.\s*(.+)$", text)
            if match:
                number, content = match.groups()
                return Element(
                    element_type=ElementType.NUMBERED_LIST,
                    content=(number, ListItem(text=content, runs=RunExtractor.extract_runs(para))),
                    raw_text=text,
                )
        
        # Default: paragraph
        return Element(
            element_type=ElementType.PARAGRAPH,
            content=Paragraph(text=text, runs=RunExtractor.extract_runs(para)),
            raw_text=text,
        )
    
    def _parse_table(self, word_table: DocxTable) -> Optional[Element]:
        """Parse a Word table."""
        # Check if it's actually a callout box
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
```

### 3.2 md_renderer.py

```python
"""
Markdown Renderer Module for Word to Markdown Converter
Handles rendering of DocumentModel to Markdown text.

Renders clean, LLM-friendly Markdown output.
"""

from dataclasses import dataclass
from typing import List, Optional
import re
import logging

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
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class RenderConfig:
    """Configuration for Markdown rendering."""
    
    # Frontmatter
    include_frontmatter: bool = True
    
    # Images
    embed_images_base64: bool = False
    image_path_prefix: str = ""
    
    # Tables
    table_alignment: bool = True  # Include alignment markers
    
    # Formatting
    bold_marker: str = "**"
    italic_marker: str = "*"
    heading_style: str = "atx"  # atx (###) or setext (underline)
    
    # LLM optimization
    strip_formatting: bool = False  # Remove all bold/italic for cleaner LLM input
    max_line_length: int = 0  # 0 = no wrapping


# ═══════════════════════════════════════════════════════════════════════════════
# FRONTMATTER RENDERER
# ═══════════════════════════════════════════════════════════════════════════════


class FrontmatterRenderer:
    """Renders YAML frontmatter from document metadata."""
    
    @staticmethod
    def render(metadata: DocumentMetadata) -> str:
        """Render metadata as YAML frontmatter."""
        lines = ["---"]
        
        if metadata.title and metadata.title != "IB Report":
            lines.append(f"title: \"{metadata.title}\"")
        if metadata.subtitle:
            lines.append(f"subtitle: \"{metadata.subtitle}\"")
        if metadata.company and metadata.company != "Korea Development Bank":
            lines.append(f"company: \"{metadata.company}\"")
        if metadata.ticker:
            lines.append(f"ticker: \"{metadata.ticker}\"")
        if metadata.sector and metadata.sector != "SECTOR":
            lines.append(f"sector: \"{metadata.sector}\"")
        if metadata.analyst and metadata.analyst != "DCM Team 1":
            lines.append(f"analyst: \"{metadata.analyst}\"")
        
        # Extra fields
        for key, value in metadata.extra.items():
            lines.append(f"{key}: \"{value}\"")
        
        lines.append("---")
        lines.append("")
        
        # Only return if we have actual content
        if len(lines) > 3:
            return "\n".join(lines)
        return ""


# ═══════════════════════════════════════════════════════════════════════════════
# TEXT RUN RENDERER
# ═══════════════════════════════════════════════════════════════════════════════


class TextRunRenderer:
    """Renders TextRun objects to Markdown inline formatting."""
    
    def __init__(self, config: RenderConfig):
        self.config = config
    
    def render(self, runs: List[TextRun]) -> str:
        """Render text runs with formatting."""
        if self.config.strip_formatting:
            return "".join(run.text for run in runs)
        
        result = []
        for run in runs:
            text = run.text
            
            if run.bold:
                text = f"{self.config.bold_marker}{text}{self.config.bold_marker}"
            if run.italic:
                text = f"{self.config.italic_marker}{text}{self.config.italic_marker}"
            if run.superscript:
                text = f"^{text}^"
            
            result.append(text)
        
        return "".join(result)


# ═══════════════════════════════════════════════════════════════════════════════
# ELEMENT RENDERERS
# ═══════════════════════════════════════════════════════════════════════════════


class HeadingRenderer:
    """Renders heading elements."""
    
    @staticmethod
    def render(heading: Heading) -> str:
        """Render a heading to Markdown."""
        prefix = "#" * heading.level
        return f"{prefix} {heading.text}\n"


class ParagraphRenderer:
    """Renders paragraph elements."""
    
    def __init__(self, config: RenderConfig):
        self.run_renderer = TextRunRenderer(config)
    
    def render(self, para: Paragraph) -> str:
        """Render a paragraph to Markdown."""
        if para.runs:
            text = self.run_renderer.render(para.runs)
        else:
            text = para.text
        
        return f"{text}\n"


class ListRenderer:
    """Renders list elements."""
    
    def __init__(self, config: RenderConfig):
        self.run_renderer = TextRunRenderer(config)
    
    def render_bullet(self, item: ListItem) -> str:
        """Render a bullet list item."""
        if item.runs:
            text = self.run_renderer.render(item.runs)
        else:
            text = item.text
        
        indent = "  " * item.indent_level
        return f"{indent}- {text}\n"
    
    def render_numbered(self, number: str, item: ListItem) -> str:
        """Render a numbered list item."""
        if item.runs:
            text = self.run_renderer.render(item.runs)
        else:
            text = item.text
        
        indent = "  " * item.indent_level
        return f"{indent}{number}. {text}\n"


class TableRenderer:
    """Renders table elements."""
    
    def __init__(self, config: RenderConfig):
        self.config = config
    
    def render(self, table: Table) -> str:
        """Render a table to Markdown."""
        if not table.rows:
            return ""
        
        lines = []
        
        # Header row
        if table.rows:
            header_row = table.rows[0]
            header_cells = [self._escape_cell(c.content) for c in header_row.cells]
            lines.append("| " + " | ".join(header_cells) + " |")
            
            # Separator row
            if self.config.table_alignment and table.alignments:
                sep_cells = []
                for i, align in enumerate(table.alignments):
                    if align == "center":
                        sep_cells.append(":---:")
                    elif align == "right":
                        sep_cells.append("---:")
                    else:
                        sep_cells.append("---")
                lines.append("| " + " | ".join(sep_cells) + " |")
            else:
                sep = ["---"] * len(header_cells)
                lines.append("| " + " | ".join(sep) + " |")
        
        # Data rows
        for row in table.rows[1:]:
            cells = [self._escape_cell(c.content) for c in row.cells]
            lines.append("| " + " | ".join(cells) + " |")
        
        lines.append("")
        return "\n".join(lines)
    
    @staticmethod
    def _escape_cell(text: str) -> str:
        """Escape pipe characters in cell content."""
        return text.replace("|", "\\|").replace("\n", " ")


class BlockquoteRenderer:
    """Renders blockquote/callout elements."""
    
    @staticmethod
    def render(blockquote: Blockquote) -> str:
        """Render a blockquote to Markdown."""
        lines = []
        
        # Title line
        lines.append(f"> **[{blockquote.title}]**")
        
        # Content lines
        for line in blockquote.text.split("\n"):
            lines.append(f"> {line}")
        
        lines.append("")
        return "\n".join(lines)


class ImageRenderer:
    """Renders image elements."""
    
    def __init__(self, config: RenderConfig):
        self.config = config
    
    def render(self, image: Image) -> str:
        """Render an image to Markdown."""
        alt = image.alt_text or "Image"
        
        if image.base64_data and self.config.embed_images_base64:
            return f"![{alt}](data:{image.mime_type};base64,{image.base64_data})\n"
        
        path = image.path
        if self.config.image_path_prefix:
            path = f"{self.config.image_path_prefix}/{path}"
        
        return f"![{alt}]({path})\n"


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN RENDERER
# ═══════════════════════════════════════════════════════════════════════════════


class MarkdownRenderer:
    """Main renderer for converting DocumentModel to Markdown."""
    
    def __init__(self, config: Optional[RenderConfig] = None):
        self.config = config or RenderConfig()
        
        # Initialize sub-renderers
        self.frontmatter_renderer = FrontmatterRenderer()
        self.heading_renderer = HeadingRenderer()
        self.paragraph_renderer = ParagraphRenderer(self.config)
        self.list_renderer = ListRenderer(self.config)
        self.table_renderer = TableRenderer(self.config)
        self.blockquote_renderer = BlockquoteRenderer()
        self.image_renderer = ImageRenderer(self.config)
    
    def render(self, model: DocumentModel) -> str:
        """Render DocumentModel to Markdown string."""
        output_parts = []
        
        # Frontmatter
        if self.config.include_frontmatter:
            frontmatter = self.frontmatter_renderer.render(model.metadata)
            if frontmatter:
                output_parts.append(frontmatter)
        
        # Elements
        for element in model.elements:
            rendered = self._render_element(element)
            if rendered:
                output_parts.append(rendered)
        
        return "\n".join(output_parts)
    
    def _render_element(self, element: Element) -> str:
        """Render a single element."""
        content = element.content
        
        if element.element_type in (
            ElementType.HEADING_1,
            ElementType.HEADING_2,
            ElementType.HEADING_3,
            ElementType.HEADING_4,
            ElementType.NUMBERED_HEADING,
        ):
            return self.heading_renderer.render(content)
        
        elif element.element_type == ElementType.PARAGRAPH:
            return self.paragraph_renderer.render(content)
        
        elif element.element_type == ElementType.BULLET_LIST:
            return self.list_renderer.render_bullet(content)
        
        elif element.element_type == ElementType.NUMBERED_LIST:
            number, item = content
            return self.list_renderer.render_numbered(number, item)
        
        elif element.element_type == ElementType.TABLE:
            return self.table_renderer.render(content)
        
        elif element.element_type == ElementType.BLOCKQUOTE:
            return self.blockquote_renderer.render(content)
        
        elif element.element_type == ElementType.IMAGE:
            return self.image_renderer.render(content)
        
        elif element.element_type == ElementType.SEPARATOR:
            return "---\n"
        
        else:
            logger.warning("Unknown element type: %s", element.element_type)
            return ""


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════


def render_to_markdown(
    model: DocumentModel,
    include_frontmatter: bool = True,
    strip_formatting: bool = False,
) -> str:
    """
    Render a DocumentModel to Markdown string.
    
    Args:
        model: The DocumentModel to render
        include_frontmatter: Whether to include YAML frontmatter
        strip_formatting: Remove bold/italic for cleaner LLM input
    
    Returns:
        Markdown string
    """
    config = RenderConfig(
        include_frontmatter=include_frontmatter,
        strip_formatting=strip_formatting,
    )
    renderer = MarkdownRenderer(config)
    return renderer.render(model)
```

### 3.3 word_to_md.py (CLI)

```python
"""
Word to Markdown Converter
Main entry point for converting Word documents to Markdown.

Usage:
    uv run word_to_md.py input.docx [output.md]
    uv run word_to_md.py input.docx --strip           # Remove formatting for LLM
    uv run word_to_md.py input.docx --no-frontmatter  # Skip YAML metadata
    uv run word_to_md.py input.docx --extract-images  # Save images to folder
"""

import sys
import argparse
import logging
from pathlib import Path
from typing import Optional

from word_parser import parse_word_file
from md_renderer import render_to_markdown, RenderConfig, MarkdownRenderer

logger = logging.getLogger("word_to_md")


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Convert Word documents to Markdown for LLM consumption",
    )
    
    parser.add_argument("input_file", help="Input Word document (.docx)")
    parser.add_argument("output_file", nargs="?", help="Output Markdown file (optional)")
    
    parser.add_argument(
        "--strip", "-s",
        action="store_true",
        help="Strip formatting (bold/italic) for cleaner LLM input",
    )
    parser.add_argument(
        "--no-frontmatter",
        action="store_true",
        help="Skip YAML frontmatter generation",
    )
    parser.add_argument(
        "--extract-images",
        action="store_true",
        help="Extract images to separate folder",
    )
    parser.add_argument(
        "--image-dir",
        help="Directory for extracted images (default: <input>_images/)",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Verbose output",
    )
    
    args = parser.parse_args()
    
    # Setup logging
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format="[%(levelname)s] %(message)s")
    
    # Resolve paths
    input_path = Path(args.input_file)
    if not input_path.exists():
        logger.error("File not found: %s", input_path)
        sys.exit(1)
    
    if args.output_file:
        output_path = Path(args.output_file)
    else:
        output_path = input_path.with_suffix(".md")
    
    # Image directory
    image_dir = None
    if args.extract_images:
        image_dir = args.image_dir or f"{input_path.stem}_images"
    
    # Parse Word document
    logger.info("Parsing: %s", input_path.name)
    model = parse_word_file(
        str(input_path),
        extract_images=args.extract_images,
        image_output_dir=image_dir,
    )
    logger.info("Parsed %d elements", len(model.elements))
    
    # Render to Markdown
    config = RenderConfig(
        include_frontmatter=not args.no_frontmatter,
        strip_formatting=args.strip,
    )
    renderer = MarkdownRenderer(config)
    markdown = renderer.render(model)
    
    # Save output
    output_path.write_text(markdown, encoding="utf-8")
    logger.info("Saved: %s", output_path)
    
    print(f"\nConversion complete!")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")


if __name__ == "__main__":
    main()
```

---

## 4. 구현 순서

### Phase 1: Core Infrastructure (1-2일)
1. [ ] `word_parser.py` 기본 구조 생성
2. [ ] `md_parser.py`의 데이터 모델 재사용 확인
3. [ ] 기본 paragraph/heading 파싱 구현
4. [ ] `md_renderer.py` 기본 구조 생성
5. [ ] 간단한 테스트 케이스 생성

### Phase 2: Element Parsing (2-3일)
1. [ ] Table 파싱 구현
2. [ ] List (bullet/numbered) 파싱 구현
3. [ ] Callout box 감지 및 파싱
4. [ ] TextRun 포맷팅 추출 (bold, italic)

### Phase 3: Advanced Features (2-3일)
1. [ ] Image 추출 (embedded images)
2. [ ] Base64 이미지 지원
3. [ ] Metadata 추출 (document properties)
4. [ ] 스타일 기반 heading 레벨 감지 개선

### Phase 4: CLI & Integration (1일)
1. [ ] `word_to_md.py` CLI 완성
2. [ ] `pyproject.toml` 엔트리포인트 추가
3. [ ] 에러 핸들링 강화
4. [ ] README 업데이트

### Phase 5: Testing & Polish (1-2일)
1. [ ] 실제 Word 문서로 테스트
2. [ ] Edge case 처리
3. [ ] Round-trip 테스트 (Word → MD → Word)
4. [ ] LLM 출력 품질 검증

---

## 5. 고려사항

### 5.1 기술적 도전
| 문제 | 해결 방안 |
|------|----------|
| Word 스타일 다양성 | 스타일 이름 패턴 매칭 + 폴백 휴리스틱 |
| 이미지 추출 복잡성 | python-docx의 inline shapes API 활용 |
| Callout 감지 | 1x1 테이블 + 배경색/보더 스타일 분석 |
| 중첩 리스트 | 들여쓰기 레벨 분석 |

### 5.2 LLM 최적화
- `--strip` 옵션: Bold/italic 제거로 토큰 절약
- `--no-frontmatter`: 메타데이터 생략
- Clean paragraph 분리
- 불필요한 공백 제거

### 5.3 의존성
```toml
# pyproject.toml 추가
[project.optional-dependencies]
full = [
    "pillow",        # Image processing
    "charset-normalizer",  # Encoding detection
]
```

### 5.4 CLI 엔트리포인트
```toml
# pyproject.toml
[project.scripts]
word-to-md = "word_to_md:main"
```

---

## 6. 예상 결과물

### Input: report.docx
```
[Title: 네페스 기업분석]
[Heading 1 style]
[Table with financial data]
[Callout box with key insights]
```

### Output: report.md
```markdown
---
title: "네페스 기업분석"
company: "Korea Development Bank"
analyst: "DCM Team"
date: "2024-01-15"
---

# 네페스 기업분석

## 1. 회사 개요

네페스는 반도체 후공정 장비 전문 기업으로...

| 항목 | 2023 | 2024E | 2025E |
|------|------|-------|-------|
| 매출액 | 1,234 | 1,567 | 1,890 |
| 영업이익 | 123 | 156 | 189 |

> **[시사점]**
> 반도체 업황 회복과 함께 실적 개선이 예상됨
```

---

## 7. 변경 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|----------|
| v1.0 | 2024-xx-xx | 초기 계획 수립 |
