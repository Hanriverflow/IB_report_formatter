# AGENTS.md - IB Report Formatter

> Bidirectional Markdown ↔ Word Document Converter for IB-Style Reports

## Quick Reference

```bash
# Markdown → Word conversion
uv run md_to_word.py input.md [output.docx]
uv run md_to_word.py --list                    # List available .md files
uv run md_to_word.py --list -i                 # Interactive file selection
uv run md_to_word.py input.md --format         # Pre-format single-line markdown

# Word → Markdown conversion (for LLM consumption)
uv run word_to_md.py input.docx [output.md]
uv run word_to_md.py --list                    # List available .docx files
uv run word_to_md.py --list -i                 # Interactive file selection
uv run word_to_md.py input.docx --strip        # LLM-optimized (no bold/italic)
uv run word_to_md.py input.docx --no-frontmatter  # Skip YAML metadata header
uv run word_to_md.py input.docx --extract-images  # Extract images to folder

# Pre-format only (Gemini Deep Research clipboard output)
uv run md_formatter.py input.md [output.md]
uv run md_formatter.py --check input.md        # Check if formatting needed

# No formal test suite exists - verify changes manually
```

## Project Structure

```
IB_report_formatter/
├── md_to_word.py      # Main CLI entry point, orchestrates conversion
├── md_parser.py       # Markdown parsing, frontmatter, elements, LaTeX, Base64 images
├── md_formatter.py    # Pre-processor for single-line clipboard markdown
├── ib_renderer.py     # Word document rendering with IB bank styling
├── word_to_md.py      # Word → Markdown CLI entry point
├── word_parser.py     # Word document parsing into DocumentModel
├── md_renderer.py     # DocumentModel → Markdown text rendering
└── docs/              # Documentation/memory files
```

## Architecture

**MD → Word Pipeline:** `Markdown → Parser → DocumentModel → Renderer → Word Document`

**Word → MD Pipeline:** `Word Document → Parser → DocumentModel → Renderer → Markdown`

| Module           | Responsibility                                          |
|------------------|---------------------------------------------------------|
| `md_to_word.py`  | CLI, path resolution, MD→Word conversion orchestration  |
| `md_parser.py`   | Parse frontmatter, headings, tables, LaTeX, images      |
| `md_formatter.py`| Convert single-line text to structured markdown         |
| `ib_renderer.py` | Apply IB styling, generate Word via python-docx         |
| `word_to_md.py`  | CLI, path resolution, Word→MD conversion orchestration  |
| `word_parser.py` | Parse Word doc properties, paragraphs, tables, images   |
| `md_renderer.py` | Render DocumentModel to clean Markdown text             |

### Markdown Paragraph Normalization (MD -> Word)

- Soft-wrapped lines inside the same markdown paragraph are merged with spaces.
- Hard breaks are preserved only for explicit markers (`<br>` or trailing `\`).
- Trailing-double-space hard break is **opt-in** only: `MarkdownParser(preserve_trailing_double_space_break=True)`.
- Paragraph text normalization includes:
  - collapse repeated spaces/tabs to a single space
  - trim unnecessary spaces inside parentheses (e.g., `( PFV )` -> `(PFV)`)

Regression coverage is in `tests/test_md_parser.py` (soft wrap merge, hard-break policy, opt-in legacy mode, spacing normalization).

## Dependencies

- **pyyaml** - YAML frontmatter parsing
- **python-docx** - Word document generation
- **charset-normalizer** (optional) - Better encoding detection for Korean text

---

## Code Style Guidelines

### Python Version
- **Python 3.8+ compatible** - Do NOT use features like `Path.with_stem()` (3.9+)
- Use `Path.with_name()` patterns instead for 3.8 compatibility

### Type Hints
```python
from typing import List, Dict, Optional, Tuple, Union

def parse(lines: List[str]) -> Tuple[DocumentMetadata, List[str]]:
    ...

def render(self, model: DocumentModel) -> Document:
    ...
```

### Docstrings
Use Google-style docstrings with Args/Returns sections:
```python
def parse_cell(text: str, is_header: bool) -> TableCell:
    """
    Parse a single table cell.

    Args:
        text: Raw cell content from markdown
        is_header: Whether this cell is in the header row

    Returns:
        TableCell with content, runs, and detected properties
    """
```

### Data Models
Use `@dataclass` with type hints. Use `frozen=True` for immutable config:
```python
@dataclass(frozen=True)
class IBStyle:
    """IB Bank styling constants"""
    NAVY: RGBColor = RGBColor(0, 51, 102)
    BODY_FONT: str = "Calibri"

@dataclass
class TableCell:
    """A cell in a table"""
    content: str
    runs: List[TextRun] = field(default_factory=list)
    is_numeric: bool = False
```

### Enums
Use `Enum` with `auto()` for type-safe element classification:
```python
class ElementType(Enum):
    HEADING_1 = auto()
    HEADING_2 = auto()
    PARAGRAPH = auto()
    TABLE = auto()
    LATEX_BLOCK = auto()
```

### Regex Patterns
Compile patterns as class attributes for performance:
```python
class TextParser:
    # Compiled regex - class-level cache
    _BOLD_SPLIT_RE = re.compile(r'(\*\*.*?\*\*)')
    _ESCAPE_RE = re.compile(r'\\([~.*"\'()\[\]{}|_-])')

    @classmethod
    def parse_runs(cls, text: str) -> List[TextRun]:
        parts = cls._BOLD_SPLIT_RE.split(text)
        ...
```

### Class Organization
- Use `@staticmethod` for utility methods without self dependency
- Use `@classmethod` for factory methods or methods using class-level data
- Group related functionality into focused classes (single responsibility)

### Section Comments
Use Unicode box-drawing for major sections:
```python
# ═══════════════════════════════════════════════════════════════════════════════
# SECTION NAME
# ═══════════════════════════════════════════════════════════════════════════════
```

### Changelog in Docstrings
Document version changes in module docstrings:
```python
"""
MD Parser Module for IB Style Word Report Converter

Changelog (v3):
    - NEW: LaTeX block equation parsing ($$ ... $$)
    - ENHANCED: Encoding detection with charset_normalizer fallback
    - FIXED: heading level mapping (## -> level=2)
"""
```

---

## Error Handling

### Element-Level Resilience
Render continues even if individual elements fail:
```python
for idx, element in enumerate(model.elements):
    try:
        self._render_element(element)
    except Exception as e:
        logger.warning("Failed to render element %d: %s", idx, e)
        # Insert visible error marker in document
        p = self.doc.add_paragraph()
        err_run = p.add_run(f"[Render Error: {element.element_type.name}]")
        FontStyler.apply_run_style(err_run, italic=True, color=STYLE.RED)
```

### File I/O with Encoding Fallback
Handle Korean text encoding gracefully:
```python
def _read_with_encoding(file_path: Path) -> str:
    encodings = ["utf-8", "utf-8-sig", "euc-kr", "cp949"]
    for enc in encodings:
        try:
            return file_path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(...)
```

### Safe Save with Lock Handling
```python
try:
    doc.save(str(output_path))
except PermissionError:
    # File is open in Word - save with timestamp suffix
    new_name = f"{output_path.stem}_{timestamp}{output_path.suffix}"
    doc.save(str(new_path))
```

---

## Naming Conventions

| Category       | Convention          | Example                        |
|----------------|---------------------|--------------------------------|
| Classes        | PascalCase          | `TableParser`, `IBDocumentRenderer` |
| Functions      | snake_case          | `parse_markdown_file`, `render_runs` |
| Constants      | UPPER_SNAKE_CASE    | `PARENT_DIR`, `OUTPUT_SUFFIX`  |
| Private        | Leading underscore  | `_BOLD_SPLIT_RE`, `_parse_cell` |
| Type aliases   | PascalCase          | `ElementContent`               |

---

## Import Order

```python
# 1. Standard library
import re
import sys
import time
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Union
from enum import Enum, auto

# 2. Third-party
import yaml
from docx import Document
from docx.shared import Inches, Pt, RGBColor

# 3. Local modules
from md_parser import DocumentModel, parse_markdown_file
from ib_renderer import IBDocumentRenderer
```

---

## Logging

Use the `logging` module, not print statements:
```python
logger = logging.getLogger(__name__)

logger.info("Parsing: %s", self.md_file_path.name)
logger.warning("Table row %d has %d columns (expected %d)", i, len(cells), col_count)
logger.debug("Encoding detected: %s", result.encoding)
```

Custom formatter for CLI output:
```python
class LogFormatter(logging.Formatter):
    PREFIXES = {
        logging.INFO: "[INFO]",
        logging.WARNING: "[WARNING]",
        logging.ERROR: "[ERROR]",
    }
```

---

## Adding New Element Types

1. Add to `ElementType` enum in `md_parser.py`
2. Create dataclass for the element data
3. Add parsing logic in `MarkdownParser._parse_elements()`
4. Add rendering logic in `ib_renderer.py` (create Renderer class)
5. Register in `IBDocumentRenderer._render_element()`

---

## Korean Text Handling

- Use East Asian font setting for Korean text support:
```python
def set_east_asian_font(element, font_name: str = "Malgun Gothic"):
    rPr.rFonts.set(qn('w:eastAsia'), font_name)
```

- Korean sentence endings for boundary detection:
```python
SENTENCE_END_RE = re.compile(
    r"(다|요|음|함|임|됨|것|수|점|니다|입니다|습니다)\."
    r"(?=[가-힣A-Z\[])"
)
```

---

## Common Pitfalls

1. **Path.with_stem()** - Not available in Python 3.8. Use `with_name()` pattern
2. **Empty catch blocks** - Always log or handle errors explicitly
3. **Regex without compile** - Compile patterns as class attributes
4. **Missing type hints** - All public functions should have type hints
5. **Print vs logging** - Use logging module for all output

---

## Configuration

The project uses `uv` for package management with `pyproject.toml`.

```bash
# Install dependencies
uv sync

# Install with optional features (LaTeX rendering, better encoding)
uv sync --extra full

# Install dev dependencies
uv sync --extra dev
```

Claude settings are in `.claude/settings.local.json` for allowed permissions.

---

## Supported Features

### Images
- **Base64 embedded images**: Automatically decoded and inserted
- **File path images**: Local files inserted if found
- **Fallback**: Placeholder text if image cannot be loaded

### LaTeX Equations
- **Block equations**: `$$ E = mc^2 $$` rendered as centered images
- **Inline equations**: `$x^2$` detected within paragraphs
- **Requires**: `matplotlib` (optional dependency)

### Tables
- **Financial tables**: Automatic thousand separator formatting
- **Negative numbers**: Red color, parentheses support
- **Sensitivity tables**: Base case highlighting
- **Risk matrices**: Color-coded risk levels

### Callout Boxes
- **Executive Summary / 요약**: Navy background, white text
- **Key Insight / 시사점**: Blue accent border
- **Warning / 주의**: Orange accent
- **Note / 참고**: Gray accent

### Headers & Footers
- Company name in header
- CONFIDENTIAL mark
- Page numbers (Page X of Y)
