# AGENTS.md - IB Report Formatter

> Markdown → Word engine for IB reports and Korean business documents.

## Product boundary

Word→MD development is deliberately retired by the project owner. Do not restore the inverse CLI/parser/renderer/OMML converter or roundtrip audit as a roadmap item. Use external projects for reverse conversion. Historical source: `d819bbb` / `codex/archive-word-to-md-d819bbb`.

Current plan: `docs/implementation-plan-20260914.md`. Verification: `docs/verification-20260914.md`. Older bidirectional plans are historical, not active requirements.

Production-quality follow-up: `docs/improvement-plan-20260915.md` and `docs/verification-20260915.md`.

Next-step research roadmap (owner decisions D1–D5 and candidate work E1–E17; not active requirements until approved): `docs/improvement-roadmap-20260929.md`.
Input-loss, save-safety and performance hardening: `docs/verification-20260929.md`.
Term-sheet profile and term variables: design `docs/term-sheet-design-20260929.md`, plan `docs/term-sheet-plan-20260929.md` (§2 decisions override the design), verification `docs/verification-term-sheet-20260929.md`.

## Quick reference

```sh
uv sync --extra dev
uv run md-to-word input.md output.docx --strict
uv run md-to-word input.md output.docx --profile plain
uv run md-to-word samples/profiles/office-letter.md letter.docx --strict
uv run md-to-word --list-profiles
uv run md-to-word samples/profiles --batch
uv run docx-audit letter.docx
uv run md_formatter.py input.md output.md
uv run md-format converted.md term-sheet.md --converted-term-sheet
uv run pytest tests/
uv build
```

Existing `md_to_word.py` and `ib-report` remain entry points. On Windows, use a fresh absolute pytest basetemp under the actual OS temp directory. Do not delete unrelated/unreadable legacy temp directories in the repository.

## Architecture

`Markdown → MarkdownParser → DocumentModel → IBDocumentRenderer → DOCX`

| Module | Responsibility |
|---|---|
| `document_model.py` | Shared data classes and enums; re-exported by `md_parser` for compatibility |
| `document_profiles.py` | Seven profiles, immutable options, YAML validation, table semantics, themes |
| `render_styles.py` | Immutable style values and render-scoped ContextVar |
| `md_parser.py` | Profile-aware Markdown parsing, frontmatter, tables, images, equations |
| `ib_renderer.py` | One composition path for all callers; reusable element renderers |
| `office_layout.py` | Letter/report/minutes metadata and native Korean numbering |
| `term_sheet.py` | House texts and term-sheet composition; must not import `ib_renderer` |
| `ooxml_order.py` | Schema-ordered insertion for hand-built tblPr/tcPr/pPr/rPr children |
| `term_variables.py` | Term variables, content-control tags and snapshots |
| `numeric_checks.py` | `checks:` relations and table `schedule` arithmetic; reads displayed amounts, reports only |
| `docx_audit.py` | Structural diagnostics, not reverse conversion or visual QA |
| `md_to_word.py` | CLI, file/batch conversion and safe save |
| `converters.py` | Registry with Markdown input and DOCX output only |
| `md_formatter.py`, `deep_md_cleaner.py` | Optional input cleanup |
| `converted_md_cleaner.py` | Opt-in MD→MD cleanup of HWP/Word-converted term sheets (`--converted-term-sheet`); reports every change |
| `diagram_renderer.py` | Existing diagram rendering |
| `cli_utils.py`, `stream_utils.py` | Shared I/O helpers |

## Profiles and configuration

- Seven profiles: `ib-report` (legacy default), `ib-memo`, `plain`, `office-letter`, `business-report`, `meeting-minutes`, `term-sheet`.
- Keep real deal documents and real institution house files outside the repository; samples and tests use fictional names and freshly written boilerplate only.
- Explicit CLI/API options override YAML, then profile defaults. Pass profile to the parser as well as renderer when overriding.
- Four general profiles use A4, neutral metadata and native numbered lists with four-space nesting. They must not infer IB financial table semantics or promote short numbered text to headings.
- Converted-input heuristics (numbered-line headings, band tables, cover to frontmatter) live only in the explicit `--converted-term-sheet` MD→MD step, never in the parser's default path. That step changes only top-level blocks (never code, list or quote blocks, or HTML tables other than clean one-row bands), reports every moved, rewritten, removed or kept line, keeps unclear lines and existing frontmatter, and never infers `prepared_by`.
- Keep styles immutable and scoped. Do not replace a process-global style instance, capture theme-dependent values in default arguments, or cache styles across requests.
- Use one renderer per concurrent job. CLI/API/registry must share `IBDocumentRenderer.render`; do not reintroduce a second document assembler.
- Tables use ordered YAML specifications. Never guess a sensitivity table's base case: use explicit one-based body-row/column coordinates.
- Preserve blank cells, escapes, codes, dates and numeric presentation. Percent/bps/multiple roles do not rescale numeric values.
- Strict mode rejects known input loss or rendering failures before saving. Diagnostics do not establish visual or financial correctness.
- Changes to output must have regression tests through actual parsing/rendering and, for layout changes, representative page inspection.
- Office letters close before the explicitly named unique H1 in `letter.appendix_heading`; no implicit appendix inference or second assembler. Optional `appendix_label` requires that boundary. Preserve native attachment numbers.
- Inline HTML breaks must reach Word paragraphs, emphasis, headings, lists and cells without changing code/escaped literals or link URLs.
- Word's nonprinting pagination squares are not list bullets. Do not globally add or remove keep-lines/keep-next; retain necessary heading controls. Audit warnings and visual-review status are distinct from structural issues.
- Word QA requires a new/empty output directory. Record source/output hashes, actual/expected pages and Word version; `visualReview: pending` is not a visual pass. Inspect every page and record review against the manifest hash.
- Built distributions explicitly include engine code only; never package private reports, QA artifacts or user files.

## Converter registry

```python
from converters import get_default_registry

registry = get_default_registry()
model = registry.convert("report.md", profile="business-report")
saved = registry.convert(model, output_format="docx", output_path="out.docx", strict=True)
```

Only `MarkdownInputConverter` and `DocxOutputConverter` are built in.

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
- **matplotlib** - Existing math and diagram rendering
- **charset-normalizer** (optional) - Better encoding detection for Korean text

---

## Code Style Guidelines

### Python Version
- **Python 3.12+**
- Python floor raised to 3.12 by owner decision D1 on 2026-09-29

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
Non-strict rendering continues with diagnostics if an element fails. Strict rendering must reject the output before saving:
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

1. Add to `ElementType` enum in `document_model.py`
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

1. **Empty catch blocks** - Always log or handle errors explicitly
2. **Regex without compile** - Compile patterns as class attributes
3. **Missing type hints** - All public functions should have type hints
4. **Print vs logging** - Use logging module for all output

---

## Configuration

The project uses `uv` for package management with `pyproject.toml`.

```bash
# Install dependencies
uv sync

# Install with optional features (better encoding)
uv sync --extra full

# Install dev dependencies
uv sync --extra dev
```

Claude settings are in `.claude/settings.local.json` for allowed permissions.

---

## Supported Features

### Images
- **Base64 embedded images**: Automatically decoded and inserted
- **File path images**: Local files inserted if found; percent-encoded paths (`a%20b.png`) are tried decoded
- **Inline images**: In text and table cells, fitted to the cell width
- **Fallback**: Placeholder text if image cannot be loaded (strict mode rejects it)

### LaTeX Equations
- **Block equations**: `$$ E = mc^2 $$` rendered as centered images
- **Inline equations**: `$x^2$` detected within paragraphs
- **Requires**: `matplotlib` (installed by default; plain-text fallback if unavailable)

### Tables
- **Financial tables**: Automatic thousand separator formatting
- **Negative numbers**: Red color, parentheses support
- **Sensitivity tables**: Explicit-coordinate base case highlighting only
- **Risk matrices**: Color-coded risk levels
- **HTML tables** (converter output): `colspan`/`rowspan` become span markers; rows the header spans cover stay header rows
- **Header rows**: table spec `header_rows` draws, merges and repeats several header rows in every profile

### Callout Boxes
- **Executive Summary / 요약**: Navy background, white text
- **Key Insight / 시사점**: Blue accent border
- **Warning / 주의**: Orange accent
- **Note / 참고**: Gray accent

### Headers & Footers
- Company name in header
- CONFIDENTIAL mark only when configured (IB defaults)
- Page numbers (Page X of Y for IB, X / Y for general documents)
