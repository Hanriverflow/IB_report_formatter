# AGENTS.md - IB Report Formatter

> Markdown → Word engine for IB reports and Korean business documents.

## Codex document-authoring workflow

For requests to **write or revise a document from source material**, the source-project workflow uses `.agents/skills/ib-document/SKILL.md` (`$ib-document`). Installed Claude Code/Codex plugin work instead uses `skills/setup/SKILL.md` and `skills/ib-document/SKILL.md`, calling `scripts/plugin_runtime.py` to keep dependencies and outputs outside the plugin; see `docs/plugin-guide.md`. Both workflows remain supported. The host supplies the LLM and orchestration; this repository remains the deterministic MD→DOCX tool. Read the applicable skill by path if automatic discovery is unavailable. This routing does not apply to ordinary engine development.

The source-project workflow records source facts and unresolved conditions, drafts profile-aware Markdown/YAML, runs `scripts/agent_render.py` in strict mode, and reviews diagnostics with at most two repair revisions. Keep real source/house/output files outside the repository. Never change source facts or remove checks merely to obtain a passing render. Source review, structural audit and actual visual review are separate statuses. Beginner instructions: `docs/codex-start.html`.

## Product boundary

Word→MD development is deliberately retired by the project owner. Do not restore the inverse CLI/parser/renderer/OMML converter or roundtrip audit as a roadmap item. Use external projects for reverse conversion. Historical source: `d819bbb` / `codex/archive-word-to-md-d819bbb`.

Current plan: `docs/implementation-plan-20260914.md`. Verification: `docs/verification-20260914.md`. Older bidirectional plans are historical, not active requirements.

Production-quality follow-up: `docs/improvement-plan-20260915.md` and `docs/verification-20260915.md`.

Next-step research roadmap (owner decisions D1–D5 and candidate work E1–E17; not active requirements until approved): `docs/improvement-roadmap-20260929.md`.
Input-loss, save-safety and performance hardening: `docs/verification-20260929.md`.
Term-sheet profile and term variables: design `docs/term-sheet-design-20260929.md`, plan `docs/term-sheet-plan-20260929.md` (§2 decisions override the design), verification `docs/verification-term-sheet-20260929.md`.
Asset path restriction, input limits, Word schema gate, snapshot and property tests (roadmap E1–E3): verification `docs/verification-20261110.md`.
IB table rules (roadmap E6: totals/subtotal rows, per-table negative style, owner decision D4): verification `docs/verification-ib-table-rules-20261114.md`.

## Quick reference

```sh
uv sync --extra dev
uv run ruff check .
uv run ruff format --check .
uv run mypy .
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
| `md_parser.py` | Profile-aware Markdown parsing: `MarkdownParser`, file and stream reading, `parse_markdown_file`; re-exports the parsing names below and the document model |
| `md_parse_text.py` | Inline text runs (`TextParser`), local link targets, term-variable substitution |
| `md_parse_frontmatter.py` | YAML frontmatter and metadata inferred from the body |
| `md_parse_tables.py` | Pipe tables and table span resolution |
| `md_parse_html_tables.py` | HTML tables from converter output |
| `md_parse_blocks.py` | Code fences, footnotes, LaTeX blocks, Base64 images |
| `ib_renderer.py` | One composition path for all callers (`IBDocumentRenderer`), cover and disclaimer, `_metadata_text`; re-exports the element renderers below |
| `ib_render_fonts.py` | Run font styling and the raster font choice (`FontStyler`, `FontPolicy`) |
| `ib_render_document_style.py` | Document styles, table styles, generator signature |
| `ib_render_footnotes.py` | Native footnotes and local link rebasing before save |
| `ib_render_media.py` | Images, charts, LaTeX equations |
| `ib_render_text.py` | Inline text runs (`TextRenderer`) |
| `ib_render_blocks.py` | TOC, headings, paragraphs, lists, callouts |
| `ib_render_tables.py` | Tables in every profile (`TableRenderer`) |
| `office_layout.py` | Letter/report/minutes metadata and native Korean numbering |
| `term_sheet.py` | House texts and term-sheet composition; must not import `ib_renderer` |
| `ooxml_order.py` | Schema-ordered insertion for hand-built tblPr/tcPr/pPr/rPr and settings children |
| `input_safety.py` | Allowed asset folders, input limits, the six rejection codes and the rejection exception types |
| `term_variables.py` | Term variables, content-control tags and snapshots |
| `numeric_checks.py` | `checks:` relations and table `schedule` arithmetic; reads displayed amounts, reports only |
| `docx_audit.py` | Structural diagnostics, not reverse conversion or visual QA |
| `md_to_word.py` | CLI, file/batch conversion and safe save |
| `converters.py` | Registry with Markdown input and DOCX output only |
| `md_formatter.py`, `deep_md_cleaner.py` | Optional input cleanup |
| `converted_md_cleaner.py` | Opt-in MD→MD cleanup of HWP/Word-converted term sheets (`--converted-term-sheet`); reports every change |
| `diagram_renderer.py` | Existing diagram rendering |
| `cli_utils.py`, `stream_utils.py` | Shared I/O helpers |

The `md_parse_*.py` and `ib_render_*.py` files are internal. Import from `md_parser` and `ib_renderer`, which re-export every name as before. A split file does not import its original module at module level, and logs under the original module's logger name; `tests/test_module_layout.py` checks the layout.

## Profiles and configuration

- Seven profiles: `ib-report` (legacy default), `ib-memo`, `plain`, `office-letter`, `business-report`, `meeting-minutes`, `term-sheet`.
- Keep real deal documents and real institution house files outside the repository; samples and tests use fictional names and freshly written boilerplate only.
- Explicit CLI/API options override YAML, then profile defaults. Pass profile to the parser as well as renderer when overriding.
- Four general profiles use A4, neutral metadata and native numbered lists with four-space nesting. They must not infer IB financial table semantics or promote short numbered text to headings.
- Converted-input heuristics (numbered-line headings, band tables, cover to frontmatter) live only in the explicit `--converted-term-sheet` MD→MD step, never in the parser's default path. That step changes only top-level blocks (never code, list or quote blocks, or HTML tables other than clean one-row bands), reports every moved, rewritten, removed or kept line, keeps unclear lines and existing frontmatter, and never infers `prepared_by`.
- Keep styles immutable and scoped. Do not replace a process-global style instance, capture theme-dependent values in default arguments, or cache styles across requests.
- Use one renderer per concurrent job. CLI/API/registry must share `IBDocumentRenderer.render`; do not reintroduce a second document assembler.
- Tables use ordered YAML specifications. Never guess a sensitivity table's base case: use explicit one-based body-row/column coordinates.
- Totals and subtotal rules (`total_rows`, `subtotal_rows`) apply only to the one-based body rows a spec lists, in every profile; never infer them from text such as `합계` or from `schedule.total`. A spec's `negative` style (`paren`, `minus`, `triangle`, `type: financial` only) changes how negative amounts look, never their digits or units; without it negatives keep their written form and red (owner decision D4, 2026-11-14). It reads a cell's whole text once and restyles only a cell that is exactly one negative amount; an amount split across differently formatted runs keeps its text and gets a rendering warning (`IBDocumentRenderer.warnings`) that never rejects. It turns off only the automatic red: author-written colours and link colours stay. Alignment, column kinds and widths are judged on the written text. Hand-built cell borders go through `TableStyler.set_cell_border` (schema-ordered `w:tcBorders`).
- Preserve blank cells, escapes, codes, dates and numeric presentation. Percent/bps/multiple roles do not rescale numeric values.
- Strict mode rejects known input loss or rendering failures before saving. Diagnostics do not establish visual or financial correctness.
- Changes to output must have regression tests through actual parsing/rendering and, for layout changes, representative page inspection.
- Office letters close before the explicitly named unique H1 in `letter.appendix_heading`; no implicit appendix inference or second assembler. Optional `appendix_label` requires that boundary. Preserve native attachment numbers.
- Inline HTML breaks must reach Word paragraphs, emphasis, headings, lists and cells without changing code/escaped literals or link URLs.
- Word's nonprinting pagination squares are not list bullets. Do not globally add or remove keep-lines/keep-next; retain necessary heading controls. Audit warnings and visual-review status are distinct from structural issues.
- Word QA requires a new/empty output directory. Record source/output hashes, actual/expected pages and Word version; `visualReview: pending` is not a visual pass. Inspect every page and record review against the manifest hash.
- Built distributions explicitly include engine code only; never package private reports, QA artifacts or user files. A new engine module must be added to the wheel `only-include` list, the sdist `include` list and ruff's `known-first-party` in `pyproject.toml`.

## Input safety

Document content (body, frontmatter, and the house or theme files it names) is untrusted. Operator CLI/API arguments are trusted.

- Images and logos load only from allowed folders: the source file's folder, plus folders the operator adds with `--asset-root` or `asset_roots`. Text, stream and hand-built-model input has no default folder. Nothing in a document can add a folder.
- Resolve the final location (relative parts, symlinks, junctions) and decide inside/outside **before** checking that the file exists, so the result for an outside path never depends on whether it exists.
- A house logo is restricted when the document names the house file and trusted when the operator does (`--house`, API). Hyperlinks, embedded Base64 images and `data:` URIs are not path-restricted.
- Five limits stop the conversion in both modes and leave no output: input bytes, cells per table, list depth, image bytes, image pixels (`InputLimits`, `--max-*`). Judge size before reading or decoding. Nothing in a document can raise a limit, and there is no unlimited value.
- There are exactly six rejection codes (`asset_outside_root` and five `limit_*`). A coded issue starts with `[code] ` and keeps that prefix inside joined messages. Do not add codes to existing warnings or change their wording.
- Strict rejections raise `StrictValidationError` and limit failures raise `InputLimitError`; both are `ValueError` subclasses and must stay reachable through `__cause__` when wrapped. The renderer's per-element resilience must not turn either into a placeholder.
- A validator that rejects the document's own settings (frontmatter values, table specifications, house or theme content named by the document, terms, checks) runs inside `input_safety.document_validation`, which re-raises the same message as `StrictValidationError`. An error in an operator argument keeps its plain `ValueError`. Do not wrap with a blanket `except Exception`.
- Strings that reach the document from decoded YAML go through `validate_xml_strings`; the Markdown body is checked on the raw text. Strict rendering checks the whole model. Non-strict rendering checks metadata only and keeps per-element resilience for the body; do not widen that check, or documents that convert today get rejected. A `ValueError` that escapes composition is reported as the character rejection only after operator arguments are validated and only when the model really holds such a character; never decide by the error's message.
- Metadata is not type-checked ahead of use: numbers, dates and lists of text render today and must keep rendering. Composers that consume metadata as text run inside `_metadata_text` in `ib_renderer.py`, which reports the `TypeError` or `AttributeError` of a wrong type as a rejection. Wrap a new metadata-consuming composer the same way.
- Load YAML written in or named by a document with `input_safety.load_yaml`, not `yaml.safe_load`: the safe loader raises `AttributeError`, `KeyError` or `IndexError` for some malformed scalars, and it loads alias structures that refer to themselves or expand without bound, which stall whatever walks or prints them. `load_yaml` refuses those (`YAML structure refers to itself`, `YAML aliases expand to too much content`) on the parsed nodes, before construction, because a merge key copies its mapping while values are built; keep the check ahead of `construct_document`, and keep it opening each collection once however often it is aliased. One mapping may hold at most 100 merge keys and at most 100 integer keys of 18 or more characters (merged keys included), and one integer at most 20 colons. Safety hardening of this loader was closed by the owner on 2026-11-13; handle a new finding as its own change, with a reproduction. The allowance must stay out of reach for YAML without aliases; do not replace it with a fixed cap on values.
- The plugin runner and the portable converter stay strict and take no option that adds folders or raises limits. The runner applies the input-size limit before hashing the draft.
- `docx-audit` decides which parts are XML from `[Content_Types].xml` with python-docx's own lookup rules (never from file names), and checks those and every relationship part with the hardened parser before python-docx opens the package.

## Verification gates

- Hand-built XML children go through `ooxml_order.insert_ordered`. The CI job `word-schema` renders every sample and fails on any OOXML-Validator error that `scripts/word_schema_exceptions.json` does not cover; see `scripts/word_schema_validation.md`. The script runs the validator only from an archive whose hash it has just verified. An exception needs a reproduction from a document saved by Word itself, and the reproduction document stays out of the repository.
- `tests/test_sample_snapshots.py` pins the parsed model and every XML part of every sample. A missing snapshot fails. Update only for an intended output change, with `uv run pytest tests/test_sample_snapshots.py --snapshot-update`, and review the diff; a snapshot change in a sample also needs the Word page check below.
- `tests/test_markdown_properties.py` holds four Hypothesis properties (visible text, cell count, literal code and escapes, strict save-or-reject). Do not shrink the generated input or loosen a comparison to pass. Its one strict `xfail` records the intended tree-value emphasis inside code blocks; a new known gap that needs an owner decision is recorded the same way, and fixing one means removing its mark. `tests/test_config_properties.py` applies the save-or-reject property to generated house and theme file content and to YAML aliases, and runs alias cases under a time limit so a conversion that never ends fails instead of hanging the run.

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

### Markdown reading rules (2026-11)

- Bullet markers are `-`, `*` and `+`. `#` to `######` are headings; levels 5 and 6 are emitted as level 4, the deepest styled level.
- A closed code fence is literal, including blank lines and spaces at its end. Do not add spacer text to a blank code line: the paragraph mark sets the line height.
- A backslash escapes any ASCII punctuation character (`TextParser._ESCAPE_RE`). Inside code spans, link destinations and text between dollar signs only the earlier set applies (`_LEGACY_ESCAPE_RE`), so LaTeX commands such as `\%` and `\,` survive in every context. Keep both rules together when adding a path that resolves escapes.
- The table row splitter resolves only an escaped pipe; the cell's inline parse resolves every other escape once.
- Edge cases are pinned in `tests/test_markdown_proposals.py`.

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
Use built-in generics and `X | None`; import abstract types from `collections.abc`. `ruff check` enforces this (pyupgrade rules).
```python
from collections.abc import Mapping

def parse(lines: list[str]) -> tuple[DocumentMetadata, list[str]]:
    ...

def find(options: Mapping[str, str], key: str) -> str | None:
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
    runs: list[TextRun] = field(default_factory=list)
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
    def parse_runs(cls, text: str) -> list[TextRun]:
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
from collections.abc import Mapping
from dataclasses import dataclass, field
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
3. Add parsing logic in `MarkdownParser._parse_elements()` (`md_parser.py`); a block scanner goes in `md_parse_blocks.py`
4. Add rendering logic in the matching `ib_render_*.py` file (create Renderer class; block elements go in `ib_render_blocks.py`) and re-export it from `ib_renderer.py` as `Name as Name`
5. Register in `IBDocumentRenderer._render_element()` (`ib_renderer.py`)

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
5. **`zip()` without `strict=`** - Pass `strict=True` only where the surrounding code guarantees equal lengths; otherwise write `strict=False` explicitly
6. **Unformatted code** - Run `uv run ruff format .`; CI rejects unformatted files

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
- **File path images**: Local files inside the allowed folders are inserted if found; percent-encoded paths (`a%20b.png`) are tried decoded
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
- **Totals and subtotals**: rules and bold on listed body rows only (`total_rows`, `subtotal_rows`)
- **Negative style**: optional per financial table (`negative: paren|minus|triangle`; turns off only the automatic red); default red as written
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
