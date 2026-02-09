<p align="right">
  <a href="./README.md"><img alt="lang English" src="https://img.shields.io/badge/lang-English-blue"></a>
  <a href="./README.ko.md"><img alt="lang 한국어" src="https://img.shields.io/badge/lang-한국어-orange"></a>
</p>

# IB Report Formatter

Bidirectional Markdown ↔ Word document converter for professional IB-style reports (`.docx`).

This project converts research and internal memo markdown into bank-style documents with structured headings, styled tables, callout boxes, images, equations, and footer/header formatting. It also supports the reverse: extracting clean Markdown from Word documents for LLM consumption.

## Features

- **Markdown → Word** conversion with IB-oriented document styling
- **Word → Markdown** conversion for LLM consumption (new!)
- Auto-format pass for single-line clipboard markdown (Deep Research output)
- YAML frontmatter parsing (`title`, `date`, `recipient`, `analyst`, etc.)
- Financial table rendering with number formatting and conditional styling
- Callout box rendering (`[Executive Summary]`, `[요약]`, `[시사점]`, `[주의]`, `[참고]`)
- Image rendering (local file paths and Base64 `data:image/...`)
- LaTeX support (`$inline$`, `$$block$$`, rendered via matplotlib when available)
- Header/footer support (company label, `CONFIDENTIAL`, page numbers)

## Project Structure

```text
IB_report_formatter/
├── md_to_word.py      # Markdown → Word CLI converter
├── md_parser.py       # Markdown/frontmatter/elements parser
├── md_formatter.py    # Single-line markdown pre-formatter
├── ib_renderer.py     # Word renderer and style system
├── word_to_md.py      # Word → Markdown CLI converter (new!)
├── word_parser.py     # Word document parser
├── md_renderer.py     # Markdown text renderer
├── tests/             # Pytest test suite
└── pyproject.toml     # Dependencies and tool config
```

## Requirements

- Python 3.8+
- [uv](https://docs.astral.sh/uv/) (recommended package manager)

## Installation on a New PC

Follow these steps to set up the project on any machine:

### 1. Install Python

Download and install Python 3.8 or higher from [python.org](https://www.python.org/downloads/).

Verify installation:

```bash
python --version
```

### 2. Install uv (Package Manager)

**Windows (PowerShell):**

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

**macOS / Linux:**

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Verify installation:

```bash
uv --version
```

### 3. Copy the Project

Copy the entire `IB_report_formatter` folder to your target machine, or clone from your repository:

```bash
git clone <your-repo-url> IB_report_formatter
cd IB_report_formatter
```

### 4. Install Dependencies

Navigate to the project folder and run:

```bash
uv sync
```

This creates a virtual environment and installs all required packages automatically.

**Optional:** Install full features (LaTeX rendering + robust encoding):

```bash
uv sync --extra full
```

**Optional:** Install dev/test tooling:

```bash
uv sync --extra dev
```

### 5. Verify Installation

```bash
uv run md_to_word.py --list
```

If successful, you'll see a list of available markdown files.

## What to Upload to GitHub

Upload only source/config files required to run the project on another PC.

Include:

- `md_to_word.py`
- `md_parser.py`
- `md_formatter.py`
- `ib_renderer.py`
- `word_to_md.py`
- `word_parser.py`
- `md_renderer.py`
- `tests/`
- `pyproject.toml`
- `uv.lock`
- `README.md`
- `README.ko.md`
- `AGENTS.md` (optional for contributor guidance)
- `docs/` (optional, after removing internal/private notes)

Do not include:

- `.venv/`, `__pycache__/`, `.pytest_cache/`, `.mypy_cache/`, `.ruff_cache/`
- `*.docx` output files
- local tool state (`.claude/`, `.sisyphus/`)
- private/raw business markdown files containing sensitive internal content

This repository now includes a root `.gitignore` configured to exclude those files by default.

## Quick Troubleshooting

| Issue | Solution |
|-------|----------|
| `uv: command not found` | Restart terminal after installing uv, or add uv to PATH |
| `python: command not found` | Install Python and ensure it's added to PATH |
| Permission errors on Windows | Run PowerShell as Administrator for uv install |
| Encoding errors with Korean files | Use `uv sync --extra full` for better encoding support |

## Quick Start

Convert markdown to Word:

```bash
uv run md_to_word.py input.md
```

Or use the script entrypoint:

```bash
uv run ib-report input.md
```

Specify output path:

```bash
uv run md_to_word.py input.md output.docx
```

Auto-format then convert:

```bash
uv run md_to_word.py input.md --format
```

What does `--format` (pre-formatting) do?

- It restores document structure from compressed or single-line markdown.
- Internally, it runs `md_formatter.py` first, creates an intermediate `*_formatted.md` file, then converts that file to Word.
- It is especially useful for copied/pasted Deep Research output where:
  - headings and paragraph boundaries are collapsed,
  - callout labels (`[요약]`, `[시사점]`, `[NOTE]`, etc.) are embedded mid-line,
  - LaTeX (`$...$`, `$$...$$`) and bold markers (`**...**`) are mixed in broken line flow.

Pre-formatting mainly performs:

- heading/subheading boundary detection and line-break insertion
- paragraph splitting based on sentence boundaries
- callout and bullet normalization
- LaTeX and bold token protection/restoration
- metadata extraction into YAML frontmatter

When should you use it?

- Input is mostly 1-5 long lines (or otherwise poorly structured)
- Direct conversion produces merged paragraphs/headings

When can you skip it?

- Markdown is already cleanly structured (headings, paragraphs, tables already separated)

Quick check before converting:

```bash
uv run md_formatter.py --check input.md
```

## Converter CLI (`md_to_word.py`)

```bash
uv run md_to_word.py [input_file] [output_file] [options]
```

Options:

- `-l, --list`: list markdown files in parent folder
- `-i, --interactive`: interactive selection mode (with `--list`)
- `-f, --format`: run formatter before conversion
- `--no-cover`: skip cover page
- `--no-toc`: skip table of contents
- `--no-disclaimer` / `--no-disc`: skip disclaimer page
- `-v, --verbose`: debug logs

Examples:

```bash
uv run md_to_word.py --list
uv run md_to_word.py --list -i
uv run md_to_word.py "네페스_기업분석2026.md"
uv run md_to_word.py report.md --format --no-toc
```

## Formatter CLI (`md_formatter.py`)

Format a raw markdown file:

```bash
uv run md_formatter.py input.md
```

Format into a specific output file:

```bash
uv run md_formatter.py input.md output_formatted.md
```

Check if formatting is needed:

```bash
uv run md_formatter.py --check input.md
```

Or use the script entrypoint:

```bash
uv run md-format --check input.md
```

## Word to Markdown CLI (`word_to_md.py`)

Convert Word documents to clean Markdown for LLM consumption:

```bash
uv run word_to_md.py [input_file] [output_file] [options]
```

Options:

- `-l, --list`: list Word files in parent folder
- `-i, --interactive`: interactive selection mode (with `--list`)
- `-s, --strip`: strip formatting (no bold/italic) for LLM optimization
- `--no-frontmatter`: skip YAML metadata header
- `--extract-images`: extract embedded images to folder
- `-v, --verbose`: debug logs

Examples:

```bash
uv run word_to_md.py --list
uv run word_to_md.py --list -i
uv run word_to_md.py report.docx
uv run word_to_md.py report.docx output.md
uv run word_to_md.py report.docx --strip              # LLM-optimized output
uv run word_to_md.py report.docx --strip --no-frontmatter
uv run word_to_md.py report.docx --extract-images     # Save images to folder
```

When to use `--strip`?

- When feeding the output to an LLM that doesn't benefit from bold/italic markers
- When you want cleaner, more compact text
- For RAG/embedding pipelines where formatting is noise

## Supported Markdown Patterns

- Headings: `#`, `##`, `###`, `####`
- Numbered heading-like lines (handled in parser/formatter)
- Paragraphs and list items
- Tables (generic, financial, risk/sensitivity patterns)
- Blockquotes for callouts
- Images:
  - `![alt](path/to/image.png)`
  - Base64: `![alt](data:image/png;base64,...)`
- LaTeX:
  - Inline: `$E=mc^2$`
  - Block: `$$\\int_a^b f(x)dx$$`

## Typical Workflow

1. If needed, normalize one-line markdown:
   `uv run md_formatter.py raw.md`
2. Convert formatted markdown to Word:
   `uv run md_to_word.py raw_formatted.md`
3. Open `.docx` in Word and update TOC field if required.

## Testing and Quality Checks

Run tests:

```bash
uv run pytest tests/ -v
```

Run type checking:

```bash
uv run mypy ib_renderer.py md_formatter.py md_parser.py md_to_word.py
```

## Notes

- If output file is locked (open in Word), the converter auto-saves with a timestamp suffix.
- LaTeX rendering requires `matplotlib` (`uv sync --extra full`). Without it, equations fall back gracefully.
- Encoding fallback includes `utf-8`, `utf-8-sig`, `euc-kr`, and `cp949` for Korean text robustness.
