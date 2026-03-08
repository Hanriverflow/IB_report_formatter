# IB Report Formatter Stability Plan

## Goal

Stabilize the existing Markdown ↔ Word pipelines before large-scale modularization.

## Completed In This Pass

- Added shared CLI helpers in `cli_utils.py` and removed duplicated logging/path/save/listing logic from `md_to_word.py` and `word_to_md.py`.
- Normalized formatter observability so `md_formatter.py` uses `logging` for operational messages while keeping explicit CLI result output.
- Fixed the Word → Markdown ordering bug by iterating Word document blocks in source order inside `word_parser.py`.
- Connected image extraction to actual markdown rendering so extracted images now produce `Image` elements and markdown links.
- Added platform-aware Korean font selection in `ib_renderer.py` with a macOS-first default (`Apple SD Gothic Neo`) and Windows-safe fallback behavior.
- Cleaned the current Ruff and mypy issues and updated `pyproject.toml` / `uv.lock` so the dev toolchain stays green.
- Added reverse-pipeline regression coverage in:
  - `tests/test_word_parser.py`
  - `tests/test_md_renderer.py`
  - `tests/test_word_to_md_cli.py`

## Verified

- `uv run --extra dev ruff check .`
- `uv run --extra dev mypy .`
- `uv run --extra dev pytest`
- `uv run --extra dev python md_to_word.py --help`
- `uv run --extra dev python word_to_md.py --help`
- `uv run --extra full python -c "import matplotlib, charset_normalizer; ..."`

## Deferred

- Split `ib_renderer.py`, `md_parser.py`, `md_formatter.py`, and `md_to_word.py` into smaller behavior-focused modules after the stabilized baseline is accepted.
