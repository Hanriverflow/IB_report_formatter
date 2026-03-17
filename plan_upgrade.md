# Functional Upgrade Plan

## Goal

Improve bidirectional Markdown ↔ Word fidelity without breaking the current CLI flow.

## Implemented In This Pass

- Added IB-generated vs generic Word parsing flow in `word_parser.py`.
- Recovered IB cover metadata from generated documents:
  - `title`
  - `subtitle`
  - `ticker`
  - `company`
  - `sector`
  - `analyst`
  - `extra.date`
  - `extra.recipient`
  - `extra.report_type`
- Recovered generated `ENDNOTES` into `DocumentModel.footnotes` and excluded boilerplate sections such as cover tables, TOC, and disclaimer from parsed body content.
- Improved generic Word fidelity:
  - nested list level detection from indentation/numbering hints
  - table alignment extraction
  - table semantic reuse through shared markdown table heuristics
  - stricter 1x1 callout detection to avoid generic table misclassification
- Added italic and superscript parsing in `md_parser.py`.
- Added markdown endnotes rendering in `md_renderer.py`.
- Added batch directory mode to `md_to_word.py` and `word_to_md.py`.
- Added `roundtrip_audit.py` and the `roundtrip-audit` script entrypoint.
- Updated `README.md` and `README.ko.md` with batch and audit workflows.

## Verified

- `uv run --extra dev ruff check .`
- `uv run --extra dev mypy .`
- `uv run --extra dev pytest`

## Deferred

- Stronger image-caption binding beyond order preservation
- Generic Word header/footer metadata recovery
- Manual preview/diff UI beyond the current CLI audit summary
