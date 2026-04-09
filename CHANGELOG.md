# Changelog

All notable changes to this project are documented in this file.

## [1.0.2] - 2026-04-03

### Fixed
- Preserved intentional blank markdown table cells during parsing so Word table headers and body fields stay aligned.
- Forced UTF-8 stdout writes for piped `word_to_md.py` output on Windows so Korean markdown round-trips cleanly through CLI pipes.

## [1.0.1] - 2026-03-24

### Changed
- Improved Markdown-to-Word table rendering so body columns use content-aware widths instead of near-uniform spacing.
- Inferred table body alignment from the first 2-3 data rows so textual columns stay left-aligned while numeric columns align right.
- Switched cover page and table-of-contents typography to `Malgun Gothic`, including Word-generated `TOC 1` to `TOC 4` styles.

### Fixed
- Reduced unnecessary line wrapping in narrow numeric columns and long descriptive table columns.
- Prevented mixed digit-text codes such as `A-101` from being misclassified as numeric body columns.

### Documentation
- Synced project progress notes in `next_step.md` and `roadmap.md` with the current `main` implementation state.
