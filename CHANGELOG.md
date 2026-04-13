# Changelog

All notable changes to this project are documented in this file.

## [1.0.3] - 2026-04-13

### Fixed
- Improved unicode LaTeX fallback rendering so mixed Korean equations preserve readable Greek letters and math operators instead of emitting raw LaTeX commands.
- Added regression coverage for Korean math fallback cases such as `\alpha` and `\sum`.

## [1.0.2] - 2026-04-13

### Changed
- Installed `matplotlib` through the default dependency set so Markdown LaTeX renders during standard Markdown-to-Word conversion.
- Added a readable plain-text image fallback for LaTeX expressions that include Korean text or other non-ASCII content.
- Normalized CLI stdout and stderr to UTF-8 on Windows so logging and piped output stay stable with report titles and Unicode text.

### Fixed
- Removed standalone HTML anchor tags such as `<a id="제3장"></a>` before Markdown paragraphs are rendered into Word.
- Preserved smoke-test report generation on Windows even when console encoding cannot represent some characters directly.

### Documentation
- Updated installation guidance to reflect that LaTeX rendering is part of the default install path.

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
