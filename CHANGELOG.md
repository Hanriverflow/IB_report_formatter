# Changelog

All notable changes to this project are documented in this file.

## [Unreleased] - 2026-09-15

### Office-letter production quality
- Render inline HTML breaks in paragraphs, emphasis, headings, lists and tables while preserving escaped/code literals and link destinations.
- Compose formal letter metadata, native attachment numbering, closing/signatory/issue/contact details, and a validated explicit `letter.appendix_heading` boundary through the shared renderer.
- Align wrapped metadata values and suppress the redundant company page header. Keep meaningful heading pagination controls; do not blanket-enable them on body text.
- Separate structural errors, pagination-marker warnings and visual-review status in `docx-audit`.
- Harden Windows Word QA: reject stale output directories, validate expected page counts, record file hashes and Word version, and leave visual review pending until pages are inspected.
- Add synthetic regressions and a complete appendix-letter example. Private transaction files are not fixtures. Word→Markdown remains retired.

## [2.0.0] - 2026-09-14

### Breaking changes
- Retired in-house Word→Markdown development by product decision. Removed the inverse CLI, parser, Markdown renderer, OMML reverse converter, roundtrip audit, and their dedicated tests. Recovery point: `d819bbb` / `codex/archive-word-to-md-d819bbb`.
- The built-in registry now supports only Markdown input and DOCX output. `docx-audit` replaces roundtrip checks with structural checks, not semantic reverse conversion.
- Sensitivity base-case highlighting now requires explicit coordinates; the engine no longer guesses the centre cell.

### Added
- Six document profiles: `ib-report`, `ib-memo`, `plain`, `office-letter`, `business-report`, `meeting-minutes`.
- Validated YAML metadata, layout settings, small YAML themes, CLI overrides, and strict diagnostics.
- A4 office layouts, recipients/sender/attachments, report and meeting metadata, native Korean multilevel numbering.
- Explicit table column roles, captions, units, sources, as-of dates, landscape sections, repeating headers, and native external hyperlinks/numeric footnotes.
- Shared model module, isolated request-local styles, examples, structural audit and profile regression tests.
- Repeatable Windows Word page-image QA script and long-table/landscape fixtures.

### Fixed
- Empty cells and escaped pipes keep their original columns; all-empty rows are retained.
- Financial number formatting reaches the actual output runs, including bold numbers and negative colouring.
- CLI, API and registry use one composition path with consistent headers, page fields and disclaimer toggles.
- General documents retain their References sections and do not use IB financial/header inference.
- Distribution packages explicitly include engine modules and exclude private reports and retired modules.
- Valid UTF-8 Korean is decoded before statistical detection; legacy EUC-KR/CP949 remains supported. Local images resolve relative to the Markdown file.
- Declared malformed YAML fails clearly; long YAML titles and explicit profile overrides retain the intended profile.
- Explicit footnotes work in headings and table headers; plain superscripts are not mistaken for footnotes.
- TOC preview entries are inside the field result and are replaced on update, without duplicate entries or the TOC listing itself.
- Neutral office titles no longer inherit the default Word title border; explicit fonts win over theme fonts. IB memo spacing and section-aware header alignment are refined.

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
