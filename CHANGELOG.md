# Changelog

All notable changes to this project are documented in this file.

## [Unreleased] - 2026-09-29

### Python 3.12 floor (owner decision D1, 2026-09-29)
- Require Python 3.12+ (`requires-python >=3.12`); CI tests Ubuntu/Windows on Python 3.12 and 3.13, and the syntax gate checks 3.12. The lock file only drops old-Python resolution forks; dependency versions are unchanged.
- Python 3.8–3.11 are no longer supported; use an earlier source revision for those runtimes.

### Themes, section presets and opt-in charts
- Port PR #5 chart fences as proper model elements through the shared renderer: grouped bar, line and cumulative waterfall, retained YAML syntax plus `type`, `unit` and unscaled number formats. Charts default off; CLI/API options override frontmatter.
- Render charts to in-memory PNGs with Figure/Agg and per-artist Korean fonts/colors, without pyplot or rcParams mutation. Strict failures reject before saving; non-strict failures retain the source code panel and a named diagnostic.
- Add immutable `ib-report`, `termsheet`, `legal-memo` and `lecture-note` section bundles, `--preset` / `--list-presets`, and frontmatter `preset`. Explicit fields override caller presets, then YAML fields/presets, then profile defaults.
- Extend request-local themes with typed uppercase presentation fields, stricter color/size validation, paired RGB/hex handling, and themed code-panel backgrounds. Preserve the default/mono output and general-profile semantics.
- Package the chart engine explicitly and add a fictional Korean chart sample, usage documentation and parse/render/CLI regressions. No retired Word-to-Markdown code or global-style loader is restored.

### Input-loss and correctness hardening (2026-09-29)
- Stop silent content loss: escaped dollars (`\$5`), escaped emphasis, prose/tables after a `References` list, unrecognized leading `**Label:** value` paragraphs, dash-only table body rows, list continuation lines, and CRLF list items.
- Numeric table cells keep native footnote references and are no longer re-formatted from concatenated run text (`1234[^1]` stays `1,234` plus the footnote).
- Equation rendering failures are renderer errors: strict mode rejects them before saving; non-strict keeps visible fallback text.
- Shared fence scanner (backtick/tilde, 4+ fences, paragraph interruption); setext `===` H1; HTML comments are not printed; reference-style links resolve to hyperlinks; image titles and `<angle-bracket>` destinations parse correctly; literal `[^n]` inside code is not a footnote.
- Consecutive leading metadata lines (`**Date:**`/`**Analyst:**`) populate separate fields. Recognized header labels are canonicalized (`date`, `analysis_period`, `analysis_basis`). An inferred IB subtitle stays on the cover and becomes a body heading when the cover is disabled.
- Nested ordered lists restart under each parent item; start overrides target the actual level.

### Robustness and performance (2026-09-29)
- Atomic saves through a sibling temp file and replace (existing reports survive failed writes; file mode preserved); exclusive, collision-resistant locked-file fallback names.
- Batch mode rejects colliding output names before writing and exits non-zero when any input fails. Missing qualified input paths no longer fall back to an unrelated same-named file.
- Thread-safe default converter registry initialization. Diagram/equation rendering no longer mutates global Matplotlib state and releases figures/temp files on failure.
- Render-path structural validation no longer resolves styles per paragraph: a 2,000-element document renders about 60% faster. `docx-audit` validates numbering references and reports placeholder-like user text as warnings.
- GitHub Actions workflow (Ubuntu/Windows): ruff, mypy, pytest, build, exact wheel payload, Python 3.8 syntax gate.

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
