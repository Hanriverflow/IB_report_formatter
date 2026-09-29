# Changelog

All notable changes to this project are documented in this file.

## [Unreleased] - 2026-09-29

### Converted term-sheet input (found on a second real term sheet)
- `md-format --converted-term-sheet` cleans Markdown from external HWP/Word converters into term-sheet input: bold numbered lines, or numbered lines followed by a table, become `##` chapter headings; one-row band tables become headings; cover lines before the first chapter move to frontmatter (`confidential_label`, `title`, `subtitle`, `date`, `disclaimer`) and cover images are removed for the house `style.logo`.
- Every moved, rewritten, removed or kept line is reported with its input line number (a range for multi-line changes), including one-line numbered paragraphs left as text. Unclear cover lines stay in place, existing frontmatter lines and keys (case-insensitive) are kept, `prepared_by` is never inferred, and only top-level blocks change: indented code, list and quote blocks, fences and HTML tables (nested ones included; only clean one-row band tables become headings) stay as written. A YAML `...` closer becomes `---` and flow-style frontmatter is rewritten as block YAML, both reported. The parser's default path and the formatter's default output are unchanged.
- New engine module `converted_md_cleaner.py`.

### Term-sheet house style (found on a second real term sheet)
- House (and frontmatter) `style` options: a cover page (`cover: page`) with an optional logo (path or `data:` URI), a boxed disclaimer, header and footer accent rules, the confidentiality label colour, the footer page format (`- {page} -`) and position, a dark table header and open-sided tables. Frontmatter settings override the house file key by key; the defaults keep the standard layout unchanged.
- A null `style` value is a setting error (omit the key for the default), so it no longer passes validation and then stops the footer or drops the logo. `page_number` rejects every line break, including carriage returns. `label_color: ""` keeps the default grey, and frontmatter `logo: ""` removes a house logo.
- A landscape table that directly follows a lone page break (cover page, TOC) drops that page break, because its section break already starts a new page. Word 16 rendered the same pages before and after; the redundant break could show as a blank page in other viewers. A paragraph with several page breaks is kept.

### Numeric checks (found on a second real term sheet)
- `checks:` declares relations between term values (`all_in = issue_rate + credit_fee + running_cost`, `facility = amount * 1.05`). Values are read as displayed: Korean money units, `%`/`%p`, bp, `개월`/`년` and plain numbers. The default tolerance is half of the left value's display step, and failures are warnings that strict mode rejects.
- The checks and their generated values are stored in the DOCX (`ibrep.checks`), and `docx-audit` re-evaluates them with the current tagged values (`terms.failed_checks`).
- Table spec `schedule` checks a repayment schedule's arithmetic: balance steps, final zero, repayments equal to the principal (a number, or a money term converted with the table `unit`), the totals row and an optional stated weighted average life.
- New engine module `numeric_checks.py`. The term-sheet sample declares its checks and its schedule.

### Table structure (found on a second real term sheet)
- Table spec `header_rows` draws several header rows in every profile: shaded, merged where spans say so, and repeated on every page. HTML tables keep the rows their header spans cover as header rows instead of joining them into one. `base_case` rows count from the first body row, and zebra shading starts on the first body row.
- Term sheets: a cell that starts in the second label column and spans into the content columns is laid out as content; a first-column label stays a label however far it spans.
- Term-sheet grid tables size each column to fit its lines and share the spare width equally, so short amount columns are no longer starved while date and label columns balloon. Tables whose content cannot fit side by side keep the content estimate.
- Column kinds skip dash placeholders (`-`, `–`), so an amount column that starts with them is still numeric and right-aligned.

### Input loss in converted documents (found on a converted real term sheet)
- `*` and `_` emphasis needs flanking delimiters (CommonMark): an opening delimiter followed by whitespace, or a closing one preceded by it, stays literal. `2 * 3 * 4` and spaced note markers (`매출처* ...`, `* 주요 매출처`) keep their asterisks instead of losing them to a false italic; `** text **` is no longer bold.
- HTML `<table>` blocks, as HWP/Word converters write them, become Word tables in every profile: `colspan`/`rowspan` merge cells (empty rows under a span keep their place), block tags (`<p>`, `<div>`, `<li>`, nested rows and tables) and `<br>` break cell lines, `<b>`/`<strong>`, `<i>`/`<em>`, `<code>`, `<sup>`/`<sub>`, colour spans and `<a href>` format text and combine when nested, `<img>` becomes a cell image (`data:` URIs included), `{{key}}` references are substituted, and a nested table is flattened into its cell. Cells are built directly as runs, so HTML text (including decoded `&lt;tags&gt;` and Markdown characters) is never re-read as Markdown. Several header rows merge into one header row. Text outside cells, extra tables on a closing line and unclosed tables are reported, and strict mode rejects them.
- Images inside text and table cells (`![alt](path)`) are inserted inline and fitted to the cell width; a failed image leaves a visible marker that strict mode rejects. An image is one unit: `$` in its path is not math, surrounding emphasis still applies, its fields are never term-substituted, and image-like text inside a link destination stays part of the link. A standalone `<img>` line is an image.
- Image paths written percent-encoded (`images/a%20b.png`) find their file; a name that really contains `%` is still tried first.

### Internal memo rendering (found on a real memo)
- Inline code is rendered without its backticks in the code font (`CODE_FONT`); its text, including edge spaces in term-sheet lines, stays literal and outside table number formatting. Plain-profile callouts render inline code and links like body text.
- Local file links (angle-bracket destinations, drive or `./`/`../` paths, document extensions) become Word hyperlinks. Saving through the CLI, converter or registry rebases links on the DOCX folder: relative links keep reaching their file, absolute paths inside the DOCX folder become relative, and other absolute paths stay `file:///` links with a log warning. `#` and `%` in absolute paths are literal; a `#` after a document extension in a relative link is a fragment, and `file:` fragments such as `#page=2` are kept. UNC paths become `file://server/share` links; targets that are not valid local paths are kept as written with a warning. Reference definitions with spaces, brackets or any path form resolve correctly.
- General profiles turn off Word's automatic Korean/Latin and Korean/number spacing (`SPC는`, `제2종`, `300억원`), as the term-sheet profile already did.

### Existing-output fixes (schema order, forced TOC, audit encoding)
- Insert table borders, cell fills and margins, paragraph borders and run styles in ECMA-376 schema order in every profile (`ooxml_order.insert_ordered`). Output changes only in element order; content is identical.
- `business-report`, `meeting-minutes`, `office-letter` and `ib-memo` without a cover now render their title opening before a forced TOC, and the TOC preview no longer lists the title heading (matching `ib-report` and `term-sheet`).
- `docx-audit` writes its JSON report as UTF-8 regardless of the console code page, so redirected output parses on Windows (cp949).

### Term-sheet profile, explicit cell spans and house boilerplate
- Add the seventh profile, `term-sheet`, with validated house/frontmatter text, source-relative house paths and a `--house` override. Real deal documents and institution wording stay outside the repository.
- Resolve explicit `^^`/`<<` cell spans and label columns; support escaped literal markers, rectangular merge validation and per-table opt-in for other profiles.
- Render term-sheet documents: opening block (title, subtitle, date, prepared-by, disclaimer), confidentiality header and version/page footer on every section, labelled term tables with fixed label widths, label tiers, merged cells, per-line hanging indents and estimated row splitting, a two-row confirmation box, and Korean word-boundary wrapping without automatic Latin/number spacing.
- Merge validated spans in every profile's tables; add the optional table `note` (below the table, right-aligned) for every profile.
- Add a fictional Korean ABCP sample, repayment schedule and freshly written house boilerplate, plus bilingual authoring documentation.

### Term variables and consistency checks
- Add `terms:` values referenced as `{{key}}` in every profile, substituted at parse time as literal runs (never re-parsed, excluded from number formatting). Undefined keys warn and fail strict mode.
- Wrap substituted values in Word content controls tagged `ibrep:term:<key>` and snapshot the generated values in `ibrep.term.<key>` custom properties; `--no-term-tags` / `layout.term_tags: false` emits plain text.
- `docx-audit` reports `mismatched`, `changed`, `missing` and `indicative` term values after Word edits as warnings; documents without terms keep the previous JSON output.

### Title block without a cover (owner decision, 2026-09-29)
- When the cover is not rendered (`--no-cover`, `termsheet`/`legal-memo` presets), `ib-report` now starts with a title block (title, subtitle, date/author) before the TOC instead of omitting the title. The block is not a TOC entry; a matching H1 or inferred subtitle appears once. Cover-on output is unchanged.

### Korean glyphs in chart, equation and diagram images
- Rasterized images now use an installed CJK-capable font (preferred theme font first, then Malgun Gothic, Apple SD Gothic Neo, Nanum, Noto/Source Han Sans KR). Previously Linux fell back to DejaVu Sans and dropped every Hangul glyph. DOCX font declarations are unchanged.
- If Hangul must be rasterized and no CJK font is installed, strict mode rejects before saving; non-strict mode warns. Linux CI installs `fonts-nanum`.

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
