# Themes, presets and charts port verification — 2026-09-29

## Scope and design

Implemented in the existing worktree on `codex/port-themes-presets-charts`, based on `bd1139c`. No commit, push, stash, reset, other-checkout writes, GitHub access or subagents were used. Historical implementations and all six requested historical test files were read with local `git show`; the old renderer/CLI diff was also inspected.

### Charts

- `Chart`, `ChartSpec`, `ChartSeries` and `ElementType.CHART` represent chart fences. Parsing retains the source and a validated specification or deferred diagnostic, independently of the rendering switch. This lets an explicit render option override frontmatter on the same parsed model.
- Rendering remains inside `IBDocumentRenderer.render`, dispatched to `ChartRenderer`. `--charts`, `--no-charts`, `RenderOptions.charts` and top-level YAML `charts` resolve explicitly; the default is off for all six profiles.
- Historical PR #5 actually uses `chart_type` and `y_label`. Both are preserved; `type` is an alias. `unit` and presentation-only `number_format` extend that schema. Waterfall inputs remain sequential deltas from zero, with an appended total; `[100, -30, 20]` produces `90`, without rescaling.
- Charts use `Figure` + `FigureCanvasAgg` + `BytesIO`, with per-artist fonts/colors. No pyplot, global rcParams changes, image temp files or cached style values are introduced. The existing theme-dependent FontPolicy cache was removed. General/mono palettes are neutral; explicit themes control chart colors.
- Invalid specs or render/insertion failures produce a named renderer diagnostic. Non-strict output contains the source code panel; strict rendering raises before the converter can save, including when an existing output must survive. Disabled fences use the exact existing code-panel path.

### Presets

- `PRESETS` is a read-only mapping of frozen `RenderOptions` bundles: `ib-report` (no overrides), `termsheet` (all three sections off), `legal-memo` (TOC only), `lecture-note` (cover and TOC).
- Resolution per section: explicit CLI/API field > caller preset field > YAML layout field > YAML preset field > profile default. The no-op `ib-report` preset leaves lower-precedence values intact.
- CLI listing, named selection, frontmatter and registry output all use this shared resolver. Presets do not select profiles, alter parsing/table semantics, add metadata, or replace the assembler.

### Themes

- `load_style` still returns a fresh frozen style consumed through `render_styles.use_style` / ContextVar. Existing lowercase aliases, default and mono remain available.
- Uppercase presentation fields now cover colors, fonts, point sizes/spacing, inch lengths, list presentation, table/heading/body toggles and labels. Validation rejects unknown fields, non-string colors, nonfinite sizes, booleans used as numbers, invalid ranges and conflicting aliases.
- Existing heading/callout properties already resolved styles dynamically; they were retained and tested. RGB/hex pairs synchronize when only one is supplied. IB table-header background follows NAVY unless explicit; waterfall negative color follows RED unless explicit. Code panels now honor CODE_BG.
- Internal `STYLE_*` names and `NATIVE_NUMBERING` are intentionally excluded from theme YAML: they are engine identifiers/profile policy, not presentation choices. External preset YAML files and the old global `theme_loader`/`preset_loader` are not restored; named immutable bundles and the current loader provide the requested capabilities.

## Changed files

Line numbers refer to this uncommitted implementation.

| File:line | Change |
|---|---|
| `document_model.py:104` | Chart series/specification/source models and element union; CHART enum at line 35 |
| `chart_renderer.py:41` | Schema validation; waterfall arithmetic at 128; Agg rendering at 169 |
| `md_parser.py:1470` | Proper chart elements; typed charts/preset frontmatter at 207 |
| `ib_renderer.py:2691` | ChartRenderer; CHART dispatch at 2985; themed code background and uncached font policy |
| `document_profiles.py:60` | Options; immutable presets at 75; precedence at 114; theme extension at 164/286 |
| `render_styles.py:26` | Request-local chart negative color |
| `md_to_word.py:420` | CLI chart/preset flags and listing; shared option forwarding at 497 |
| `converters.py:197` | Registry chart/preset option forwarding |
| `pyproject.toml:57` | Explicit chart module in wheel only-include and sdist include at 66 |
| `tests/test_chart_port.py:1` | 45 chart cases through parsing, DOCX serialization, converter, registry and CLI |
| `tests/test_preset_port.py:1` | 38 preset cases: all 24 preset/profile combinations, precedence, CLI and immutability |
| `tests/test_theme_port.py:1` | 24 theme cases: actual output, validation, concurrency and CLI override/save safety |
| `tests/test_render_hardening.py:367` | Add charts to the existing Matplotlib global-state regression |
| `samples/qa/charts.md:1` | Fictional Korean bar/line/waterfall sample; expected three pages for native QA |
| `README.md:126`, `README.ko.md:101` | Usage, schema, precedence, extended keys and limitations |
| `CHANGELOG.md:7` | Unreleased feature-port entry |

## TDD and automated results

Used only `.venv-codex\Scripts\python.exe` (Python 3.12.13). Before each Python command, TEMP/TMP pointed to the resolved `.uv-cache` directory. Every pytest run used a fresh `.uv-cache\run-*\bt` base directory with `-p no:cacheprovider`.

1. Before implementation, the three new suites produced **68 failed, 22 passed**: charts 30 failures, presets 33, themes 5. Missing options/model/flags and unsupported uppercase theme fields caused the expected failures. Evidence: `.uv-cache/port-red.log`.
2. After implementation and correcting test fixture syntax/log capture, those suites plus rendering hardening produced **129 passed**. Evidence: `.uv-cache/port-green.log`.
3. Additional tests first reproduced two gaps (no-op caller preset inheritance and CODE_BG not honored), then one legacy RED/waterfall mismatch. All were fixed and pass in the final suite. Evidence: `.uv-cache/port-red-followup.log`, `.uv-cache/theme-chart-red.log`.
4. Final suite: **562 passed, 1 skipped**, 21.85 seconds. The skip is the existing POSIX permission-bit test on Windows. Evidence: `.uv-cache/port-full-final.log`. Compared with the starting 453 passed / 1 skipped, there are 109 additional passing cases (107 in new suites, one Matplotlib case, one newly discovered sample).
5. Final `python -m ruff check .`: passed. Final `python -m mypy .`: passed, 15 source files; pre-existing notes about untyped formatter bodies remain.
6. The existing Python 3.8 AST gate and exact wheel-payload gate tests passed as part of the suite. No Python 3.8 interpreter runtime was executed here.

Coverage includes PNG insertion via reopened DOCX, all three chart types and both type spellings, Korean glyph-warning checks when the font is installed, no scale conversion, exact waterfall coordinates/negative totals, neutral/theme palettes, saved code fallbacks, strict destination preservation, backend/rcParams preservation, failed-figure cleanup and six-profile disabled-chart XML equivalence.

## Output parity

Before editing, the required CLI rendered all nine existing `samples/profiles/*.md` and `samples/qa/*.md` into `.uv-cache/parity-before`. Additional baselines used explicit `default` and `mono` themes in sibling directories. Final code regenerated the same inputs in `.uv-cache/parity-after` and corresponding theme directories.

Result: **81/81 parts identical across 27 documents** (9 samples × 3 theme choices × 3 parts):

- `word/styles.xml` and `word/numbering.xml`: byte-for-byte identical without normalization.
- `word/document.xml`: byte-for-byte identical after replacing only the four random hex characters at the end of `_ibrep_...` semantic bookmark names. No layout, content, style, field or numbering data was excluded.
- Chart-off output was also compared directly against the legacy CodeBlock render path for every profile with a seeded bookmark generator: all three XML parts matched.
- The new charts sample has no pre-port baseline and is excluded from the nine-sample equality claim.

Comparison script: `.uv-cache/verify_parity.py`; normalized part hashes and results: `.uv-cache/parity-result.json`.

## Visual verification and remaining environment limitations

Three review DOCX files were generated in `.uv-cache/visual-input`: general charts, custom-theme IB charts with the termsheet preset, and a plain lecture-note preset document. Native QA was attempted with the repository's `scripts/word_visual_qa.ps1`, a new empty `.uv-cache/visual-pages` directory and `-ExpectedPages 3`.

**Word page review remains pending.** Word.Application could not start: COM `0x80070520`, “A specified logon session does not exist.” There was no usable LibreOffice fallback. Actual page counts, Word version and page images are unavailable; no page-layout pass is claimed. The expected count is an assertion to verify later, not an observation. The attempt record `.uv-cache/visual-attempt.json` records DOCX hashes, expected counts, unavailable fields and the failure.

All six embedded chart PNGs (three general, three custom-theme IB) were individually inspected: Korean labels, legends, source labels, plotted values and the 90 waterfall total are legible, with no visible clipping or overlap in those chart images. This is asset inspection only. `.uv-cache/chart-asset-review.json` records image hashes and ties that inspection to the attempt-manifest SHA256; it leaves page review pending.

**Distribution build was skipped as requested when unavailable:** neither `build` nor `hatchling` is installed in `.venv-codex`. No wheel/sdist or wheel-install result is claimed. Both package include lists contain the new engine module, and CI gate regression tests pass, but an actual distribution build remains to be run in a provisioned environment.

All logs, generated documents, PNGs and probes remain in excluded `.uv-cache`; only synthetic source samples and this verification record are tracked candidates. Word-to-Markdown, version metadata and excluded owner files were not changed.
