# IB Report Formatter 2.0

Markdown → editable Word for investment-banking reports and Korean business documents.

[한국어 설명](README.ko.md) · [Current integration record](docs/pr-consolidation-20260926.md) · [Release checklist](#release-checklist)

## Status and documentation map

Repository review on **2026-09-27**: the default branch is `main`, the README calls the writer-only architecture **2.0**, and [package metadata](pyproject.toml) declares **2.0.0 (Beta)**. GitHub showed **0 tags and no published releases**. These version labels describe source/package metadata, not an already published or newly certified v2.0 release.

- **Current usage and implementation:** this README, [Korean usage](README.ko.md), [runnable profiles](samples/profiles), [CLI source](md_to_word.py), [profile definitions](document_profiles.py), and [tests](tests). The built-in direction is Markdown to DOCX; Word-to-Markdown remains retired.
- **Latest integration and reported verification:** [2026-09-26 consolidation](docs/pr-consolidation-20260926.md), including promoted functionality and deliberately deferred old-PR features. It reports 311 tests passed at product-code commit `cadc0ae515d2f5757d5006e64bec1acaff56ccba`, package/install checks, and nine synthetic documents / 16 pages of Word evidence. These are that record's results, not tests rerun during this documentation review or human/business approval.
- **Design and earlier evidence:** [September implementation plan](docs/implementation-plan-20260914.md), [September 14 verification](docs/verification-20260914.md), [letter improvement plan](docs/improvement-plan-20260915.md), and [September 15 verification](docs/verification-20260915.md). Plans explain intended work; dated verification records describe only their recorded scope and environment.
- **Preserved historical plans:** [stability plan](plan.md), [upgrade plan](plan_upgrade.md), [next steps](next_step.md), and [roadmap](roadmap.md). Their bidirectional features, old test counts, and checkmarks are not the current feature list or a release guarantee.
- **Preserved upload checklist:** [git_checklist.md](git_checklist.md) is an older repository-upload guide, not evidence that release gates have passed. Use the release checklist below for a future release. Do not reinitialize this repository or blindly stage all local files.
- **Change history:** [CHANGELOG.md](CHANGELOG.md). A dated entry or completed plan alone does not prove a tag, published artifact, or successful validation of today's checkout.

## Release checklist

This is a **pending maintainer workflow**, not a completed deployment. This documentation-only review did not run tests, build/install packages, render Word pages, publish packages, or create tags/releases.

### Local validation commands

Run from a clean, isolated checkout of the exact proposed release commit, using both Python 3.12 and 3.13. Preserve command output, tool versions, exit codes, and artifact hashes. These commands are instructions, not fresh pass results:

```sh
uv sync --locked --extra dev
uv run pytest tests/
uv run ruff check .
uv run mypy .
uv run md-to-word --help
uv run md-to-word --list-profiles
uv run md-to-word samples/profiles/office-letter.md release-letter.docx --strict
uv run docx-audit release-letter.docx
uv build
```

Keep generated output local and outside commits. For repeated runs, choose a new output path. Inspect the audit JSON; successful DOCX generation is not proof of acceptable page layout. See [Word page QA on Windows](#word-page-qa-on-windows) for native rendering with new/empty evidence directories and separate visual review.

- [ ] Record the exact release commit and reconcile both READMEs, package version and changelog with its implemented scope, not historical plans.
- [ ] Run the checks above; investigate every failure. Recheck the declared Python/OS support on actual target runtimes; Python 3.12 syntax checks alone are not runtime validation.
- [ ] Inspect wheel/sdist contents for missing engine modules and accidental credentials, private reports, generated DOCX files, or retired inverse-conversion modules. Install the built wheel in a separate clean environment and smoke-test its installed CLI outside the source checkout.
- [ ] Regenerate representative synthetic samples; inspect structural audit results and every rendered Word page. Record Word/font environment, manifests, hashes, and separate visual-review/business-approval status.
- [ ] Resolve the rights and notice gate below, including third-party dependencies, fonts, images, diagrams and sample content.
- [ ] Only after those gates pass, have the maintainer choose a version consistent with package metadata and changelog, create a tag on the **validated commit**, and publish release notes with the exact validation evidence, limitations and artifact hashes. Do not infer or create `v2.0`/`v2.0.0` merely from the README heading. Never move an existing release tag to hide a failed validation; document a corrected release instead.

### License and rights gate

The project is licensed under the [MIT License](LICENSE), matching the existing `license = {text = "MIT"}` declaration and classifier in [pyproject.toml](pyproject.toml). On 2026-09-29 the rights holder confirmed the copyright notice **Copyright (c) 2026 Hank**; the `LICENSE` file carries it and is included in the sdist and wheel metadata.

**[open] Third-party notices:** dependency, font, image, diagram and sample-content licenses are separate from the project license. Confirm them before a public release; the MIT notice does not cover third-party material.

## Product direction

This is a **Markdown-to-Word-only engine**. In-house Word→Markdown development has been retired deliberately: established external projects already address that problem. Version 2 removes the inverse parser/CLI, Markdown output converter, OMML reverse converter and roundtrip audit. There is no plan to rebuild them.

## Quick start

Requires Python 3.12+. CI checks Python 3.12 and 3.13 on Ubuntu and Windows. Install with [uv](https://docs.astral.sh/uv/):

```sh
uv sync
uv run md-to-word samples/profiles/office-letter.md letter.docx --strict
uv run md-to-word samples/profiles/ib-memo.md memo.docx --strict
uv run md-to-word input.md output.docx --profile plain
uv run md-to-word --list-profiles
uv run docx-audit letter.docx
```

The existing `uv run md_to_word.py ...` and `uv run ib-report ...` commands still work. Word is not required to generate DOCX; it is useful for reviewing pagination and updating fields. Files locked by Word are saved to a timestamp-suffixed path, reported in the log.

## Profiles

| Profile | Default composition | Page |
|---|---|---|
| `ib-report` | Legacy IB cover, TOC, disclaimer, confidentiality label | Letter |
| `ib-memo` | Compact IB title/author/date, confidentiality label; no cover/TOC/disclaimer | A4 |
| `plain` | Body only; no IB metadata inference | A4 |
| `office-letter` | Company, document number, recipients, title, attachments, sender | A4 |
| `business-report` | Title, date, department, author, body | A4 |
| `meeting-minutes` | Title, time, place, attendees, author, body | A4 |
| `term-sheet` | Deal opening, house boilerplate, merged term tables, confirmation box | A4 |

Profile defaults do not supply factual document content. The four general profiles do not infer financial tables or turn short numbered items into headings. They use neutral styling and editable Word multilevel numbering (decimal → Korean 가나다 → parenthesized decimal). Nest lists using four spaces per level. Legacy IB list behaviour is retained.

## Frontmatter

Use [the profile examples](samples/profiles) as starting points:

```yaml
---
profile: office-letter
title: 자료 제출 요청
document_no: 기획-2026-015
date: "2026-09-14"
sender:
  organization: 주식회사 예시
  department: 경영기획팀
  signatory: 주식회사 예시 대표이사
  contact: planning@example.com
recipients: [협력회사 담당부서장]
cc: [재무담당자]
attachments: [제출 양식 1부]
---
```

Letters require `recipients` and `sender.organization`; no seal/signature image is manufactured. Reports use `analyst`, `date`, `sender.department`; minutes also accept `meeting_time`, `location`, `attendees`. Quote dates, codes and numeric-looking identifiers to preserve their text. A matching first H1 is not duplicated when the profile already supplies a title.

### Letters with an appendix

The letter profile uses a centred company name, aligned recipient/reference/subject fields, native numbered attachments, a closing marker, signatory and issue/contact details. Wrapped metadata aligns below the value. The company name is not repeated in the page header. `sender.signatory` accepts a YAML multiline string; no signer name is invented.

To place the closing **before** an appendix, specify its exact H1 text:

```yaml
letter:
  appendix_heading: 운영자료 확인 내역
  appendix_label: 붙임 1
```

That H1 must appear exactly once after the letter body and differ from the document title. Missing/duplicate/empty-body boundaries and unknown `letter` keys are rejected before saving. `appendix_label` is optional and requires `appendix_heading`. Without these settings the entire Markdown remains the letter body. Attachments are a declared list, not proof that external files exist or have been embedded. See [the complete synthetic example](samples/qa/office-letter-appendix.md).

Configuration precedence is **explicit CLI/API option > YAML > profile default**. YAML `layout` supports boolean `cover`, `toc`, `disclaimer`, `confidential`, `strict`, plus `separator_mode: auto|rule|page-break`. Unknown layout/table/theme keys fail validation. CLI overrides: `--profile`, `--theme`, `--strict`, `--no-cover`, `--no-toc`, `--no-disclaimer`, `--no-confidential`, `--separator-mode`. Enable cover/TOC through YAML or API when needed.

Themes: `default`, `mono`, or a small YAML file using `body_font`, `heading_font`, `korean_font`, `body_size`, `primary_color`, `margin_mm`. See [company-theme.yaml](samples/profiles/company-theme.yaml). YAML theme paths are relative to the input Markdown; CLI paths are relative to the working directory. This is not an arbitrary DOCX template importer.

## Term sheet

Use `profile: term-sheet` for Korean structured-finance terms. See the [fictional ABCP sample](samples/profiles/term-sheet.md) and [fictional house file](samples/profiles/term-sheet-house.yaml).

```yaml
---
profile: term-sheet
title: "가나다머티리얼즈㈜ ABCP 300억원"
date: "2026-10-16"
version: v1
house: term-sheet-house.yaml
tables:
  - {}
  - {columns: [text, date, date, number, number, number], unit: 억원}
---
```

`title` is required; `subtitle` defaults to `Term Sheet`. Optional `date` is displayed as written and `version` labels the footer. `prepared_by` and `disclaimer` must be nonblank in frontmatter or the house file. House YAML accepts only `prepared_by`, `disclaimer`, `confidential_label`, and `confirmation` (`intro`, a list of `items`, `signature`). Frontmatter keys override house values, including an explicitly empty value. House prose supports inline emphasis such as `( *주요내용* )`.

Keep real deal documents and real institution house files **outside the repository**. Frontmatter `house` paths resolve relative to the source Markdown; `--house /absolute/path/to/house.yaml` overrides that file, with relative CLI paths resolved from the working directory. String/stream input without a source path needs an absolute house path. The checked-in house file contains fictional wording only.

Write numbered sections explicitly with `##` and subsections with `###`. Table specifications follow body-table order; retain `{}` for tables with no overrides. The snippet above illustrates a two-table document; use the complete sample's ordered specifications when copying its tables.

- A cell containing only `^^` merges upward; `<<` merges left. Merges must form rectangles and cannot cross the header/body boundary. Empty cells stay empty. Use `\^^` and `\<<` for literal markers. Invalid groups retain their source markers with warnings; strict mode rejects them.
- Spans default on for `term-sheet`; `spans: false` disables them. Other profiles enable them only with `spans: true`.
- `label_columns` overrides the leading label count (integer from 0 to one less than the column count). Otherwise the first header cell's resolved span determines it, capped at column count minus one. Thus `| 구 분 | << | 내 용 |` gives two labels; `| 구 분 | 내 용 |` gives one.
- Key-value tables have one or two label columns and exactly one content column. The first label is shaded/bold, with a fixed 33.5 mm width; the second is white, regular/muted, 30 mm wide. Other grid tables use content-based widths and regular-weight shaded labels. Use `label_columns: 0` for no label shading.
- Use `columns` roles for dates, codes and numeric values, and `unit` for units above the table. `note` is optional text placed below the table, right-aligned and muted (8 pt); every profile supports it. Financial meanings and repayment schedules are not inferred or calculated.

Use `<br>` inside cells; in term-sheet body paragraphs it also starts a separate Word paragraph. Line markers receive hanging indents: `•`, `-`, `·` start at 0/3/6 mm with a 3 mm hang; `①`–`⑳` start at 0 with a 4.5 mm hang; `※` starts at 0 with a 4 mm hang and 8 pt text. Wrapped text aligns after the marker. A body Markdown `-` list retains normal list handling.

Place an empty, closed fence where the house/frontmatter confirmation should appear:

````markdown
```confirmation
```
````

The box has an introduction/checklist row and a shaded, centred signature row kept on one page. Missing text, a nonempty fence or an unclosed fence preserves the source in a diagnostic code panel; strict rendering rejects it. Other profiles treat the fence as code.

Every page uses `confidential_label` (default `Strictly Confidential`) in the header. An empty label, `layout.confidential: false`, or `--no-confidential` suppresses it. The footer shows `subtitle version` when `version` is present, plus `PAGE / NUMPAGES`. The opening disclaimer remains required even with `--no-disclaimer`. The existing **`termsheet` preset** only toggles cover/TOC/end-disclaimer sections; it does not select this profile, load house text or apply term-sheet table formatting.

Term-sheet documents wrap Korean at word boundaries (`w:wordWrap=1`) and turn off Word's automatic spacing between Korean and Latin text or digits (`w:autoSpaceDE/DN=0`), so amounts print as `300억원` and `SPC에`.

### Term variables and consistency checks

Define values that repeat across the document (amounts, rates, dates) once under frontmatter `terms:` and reference them as `{{key}}`. Every profile supports this.

```yaml
terms:
  amount: "300억원"
  cap_spread: "[1.10]%p"
```

- Keys start with a lowercase letter and use lowercase letters, digits and underscores (at most 40 characters). Values must be single-line strings; quote anything YAML would read as a number (`1.10` would otherwise become `1.1`).
- References are substituted in the title, subtitle and date, headings, paragraphs, lists, table cells and blockquotes. Table captions, units, sources and as-of dates receive plain, untagged text. Inline code, code blocks, math and link destinations are never substituted. `\{{` stays literal.
- Values are inserted exactly as written: they are not parsed as Markdown and table number formatting does not apply. Surrounding bold/italic is inherited.
- An undefined key is left as `{{key}}` with a warning, and strict mode rejects it. Without `terms:`, `{{…}}` is ordinary text; `terms: {}` turns the feature on and checks every reference.
- Each substituted value is wrapped in a Word content control tagged `ibrep:term:<key>`, and the generated value is recorded in the custom document property `ibrep.term.<key>`. For a clean copy without controls, use `--no-term-tags` or `layout: {term_tags: false}`.
- After editing in Word, `docx-audit` adds a `terms` object to its JSON: `mismatched` (one key with different values), `changed` (values edited since generation, to carry back into the YAML), `missing` (keys whose controls were all removed) and `indicative` (values still in brackets). These are warnings; they do not affect `issues` or the exit code. The check covers consistency between tagged values only, not financial correctness.

## Charts (opt-in)

```sh
md-to-word report.md report.docx --charts --strict
md-to-word samples/qa/charts.md charts.docx --strict
md-to-word report.md code-panels.docx --no-charts
```

Charts are **off by default for all seven profiles**. Enable with `--charts`, `RenderOptions(charts=True)`, or top-level frontmatter `charts: true`. An explicit CLI/API value wins over frontmatter; `--no-charts` / `charts=False` forces code panels. The parser retains the original fence in a chart model element so the same parsed model supports either setting.

````markdown
```chart
chart_type: bar
title: 가상 매출
labels: [상반기, 하반기]
series:
  - name: 매출
    values: [100, 120]
y_label: 백만원
source: 기능 검증용 가상 자료
number_format: ',.1f'
```
````

Supported types are `bar` (grouped series), `line`, and `waterfall` (exactly one series). PR #5's `chart_type`, `y_label`, `title`, `labels`, `series`, `source`, and `total_label` fields remain supported. `type` aliases `chart_type`; conflicting values are rejected. Optional `unit` supplies the axis label when `y_label` is absent. `number_format` adds `number` (the legacy default), `percent`, `bps`, `multiple`, or fixed-point formats such as `',.2f'` / `'.1f'`. Formats never rescale values: `12.5` with `percent` displays `12.5%`.

Each series needs a text `name` and finite numeric `values` matching the label count. Labels accept text or numbers; booleans are rejected. Waterfalls start at zero, cumulatively add **every** supplied value, then append a total: `[100, -30, 20]` ends at `90`. Negative totals work; `total_label` names the appended bar. Do not supply a precomputed total as another delta.

Invalid YAML/specifications and rendering failures name the chart in diagnostics. Strict mode rejects before creating/replacing the output; non-strict mode displays the original code panel and logs a warning. Disabled charts remain ordinary code panels even if their YAML is invalid. Charts are PNG images, sized to the page content width, with request-local colors and Korean font policy. Korean text needs the configured font installed (Windows default: Malgun Gothic). See the [fictional three-chart sample](samples/qa/charts.md).

## Section presets

```sh
md-to-word --list-presets
md-to-word report.md terms.docx --preset termsheet
md-to-word notes.md notes.docx --profile plain --preset lecture-note --no-cover
```

| Preset | Cover | TOC | Disclaimer |
|---|---|---|---|
| `ib-report` | Inherit | Inherit | Inherit |
| `termsheet` | Off | Off | Off |
| `legal-memo` | Off | On | Off |
| `lecture-note` | On | On | Off |

Also available through `RenderOptions(preset="termsheet")` or top-level frontmatter `preset: termsheet`. Resolution **per section** is: explicit CLI/API field > CLI/API preset field > individual YAML `layout` field > YAML preset field > profile default. `ib-report` supplies no overrides. Unknown preset names fail without saving.

All presets work with all seven profiles; they change only sections. The general profiles retain neutral metadata, native numbering and general table semantics. `termsheet` is useful with `ib-report`/`ib-memo`, `legal-memo` with `ib-memo`/`plain`, and `lecture-note` with `plain`/`business-report`. Covers/TOCs are usually unnecessary for an `office-letter` or short `meeting-minutes` document. Presets supply no legal wording or document-type metadata.

Without a cover (`--no-cover`, API `include_cover=False`, YAML `layout.cover: false`, or `termsheet`/`legal-memo`), `ib-report` begins the document with a theme-aware title, optional subtitle, and the same date/author rows as `ib-memo`; any TOC follows on the same page, retaining its page break before the body. A matching body H1 and inferred subtitle appear only in that block and do not enter the TOC; cover-on output and other profiles keep their existing title behavior.

## Extended themes

The six existing lowercase theme keys remain supported. YAML themes also accept these typed, uppercase presentation fields from `IBStyle` (PR #5 spelling):

| Category | Keys / units |
|---|---|
| Colors | `NAVY`, `DARK_GRAY`, `LIGHT_GRAY`, `ACCENT_BLUE`, `WHITE`, `RED`, `GREEN`, `ORANGE`, `MEDIUM_GRAY`, `CODE_BG`, `CHART_NEGATIVE_COLOR`, `TABLE_HEADER_COLOR` |
| OOXML colors | `NAVY_HEX`, `LIGHT_GRAY_HEX`, `ACCENT_BLUE_HEX`, `GRAY_BORDER_HEX`, `YELLOW_HEX`, `TABLE_HEADER_BG` |
| Fonts | `HEADING_FONT`, `BODY_FONT`, `KOREAN_FONT`, `COVER_FONT`, `TOC_FONT` |
| Point sizes | `H1_SIZE`–`H4_SIZE`, `BODY_SIZE`, `SMALL_SIZE`, `TABLE_HEADER_SIZE`, `TABLE_BODY_SIZE` |
| Point spacing | `H1_SPACE_BEFORE/AFTER`, `H2_SPACE_BEFORE/AFTER`, `H3_SPACE_BEFORE/AFTER`, `BODY_SPACE_AFTER`, `BULLET_SPACE_AFTER` |
| Inch lengths | `TOP_MARGIN`, `BOTTOM_MARGIN`, `LEFT_MARGIN`, `RIGHT_MARGIN`, `BULLET_INDENT`, `DEEP_LIST_INDENT`, `MAX_LIST_INDENT` |
| Other | `BODY_LINE_SPACING` (positive multiplier), `FULL_LIST_INDENT_LEVELS` (integer 0–9), `BULLET_CHAR`, `TABLE_ZEBRA`, `BODY_JUSTIFY`, `HEADING_BORDER` (booleans), `TOC_TITLE`, `PAGE_LABEL`, `PAGE_OF_LABEL` (text) |

```yaml
NAVY: "234567"
H1_SIZE: 18
BODY_FONT: Calibri
KOREAN_FONT: Malgun Gothic
BODY_SPACE_AFTER: 6
TOP_MARGIN: 0.8
CHART_NEGATIVE_COLOR: "995544"
TABLE_ZEBRA: false
```

Quote six-digit hex colors (`"#234567"` also works). Numeric colors, booleans used as sizes, nonfinite numbers, unknown fields and conflicting aliases are rejected before saving. Sizes allow 1–144 pt; spacing allows 0–144 pt; inch lengths allow 0–3; line spacing allows greater than 0 through 5. The existing `body_size` alias retains its 6–30 pt range and `margin_mm` retains 5–60 mm. Font names must be nonempty. Internal `STYLE_*` identifiers and the profile's `NATIVE_NUMBERING` policy are not theme fields.

RGB/hex color pairs are synchronized when only one is supplied; both may be set independently. `NAVY` also colors IB table headers unless `TABLE_HEADER_BG` is specified. `RED` supplies the waterfall negative color unless `CHART_NEGATIVE_COLOR` is specified. General/mono charts use a neutral palette by default. Headings, callouts, code panels and charts read the immutable style for the current request; `default` and `mono` retain their existing document output.

## Explicit financial tables

```yaml
tables:
  - type: financial
    columns: [text, money, percent, code]
    caption: 실적 요약
    unit: 금액 백만원
    as_of: "2026-06-30"
    source: "[회사 자료](https://example.com)"
    landscape: false
  - type: sensitivity
    base_case: {row: 2, column: 3}
```

Specifications correspond to body tables in order; use `{}` to skip one. Supported types: `generic`, `financial`, `sensitivity`, `risk`. Supported column roles: `text`, `code`, `date`, `number`, `money`, `percent`, `bps`, `multiple`. If supplied, roles must cover every column.

Number/money add thousands separators. Percent/bps/multiple append display suffixes to plain numbers; **they do not rescale values** (`12.5` becomes `12.5%`, not `1250%`). Code/date/text roles preserve values such as `001234`. Blank cells and escaped pipes `A\|B` retain their positions. Headers repeat across pages; rows are kept together when Word permits. `landscape: true` places that table in a separate landscape section, then restores portrait geometry.

Sensitivity highlighting is **explicit only**: row 1 is the first data row; columns are one-based and include the label column. The old guessed centre-cell highlight is removed. Risk highlighting recognises supported risk headers and levels such as high/medium/low or 높음/중간/낮음.

## Supported Markdown and limitations

- Headings, paragraphs, emphasis, pipe tables, lists, blockquotes/callouts, local/base64 images, code blocks and existing diagram/LaTeX rendering.
- External HTTP(S)/mailto links become Word hyperlinks. Local file links become hyperlinks too when the destination is in angle brackets, starts with a drive (`C:/`) or `./`/`../`, or ends in a document extension (`.md`, `.docx`, `.xlsx`, `.pdf`, `.hwp(x)`, images...). When converting a file, absolute paths inside the Markdown file's folder are written as relative links, so they open for colleagues when the folder is shared and do not reveal your local directory; paths elsewhere stay absolute (`file:///`) with a log warning.
- Inline code is shown without its backticks, in the code font (`CODE_FONT`, default Consolas); its content stays literal (no emphasis, breaks, footnotes or term substitution) and is excluded from table number formatting.
- General profiles (`plain`, `office-letter`, `business-report`, `meeting-minutes`) and `term-sheet` turn off Word's automatic spacing between Korean and Latin text or digits, so `SPC는`, `제2종`, `300억원` print tight.
- Numeric single-line footnotes `[^1]` / `[^1]: Note` become native notes when referenced; no named or multiline footnotes.
- Soft-wrapped paragraph lines merge with spaces. `<br>`, `<br/>`, `<BR />` work inside paragraphs, emphasis, headings, lists and table cells. A trailing backslash also preserves a paragraph hard break. Escaped `\<br>` and code remain literal; link destinations are not rewritten. Trailing two spaces require the parser's opt-in legacy flag.
- IB legacy reference-section extraction remains, but only actual reference labels trigger it. General documents keep References as ordinary content.
- This is not a full CommonMark/GFM implementation, a financial calculation engine, a compliance validator, or a lossless layout converter. Raw HTML, complex nested Markdown and arbitrary Word templates are not guaranteed.
- Math/diagram/chart output may be images, not editable equations/charts. New valuation models, native editable charts and PDF distribution are outside this release.
- Update fields/TOC in Word if necessary. Pagination, font substitution and very tall/wide tables still require a visual review. Strict mode is **not** visual QA.

## API

```python
from md_parser import parse_markdown_file
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer

model = parse_markdown_file("input.md", profile="plain")
doc = IBDocumentRenderer(options=RenderOptions(strict=True)).render(model)
doc.save("output.docx")
```

Pass the same profile at parsing time to apply its parsing policy. Parsed models record that profile, and rendering under a different profile raises `ValueError` in both strict and non-strict modes: reparse the Markdown with the desired profile instead of silently losing transformed content. Hand-built models without parser provenance may select their rendering profile. The renderer copies its input model and uses request-local immutable styles. Use one renderer instance per concurrent job.

```python
from converters import get_default_registry

registry = get_default_registry()
model = registry.convert("input.md", profile="business-report")
saved_path = registry.convert(model, output_format="docx",
                              output_path="output.docx", strict=True)
```

Only Markdown input and DOCX output are built in. Models live in `document_model.py`; legacy model imports from `md_parser` remain supported.

## Validation and development

`--strict` rejects known input loss (e.g. extra table cells), invalid footnote references, element render errors and unresolved image placeholders before saving. Without strict mode the converter may emit a partial document with warnings. `docx-audit` reports structural observations and known problems as JSON; it never converts Word back to Markdown.

Audit `warnings` distinguish Normal-style pagination constraints from structural `issues`. `pagination_marked_paragraphs` counts effective keep-lines/keep-next/page-break-before settings, including table paragraphs and inherited styles; it is not a bullet count. Word may show these as **nonprinting black squares** when formatting marks are visible. Keep necessary heading pagination controls; do not strip every flag or silently change the user's Word display settings. `visual_review: not_performed` means that structural inspection did not review rendered pages.

```sh
uv sync --extra dev
uv run pytest tests/
uv run mypy document_model.py document_profiles.py render_styles.py office_layout.py docx_audit.py md_parser.py ib_renderer.py md_to_word.py converters.py --follow-imports=silent
uv build
```

Batch conversion: `uv run md-to-word samples/profiles --batch`. Optional clipboard cleanup: `--format`; optional DeepResearch cleanup: `--deepresearch-cleaner auto`. Preprocessors can change content, so review the cleaned input.

UTF-8/BOM, EUC-KR and CP949 Korean input is decoded deterministically before optional statistical detection. Local image paths resolve from the Markdown directory. In plain/office profiles, superscripts such as `m^2^` remain text; use `[^2]` for a footnote.

### Word page QA on Windows

With Microsoft Word installed, generate examples and render all pages as PNGs:

```powershell
uv run md_to_word.py samples/profiles .dryforge/qa/docx --batch --strict
uv run md_to_word.py samples/qa .dryforge/qa/docx --batch --strict
powershell.exe -NoProfile -File scripts/word_visual_qa.ps1 -InputDirectory .dryforge/qa/docx -OutputDirectory .dryforge/qa/pages
```

The local QA script opens source DOCX files read-only, updates fields in QA copies, warms Word's page cache, and writes page PNGs plus `manifest.json`. It does not change the default printer or add-in settings. Output directories **must be new or empty**; existing evidence is never overwritten. Optional `-ExpectedPages 2` requires every input to have two pages (omit for mixed page counts). A mismatch or source change produces a nonzero exit after recording the completed evidence.

The manifest is always an array and records Word version/build, timestamps, source/updated-DOCX/PNG SHA-256 hashes, actual/expected page counts, and `visualReview: pending`. A successful render does not mark visual review complete: inspect every PNG and record observations separately against the manifest hash. This is not unattended Office hosting or PDF delivery. See the [latest verification record](docs/verification-20260915.md).

On Windows, pass Korean paths in quotes using the literal filesystem spelling. Markdown-escaped paths such as `file\_draft.md` are not the same as `file_draft.md`; do not blindly remove backslashes, since `_draft.md` could be a real child path. Resolve the actual file first; input/output paths are not guessed or rewritten by the CLI.

## Migration from 1.x

- Default `ib-report` and the existing Markdown→Word CLI stay available. Corrected empty cells, number formatting, CLI header/footer composition and disclaimer toggles intentionally change affected output.
- Removed modules: `word_to_md`, `word_parser`, `md_renderer`, `omml_latex`, `roundtrip_audit`; removed built-ins: `DocxInputConverter`, `MarkdownOutputConverter`; removed command: `roundtrip-audit`.
- Historical source is recoverable at commit `d819bbb` / local branch `codex/archive-word-to-md-d819bbb`. Use external tools for Word→Markdown; those tools are not dependencies of this engine.
- Historical plans may discuss bidirectional conversion; this README and the September 2026 implementation plan supersede them.
