# Cover-free title block verification — 2026-09-29

Owner decision, clarified after coordinator Word review: when there is no cover, show
the document title at the very top of the document, before any TOC.
Branch: `feat/title-block-without-cover`. Work and probes stayed in this worktree; no commits,
remote operations, or edits to `CHANGELOG.md` were performed.

## Behavior and scope

| Profile | Existing body title | Result |
|---|---|---|
| `ib-report` | None unless a body H1 survived parsing | Add a Title paragraph, optional subtitle, then memo-style metadata when the cover is off |
| `ib-memo` | Shared office opening, Title style | Unchanged, including existing cover-off inferred-subtitle H2 behavior |
| `plain` | Body H1 in the sample; no synthetic frontmatter-title opening | Unchanged |
| `office-letter` | Office Subject paragraph after sender/recipient metadata | Unchanged |
| `business-report` | Shared office opening, Title style | Unchanged |
| `meeting-minutes` | Shared office opening, Title style | Unchanged |

The new report opening reuses `render_office_opening` inside the existing request-scoped
style context and single `IBDocumentRenderer.render` composition path. It is the first
document content, followed by any TOC on the same page. The existing page break after
the TOC still separates it from the body. Its title typography is the same as the memo title; the optional
subtitle uses the current body style. Date and author are separate `작성일` and `작성자`
rows, matching `ib-memo`, rather than a new combined metadata line. Company remains in
the existing header. Empty or whitespace-only titles create no opening block.

For cover-free `ib-report`, the first matching body H1 is transferred to the Title block.
The inferred-subtitle element is also omitted from the body outline. Both adjustments
occur only on the renderer's deep copy, before TOC preview creation. Title/subtitle
paragraphs have no heading outline level and do not enter an updated Word TOC.
Explicit, unflagged body H2 headings remain body headings, even with a YAML subtitle.
Cover-on composition and the other five profiles are unchanged.

## Initial TDD and checks

All Python commands used `.venv-codex\Scripts\python.exe`, with `TEMP` and `TMP` set to
the resolved `.uv-cache` directory. Pytest used a fresh `.uv-cache\run-*\bt` directory
and `-p no:cacheprovider`.

| Check | Result | Evidence |
|---|---|---|
| New regressions plus parser hardening, before implementation | **24 failed, 86 passed** | `.uv-cache/title-tdd-red.log` |
| Same focused set after implementation | **110 passed** | `.uv-cache/title-tdd-green.log` |
| Full suite, initial title feature | **590 passed, 1 skipped** | `.uv-cache/title-full-suite.log` |
| `python -m ruff check .` | Passed | No violations |
| `python -m mypy .` | Passed | No issues in 15 source files; existing untyped-function notes only |
| `git diff --check` | Passed | No whitespace errors |

`tests/test_title_block.py` adds 28 parse/render/serialized-DOCX cases covering explicit
and inferred title/subtitle content, strict and non-strict mode, empty titles, absent
subtitles, matching H1 deduplication, TOC fields and outline levels, metadata ordering,
source-model immutability, YAML settings and presets, CLI `main()`, concurrent themes,
all five existing body-title profiles, landscape section restoration, header/footer
fields, and strict input/render failure preservation of an existing destination.

Existing tests changed:

- `test_d1_inferred_subtitle_follows_cover_and_toc`: cover-free **ib-report** now expects
  one non-Heading-2 subtitle, even with a TOC. The memo expectations are unchanged.
  Rendering the same model with the opposite cover setting still checks source preservation.
- `test_d1_explicit_yaml_subtitle_keeps_h2_in_body`: explicit report subtitles are now
  visible without a cover; the independent body H2 remains. Memo expectations are unchanged.
- `test_extended_theme_cli_overrides_yaml_and_preserves_failed_destination`: the CLI's
  cover-free title is now a Title paragraph, so its theme color is checked on the Title
  style rather than as direct H1 run formatting. The font-size and failed-save checks remain.

The first implementation run exposed four over-specific assertions in the new tests:
existing footer instructions are `PAGE` and `NUMPAGES` without surrounding spaces.
The tests now inspect stripped instruction tokens. The first full run exposed the existing
H1 direct-color assertion described above; its corrected Title-style check passes.

## Initial XML parity (before ordering correction)

Before any source edits, `.uv-cache/title_parity.py before` rendered all ten
`samples/profiles/*.md` and `samples/qa/*.md` files with default CLI options, plus ten
cover-on cases using `--preset lecture-note`, and the three requested report variants.
The same cases were rendered afterward with the same script.

Artifacts: `.uv-cache/parity-before` and `.uv-cache/parity-after`. Each directory has a
manifest with source/output SHA-256 values and normalized XML-part SHA-256 values.
Comparison: `.uv-cache/parity-after/comparison.json`.

Only the random four-hex-character suffix of generated `_ibrep_` bookmark names is
normalized. No other XML content, formatting, attributes, or ordering is normalized.
`word/document.xml`, `word/styles.xml`, and `word/numbering.xml` match byte-for-byte
after that normalization for **all 20 unchanged-output cases** (60 part comparisons).

| `samples/profiles/ib-report.md` variant | `document.xml` | `styles.xml` / `numbering.xml` |
|---|---|---|
| `--no-cover` | Four paragraphs inserted after the TOC, before the body | Identical |
| `--preset termsheet` | Four paragraphs inserted at the start of the body | Identical |
| `--preset legal-memo` | Four paragraphs inserted after the TOC, before the body | Identical |

The inserted text is the sample's title, subtitle, `작성일: 2026-09-14`, and
`작성자: 기업금융팀`. A structural XML diff confirms that **all other body XML and every
header/footer part remain unchanged** in these three variants. Thus the existing TOC
field, section settings, tables, disclaimer selection, and page fields are retained.
Evidence: `.uv-cache/parity-after/intended-differences.json` and
`.uv-cache/title_diff_evidence.py`.

## Ordering correction after coordinator Word review

The coordinator's Word review found that the first implementation placed the title
block on page 2, after a titleless TOC page. The corrected composition renders the
cover-free `ib-report` opening before `TOCRenderer.render`. Other profiles keep their
existing opening position. No page break is inserted between the title block and TOC,
and the original TOC field and page break remain unchanged. Termsheet has no TOC and
retains its original output.

TDD evidence for this correction:

| Check | Result | Evidence |
|---|---|---|
| Title-block tests before the order fix | **8 failed, 24 passed** | `.uv-cache/title-order-red.log` |
| Same tests after the order fix | **32 passed** | `.uv-cache/title-order-green.log` |
| Full suite, final | **594 passed, 1 skipped** | `.uv-cache/title-order-full-suite.log` |
| `python -m ruff check .` | Passed | No violations |
| `python -m mypy .` | Passed | No issues in 15 source files; existing untyped-function notes only |
| `git diff --check` | Passed | No whitespace errors |

Four new CLI cases in `test_cli_title_block_precedes_toc_on_first_page` exercise
`--no-cover` and `--preset legal-memo`, each with explicit YAML and inferred titles.
They reopen the saved DOCX and assert that title/subtitle/date/author occupy the first
four paragraphs, before both the TOC heading and field; no page/section break or
inherited page-break-before separates the title from the TOC; the original page break
precedes the body; and title/subtitle remain absent from the TOC outline and preview.
The four TOC-enabled cases of `test_cover_off_title_subtitle_metadata_once` were also
updated from the previous, incorrect title-after-TOC expectation. All eight failed before
the renderer edit and passed afterward. No other existing tests changed in this correction.

The parity script `.uv-cache/title_order_parity.py` captured 33 cases before and after
the order fix in `.uv-cache/parity-order-before` and `.uv-cache/parity-order-after`.
It repeats the original 23 cases and adds `--no-cover` and `legal-memo` runs for each of
the other five profiles. The same bookmark normalization and three-part comparison apply.

- **31 unchanged cases** match byte-for-byte across `document.xml`, `styles.xml`, and
  `numbering.xml`, including termsheet, all default/cover-on samples, and the other
  profiles with and without TOC.
- Only `word/document.xml` changes for the report's `--no-cover` and `legal-memo` cases.
  In each, exactly four unchanged title-block paragraphs move from body-child index 6
  to index 0. Every other body XML node, including the TOC field, post-TOC page break,
  body content and section settings, is identical, as are all header/footer parts.
- Compared with the original pre-feature baseline, all three cover-free report variants
  now differ only by four paragraphs added at document start.

Evidence: `.uv-cache/parity-order-after/comparison.json`,
`.uv-cache/parity-order-after/intended-differences.json`, and
`.uv-cache/title_order_diff_evidence.py`. Both READMEs now describe the title-before-TOC order.

## Deferred visual verification

Word page rendering was attempted with the repository's `scripts/word_visual_qa.ps1`
on the termsheet output, using a new, empty `.uv-cache/title-qa-termsheet` output directory
and an expected page count of 1. Word COM activation failed at script line 55 with
`0x80070520`: "A specified logon session does not exist. It may already have been terminated."
No LibreOffice fallback executable was available at the usual installed path or on PATH.

**Visual review is pending, not passed.** No pages were rendered or inspected; actual
page counts and Word version/build could not be obtained. The failure record, input
DOCX hash, expected page count, unavailable fields, and exact command are in
`.uv-cache/title-qa-termsheet/blocked.json`. That directory now contains evidence and
must not be reused for a later QA run.

After the ordering correction, Word QA was retried on the corrected legal-memo sample
with a new `.uv-cache/title-order-word-qa` directory and an expected page count of 2.
COM activation again failed with `0x80070520`. The command/error log is
`.uv-cache/title-order-word-qa.log`; the input hash, expected count, and unavailable
actual-count/Word-version fields are in `.uv-cache/title-order-word-qa/blocked.json`.
The coordinator's reported review established the earlier defect; corrected output still
requires a fresh page review. No visual pass is claimed for this correction.

Follow-up in an interactive Word-capable session: render the three cover-free variants
and a themed inferred-title example into new output directories, record actual/expected
page counts and Word version, inspect every rendered page, and record the review against
the resulting manifest hash. XML parity and automated structure checks do not establish
visual correctness.
