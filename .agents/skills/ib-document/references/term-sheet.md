# Term-sheet authoring decisions

Use `samples/profiles/term-sheet.md` for supported syntax and `samples/harness/` for a smaller source-to-document exercise. The full ABCP sample contains a fictional transaction structure; copy layout selectively, not its commercial substance.

## Minimum frontmatter

```yaml
---
profile: term-sheet
title: "{{company}} 금융조건 협의안"
date: "2026-10-08"
version: "v0.1 (검토용)"
house: /absolute/job/house.yaml
terms:
  company: "가상회사㈜"
  amount: "300억원"
tables:
  - {}
---
## 1. 주요 조건

| 구 분 | 내 용 |
|---|---|
| 조달금액 | {{amount}} |
| 미확정 사항 | [확인 필요: 실행일] |
```

The displayed date and “검토용” are examples; use the user's date/status. `title`, `prepared_by` and `disclaimer` must be nonblank; the latter two may come from the house. `house` resolves relative to the Markdown. A logo declared in the house resolves relative to the house file. Frontmatter overrides house values, including explicitly empty values, so do not insert blank overrides unintentionally.

The house accepts `prepared_by`, `disclaimer`, `confidential_label`, `confirmation` (`intro`, `items`, `signature`) and `style`. Keep approved institution wording as supplied. If no approved wording is available, ask or deliver an incomplete draft; only demonstrations may use the supplied fictional house. The old `termsheet` preset is not the `term-sheet` profile.

## Repeated terms and checks

`terms` is a mapping of keys matching `^[a-z][a-z0-9_]{0,39}$` to quoted single-line strings. Use `{{amount}}`, not an unrelated hardcoded copy, throughout headings, prose and tables. Values are exact displayed text, not Markdown. Quotes protect precision (`"1.10%"`), date spelling and numeric codes. Variables in code, math or link URLs are not substituted. Undefined keys trigger strict failure. An unknown condition can be an explicit string such as `"[확인 필요: 금리]"`; never manufacture a numeric value for it.

Add `checks:` only for relationships actually specified by the source or clearly disclosed derivations. Example: `"all_in = base_rate + spread + fee"`. Existing declared checks must not be dropped to hide contradictions. Checks read displayed units and report disagreements; they do not calculate a schedule or establish financial validity. A placeholder used in numeric arithmetic cannot pass: preserve the unresolved condition and report why a finished render is blocked.

## Tables and closing

- Use explicit `##`/`###` headings. Each entry in `tables:` corresponds to the next Markdown table; keep `{}` for tables without overrides.
- Keep every row's cell count equal to its header. Use `<br>` for multiple lines inside a cell and escape literal pipes. Preserve blanks.
- `| 구 분 | << | 내 용 |` creates two label columns. `<<` merges left and `^^` merges upward; only rectangular groups within the header or body are allowed. Do not merge across that boundary. Escape literal markers with backslash.
- `header_rows` repeats the declared leading rows. Set `columns` roles explicitly for codes, dates or numeric schedules; do not rely on financial meaning inference.
- A `schedule` uses one-based column indexes and the stated table unit. Author only source-supported dates and repayment amounts; holiday adjustment and repayment terms are not inferred from the sample.
- A customer-confirmation box uses an empty closed `confirmation` fence only when the house/frontmatter contains its wording. Omit the box if not requested; never invent a signature, consent or approval.

Review the resulting Word document for table continuations, repeated headers, line wrapping, closing/confirmation placement and page breaks. Mark visual review pending until actual pages are inspected.
