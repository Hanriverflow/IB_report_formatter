# Handoff — Interactive Report Selector + IB Style Preset

> **Date**: 2026-06-02
> **Author**: Claude (Opus) session
> **Status**: Shipped to branches; ready for review/PR

## What shipped

Two pieces of work, on two stacked branches.

### 1. Interactive report selector (branch `feat/interactive-report-selector`)
- `report_selector.py` (`report-select`): interactive dashboard to pick a `.md`/`.docx`
  in `reports/` and convert it via submenus.
- `reports_md_to_word.py` (`convert-reports`): dedicated MD→Word converter for
  `reports/` with interactive / single-file / batch modes.
- `howtouse.md`: Korean usage guide. Tests: `tests/test_reports_md_to_word_cli.py`.
- `.gitignore`: now excludes `reports/`, `output/`, `IB_CPSTB_reader/` (local/sensitive).
- Commit: `d1b5539`.

### 2. IB style preset + audit workflow (branch `feat/ib-style-preset`)
Selectable output style **without touching renderer call sites**, plus a docx
structural audit tool.

- **Preset system**: `IBStyle.classic()/ib_pro()` + `style_profiles.set_active_profile()`
  rebinds the module-global `ib_renderer.STYLE` singleton. Safe because no
  define-time `STYLE.` evaluation exists in `ib_renderer.py` (verified).
- **CLI**: `--style {classic,ib-pro}` (default `classic`) on `md_to_word.py` and
  `reports_md_to_word.py`. `run_conversion()` activates the profile, so single /
  batch / reports paths all honor it.
- **Audit tool**: `tools/ib_style_audit.py` extracts docx structure (margins,
  style fonts/sizes/colors/alignment, paragraph alignment, table border style by
  `insideV`, header shading, color inventory), scores against an IB rubric, and
  diffs classic vs ib-pro. Run: `uv run python tools/ib_style_audit.py <file.md>`.
- **ib-pro improvements** (all behind toggles; classic byte-for-byte unchanged):
  - `BODY_JUSTIFY=False` → body left-aligned (Korean readability)
  - `TABLE_BORDER_STYLE="horizontal"` → no vertical rules, open sides, navy
    top/bottom rules, gray inner horizontals
  - header rule → thick navy bottom border under header row
    (`TableStyler._apply_header_bottom_rule`, sz=18)
  - zebra shading intentionally **omitted** (user preference)
- Commits: `e17dfc2` (design), `b6efe6b` (P1), `6f46052` (P2–P4), `cb64e2e`
  (lint), `6d5da91` (header rule).

## Branch lineage (important for PR)

```
main
 └─ codex/latex-anchor-fixes        (8 latex/semantic commits, already on origin)
     └─ feat/interactive-report-selector   (+ d1b5539 report selector)
         └─ feat/ib-style-preset           (+ 5 IB style commits)  ← current
```

`feat/ib-style-preset` therefore contains latex + report-selector + IB commits.
**PR options:**
- PR `feat/ib-style-preset` → `main`: includes everything (large).
- PR with base `codex/latex-anchor-fixes`: shows only report-selector + IB.
- Cleanest long-term: merge latex branch first, then PR report-selector, then IB.

## Verification (all green at handoff)

- `uv run pytest tests/ -q` → **306 passed** (286 baseline + 20 new)
- `uv run --extra dev ruff check <new files>` → clean
- classic output unchanged (audit diff empty before P3; baseline tests green)
- ib-pro verified visually via Word render (see Temp artifacts)
- **Known/out-of-scope**: `uv run --extra dev mypy ib_renderer.py` reports 6
  pre-existing errors in table column-width / Length code. Confirmed present
  before this work (stash test). Not touched here.

## How to use

```bash
uv run md_to_word.py report.md --style ib-pro       # single
uv run convert-reports --batch --style ib-pro       # batch over reports/
uv run python tools/ib_style_audit.py report.md     # quality audit + classic/ib-pro diff
```

## Remaining / next candidates (from P2 audit findings)

1. Callout & disclaimer paragraphs still use direct `JUSTIFY` (3 per doc) — body
   *style* is left-aligned but these direct-set paragraphs are not. Audit still
   flags `[WARN] body_alignment`. Next ib-pro round.
2. Navy tone unification: headings `003366` vs cover Title `17365D`.
3. Margin symmetry: left 1.0in vs right 0.8in.
4. Optionally fix the 6 pre-existing mypy errors (separate cleanup).

Full gap analysis: `docs/superpowers/specs/2026-06-02-p2-audit-findings.md`.
Design: `docs/superpowers/specs/2026-06-02-ib-style-preset-audit-design.md`.

## Local-only artifacts (gitignored, safe to delete)

- `output/_*.py` (audit/render/compare scripts), `output/_render/*` (PDF/PNG)
- `reports/*__classic.docx`, `reports/*__ib-pro.docx` (A/B comparison renders)
- Render deps installed into `.venv`: `pywin32`, `pymupdf` (not in pyproject;
  add to `[dev]` extras if the render workflow should be reproducible).
