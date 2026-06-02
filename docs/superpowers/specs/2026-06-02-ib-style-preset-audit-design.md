# IB Style Preset + Quality Audit Workflow — Design

> **Date**: 2026-06-02
> **Status**: Approved (user instructed to proceed)
> **Branch**: `feat/ib-style-preset`
> **Scope**: MD→Word rendering only. Word→MD pipeline and existing CLIs unchanged.

## Goal

On top of the mature MD→Word pipeline (Step 1–16, 286 tests passing), build an
**IB-grade quality audit workflow** and capture the resulting improvements in a
selectable **`--style ib-pro`** preset. The default **`classic` preset must
reproduce current output exactly** (no regression).

## Decision Log (brainstorming)

| Axis | Decision |
|---|---|
| Primary focus | Inspect → improve in one pass |
| Quality baseline | Claude defines an IB-standard rubric |
| Selection unit | Global style preset `--style classic\|ib-pro` (default `classic`) |
| Audit method | python-docx structural inspection (reusable as regression gate) |
| Integration | Global profile swap (rebind the `STYLE` singleton) |

## Architecture — Preset System (global profile swap)

```
IBStyle (frozen dataclass)
   ├── IBStyle.classic()  classmethod  → current values verbatim (regression-safe)
   └── IBStyle.ib_pro()   classmethod  → improved values + behavior toggles

style_profiles.py (new, thin module)
   ├── set_active_profile("classic"|"ib-pro")  → rebind ib_renderer.STYLE
   └── get_active_profile()

ib_renderer.py
   - STYLE.XXX references stay as-is (no mass rewrite)
   - IBStyle gains behavior-toggle fields:
       PROFILE: str, TABLE_BANDED: bool, TABLE_BORDER_STYLE: str,
       COVER_RULE: bool, NUMBERED_SECTIONS: bool, ...
   - Structural differences become `if STYLE.<toggle>:` at a few points only

CLI: md_to_word.py / reports_md_to_word.py
   - --style {classic,ib-pro}  (default: classic)
   - call set_active_profile() at the start of run_conversion()
```

**Principle**: constant differences live in profile values; structural/behavioral
differences live in boolean toggles on `STYLE`. No renderer surgery → regression-safe.

## IB Quality Rubric (Claude-defined)

Each item is measured automatically from the generated docx:

| Area | Checks |
|---|---|
| Type hierarchy | H1>H2>H3 size/color/weight contrast, spacing consistency |
| Color palette | restrained navy + charcoal + one accent; no color overload |
| Tables | full-grid vs horizontal-rule, header emphasis, numeric right-align / thousands / negatives, cell padding, optional banding |
| Cover | company/title/date/confidential placement, horizontal rule, margin balance |
| Section structure | numbering, header/footer consistency, page numbers |
| Body | Korean left-align, line spacing, paragraph spacing, margins |

## Audit Tool (new `tools/ib_style_audit.py`)

- Extract structural attributes via python-docx (font/size/color/alignment/borders/margins/styles)
- Evaluate against rubric → per-item pass/warn/fail + metrics
- `classic` vs `ib-pro` diff report (markdown/JSON): convert the same input with
  both presets, compare attributes
- Sample inputs: `tests/일동제약_수익성분석.md`, `tests/웅진_계열사.md`
- Callable from pytest → reusable as a **regression gate**

## ib-pro Improvement Candidates (finalized in Phase 2 audit)

Candidates only; concrete values come from the audit:

- **Tables**: full grid → header/top/bottom rules + stronger header shading +
  optional row banding + robust numeric-column alignment
- **Typography**: refine H1–H3 contrast/spacing, tune body line spacing
- **Color**: refine navy/charcoal tones, restrain accent
- **Cover/Sections**: refine horizontal rules, margins, consistent section numbering

## Regression Safety & Testing

- **classic invariant**: existing 286 tests must pass with `classic` default (zero change)
- **profile fixture**: reset `set_active_profile("classic")` after each test (global-state isolation)
- **ib-pro tests**: verify toggle behavior + rubric score improvement
- **audit tests**: attribute extraction / diff correctness

## Phased Execution Plan

| Phase | Work | Output |
|---|---|---|
| **P1** | Preset skeleton (`IBStyle.classic/ib_pro`, `set_active_profile`, CLI `--style`) + audit tool | Working preset switch (ib-pro initially == classic) + audit report |
| **P2** | Run rubric audit on current output → identify gaps | Audit findings doc |
| **P3** | Implement ib-pro values/toggles (tables → typography → color → cover) | Improved ib-pro preset |
| **P4** | Regression (classic 286) + ib-pro tests + final verification | Green tests, before/after report |
