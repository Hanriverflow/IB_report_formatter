# P2 Audit Findings — current (classic) output vs IB rubric

> Tool: `tools/ib_style_audit.py`
> Audited: `tests/일동제약_수익성분석.md` (9 tables), `tests/웅진_계열사.md`

## Gaps found

1. **Tables are full-grid.** All data tables render with both inside-horizontal
   and inside-vertical borders (`border_style: grid`). Sell-side IB tables
   favor horizontal rules — header top/bottom rule plus a bottom rule, with
   vertical lines minimized. → ib-pro: `TABLE_BORDER_STYLE = "horizontal"`.

2. **Body has justified paragraphs** (3 in 일동제약). Korean text reads best
   left-aligned; justification opens uneven inter-word gaps. → ib-pro: left-align
   body paragraphs.

3. **Navy tone inconsistency.** Headings use `003366`; the cover Title style uses
   `17365D`. → ib-pro: unify on the house navy (optional, lower priority).

4. **Margins asymmetric** (left 1.0in vs right 0.8in). → ib-pro: symmetric
   margins (optional, lowest priority).

## P3 priority (matches design: tables → typography → color → cover)

1. Table border style: full-grid → horizontal-rule  **[highest impact]**
2. Body alignment: justify → left
3. Navy palette unification (optional)
4. Margin symmetry (optional)

Items 1–2 are the high-leverage IB improvements and will be implemented first,
each behind an `ib-pro`-only toggle so `classic` stays byte-for-byte unchanged.
