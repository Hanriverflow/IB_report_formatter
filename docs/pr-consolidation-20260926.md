# Writer-profile promotion and PR consolidation (2026-09-26)

## Approved scope

The owner selected the September writer-only implementation as the successor to the older open PRs. The changes were assembled in a separate worktree from `main` at `d819bbb3646554b3cdce6499bd6bdd6113513b58`; the dirty user checkout was not staged, stashed, reset, or cleaned.

Promoted functionality includes six document profiles, shared document models, render-scoped immutable styles, office-letter/appendix composition, structural DOCX auditing, and a single CLI/API/registry rendering path. Word-to-Markdown remains deliberately retired. Its historical source remains recoverable from `d819bbb` in Git history.

This is source integration, not a formal release or human/business approval. No CI, account billing, branch protection, or publication settings were changed.

## Validation

Product-code HEAD: `cadc0ae515d2f5757d5006e64bec1acaff56ccba`.

- Isolated environment: Python 3.12.12; dependencies from `uv sync --locked --extra dev`.
- Full suite: **311 passed, 0 skipped**.
- `ruff check .`: passed.
- `mypy .`: passed for 14 source files under the repository configuration. Existing notes about untyped function bodies remain; this is not a claim that every function body is fully type-checked.
- Wheel and sdist build: passed. Wheel inventory: 18 files; sdist: 37 files. Neither includes retired inverse modules, private reports, raw documents, QA images, or agent state.
- Wheel installed into a separate locked-dependency environment and imported outside the repository. Six-profile listing, strict office-letter conversion, and structural audit passed.
- Nine synthetic DOCX samples: ZIP integrity, reopening, and known structural checks passed.
- Staged whitespace checks passed. Credential-pattern and generated/private-file checks found no unexpected new payloads; these checks are not a universal secret-detection guarantee.

### Regression fixes added during consolidation

1. Blank leading header cells, interior cells, and trailing cells retain their columns through parsing, DOCX saving, and reopening. The first two probes fail on old PR #5; the September parser preserves them. Three saved-document regressions now protect this behavior from PR #2.
2. Small existing helper issues were cleaned up: resolvable `pathlib.Path` annotations, optional matplotlib import probing, import ordering, equivalent dictionary literals, and the typed `tight_layout` rectangle. Three helper regressions were added.
3. Bounded AI review found that changing only the render profile after profile-sensitive parsing could silently discard content. Eight renderer/registry strict/non-strict cases reproduced the missing guard. Parsed models now retain `parsed_profile`; mismatched render profiles raise an actionable error requiring reparsing, in both strict and non-strict modes. Hand-built models remain supported. Ten boundary tests cover rejection and valid use. Re-review of `cadc0ae` found the reported blocker resolved and no direct blocker introduced by this fix.

### Native Word and visual evidence

Microsoft Word 16.0, build `16.0.20326`, freshly rendered all pages. Source DOCX hashes were unchanged. Observed counts:

| Synthetic document | Pages |
|---|---:|
| business-report | 1 |
| ib-memo | 1 |
| ib-report | 4 |
| landscape | 3 |
| long-table | 2 |
| meeting-minutes | 1 |
| office-letter | 1 |
| office-letter-appendix | 2 |
| plain | 1 |
| **Total** | **16** |

All 16 full-page images were inspected in an AI visual review: no visible merge-blocking clipping, overlap, missing table columns, unintended blank pages, broken Korean glyphs, or footer/page-number errors were found. This is not human or business approval. `ExpectedPages=0` meant no fixed-count assertion; the counts above are observations.

Native manifest SHA256: `E9B01249AD17FFCEB192E65B99530C6310831244E4D40A86AE12B2861162BB88`.

The native render was generated at `c323d7858a36e4dda2240663cae7de35a09a388b`. After the profile-boundary fix at `cadc0ae`, all nine samples were regenerated and every uncompressed OOXML package entry name and byte was compared with the rendered packages. All were identical, with no excluded parts or ignored differences. Thus the visual evidence is reused for identical document content, not inferred from an old test count. ZIP container hashes can differ because of archive metadata.

## Old PR disposition after successor merge

The following PRs are superseded, not claimed to have been merged wholesale. Their branches and review history are retained.

| PR | Preserved source | Disposition / deliberately deferred content |
|---|---|---|
| #2 | `codex/ship-rv5` at `7ef3c664fbc9e13aec77c075c1b4befbf15543b2` | Blank-cell correctness and regressions are carried forward; local output/data ignore rules are included. The retired inverse CLI's UTF-8 work and old version metadata are not ported. |
| #4 | `feat/ib-style-preset` at `0b9ff68896d269af1fed0a0de939092e48eff2a2` | The reports selector/dedicated CLI, classic/ib-pro visual policy, and comparison audit are deferred, not present in this successor. The original target was `codex/latex-anchor-fixes`, not main. Its inherited-alignment audit review remains relevant if this work is revived. |
| #5 | `codex/remove-word-to-md` at `f78a05206fb78a67397325c81921a1d37797ac16` | Superseded by the newer writer-only profile architecture. Its chart engine and original theme/preset implementation are not silently treated as ported. A future port must adapt to request-scoped immutable styles. |

The earlier September verification files remain historical records with their original counts. Current validation is the 311-test result above. Raw test logs, package hashes, the native Word manifest/images, output-equivalence checks, and original-file preservation checks are retained outside the repository rather than publishing user files or full agent sessions.
