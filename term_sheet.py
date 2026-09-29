"""Term-sheet profile composition: house boilerplate, opening block and term tables.

Called from `IBDocumentRenderer.render` (the single composition path) when the
resolved profile is `term-sheet`. Design: docs/term-sheet-design-20260929.md.
This module may import from `ib_renderer`; `ib_renderer` imports it lazily.
"""
