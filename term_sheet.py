"""Term-sheet profile composition: house boilerplate, opening block and layout helpers.

Called from `IBDocumentRenderer.render` (the single composition path) when the
resolved profile is `term-sheet`. Design: docs/term-sheet-design-20260929.md.
This module depends only on the model, styles, YAML and python-docx primitives;
it must not import `ib_renderer` (the renderer injects run-rendering callbacks).
"""
