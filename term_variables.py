"""Term variables: `terms:` frontmatter values referenced as `{{key}}`.

Design: docs/term-sheet-design-20260929.md §6. Substituted values are tagged
as Word content controls and snapshotted in custom document properties so
`docx-audit` can report inconsistent edits made later in Word.
"""

# Word content-control tag and custom-property name prefixes shared by the
# renderer (writer) and docx_audit (reader).
TERM_TAG_PREFIX = "ibrep:term:"
TERM_PROPERTY_PREFIX = "ibrep.term."
