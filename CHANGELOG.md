# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

Changes since 2.0.0-beta.1 (2026-11). Engine/package metadata stays **2.0.0 (Beta)** and the plugin stays **0.1.0**; nothing here has been tagged or published. Verification: `docs/verification-20261110.md`.

### Office letters: attachments, tables at the end and item symbols (roadmap E7)
Existing office letters change appearance where they have attachments or items nested three or more levels deep. Verification: `docs/verification-office-letter-20261114.md`.
- A letter body that ends with a full table keeps `끝.` below the table; this is now pinned by tests (no output change).
- New `letter.table_end: partial`: when the table ending the letter body is filled only partway, `이하 빈칸` goes in the first cell of the first blank row after the last written row and `끝.` is omitted (it still follows attachments). Never inferred from blank rows; an invalid value, a body that does not end with a table, or a table without a blank row after its last written row is rejected before saving. Letters without the setting are unchanged.
- **Changed appearance:** attachments follow the administrative handbook. The separate `붙임` line is gone: one attachment is one unnumbered line (`붙임  항목.  끝.`), several start on the `붙임  1.` line with later numbers aligned under `1.`, still as native Word numbering.
- **Changed appearance:** office-letter numbered items use `1.` `가.` `1)` `가)` `(1)` `(가)` `①` for levels one to seven (levels one and two are as before), with wider hanging indents for `(1)` and `(가)`. Word has no number format for `㉮`, so an eighth level is rejected in strict mode and rendered with `가` and a warning otherwise. Other profiles keep `1.` `가.` `(1)`.

### Breaking: image and logo paths are limited to allowed folders
- Images and logos that a document names by path load only from the Markdown file's folder (subfolders included) and from folders the operator adds. A path that resolves outside them (`../`, an absolute path, a UNC path, or a symlink or junction leading out) is reported as `[asset_outside_root]`: strict mode stops before saving, and non-strict mode keeps the image placeholder and a warning. Earlier versions loaded any readable local path, so a document that relied on that now needs one of the remedies below.
- Remedies: add the folder with `--asset-root FOLDER` (repeatable) or the API argument `asset_roots`, or copy the files into the Markdown file's folder and use relative paths. Text input, stream input and hand-built models have no default folder; pass `asset_roots` when they use path images.
- The location is resolved first and judged before the file is looked up, so an outside path gives the same result whether or not the file exists. A path written as absolute is allowed when it resolves inside an allowed folder.
- A house-file logo is restricted when the document's frontmatter names the house file and trusted when the operator names it (`--house`, API). Nothing in a document, house file or theme file can add a folder. Hyperlinks, Base64 images and `data:` URIs are not path-restricted.
- The plugin runner and the portable converter have no option to add folders: copy the images, the house file and its logo into the job folder. Their guidance files now say so.

### Input limits
- Five limits stop a conversion in both strict and non-strict mode and leave no output file: input size (`--max-input-bytes`, default 16 MiB), cells in one table (`--max-table-cells`, 50,000), list depth (`--max-list-depth`, 32), bytes of one image (`--max-image-bytes`, 20 MiB) and pixels of one image (`--max-image-pixels`, 50,000,000). The API takes the same values as an immutable `InputLimits`.
- Size is judged before reading or decoding: a file by its size, a stream as soon as it passes the limit, Base64 from its encoded length, pixels from the image header. Values must be positive integers; there is no unlimited value, and a document cannot raise a limit.
- In batch conversion a file over a limit fails and the other files continue.

### Rejection codes and exception types
- Exactly six codes: `asset_outside_root`, `limit_input_bytes`, `limit_table_cells`, `limit_list_depth`, `limit_image_bytes`, `limit_image_pixels`. A coded issue starts with `[code] ` and keeps the prefix inside joined messages. Existing warning sentences are unchanged.
- Strict rejections raise `StrictValidationError` (`.codes`, `.issues`) and limit failures raise `InputLimitError` (`.code`, `.limit`, `.measured`). Both are `ValueError` subclasses, so existing `except ValueError` code keeps working, and both stay reachable as the cause of wrapping errors such as `Rendering failed: ...`.
- Errors that reject a document's own settings raise `StrictValidationError` with their earlier message. Where a validator or the YAML loader raised the error, it is kept as the cause: malformed declared YAML frontmatter (including a scalar YAML cannot construct, such as the date `2026-99-01` or `!!timestamp x`), invalid frontmatter values (profile, layout, style, table specifications, terms, checks, letter settings), a metadata value of the wrong type where the cover, title block, closing or header needs text (such as `sender: {organization: 12}`; values that rendered before, including numbers, dates and lists of text, still render), and invalid content of a house or theme file named by the document. A footnote ID that is not positive and an empty input file are raised directly, without a cause. Errors in operator arguments, including a path error in an operator-named house file, keep their type.
- The plugin runner's `result.json` adds one diagnostic per coded issue (`code` is the rejection code); its other `code` values and keys are unchanged. The runner applies the input-size limit before it hashes the Markdown, so an oversized draft is rejected at stage `input`. The portable converter window explains each code in Korean.
- Control characters that XML cannot hold are diagnosed before rendering: in the Markdown body as an input warning that strict mode rejects, and in metadata or house/theme values (including YAML escapes) as a rejection. Strict rendering checks every string of the model, so a hand-built model is covered too. Non-strict rendering still replaces an affected body element with a marker; where such a character is in text outside that per-element handling (a footnote, a report-profile heading), the failure is now a `StrictValidationError` naming the character.
- A chart fence whose YAML holds a value that cannot be constructed is reported as that chart's error, like a YAML syntax error, instead of stopping the parser.
- A footnote or reference number, or a YAML integer, longer than Python converts (more than 4,300 digits) is a document rejection instead of a plain `ValueError`.
- YAML whose aliases expand to far more than the YAML holds is refused with `YAML aliases expand to too much content`: alias lists nested in alias lists, one long text repeated thousands of times, or one large mapping pulled in thousands of times with a merge key. A document of a few hundred characters to about a hundred kilobytes could keep the conversion from finishing this way. YAML that refers to itself through an alias (`loop: &x [*x]`), which used to load, is refused with `YAML structure refers to itself`. Both are decided on the parsed YAML before any value is built. In frontmatter and in a house or theme file this is a document rejection, in a chart fence it is that chart's error, and a diagram fence stays a code block. The allowance is the larger of 100,000 and four times the length of the YAML, counting one for each value and each character of text, with an aliased value counted wherever it is used. A mapping that repeats the merge key more than 100 times is refused with `YAML mapping repeats the merge key too many times`; one merge key is all YAML needs, and its value may list several mappings. Ordinary anchors, aliases and merge keys load as before, and YAML without aliases cannot reach the allowance. Two more forms that slowed loading without any alias are refused the same way: a mapping that ends up with more than 100 integer keys of 18 or more characters, counting keys merged into it (`YAML mapping has too many very long integer keys`; quote such keys to keep them as text), and an integer written with more than 20 colons (`YAML integer has too many parts separated by ':'`). Parsing YAML still takes seconds per megabyte, so an operator who converts untrusted documents should put a time limit on the conversion.

### Korean typesetting (output changes)
- `ib-report` and `ib-memo` body text (the `IB Body` style) is left-aligned instead of justified. Korean lines break only between words, so justified Korean body text stretched word gaps to up to 4.6 times a normal space. English IB body text loses its even right edge too. The cover and disclaimer keep their alignment, and a theme with `BODY_JUSTIFY: true` still justifies the body. Page counts of the samples are unchanged.
- Every profile declares Korean (`ko-KR`) as the East Asian language in the document defaults; term-sheet already did. Word no longer breaks a line inside a Korean word after `㈜` or a closing `」` (`가나다㈜⏎는`, `「신탁법」⏎에`). Latin text keeps `en-US`, and term-sheet output is unchanged. Verification: `docs/verification-korean-typesetting-20261114.md`.
- `ib-report` and `ib-memo` turn off Word's automatic spacing between Korean and Latin text or digits, as the other five profiles already did: `2025 년`, `300 억원`, `SPC 는` now print as `2025년`, `300억원`, `SPC는`. Page counts of the samples are unchanged.

### IB table rules (opt-in, no change to existing output)
- Table specs take `total_rows` and `subtotal_rows`: lists of one-based body rows (the first row after the header rows is 1). A subtotal row gets a single rule above and bold text; a totals row also gets a double rule below. Rules use the profile's table accent colour and work in every profile. Rows are never inferred from text such as `합계`; `schedule.total` still only checks. Row numbers that are not integers, are out of range, repeat, or appear in both lists are rejected in strict and non-strict mode.
- A `type: financial` table spec can set `negative: paren`, `minus` or `triangle` to show its negative amounts as `(1,234)`, `-1,234` or `△1,234` in black. Amounts written `(…)`, `-…` or, in such a table only, `△…` in numeric columns are restyled; the digits and unit suffixes stay (a suffix goes inside parentheses: `(3.5%)`), and text, code and date columns, `▲`/`▼` and tokens such as `NM` are untouched. Without the setting nothing changes: negatives keep their written form and red (owner decision D4, 2026-11-14). Other values, or the setting without `type: financial` in the same spec, are rejected.
- Fixed after an independent review (2026-11-14): the negative style reads a cell's whole text once and restyles only a cell that is exactly one negative amount, so `-1200**.50**` no longer becomes `(1,200).50` and `1**-2**` no longer becomes `1(2)`. A unit inside the parentheses (`(3.5%)`, `(1200억원)`) and a U+2212 minus are read in every style, and `△3.5` gets a percent column's `%` like `-3.5`. Ranges, several numbers, words and negative zero keep their written form; an amount split across differently formatted text keeps it too and logs a warning that strict mode does not reject (`IBDocumentRenderer.warnings`). Alignment, column kinds and widths are judged on the written text, as without the option. The option still turns off only the automatic red; written colours and link colours stay. Tables without the option are unchanged.
- Fixed after a second independent review (2026-11-14): the negative style no longer rewrites an amount whose integer part starts with a zero (`(0001.20%)`, `−0001200` used to become `(1.20%)`, `-1,200`); it keeps the existing number formatting and shows the same text as in a table without the option, which drops leading zeros in some forms (`-0001200`, `(0001.20)%`, `-00.5`) as it always has. An amount split by a superscript or subscript piece (`-1200^.50^`) now logs the split-amount warning too. Full-width digits are read as the formatting without the option reads them, so `-１２３４` becomes `(1,234)` with `paren` instead of keeping `-1,234`. `IBDocumentRenderer.warnings` is emptied when `render` starts, so a render rejected before composition no longer shows the previous document's warnings. Without the option the automatic red still covers a colour written in the document; with it that colour shows.
- Fixed after a third independent check (2026-11-14): an amount split across differently formatted text warns even when a leading zero keeps it from being restyled (`-0001200**.50**`, `-00^.5^`). In an office letter with `letter.table_end: partial`, `이하 빈칸` on a row named in `total_rows` or `subtotal_rows` is bold like the rest of that row; letters without those lists are unchanged.
- Fixed after a fourth independent check (2026-11-14): a run holding only spaces before or after an amount (`<b>-1200</b>&nbsp;`, `**-1200** [^1]`, `<br>**-1200**`) no longer counts as splitting it, so no split-amount warning is logged and the amount is restyled while the space keeps its own text and format. A split inside the amount still warns.
- Fixed after a fifth independent check (2026-11-14): space runs and footnote markers around an amount are set apart in any order and number, not only footnote markers at the very end. `-1200**.50**[^1]<br>` now warns like `-1200**.50**[^1]`, and `**-1200** [^1]<br>` is restyled. An amount after a footnote marker is still not restyled, but a split one now warns. Tables without the option are unchanged.
- New fictional sample `samples/qa/ib-table-rules.md` shows both settings; README, AGENTS.md and the term-sheet skill reference describe them. Existing samples render pixel-identically in Word and their XML snapshots are unchanged; the model snapshots of the 11 samples with tables gain only the three new default fields. Verification: `docs/verification-ib-table-rules-20261114.md`.

### Document properties and DRAFT marking (output changes)
- Frontmatter `version` is written as the custom document property `Version` and shown after the company in the page header as a `DOCPROPERTY Version` field holding the current value (office letters show no version; the term-sheet footer's version becomes the same field with unchanged text). After editing the property in Word, F9 inside a header updates that section and Ctrl+P or print preview updates every section. An empty `version:` means no version. Verification: `docs/verification-document-properties-20261114.md`.
- `status: draft` puts a DRAFT watermark in Word's own format into every header of every section, so Design > Watermark > Remove Watermark clears every page, including the IB cover and landscape-table sections; `final` or no status adds nothing and any other value is rejected. New QA sample `samples/qa/draft-version.md`.
- File > Info now shows the document's own title, an empty author and comment, and the conversion time instead of python-docx's template values (author `python-docx`, comment `generated by python-docx`, dates of 2013). Only `docProps/core.xml` changes; page images of all samples are unchanged. A profile's stand-in title is left out, but a title the document writes is kept even when it reads `Document` or `IB Report`. A title longer than 255 characters, Word's limit for a file property, is shortened there only, with a warning; the document text keeps it whole.
- A title written as `Document` or `IB Report` is now kept as the document title. With a different first H1 heading it used to be replaced by that heading in general profiles (`title: Document` with `# Chapter` gave the title `Chapter`) and, for `title: IB Report` in IB profiles, when read through `parse_markdown_file` (the CLI and converter path); and a hand-built model rendered under another profile used to take that profile's stand-in title. How a document without a written title gets its title is unchanged.

### Table and figure numbers
- A table spec `id: tbl-…` or an image file line ending in `{#fig-…}` numbers that table or figure: its caption reads `표 1. 제목` / `그림 1. 설명`, where the number is a Word `SEQ` field holding the number already worked out (tables and figures counted separately in document order) and a hidden bookmark spans `표 1`. Every profile, the term-sheet caption/unit line included. Documents without IDs are unchanged. A malformed, repeated or caption-less table `id` rejects the document; an unusable figure ID keeps the line as text and is rejected in strict mode.
- `@tbl-…` / `@fig-…` in paragraphs, lists, quotes, headings, table cells and table caption/unit/source/note text become Word cross-references (`REF _Ref_<id> \h \* CHARFORMAT`) showing the cached `표 1`; the result keeps the surrounding body formatting after Word updates it. Code, math, link destinations, `\@` and e-mail-like `a@tbl-x` stay literal. Strict mode rejects a reference to an unknown ID, a malformed reference, one inside link text, and one in the title, subtitle, an image description or a footnote; the default mode keeps such text and reports it. `docx-audit` reports a reference whose bookmark is missing as an issue and a stale reference or caption number as a warning.
- Author usage for table and figure numbers is in `README.md`, `README.ko.md` and `AGENTS.md`, including the one-time `그림` label needed to continue figure numbers with Word's Insert Caption; Word, page-image, schema and build results are in `docs/verification-captions-crossrefs-20261114.md`.
- Review fixes, behaviour change: strict mode now rejects every `@tbl-…` / `@fig-…` that the grammar reads as a reference and that does not resolve — an unknown ID, a malformed ID read whole (`@tbl-a-`, `@tbl-`, never shortened to `@tbl-a`), an ID split by formatting, and one in link text, a title, subtitle, footnote or any image description (inline images included). The `@` must follow the start of the text, whitespace or an opening bracket or quote, and `@tbl-x.example` is not a reference. One recognizer serves paragraphs, callouts, headings, the TOC preview, pipe and HTML table cells and table texts, with the same protection of code, math, escapes and link destinations. Text that does not resolve renders exactly as before (one run, as written), and the term-sheet title and subtitle write no field. A table that is not drawn gets no number (its `id` is reported, and rejected in strict mode), written caption numbers are checked against the numbers worked out, and a `{#…}` after any image that is not a numbered figure line is reported.
- `docx-audit` reads a field nested in another field's instruction and a quoted bookmark name (`REF "_Ref1"`) correctly, and judges `SEQ` numbering order only for sequences with no switch or `\* ARABIC` alone; it checks the main body only.
- Second review fixes: the parser now splits a run only around a reference that resolves (it learns the numbered IDs from a first parse) and reports every other one without touching the runs, so an unresolved reference no longer merges the pictures or term values after it, and reading an ID next to a colour span always ends. Text is read as shown, across emphasis and colour (an empty colour span adds nothing, `@tbl-a**b**.com` is a domain); a heading is judged with its `**` markers; term values are never read, in table captions and notes included. Behaviour change: strict mode also rejects a reference in a frontmatter value the document shows (the date, sender fields, cover and letter fields, term-sheet texts; the list is in `README.md`), and term-sheet date and preparer lines no longer hold a field. `docx-audit` reads a `w:fldSimple` with a field nested in it correctly.
- Third review fixes: the private-use characters U+E010 and U+E011, which the engine had used as marks for an escaped `@` and `$`, are kept as written again (they had become `@` and `$` everywhere, code included); marks are now characters the text does not hold, and none is used outside a reference scope. Behaviour change: strict mode also rejects a reference that formatting splits or that only the parsed text spells (`@**tbl-x**`, `&#64;tbl-x` in an HTML table, a YAML escape for `@`), a reference in the description of an HTML `data:` image, and, for a model built without the parser, one in an image description or footnote. A term value in a heading that becomes the document title is no longer read as a reference.
- Fourth check fixes: a reference that resolves changes nothing but its own text (P6). A table no longer formats an amount that shares a run with a resolved reference (`-001200 @tbl-a` stays `-001200 표 1`, not `-1,200 표 1`), and the IB static TOC preview keeps the heading as written (its `**`, escapes and `<br>`) with only the reference as `표 1`. The file and stream parser now fills IB metadata from the body while it parses, so a first H1 that becomes the title is judged (`# Review @tbl-missing` with `profile: ib-memo` is rejected in strict mode, as `MarkdownParser.parse` already rejected it), and a label value such as `**작성자:** {{x}}`, `\@…` or code is no longer rejected as a reference; one that resolves in its paragraph is reported where the value is shown.
- Fifth check fixes: a reference written right after other text no longer stays text without a word. `@tbl-a,@tbl-b`, `@tbl-a~@tbl-b`, `참조:@tbl-a` and `다음@tbl-a` are still not references by the grammar and render as before, but when the ID is the ID of a table or figure in the document, strict mode now rejects it with the advice to put a space before the `@` or write `\@`; an ID that names nothing, an e-mail-like `user@tbl-x` and a doubled `@@` are untouched. A prefix with capitals (`@Tbl-a`, `@FIG-a`) is reported as a malformed reference instead of being ignored. Strict mode also rejects a reference in `letter.appendix_label` and in the cover's `report_type`, which were shown but not checked.
- With E8's version in the page header, strict mode also rejects a reference in `version` wherever the header shows it (not in an office letter, whose header has no version; the term-sheet footer was checked already).
- The reference grammar, the rejected inputs and the audit scope are described in `README.md`, `README.ko.md` and `AGENTS.md`; the two independent reviews' findings (twelve and nine), their fixes and the re-run checks are recorded in `docs/verification-captions-crossrefs-20261114.md`.

### Fixes
- `w:updateFields` is inserted at its schema position in `settings.xml` in every profile. Output changes only in element order.
- Logo paths written with percent-encoding are resolved and judged candidate by candidate, like body images. A relative logo path is made absolute without following links, so a Windows drive-relative path (`Z:logo.png`) keeps working.
- A line that starts with `+ ` is a bullet item, like `- ` and `* `. Such a line also ends the paragraph before it.
- `#####` and `######` headings are headings, rendered with the level 4 style; they used to be paragraphs that kept their hash marks. A references header at those levels (`##### References`) starts the reference section like one at levels 1 to 4, so its items appear once, as endnotes.
- A backslash escapes every ASCII punctuation character, as in CommonMark. The fourteen that kept their backslash (`! # % & + , / : ; < = > ? @`) now lose it in paragraphs, headings, lists, quotes, callouts, emphasis, link text and table cells. Unchanged: code spans and blocks, link destinations, an escaped `<br>`, footnote text, and everything between dollar signs (`$5\%$`, `$a\,b$`), whether or not that context renders math. The engine cannot tell currency from math there, so an escape between two currency dollar signs on one line also keeps its backslash. In quotes and numbered headings the link-destination rule covers destinations without spaces or parentheses. With table spans off, an escaped `<<` shows as `<<`, as an escaped `^^` already showed as `^^`.
- In a table cell an escaped backslash is resolved once. `x\\*` shows `x\*`; it used to be unescaped twice and show `x*`. Paths written with doubled backslashes (`C:\\Users`) show as before.
- A closed code fence keeps its text exactly: blank lines and spaces at its end are no longer dropped, and a blank line no longer carries a one-space run. Line heights in Word are unchanged.
- Headings keep the characters of code spans and escaped punctuation (`` # `**kwargs` `` and `# \*literal\*` used to lose their asterisks). Other headings render as before.
- A whitespace-only line in a fenced code block keeps its own whitespace.
- `docx-audit` reads the package's content-type list first and checks every part it declares as XML, and every relationship part, with entity resolution, DTD loading and network access off, before opening the document. Part extensions are read by python-docx's own rule, so a part without an extension is covered. A package whose content-type list is missing or unreadable is reported and not opened. A package with a part that declares a DTD or entities is reported as an issue naming the part and is not opened or expanded. Results for ordinary documents are unchanged.

### Verification infrastructure
- CI gate `word-schema`: every sample (12 Markdown samples, 13 documents with the charts-on variant) is rendered in strict mode and checked with OOXML-Validator v2.1.6, downloaded and verified against a pinned SHA-256. The validation script verifies the downloaded archive in the same run that executes it, so an unverified validator never runs. An error not covered by `scripts/word_schema_exceptions.json` (currently empty) fails the job, and an exception covers an error only when the error type and description also match its evidence. See `scripts/word_schema_validation.md`.
- Snapshot tests pin the parsed model and every XML part of every sample; property tests (Hypothesis, fixed examples) check visible text, table cells, literal code and escapes, and that strict conversion either saves or rejects. The gaps they found in Markdown reading are fixed (see Fixes); the one strict `xfail` left in `tests/test_markdown_properties.py` records the intended tree-value emphasis inside code blocks. `tests/test_config_properties.py` applies the save-or-reject property to generated house and theme file content and to YAML built with aliases, and checks that each conversion finishes.
- CI also runs `ruff format --check`. `syrupy` and `hypothesis` are development-only dependencies.
- New engine module `input_safety.py`.

### Repository upkeep
- Removed two sample files with real company names and the prototype script that read them; test names were replaced with fictional ones. Earlier commits are not rewritten.
- Applied `ruff format` to the whole tree (the commit is listed in `.git-blame-ignore-revs`) and modernized type annotations. Sample output was compared before and after: all XML parts identical.

### Internal structure
- `ib_renderer.py` and `md_parser.py` are split into twelve internal files (`ib_render_*.py`, `md_parse_*.py`) by moving code only; every name is still imported from `ib_renderer` and `md_parser`, and output, commands, messages and log names are unchanged. Verification: `docs/verification-large-file-split-20261113.md`.

## [2.0.0-beta.1] - 2026-10-08

First public distribution prerelease. The release tag identifies this delivery set; engine/package metadata remains **2.0.0 (Beta)** and the separately versioned plugin remains **0.1.0**.

- Ship the same Markdown-to-Word engine as a Windows portable converter, a Claude Code / Codex source plugin, and an advanced Codex source-project kit.
- Put workflow selection first in both READMEs; distinguish distribution archive names and engine/plugin versions.
- Include synthetic examples and per-file manifests. The Windows package includes Python and dependency notices; source/plugin archives do not bundle a runtime.
- Keep the engine features listed below, including seven profiles, term variables, numerical checks, house styling and opt-in cleanup of externally converted Markdown.
- This is a prerelease, not financial/legal approval or a guarantee of page layout. Claude web/Cowork, native host GUI installation and a separate Python-uninstalled Windows machine are not certified. The release notes record the exact tested artifacts, environments and limitations.

## Engine source changes included in 2.0.0-beta.1 - 2026-09-29

### Converted term-sheet input (found on a second real term sheet)
- `md-format --converted-term-sheet` cleans Markdown from external HWP/Word converters into term-sheet input: bold numbered lines, or numbered lines followed by a table, become `##` chapter headings; one-row band tables become headings; cover lines before the first chapter move to frontmatter (`confidential_label`, `title`, `subtitle`, `date`, `disclaimer`) and cover images are removed for the house `style.logo`.
- Every moved, rewritten, removed or kept line is reported with its input line number (a range for multi-line changes), including one-line numbered paragraphs left as text. Unclear cover lines stay in place, existing frontmatter lines and keys (case-insensitive) are kept, `prepared_by` is never inferred, and only top-level blocks change: indented code, list and quote blocks, fences and HTML tables (nested ones included; a one-row band becomes a heading only when its text has no inline syntax characters and an HTML band has nothing outside its cells) stay as written. A YAML `...` closer becomes `---` and flow-style frontmatter is rewritten as block YAML, both reported. The parser's default path and the formatter's default output are unchanged.
- New engine module `converted_md_cleaner.py`.

### Term-sheet house style (found on a second real term sheet)
- House (and frontmatter) `style` options: a cover page (`cover: page`) with an optional logo (path or `data:` URI), a boxed disclaimer, header and footer accent rules, the confidentiality label colour, the footer page format (`- {page} -`) and position, a dark table header and open-sided tables. Frontmatter settings override the house file key by key; the defaults keep the standard layout unchanged.
- A null `style` value is a setting error (omit the key for the default), so it no longer passes validation and then stops the footer or drops the logo. `page_number` rejects every line break, including carriage returns. `label_color: ""` keeps the default grey, and frontmatter `logo: ""` removes a house logo.
- A landscape table that directly follows a lone page break (cover page, TOC) drops that page break, because its section break already starts a new page. Word 16 rendered the same pages before and after; the redundant break could show as a blank page in other viewers. A paragraph with several page breaks is kept.

### Numeric checks (found on a second real term sheet)
- `checks:` declares relations between term values (`all_in = issue_rate + credit_fee + running_cost`, `facility = amount * 1.05`). Values are read as displayed: Korean money units, `%`/`%p`, bp, `개월`/`년` and plain numbers. The default tolerance is half of the left value's display step, and failures are warnings that strict mode rejects.
- The checks and their generated values are stored in the DOCX (`ibrep.checks`), and `docx-audit` re-evaluates them with the current tagged values (`terms.failed_checks`).
- Table spec `schedule` checks a repayment schedule's arithmetic: balance steps, final zero, repayments equal to the principal (a number, or a money term converted with the table `unit`), the totals row and an optional stated weighted average life.
- New engine module `numeric_checks.py`. The term-sheet sample declares its checks and its schedule.

### Table structure (found on a second real term sheet)
- Table spec `header_rows` draws several header rows in every profile: shaded, merged where spans say so, and repeated on every page. HTML tables keep the rows their header spans cover as header rows instead of joining them into one. `base_case` rows count from the first body row, and zebra shading starts on the first body row.
- Term sheets: a cell that starts in the second label column and spans into the content columns is laid out as content; a first-column label stays a label however far it spans.
- Term-sheet grid tables size each column to fit its lines and share the spare width equally, so short amount columns are no longer starved while date and label columns balloon. Tables whose content cannot fit side by side keep the content estimate.
- Column kinds skip dash placeholders (`-`, `–`), so an amount column that starts with them is still numeric and right-aligned.

### Input loss in converted documents (found on a converted real term sheet)
- `*` and `_` emphasis needs flanking delimiters (CommonMark): an opening delimiter followed by whitespace, or a closing one preceded by it, stays literal. `2 * 3 * 4` and spaced note markers (`매출처* ...`, `* 주요 매출처`) keep their asterisks instead of losing them to a false italic; `** text **` is no longer bold.
- HTML `<table>` blocks, as HWP/Word converters write them, become Word tables in every profile: `colspan`/`rowspan` merge cells (empty rows under a span keep their place), block tags (`<p>`, `<div>`, `<li>`, nested rows and tables) and `<br>` break cell lines, `<b>`/`<strong>`, `<i>`/`<em>`, `<code>`, `<sup>`/`<sub>`, colour spans and `<a href>` format text and combine when nested, `<img>` becomes a cell image (`data:` URIs included), `{{key}}` references are substituted, and a nested table is flattened into its cell. Cells are built directly as runs, so HTML text (including decoded `&lt;tags&gt;` and Markdown characters) is never re-read as Markdown. Several header rows merge into one header row. Text outside cells, extra tables on a closing line and unclosed tables are reported, and strict mode rejects them.
- Images inside text and table cells (`![alt](path)`) are inserted inline and fitted to the cell width; a failed image leaves a visible marker that strict mode rejects. An image is one unit: `$` in its path is not math, surrounding emphasis still applies, its fields are never term-substituted, and image-like text inside a link destination stays part of the link. A standalone `<img>` line is an image.
- Image paths written percent-encoded (`images/a%20b.png`) find their file; a name that really contains `%` is still tried first.

### Internal memo rendering (found on a real memo)
- Inline code is rendered without its backticks in the code font (`CODE_FONT`); its text, including edge spaces in term-sheet lines, stays literal and outside table number formatting. Plain-profile callouts render inline code and links like body text.
- Local file links (angle-bracket destinations, drive or `./`/`../` paths, document extensions) become Word hyperlinks. Saving through the CLI, converter or registry rebases links on the DOCX folder: relative links keep reaching their file, absolute paths inside the DOCX folder become relative, and other absolute paths stay `file:///` links with a log warning. `#` and `%` in absolute paths are literal; a `#` after a document extension in a relative link is a fragment, and `file:` fragments such as `#page=2` are kept. UNC paths become `file://server/share` links; targets that are not valid local paths are kept as written with a warning. Reference definitions with spaces, brackets or any path form resolve correctly.
- General profiles turn off Word's automatic Korean/Latin and Korean/number spacing (`SPC는`, `제2종`, `300억원`), as the term-sheet profile already did.

### Existing-output fixes (schema order, forced TOC, audit encoding)
- Insert table borders, cell fills and margins, paragraph borders and run styles in ECMA-376 schema order in every profile (`ooxml_order.insert_ordered`). Output changes only in element order; content is identical.
- `business-report`, `meeting-minutes`, `office-letter` and `ib-memo` without a cover now render their title opening before a forced TOC, and the TOC preview no longer lists the title heading (matching `ib-report` and `term-sheet`).
- `docx-audit` writes its JSON report as UTF-8 regardless of the console code page, so redirected output parses on Windows (cp949).

### Term-sheet profile, explicit cell spans and house boilerplate
- Add the seventh profile, `term-sheet`, with validated house/frontmatter text, source-relative house paths and a `--house` override. Real deal documents and institution wording stay outside the repository.
- Resolve explicit `^^`/`<<` cell spans and label columns; support escaped literal markers, rectangular merge validation and per-table opt-in for other profiles.
- Render term-sheet documents: opening block (title, subtitle, date, prepared-by, disclaimer), confidentiality header and version/page footer on every section, labelled term tables with fixed label widths, label tiers, merged cells, per-line hanging indents and estimated row splitting, a two-row confirmation box, and Korean word-boundary wrapping without automatic Latin/number spacing.
- Merge validated spans in every profile's tables; add the optional table `note` (below the table, right-aligned) for every profile.
- Add a fictional Korean ABCP sample, repayment schedule and freshly written house boilerplate, plus bilingual authoring documentation.

### Term variables and consistency checks
- Add `terms:` values referenced as `{{key}}` in every profile, substituted at parse time as literal runs (never re-parsed, excluded from number formatting). Undefined keys warn and fail strict mode.
- Wrap substituted values in Word content controls tagged `ibrep:term:<key>` and snapshot the generated values in `ibrep.term.<key>` custom properties; `--no-term-tags` / `layout.term_tags: false` emits plain text.
- `docx-audit` reports `mismatched`, `changed`, `missing` and `indicative` term values after Word edits as warnings; documents without terms keep the previous JSON output.

### Title block without a cover (owner decision, 2026-09-29)
- When the cover is not rendered (`--no-cover`, `termsheet`/`legal-memo` presets), `ib-report` now starts with a title block (title, subtitle, date/author) before the TOC instead of omitting the title. The block is not a TOC entry; a matching H1 or inferred subtitle appears once. Cover-on output is unchanged.

### Korean glyphs in chart, equation and diagram images
- Rasterized images now use an installed CJK-capable font (preferred theme font first, then Malgun Gothic, Apple SD Gothic Neo, Nanum, Noto/Source Han Sans KR). Previously Linux fell back to DejaVu Sans and dropped every Hangul glyph. DOCX font declarations are unchanged.
- If Hangul must be rasterized and no CJK font is installed, strict mode rejects before saving; non-strict mode warns. Linux CI installs `fonts-nanum`.

### Python 3.12 floor (owner decision D1, 2026-09-29)
- Require Python 3.12+ (`requires-python >=3.12`); CI tests Ubuntu/Windows on Python 3.12 and 3.13, and the syntax gate checks 3.12. The lock file only drops old-Python resolution forks; dependency versions are unchanged.
- Python 3.8–3.11 are no longer supported; use an earlier source revision for those runtimes.

### Themes, section presets and opt-in charts
- Port PR #5 chart fences as proper model elements through the shared renderer: grouped bar, line and cumulative waterfall, retained YAML syntax plus `type`, `unit` and unscaled number formats. Charts default off; CLI/API options override frontmatter.
- Render charts to in-memory PNGs with Figure/Agg and per-artist Korean fonts/colors, without pyplot or rcParams mutation. Strict failures reject before saving; non-strict failures retain the source code panel and a named diagnostic.
- Add immutable `ib-report`, `termsheet`, `legal-memo` and `lecture-note` section bundles, `--preset` / `--list-presets`, and frontmatter `preset`. Explicit fields override caller presets, then YAML fields/presets, then profile defaults.
- Extend request-local themes with typed uppercase presentation fields, stricter color/size validation, paired RGB/hex handling, and themed code-panel backgrounds. Preserve the default/mono output and general-profile semantics.
- Package the chart engine explicitly and add a fictional Korean chart sample, usage documentation and parse/render/CLI regressions. No retired Word-to-Markdown code or global-style loader is restored.

### Input-loss and correctness hardening (2026-09-29)
- Stop silent content loss: escaped dollars (`\$5`), escaped emphasis, prose/tables after a `References` list, unrecognized leading `**Label:** value` paragraphs, dash-only table body rows, list continuation lines, and CRLF list items.
- Numeric table cells keep native footnote references and are no longer re-formatted from concatenated run text (`1234[^1]` stays `1,234` plus the footnote).
- Equation rendering failures are renderer errors: strict mode rejects them before saving; non-strict keeps visible fallback text.
- Shared fence scanner (backtick/tilde, 4+ fences, paragraph interruption); setext `===` H1; HTML comments are not printed; reference-style links resolve to hyperlinks; image titles and `<angle-bracket>` destinations parse correctly; literal `[^n]` inside code is not a footnote.
- Consecutive leading metadata lines (`**Date:**`/`**Analyst:**`) populate separate fields. Recognized header labels are canonicalized (`date`, `analysis_period`, `analysis_basis`). An inferred IB subtitle stays on the cover and becomes a body heading when the cover is disabled.
- Nested ordered lists restart under each parent item; start overrides target the actual level.

### Robustness and performance (2026-09-29)
- Atomic saves through a sibling temp file and replace (existing reports survive failed writes; file mode preserved); exclusive, collision-resistant locked-file fallback names.
- Batch mode rejects colliding output names before writing and exits non-zero when any input fails. Missing qualified input paths no longer fall back to an unrelated same-named file.
- Thread-safe default converter registry initialization. Diagram/equation rendering no longer mutates global Matplotlib state and releases figures/temp files on failure.
- Render-path structural validation no longer resolves styles per paragraph: a 2,000-element document renders about 60% faster. `docx-audit` validates numbering references and reports placeholder-like user text as warnings.
- GitHub Actions workflow (Ubuntu/Windows): ruff, mypy, pytest, build, exact wheel payload, Python 3.8 syntax gate.

### Office-letter production quality
- Render inline HTML breaks in paragraphs, emphasis, headings, lists and tables while preserving escaped/code literals and link destinations.
- Compose formal letter metadata, native attachment numbering, closing/signatory/issue/contact details, and a validated explicit `letter.appendix_heading` boundary through the shared renderer.
- Align wrapped metadata values and suppress the redundant company page header. Keep meaningful heading pagination controls; do not blanket-enable them on body text.
- Separate structural errors, pagination-marker warnings and visual-review status in `docx-audit`.
- Harden Windows Word QA: reject stale output directories, validate expected page counts, record file hashes and Word version, and leave visual review pending until pages are inspected.
- Add synthetic regressions and a complete appendix-letter example. Private transaction files are not fixtures. Word→Markdown remains retired.

## [2.0.0] - 2026-09-14

### Breaking changes
- Retired in-house Word→Markdown development by product decision. Removed the inverse CLI, parser, Markdown renderer, OMML reverse converter, roundtrip audit, and their dedicated tests. Recovery point: `d819bbb` / `codex/archive-word-to-md-d819bbb`.
- The built-in registry now supports only Markdown input and DOCX output. `docx-audit` replaces roundtrip checks with structural checks, not semantic reverse conversion.
- Sensitivity base-case highlighting now requires explicit coordinates; the engine no longer guesses the centre cell.

### Added
- Six document profiles: `ib-report`, `ib-memo`, `plain`, `office-letter`, `business-report`, `meeting-minutes`.
- Validated YAML metadata, layout settings, small YAML themes, CLI overrides, and strict diagnostics.
- A4 office layouts, recipients/sender/attachments, report and meeting metadata, native Korean multilevel numbering.
- Explicit table column roles, captions, units, sources, as-of dates, landscape sections, repeating headers, and native external hyperlinks/numeric footnotes.
- Shared model module, isolated request-local styles, examples, structural audit and profile regression tests.
- Repeatable Windows Word page-image QA script and long-table/landscape fixtures.

### Fixed
- Empty cells and escaped pipes keep their original columns; all-empty rows are retained.
- Financial number formatting reaches the actual output runs, including bold numbers and negative colouring.
- CLI, API and registry use one composition path with consistent headers, page fields and disclaimer toggles.
- General documents retain their References sections and do not use IB financial/header inference.
- Distribution packages explicitly include engine modules and exclude private reports and retired modules.
- Valid UTF-8 Korean is decoded before statistical detection; legacy EUC-KR/CP949 remains supported. Local images resolve relative to the Markdown file.
- Declared malformed YAML fails clearly; long YAML titles and explicit profile overrides retain the intended profile.
- Explicit footnotes work in headings and table headers; plain superscripts are not mistaken for footnotes.
- TOC preview entries are inside the field result and are replaced on update, without duplicate entries or the TOC listing itself.
- Neutral office titles no longer inherit the default Word title border; explicit fonts win over theme fonts. IB memo spacing and section-aware header alignment are refined.

## [1.0.3] - 2026-04-13

### Fixed
- Improved unicode LaTeX fallback rendering so mixed Korean equations preserve readable Greek letters and math operators instead of emitting raw LaTeX commands.
- Added regression coverage for Korean math fallback cases such as `\alpha` and `\sum`.

## [1.0.2] - 2026-04-13

### Changed
- Installed `matplotlib` through the default dependency set so Markdown LaTeX renders during standard Markdown-to-Word conversion.
- Added a readable plain-text image fallback for LaTeX expressions that include Korean text or other non-ASCII content.
- Normalized CLI stdout and stderr to UTF-8 on Windows so logging and piped output stay stable with report titles and Unicode text.

### Fixed
- Removed standalone HTML anchor tags such as `<a id="제3장"></a>` before Markdown paragraphs are rendered into Word.
- Preserved smoke-test report generation on Windows even when console encoding cannot represent some characters directly.

### Documentation
- Updated installation guidance to reflect that LaTeX rendering is part of the default install path.

## [1.0.1] - 2026-03-24

### Changed
- Improved Markdown-to-Word table rendering so body columns use content-aware widths instead of near-uniform spacing.
- Inferred table body alignment from the first 2-3 data rows so textual columns stay left-aligned while numeric columns align right.
- Switched cover page and table-of-contents typography to `Malgun Gothic`, including Word-generated `TOC 1` to `TOC 4` styles.

### Fixed
- Reduced unnecessary line wrapping in narrow numeric columns and long descriptive table columns.
- Prevented mixed digit-text codes such as `A-101` from being misclassified as numeric body columns.

### Documentation
- Synced project progress notes in `next_step.md` and `roadmap.md` with the current `main` implementation state.
