"""Markdown properties, with source-side oracles independent of the engine parser.

Generation scope: bounded documents contain paragraphs (including soft wraps,
spaces/tabs and parentheses), ATX headings, ordered/unordered lists, rectangular
pipe tables (1-5 rows including the header, 1-4 columns, empty cells), emphasis,
backtick spans, backtick/tilde fences and backslash-escaped ASCII punctuation.
Text combines Korean syllables, Latin letters, digits and punctuation. Blocks
are permuted, and inline constructs occur in paragraphs, headings, list items
and table cells. Code includes Markdown-looking text, tabs, spaces, blank lines,
backticks, HTML breaks and footnote/term-looking text. Expected text is composed
alongside the source, never obtained from parser/model text or engine helpers.
ATX levels 1-6 and all three bullet markers remain in scope; unsupported output
for levels 5/6 and '+' is retained as strict xfail parameter groups, not filtered.
Backslash escapes cover all ASCII punctuation in two groups: the characters the
engine resolves (TextParser's escapable set: backslash $ ` ^ ~ . * " ' ( ) [ ] { }
| _ -) are generated inside documents, tables and the escape property; the other
fourteen (! # % & + , / : ; < = > ? @) keep their backslash in the output today
and stay as strict xfail groups until the owner decides the escape set.

Exclusions from the structured preservation grammar (README.md, "Supported
Markdown and limitations"): raw HTML other than documented breaks, and complex
nested Markdown. Justification: "Raw HTML, complex nested Markdown and arbitrary
Word templates are not guaranteed." Neither exclusion filters the arbitrary
strings used for strict conversion. No generated example is discarded/assumed.
Unescaped syntax delimiters are grammar, not visible characters; structural
paragraph/cell boundaries are compared separately from run text. Paragraph
normalization follows AGENTS.md; code payloads bypass it. README.md: "Inline code
is shown without its backticks, in the code font (`CODE_FONT`, default Consolas);
its content stays literal (no emphasis, breaks, footnotes or term substitution)
and is excluded from table number formatting." Code font/panel styling is allowed;
Markdown emphasis, hyperlinks, footnotes and math inside literals are not.
An isolated tab is not repeated whitespace. Escape checks normalize surrounding
prose on both sides (including spaces adjacent to escaped parentheses); they do
not normalize code. Table probes explicitly disable cover/TOC/disclaimer using
documented RenderOptions, so presentation tables are not counted as input cells.

Strict input additionally includes unclosed fences/spans/emphasis/links/tables,
malformed frontmatter and arbitrary Unicode strings up to 100 characters. C0/C1
controls are included: the public text parser accepts str and attempts UTF-8
encoding. Lone surrogates are not accepted by that boundary (parse() calls
content.encode("utf-8")), so the arbitrary alphabet is Unicode scalar values.
Explicit examples cover NUL, empty input, malformed YAML and known rejections.
All strict examples exercise text->model->registry save, stream->registry save
and the CLI converter class, under both requested profiles.
Valid YAML with wrong scalar/collection types and nested settings is generated
for every known metadata key. Separate per-key cases let YAML select the profile
(including term-sheet and office-letter), so operator overrides cannot mask
profile or profile-specific validation. Escaped XML-incompatible metadata
characters, including surrogate escapes, are also in the strict source grammar.
Invalid date/month/hour scalars and explicitly tagged scalars the loader cannot
construct (int, float, timestamp, bool) are generated as raw YAML (safe_dump would quote them). Each renderer text key and
sender/letter child is also exercised under plain and ib-report with both cover
settings, so profile selection cannot hide an unvalidated IB field. Hand-built
models vary paragraph/heading/run/cell/list/footnote/metadata text, including NUL,
vertical tab, U+FFFF and a lone surrogate, and go through registry save/reject.
This is bounded coverage, not a claim about all input. It does not generate
house/theme file contents, asset bytes or filesystem graphs, arbitrary YAML tags
or aliases, every combination of nested settings, or arbitrary malformed model
shapes. Dedicated regressions cover selected named-configuration and asset cases.
Fenced-code generation partitions the original alphabet: nonempty lines and a
non-whitespace final character have an active text/formatting property; empty
lines and trailing whitespace retain generated strict xfail groups. Tree-value
emphasis retains its separate minimal strict xfail. No payload is filtered out.

Determinism: local default profile markdown-properties, derandomize=True, deadline=None,
database=None, max_examples=24. TemporaryDirectory is created per example under
the OS temp directory; no function-scoped pytest fixture is used with @given.
"""

import re
import string
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest
import yaml
from docx import Document
from docx.document import Document as WordDocument
from docx.oxml.ns import qn
from hypothesis import example, given, settings
from hypothesis import strategies as st

from converters import get_default_registry
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from input_safety import InputLimitError, InputLimits, StrictValidationError
from md_parser import MarkdownParser
from md_to_word import IBReportConverter

settings.register_profile(
    "markdown-properties", derandomize=True, deadline=None, database=None, max_examples=24
)
settings.load_profile("markdown-properties")

LETTERS = string.ascii_letters + string.digits + "가상예시한글문서값항목"
PROSE = LETTERS + " .,;:!?%/@=+()-"
CODE = LETTERS + string.punctuation + " \t"
# Written out here, not read from the engine, so the oracle stays independent.
RESOLVED_ESCAPES = "\\$`^~.*\"'()[]{}|_-"
RETAINED_ESCAPES = "".join(char for char in string.punctuation if char not in RESOLVED_ESCAPES)
RETAINED_ESCAPE_REASON = (
    "'가상\\!예시' renders '가상\\!예시', expected '가상!예시'; TextParser._ESCAPE_RE "
    "resolves only its own set. Widening it changes visible text for existing input "
    "and one pinned table test, so the owner decides."
)


@dataclass(frozen=True)
class Fragment:
    """A generated source and its independently known visible characters."""

    source: str
    visible: str


class ParagraphNormalization:
    """Only the three paragraph transformations documented in AGENTS.md."""

    spaces = re.compile(r"[ \t]{2,}")
    opening = re.compile(r"\( +")
    closing = re.compile(r" +\)")

    @classmethod
    def apply(cls, text: str) -> str:
        text = cls.spaces.sub(" ", " ".join(text.split("\n")))
        return cls.closing.sub(")", cls.opening.sub("(", text)).strip()


class CodeDelimiters:
    """Delimiter runs used only to construct unambiguous literal syntax."""

    backticks = re.compile(r"`+")
    tildes = re.compile(r"~+")


def _span(payload: str) -> Fragment:
    # A longer delimiter and padding isolate literal backticks at either end.
    delimiter = "`" * (
        max((len(run) for run in CodeDelimiters.backticks.findall(payload)), default=0) + 1
    )
    padding = " " if payload.startswith("`") or payload.endswith("`") else ""
    return Fragment(delimiter + padding + payload + padding + delimiter, payload)


@st.composite
def _inline(draw: st.DrawFn) -> Fragment:
    kind = draw(st.sampled_from(("text", "bold", "italic", "escape", "code")))
    text = draw(st.text(LETTERS, min_size=1, max_size=12))
    if kind == "text":
        return Fragment(text, text)
    if kind in {"bold", "italic"}:
        marker = "**" if kind == "bold" else draw(st.sampled_from(("*", "_")))
        return Fragment(marker + text + marker, text)
    if kind == "escape":
        punctuation = draw(st.sampled_from(tuple(RESOLVED_ESCAPES)))
        return Fragment(text + "\\" + punctuation + text, text + punctuation + text)
    return _span(text + draw(st.sampled_from(("**가**", "<br>", "[^1]", "{{예시}}", "  ( 값 )"))))


@st.composite
def _table(draw: st.DrawFn) -> tuple[str, tuple[tuple[str, ...], ...]]:
    rows = draw(st.integers(min_value=1, max_value=5))
    columns = draw(st.integers(min_value=1, max_value=4))
    source_rows = []
    expected_rows = []
    for _ in range(rows):
        cells = draw(
            st.lists(
                st.one_of(st.just(Fragment("", "")), _inline()), min_size=columns, max_size=columns
            )
        )
        source_rows.append("| " + " | ".join(cell.source for cell in cells) + " |")
        expected_rows.append(tuple(cell.visible for cell in cells))
    source_rows.insert(1, "| " + " | ".join(["---"] * columns) + " |")
    return "\n".join(source_rows), tuple(expected_rows)


@st.composite
def _code_case(draw: st.DrawFn) -> tuple[str, str, str, str]:
    # Surrounding letters keep space padding unambiguous in inline code.
    span = "가" + draw(st.text(CODE, max_size=30)) + "상"
    lines = draw(st.lists(st.text(CODE, max_size=25), min_size=1, max_size=4))
    fenced = "\n".join(lines)
    escaped = draw(st.text(RESOLVED_ESCAPES, min_size=1, max_size=12))
    retained = draw(st.text(RETAINED_ESCAPES, min_size=1, max_size=12))
    return span, fenced, escaped, retained


def _fence(payload: str, marker: str = "`") -> str:
    runs = (CodeDelimiters.backticks if marker == "`" else CodeDelimiters.tildes).findall(payload)
    delimiter = marker * (
        max(
            3,
            max((len(run) for run in runs), default=0) + 1,
        )
    )
    return delimiter + "text\n" + payload + "\n" + delimiter


@st.composite
def _document(
    draw: st.DrawFn,
    list_marker: str | None = None,
    heading_levels: tuple[int, ...] = (1, 2, 3, 4, 5, 6),
) -> Fragment:
    fragments = draw(st.lists(_inline(), min_size=4, max_size=4))
    prose = draw(st.text(PROSE, min_size=1, max_size=25))
    break_marker = draw(st.sampled_from(("", "<br>", "\\", "  ")))
    # Prefixes prevent randomly produced text being mistaken for a block marker.
    paragraph = "가상  ( " + prose + " )\t본문" + break_marker + "\n예시 " + fragments[0].source
    paragraph_text = (
        ParagraphNormalization.apply("가상  ( " + prose + " )\t본문")
        + ("\n" if break_marker in {"<br>", "\\"} else " ")
        + "예시 "
        + fragments[0].visible
    )
    heading = "#" * draw(st.sampled_from(heading_levels)) + " " + fragments[1].source
    marker = list_marker or draw(st.sampled_from(("-", "*", "+", "1.")))
    listing = marker + " " + fragments[2].source + "\n" + marker + " " + fragments[3].source
    table_source, cells = draw(_table())
    code = "가상 " + draw(st.text(CODE, max_size=20)) + " 예시"
    blocks = (
        Fragment(paragraph, paragraph_text),
        Fragment(heading, fragments[1].visible),
        Fragment(listing, fragments[2].visible + fragments[3].visible),
        Fragment(table_source, "".join(cell for row in cells for cell in row)),
        Fragment(_fence(code, draw(st.sampled_from(("`", "~")))), code),
    )
    ordered = draw(st.permutations(blocks))
    return Fragment(
        "\n\n".join(block.source for block in ordered), "".join(block.visible for block in ordered)
    )


def _render(source: str, profile: str = "plain") -> WordDocument:
    model = MarkdownParser(profile=profile).parse(source)
    return IBDocumentRenderer(
        options=RenderOptions(
            profile=profile,
            strict=True,
            include_cover=False,
            include_toc=False,
            include_disclaimer=False,
        )
    ).render(model)


def _body_text(document: WordDocument) -> str:
    # XML document order includes table cells (Document.paragraphs does not).
    values = {qn("w:tab"): "\t", qn("w:br"): "\n", qn("w:cr"): "\n"}
    return "".join(
        node.text or "" if node.tag == qn("w:t") else values.get(node.tag, "")
        for node in document.element.body.iter()
    )


def _assert_visible(document: WordDocument, expected: str) -> None:
    assert _body_text(document) == expected


def _assert_cells(
    document: WordDocument, expected: tuple[tuple[str, ...], ...], profile: str
) -> None:
    assert len(document.tables) == 1
    table = document.tables[0]
    assert len(table.rows) == len(expected)
    assert len(table.columns) == len(expected[0])
    for row, values in zip(table.rows, expected, strict=True):
        assert len(row.cells) == len(values)
        # Assert physical XML cells too: repeated merged-cell proxies can hide loss.
        assert len(row._tr.tc_lst) == len(values)
        for cell, value in zip(row.cells, values, strict=True):
            if profile == "plain" or not value:
                assert cell.text == value


def _assert_literal_runs(document: WordDocument) -> None:
    for run in document.element.body.iter(qn("w:r")):
        properties = run.find(qn("w:rPr"))
        if properties is not None:
            for name in ("b", "i", "strike", "dstrike", "u", "vertAlign"):
                prop = properties.find(qn("w:" + name))
                assert prop is None or prop.get(qn("w:val")) in {"0", "false", "none", "baseline"}
    for name in ("hyperlink", "footnoteReference", "drawing", "fldChar"):
        assert not list(document.element.body.iter(qn("w:" + name)))


@pytest.mark.parametrize(
    "marker",
    (
        "-",
        "*",
        pytest.param(
            "+",
            marks=pytest.mark.xfail(
                strict=True,
                reason="'+ 0' renders '+ 0', expected list text '0'; MarkdownParser.BULLET_PATTERN omits '+'. Fix changes visible list layout.",
            ),
        ),
        "1.",
    ),
)
@pytest.mark.parametrize(
    "levels",
    (
        (1, 2, 3, 4),
        pytest.param(
            (5, 6),
            marks=pytest.mark.xfail(
                strict=True,
                reason="'##### 0' renders '##### 0', expected heading text '0'; MarkdownParser._try_parse_heading only defines H1-H4. Adding H5/H6 needs layout design.",
            ),
        ),
    ),
)
@given(data=st.data())
def test_visible_text_preserved(marker: str, levels: tuple[int, ...], data: st.DataObject) -> None:
    """Compare the entire visible stream, not a set or subsequence."""
    case = data.draw(_document(list_marker=marker, heading_levels=levels))
    _assert_visible(_render(case.source), case.visible)


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@example(
    case=(
        "| 가상 | | 예시 | |\n| --- | --- | --- | --- |\n| 001234 | -1234.50 | 12.5% | |\n| | A\\|B | | 끝 |",
        (("가상", "", "예시", ""), ("001234", "-1234.50", "12.5%", ""), ("", "A|B", "", "끝")),
    )
)
@given(case=_table())
def test_cell_count_and_characters_preserved(
    profile: str, case: tuple[str, tuple[tuple[str, ...], ...]]
) -> None:
    source, expected = case
    _assert_cells(_render(source, profile), expected, profile)


@pytest.mark.parametrize(
    "kind",
    (
        "span",
        "escape",
        pytest.param(
            "escape-retained",
            marks=pytest.mark.xfail(strict=True, reason=RETAINED_ESCAPE_REASON),
        ),
    ),
)
@given(case=_code_case())
@example(
    case=("가상 **예시** <br> [^1] {{예시}} \\_", "가상\n", "\\$`*_[](){}|~^", "!#%&+/:;<=>?@")
)
def test_code_and_escapes_unchanged(kind: str, case: tuple[str, str, str, str]) -> None:
    span, _, escaped, retained = case
    if kind == "span":
        source, expected = _span(span).source, span
    else:
        characters = escaped if kind == "escape" else retained
        source = "가상 " + "".join("\\" + char for char in characters) + " 예시"
        expected = ParagraphNormalization.apply("가상 " + characters + " 예시")
    document = _render(source)
    if kind == "span":
        assert document.paragraphs[0].text == expected
    else:
        assert ParagraphNormalization.apply(document.paragraphs[0].text) == expected
    _assert_literal_runs(document)


@st.composite
def _working_fenced_payload(draw: st.DrawFn) -> str:
    # Partition the original line alphabet, rather than discard examples. All
    # nonempty intermediate lines (including whitespace-only lines) are active;
    # empty lines and trailing whitespace are in the strict xfail groups below.
    lines = draw(st.lists(st.text(CODE, min_size=1, max_size=25), max_size=3))
    last = draw(st.text(CODE, max_size=24)) + draw(st.sampled_from(tuple(CODE.rstrip(" \t"))))
    return "\n".join([*lines, last])


@pytest.mark.parametrize("marker", ("`", "~"))
@example(payload=" ")
@example(payload="가상\n \t \n**예시** <br> [^1] {{예시}} `x`")
@given(payload=_working_fenced_payload())
def test_generated_fenced_code_preserves_lines_and_formatting(marker: str, payload: str) -> None:
    document = _render(_fence(payload, marker))
    assert len(document.tables) == 1
    cell = document.tables[0].cell(0, 0)
    assert [paragraph.text for paragraph in cell.paragraphs] == payload.split("\n")
    _assert_literal_runs(document)


@pytest.mark.parametrize("marker", ("`", "~"))
@pytest.mark.xfail(
    strict=True,
    reason="An empty fenced line receives a one-space spacer; removing it needs a layout decision.",
)
@given(lines=st.lists(st.text(CODE, max_size=25), max_size=3), tail=_working_fenced_payload())
def test_generated_fenced_empty_lines_have_no_spacer(
    marker: str, lines: list[str], tail: str
) -> None:
    # Includes every original line kind, plus a guaranteed empty line.
    payload = "\n".join([*lines, "", tail])
    document = _render(_fence(payload, marker))
    assert document.tables[0].cell(0, 0).text == payload
    _assert_literal_runs(document)


@pytest.mark.parametrize("marker", ("`", "~"))
@pytest.mark.xfail(
    strict=True,
    reason="Trailing fenced whitespace/blank lines are dropped by rstrip/splitlines; restoring them needs a layout decision.",
)
@given(
    lines=st.lists(st.text(CODE, max_size=25), min_size=1, max_size=4),
    ending=st.text(" \t\n", min_size=1, max_size=4),
)
def test_generated_fenced_trailing_whitespace_is_preserved(
    marker: str, lines: list[str], ending: str
) -> None:
    payload = "\n".join(lines) + "예시" + ending
    document = _render(_fence(payload, marker))
    assert document.tables[0].cell(0, 0).text == payload
    _assert_literal_runs(document)


BROKEN = st.sampled_from(
    (
        "```\n가상",
        "`가상",
        "**가상",
        "[가상](",
        "| 가상 |\n| --- |\n|",
        "---\nprofile: [\n---",
        "[^0]: 가상",
        "![예시](missing.png)",
        "[^1]",
    )
)

FRONTMATTER_KEYS = (
    "profile",
    "title",
    "subtitle",
    "company",
    "ticker",
    "sector",
    "analyst",
    "date",
    "layout",
    "style",
    "tables",
    "terms",
    "checks",
    "letter",
    "sender",
    "recipients",
    "cc",
    "attachments",
    "attendees",
    "charts",
    "preset",
    "house",
    "theme",
    "prepared_by",
    "disclaimer",
    "confidential_label",
    "confirmation",
    "subject_company",
    "recipient",
    "report_type",
    "analysis_period",
    "analysis_basis",
    "meeting_time",
    "location",
    "document_no",
    "version",
)
YAML_VALUE = st.recursive(
    st.one_of(
        st.none(), st.booleans(), st.integers(-2, 5), st.sampled_from(("wrong", "", "예시", 1.5))
    ),
    lambda child: st.one_of(
        st.lists(child, max_size=2),
        st.dictionaries(st.sampled_from(("예시", "wrong", 0)), child, max_size=2),
    ),
    max_leaves=5,
)
NESTED_SETTINGS = {
    "layout": (
        "cover",
        "toc",
        "disclaimer",
        "confidential",
        "separator_mode",
        "strict",
        "term_tags",
    ),
    "style": (
        "cover",
        "disclaimer",
        "header_rule",
        "footer_rule",
        "logo",
        "logo_width_mm",
        "label_color",
        "page_number",
        "page_number_align",
        "table_header",
        "table_sides",
    ),
    "tables": (
        "type",
        "columns",
        "caption",
        "unit",
        "source",
        "as_of",
        "landscape",
        "base_case",
        "spans",
        "label_columns",
        "note",
        "header_rows",
        "schedule",
    ),
    "letter": ("appendix_heading", "appendix_label"),
    "sender": ("organization", "department", "name", "signatory", "contact"),
    "confirmation": ("intro", "items", "signature"),
    "terms": ("amount",),
}


@st.composite
def _frontmatter(draw: st.DrawFn, key: str | None = None) -> str:
    key = key or draw(st.sampled_from(FRONTMATTER_KEYS))
    value = draw(YAML_VALUE)
    if key in NESTED_SETTINGS and draw(st.booleans()):
        value = {draw(st.sampled_from(NESTED_SETTINGS[key])): value}
        if key == "tables":
            value = [value]
    data = {
        "profile": "plain",
        "title": "예시 문서",
        "prepared_by": "가상 기관",
        "disclaimer": "예시 문구",
        "recipients": ["가상 수신처"],
        "sender": {"organization": "가상 기관"},
    }
    if key in {"style", "confirmation", "prepared_by", "disclaimer", "confidential_label"}:
        data["profile"] = "term-sheet"
    elif key in {"letter", "sender", "recipients", "cc", "attachments"}:
        data["profile"] = "office-letter"
    data[key] = value
    return "---\n" + yaml.safe_dump(data, allow_unicode=True) + "---\n| 예시 |\n| --- |\n| 가상 |"


ESCAPED_METADATA = st.tuples(
    st.sampled_from(("title", "subtitle", "company", "analyst", "date")),
    st.sampled_from((r"\0", r"\x0b", r"\u001F", r"\uFFFE", r"\uFFFF", r"\uD800")),
).map(lambda pair: f'---\n{pair[0]}: "예시{pair[1]}"\n---\n가상 본문')

STRICT_SOURCE = st.one_of(
    _document().map(lambda case: case.source),
    BROKEN,
    _frontmatter(),
    ESCAPED_METADATA,
    st.sampled_from(
        (
            "2026-99-01",
            "2026-01-01T25:00:00Z",
            "!!int example",
            "!!float example",
            "!!timestamp example",
            "!!bool example",
            '!!int ""',
            '!!float ""',
        )
    ).map(lambda scalar: f"---\nprofile: plain\ndate: {scalar}\n---\n가상 본문"),
    st.text(st.characters(exclude_categories=("Cs",)), max_size=100),
)


def _is_rejection(error: Exception) -> bool:
    seen: set[int] = set()
    current: BaseException | None = error
    while current is not None and id(current) not in seen:
        if isinstance(current, (StrictValidationError, InputLimitError)):
            return True
        seen.add(id(current))
        current = current.__cause__
    return False


def _strict_save(
    source: str, output: Path, profile: str | None, route: str, limits: InputLimits
) -> str:
    registry = get_default_registry()
    if route == "cli":
        input_path = output.with_suffix(".md")
        input_path.write_text(source, encoding="utf-8")
        return IBReportConverter(
            str(input_path), str(output), RenderOptions(profile=profile, strict=True), limits=limits
        ).convert()
    if route == "stream":
        model = registry.convert(
            BytesIO(source.encode("utf-8")), extension_hint="md", profile=profile, limits=limits
        )
    else:
        model = MarkdownParser(profile=profile, limits=limits).parse(source)
    return registry.convert(
        model, output_format="docx", output_path=output, profile=profile, strict=True, limits=limits
    )


def _assert_strict_outcome(
    source: str, directory: Path, profile: str | None, route: str, limits: InputLimits
) -> None:
    output = directory / "example.docx"
    try:
        saved = _strict_save(source, output, profile, route, limits)
    except Exception as error:
        assert not output.exists(), "rejected input left an output file"
        assert not list(directory.glob("*.docx")), "rejection left an alternate output"
        assert sorted(path.name for path in directory.iterdir()) == (
            ["example.md"] if route == "cli" else []
        ), "rejection left a partial save"
        assert _is_rejection(error), f"undocumented rejection: {type(error).__name__}: {error}"
    else:
        _assert_saved_output(saved, output, route)


def _assert_saved_output(saved: str, output: Path, route: str) -> None:
    assert Path(saved) == output
    assert output.is_file()
    document = Document(output)
    assert document.element.body is not None
    assert sorted(path.name for path in output.parent.iterdir()) == (
        ["example.docx", "example.md"] if route == "cli" else ["example.docx"]
    )


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@pytest.mark.parametrize("route", ("text", "stream", "cli"))
@example(source="", small_limit=False)
@example(source="가상 예시", small_limit=False)
@example(source="\x00", small_limit=False)
@example(source="---\ntitle: [\n---", small_limit=False)
@example(source="---\nprofile: [\n---", small_limit=False)
@example(source="[^0]: 가상", small_limit=False)
@example(source="# \x00", small_limit=False)
@example(source="![예시](missing.png)", small_limit=False)
@example(source="가상" * 20, small_limit=True)
@example(source="---\nlayout: {cover: wrong}\n---\n가상", small_limit=False)
@example(source='---\ntitle: "예시\\0"\n---\n가상', small_limit=False)
@example(source="---\nprofile: plain\ndate: 2026-99-01\n---\n가상", small_limit=False)
@example(source="---\nprofile: ib-report\nsender: {organization: 12}\n---\n가상", small_limit=False)
@given(source=STRICT_SOURCE, small_limit=st.booleans())
def test_strict_mode_dichotomy(profile: str, route: str, source: str, small_limit: bool) -> None:
    limits = InputLimits(max_input_bytes=32) if small_limit else InputLimits()
    with TemporaryDirectory(prefix="ib-prop-") as temporary:
        _assert_strict_outcome(source, Path(temporary), profile, route, limits)


@pytest.mark.parametrize("key", FRONTMATTER_KEYS)
@pytest.mark.parametrize("route", ("text", "stream", "cli"))
@given(data=st.data())
def test_frontmatter_strict_mode_dichotomy(key: str, route: str, data: st.DataObject) -> None:
    """Exercise every key without an operator profile masking the YAML profile."""
    source = data.draw(_frontmatter(key))
    with TemporaryDirectory(prefix="ib-prop-frontmatter-") as temporary:
        _assert_strict_outcome(source, Path(temporary), None, route, InputLimits())


RENDER_TEXT_KEYS = (
    "title",
    "subtitle",
    "company",
    "ticker",
    "sector",
    "analyst",
    "date",
    "subject_company",
    "recipient",
    "report_type",
    "analysis_period",
    "analysis_basis",
    "meeting_time",
    "location",
    "document_no",
    "version",
    "attendees",
    "recipients",
    "cc",
    "attachments",
    "sender",
    "sender.organization",
    "sender.department",
    "sender.name",
    "sender.signatory",
    "sender.contact",
    "letter.appendix_heading",
    "letter.appendix_label",
)


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@pytest.mark.parametrize("key", RENDER_TEXT_KEYS)
@given(value=YAML_VALUE, cover=st.booleans(), route=st.sampled_from(("text", "stream", "cli")))
def test_render_metadata_strict_dichotomy(
    profile: str, key: str, value: object, cover: bool, route: str
) -> None:
    """Generate each consumed text field with profile and cover paths held explicit."""
    fields = {"profile": profile, "title": "예시 문서", "layout": {"cover": cover}}
    if "." in key:
        parent, child = key.split(".")
        fields[parent] = {child: value}
    else:
        fields[key] = value
    source = "---\n" + yaml.safe_dump(fields, allow_unicode=True) + "---\n가상 본문"
    with TemporaryDirectory(prefix="ib-prop-render-") as temporary:
        _assert_strict_outcome(source, Path(temporary), profile, route, InputLimits())


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@pytest.mark.parametrize(
    "location", ("paragraph", "heading", "run", "table", "list", "footnote", "metadata")
)
@given(
    text=st.text(PROSE, max_size=20),
    character=st.sampled_from(("", "\x00", "\x0b", "\uffff", "\ud800")),
)
def test_handbuilt_model_strict_dichotomy(
    profile: str, location: str, text: str, character: str
) -> None:
    from document_model import (
        DocumentModel,
        Element,
        ElementType,
        Heading,
        ListItem,
        Paragraph,
        Table,
        TableCell,
        TableRow,
        TextRun,
    )
    from document_profiles import default_metadata

    value = "가상" + text + character
    model = DocumentModel(metadata=default_metadata(profile))
    model.elements = [
        {
            "paragraph": Element(ElementType.PARAGRAPH, Paragraph(value)),
            "heading": Element(ElementType.HEADING_1, Heading(1, value)),
            "run": Element(ElementType.PARAGRAPH, Paragraph("예시", [TextRun(value)])),
            "table": Element(
                ElementType.TABLE, Table(rows=[TableRow(cells=[TableCell(value)])], col_count=1)
            ),
            "list": Element(ElementType.BULLET_LIST, ListItem(value)),
            "footnote": Element(
                ElementType.PARAGRAPH, Paragraph("예시", [TextRun("", footnote_id=1)])
            ),
            "metadata": Element(ElementType.PARAGRAPH, Paragraph("예시")),
        }[location]
    ]
    if location == "footnote":
        model.footnotes[1] = value
    elif location == "metadata":
        model.metadata.title = value
    with TemporaryDirectory(prefix="ib-prop-model-") as temporary:
        output = Path(temporary) / "example.docx"
        try:
            saved = get_default_registry().convert(
                model, output_format="docx", output_path=output, strict=True
            )
        except Exception as error:
            assert list(output.parent.iterdir()) == []
            assert _is_rejection(error), f"undocumented rejection: {type(error).__name__}: {error}"
        else:
            _assert_saved_output(saved, output, "model")


@pytest.mark.parametrize(
    "punctuation",
    (
        *RESOLVED_ESCAPES,
        *(
            pytest.param(char, marks=pytest.mark.xfail(strict=True, reason=RETAINED_ESCAPE_REASON))
            for char in RETAINED_ESCAPES
        ),
    ),
)
def test_ascii_punctuation_escape(punctuation: str) -> None:
    """Each escaped character appears once, without its backslash, in every block kind."""
    source = "가상\\" + punctuation + "예시"
    expected = "가상" + punctuation + "예시"
    for template in ("{}", "# {}", "- {}", "| {} |\n| --- |"):
        _assert_visible(_render(template.format(source)), expected)
    _assert_literal_runs(_render(source))


@pytest.mark.parametrize(
    "source, expected",
    (
        ("# `**가상**`", "**가상**"),
        (r"# \*\*가상\*\*", "**가상**"),
        ("## `가상  <br> 예시`", "가상  <br> 예시"),
    ),
)
def test_heading_literal_regression(source: str, expected: str) -> None:
    """A heading must not strip literal stars or unescape before tokenization."""
    _assert_visible(_render(source), expected)


@pytest.mark.parametrize(
    "source, expected",
    (
        ("## **가상 예시**", "가상 예시"),
        ("## 1. **가상** 예시", "1. 가상 예시"),
        ("## **가상 예시", "가상 예시"),
        ("## 가상 ** 예시 **", "가상 예시"),
        ("## ***가상***", "가상"),
    ),
)
def test_heading_bold_markers_still_dropped(source: str, expected: str) -> None:
    """The literal fix must leave ordinary headings as they were: one run, no stars."""
    document = _render(source)
    heading = next(p for p in document.paragraphs if p.style.name.startswith("Heading"))
    assert ParagraphNormalization.apply(heading.text) == expected
    assert len(heading.runs) == 1


def test_fenced_whitespace_regression() -> None:
    """A whitespace-only code line keeps its own characters, not one ASCII space."""
    payload = "가상\n \t \n예시"
    document = _render(_fence(payload))
    assert document.tables[0].cell(0, 0).text == payload
    _assert_literal_runs(document)


@pytest.mark.parametrize("payload", ("", "가상\n\n예시"))
@pytest.mark.xfail(
    strict=True,
    reason="'```text\\n가상\\n\\n예시\\n```' renders '가상\\n \\n예시': IBDocumentRenderer._render_code_block_line puts a one-space spacer run on an empty line. Removing it may change the panel's line height, so the owner decides.",
)
def test_fenced_empty_line_has_no_spacer(payload: str) -> None:
    document = _render(_fence(payload))
    assert document.tables[0].cell(0, 0).text == payload


@pytest.mark.parametrize("payload", ("  ", "\t", "가상\n", "가상\n\n", "가상\n \t "))
@pytest.mark.xfail(
    strict=True,
    reason="'```text\\n \\n```' renders '', expected ' '; MarkdownParser._parse_elements rstrip() also drops trailing blank lines. Full literal preservation needs the owner to approve code-panel layout changes.",
)
def test_fenced_trailing_whitespace_preserved(payload: str) -> None:
    document = _render(_fence(payload))
    assert document.tables[0].cell(0, 0).text == payload


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@pytest.mark.parametrize("route", ("text", "stream", "cli"))
@pytest.mark.parametrize("source", ("---\nprofile: [\n---", "[^0]: 가상", "# \x00", "# \ufffe"))
def test_documented_rejection_regression(profile: str, route: str, source: str) -> None:
    """Malformed source must reject before saving, preserving a typed cause."""
    with TemporaryDirectory(prefix="ib-prop-regression-") as temporary:
        directory = Path(temporary)
        with pytest.raises(Exception) as caught:
            _strict_save(source, directory / "example.docx", profile, route, InputLimits())
        assert _is_rejection(caught.value)
        assert sorted(path.name for path in directory.iterdir()) == (
            ["example.md"] if route == "cli" else []
        )


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
def test_empty_cli_documented_rejection_regression(profile: str) -> None:
    """The CLI's empty-file preflight must use a documented rejection type."""
    with TemporaryDirectory(prefix="ib-prop-regression-") as temporary:
        directory = Path(temporary)
        with pytest.raises(StrictValidationError, match="Input file is empty"):
            _strict_save("", directory / "example.docx", profile, "cli", InputLimits())
        assert sorted(path.name for path in directory.iterdir()) == ["example.md"]


@pytest.mark.xfail(
    strict=True,
    reason="'```text\\n가상 1% ── 예시\\n```' makes '1%' bold; IBDocumentRenderer._render_code_block_line deliberately styles tree values. Removing this changes visible formatting.",
)
def test_fenced_code_has_no_semantic_emphasis() -> None:
    """Retain the layout decision separately from the trailing-line property."""
    _assert_literal_runs(_render(_fence("가상 1% ── 예시")))
