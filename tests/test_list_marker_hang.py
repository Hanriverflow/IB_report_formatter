"""Lines the block parser used to take without consuming, so it never finished.

A list marker followed only by whitespace or HTML anchors (`1.   `, `-  `,
`- <a id="x"></a>`) and a table or quote line led by HTML anchors
(`<a id="x"></a>| a |`, also after `>`: `><a id="x"></a> 인용`) each kept
`MarkdownParser._parse_elements` on one line forever, in every profile.

Two principles decide the result (owner's coordinator, 2026-11-14):

A. A document that finished on `c97841c` parses exactly as it did. A Base64
   image keeps taking a line before tables and quotes do.
B. A document that never finished reads as the same document with the
   leading anchors (and the whitespace around them) removed from the lines
   that never finished, or with a blank item written as `- ` / `1. `. What
   follows a block marker stays as written; the leading anchors are dropped,
   as they are before a paragraph.

Every parse and render runs in a child process under a time limit, so a
regression fails the test instead of stopping the run. Record:
docs/verification-list-marker-hang-20261114.md.
"""

import base64
import json
import random
import struct
import subprocess
import sys
import zlib
from pathlib import Path

import pytest

from document_profiles import PROFILES
from md_parser import Base64ImageParser

ROOT = Path(__file__).resolve().parents[1]
TIME_LIMIT = 120
ANCHOR = '<a id="x"></a>'
TABLE = "| a | b |\n|---|---|\n| 1 | 2 |"
KEEP = '`<a id="keep"></a>`'

# (input that never finished, nearest input that always finished)
LIST_PAIRS = [
    ("1.   ", "1. "),
    ("1.  ", "1. "),
    ("12.   ", "12. "),
    ("1.\t\t", "1. "),
    ("1.\u3000\u3000", "1. "),
    ("-   ", "- "),
    ("-  ", "- "),
    ("*   ", "* "),
    ("*  ", "* "),
    ("+   ", "+ "),
    ("+  ", "+ "),
    ("-\t\t", "- "),
    ("- \t", "- "),
    ("-\u3000\u3000", "- "),
    ("- \u3000", "- "),
    ("-\u00a0\u00a0", "- "),
    ("- \u2003\u2003", "- "),
    ("- \x0b\x0b", "-\x0b"),
    ("- \x0c\x0c", "-\x0c"),
    ("- \x1c\x1c", "-\x1c"),
    ("- \x85\x85", "-\x85"),
    ("    -   ", "    - "),
    ("  1.   ", "  1. "),
    ("\t-  ", "\t- "),
    ("- " + ANCHOR, "- "),
    ("-  " + ANCHOR, "- "),
    ("1. " + ANCHOR, "1. "),
    # HTML comments are removed before blocks are read.
    ("+<!-- 주석 -->  ", "+ "),
    ("1.<!-- 주석 -->\t\t", "1. "),
]
BLOCK_PAIRS = [
    (ANCHOR + TABLE, TABLE),
    (ANCHOR + "> 인용", "> 인용"),
    ("<!-- 주석 -->" + ANCHOR + TABLE, TABLE),
    # Only the leading anchors go; the rest of the line stays as written.
    (ANCHOR + "| " + KEEP + " |", "| " + KEEP + " |"),
    (ANCHOR + "> " + KEEP, "> " + KEEP),
    (ANCHOR + " <a id='y'></a>  | a " + ANCHOR + " |", "| a " + ANCHOR + " |"),
    ('<A ID="z"></A>\t> 인용 \\| 끝', "> 인용 \\| 끝"),
    # Anchors right after `>` stand where the space would be.
    ('><a id="x"></a> 인용', '> <a id="x"></a> 인용'),
    # Later table rows: before and after the separator, in the middle, in a row.
    (
        "| A | B |\n" + ANCHOR + "| first |\n| second | KEEP |",
        "| A | B |\n| first |\n| second | KEEP |",
    ),
    ("| A | B |\n" + ANCHOR + "|---|---|\n| 1 | 2 |", "| A | B |\n|---|---|\n| 1 | 2 |"),
    (
        "| A | B |\n|---|---|\n" + ANCHOR + "| 1 | 2 |\n| 3 | 4 |",
        "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |",
    ),
    ("| A |\n|---|\n| 1 |\n" + ANCHOR + "| 2 |\n| 3 |", "| A |\n|---|\n| 1 |\n| 2 |\n| 3 |"),
    (
        "| A |\n" + ANCHOR + "|---|\n" + ANCHOR + "| 1 |\n" + ANCHOR + "| 2 |",
        "| A |\n|---|\n| 1 |\n| 2 |",
    ),
    # Later quote lines, callout titles and callout bodies.
    ("> 첫째\n" + ANCHOR + "> 둘째\n> 셋째", "> 첫째\n> 둘째\n> 셋째"),
    (ANCHOR + "> **요약**: 본문\n> 둘째", "> **요약**: 본문\n> 둘째"),
    ("> **요약**: 본문\n" + ANCHOR + "> 둘째", "> **요약**: 본문\n> 둘째"),
    (ANCHOR + "> [!NOTE]\n" + ANCHOR + "> 참고 본문", "> [!NOTE]\n> 참고 본문"),
]
CONTEXTS = [
    "{}",
    "본문\n\n{}",
    "{}\n\n본문",
    "본문\n{}\n본문",
    "{}\n본문",
    "- 가\n{}\n- 나",
    "- 가\n    {}\n- 나",
    "1. 가\n{}",
    "> 인용 앞\n\n{}",
    "| 앞 |\n\n{}",
    "{}\n| 뒤 |",
    "# 제목\n{}\n## 소제목",
    # Right after a table or quote: the line continues it, as without anchors.
    "| 앞 |\n{}",
    "> 인용 앞\n{}",
]

# Metadata each profile needs before it renders.
PROFILE_FRONTMATTER = {
    "office-letter": "recipients:\n  - 가상 수신처\nsender:\n  organization: 주식회사 예시\n",
    "term-sheet": "prepared_by: 예시증권\ndisclaimer: 시험용 가상 자료입니다.\n",
}


def _front(profile: str) -> str:
    return "---\ntitle: 가상 문서\n" + PROFILE_FRONTMATTER.get(profile, "") + "---\n\n"


def _png() -> str:
    """A valid one-pixel PNG as Base64, written here so a Base64 image parses."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00"))
        + chunk(b"IEND", b"")
    )
    return base64.b64encode(png).decode()


IMAGE = "![가상](data:image/png;base64," + _png() + ")"
# A Base64 image inside a table row or quote line is that row's or line's text.
BLOCK_PAIRS += [
    ("| A |\n" + ANCHOR + "| " + IMAGE + " |", "| A |\n| " + IMAGE + " |"),
    (ANCHOR + "> " + IMAGE, "> " + IMAGE),
]

# One JSON line per case: the parsed model without source text, and optionally
# the rendered body. Which fix paths ran is recorded for principle A.
_CHILD = r"""
import json
import re
import sys

from md_parser import MarkdownParser

request = json.load(sys.stdin)
taken = []
# Older code has neither method; the other checks still run against it.
anchor_led = getattr(MarkdownParser, "_anchor_led_line", None)
list_match = getattr(MarkdownParser, "_list_match", None)
recorded = anchor_led is not None and list_match is not None


def record_anchor_led(self, line, pattern):
    value = anchor_led(self, line, pattern)
    if value is not None:
        taken.append("anchor-led line")
    return value


def record_list_match(cls, pattern, raw_line):
    value = list_match.__func__(cls, pattern, raw_line)
    if value is None and pattern.match(raw_line):
        taken.append("blank list item")
    return value


if recorded:
    MarkdownParser._anchor_led_line = record_anchor_led
    MarkdownParser._list_match = classmethod(record_list_match)

if request["render"]:
    from document_profiles import RenderOptions
    from ib_renderer import IBDocumentRenderer

for profile, text in request["cases"]:
    taken.clear()
    try:
        model = MarkdownParser(profile=profile).parse(text)
        value = []
        for element in model.elements:
            item = element.content[1] if isinstance(element.content, tuple) else element.content
            value.append([element.element_type.name, repr(element.content), getattr(item, "text", None)])
        value.append([model.warnings, repr(model.footnotes), repr(model.metadata)])
    except Exception as exc:
        value = ["EXC", type(exc).__name__, str(exc)]
        model = None
    if request["render"] and model is not None:
        try:
            options = RenderOptions(
                profile=profile, include_cover=False, include_toc=False, include_disclaimer=False
            )
            document = IBDocumentRenderer(options=options).render(model)
            # Semantic bookmark names end in four random characters, and an
            # image keeps the name of the temporary file it was read from.
            body = re.sub(r'(w:name="_ibrep_[^"]*_)[0-9a-f]{4}"', r'\1"', document.element.body.xml)
            value.append(re.sub(r'name="tmp[^"]*"', 'name="tmp"', body))
        except Exception as exc:
            value.append("RENDER EXC " + type(exc).__name__ + ": " + str(exc))
    print(
        json.dumps({"value": value, "taken": sorted(set(taken)), "recorded": recorded}),
        flush=True,
    )
"""

_GUARD_CHILD = r"""
from md_parser import MarkdownParser

# A branch that takes a line and consumes none must stop the parse, not loop.
MarkdownParser._collect_paragraph = lambda self, lines, start, first_line=None: ("", start)
try:
    MarkdownParser(profile="plain").parse("가상 본문")
except RuntimeError as exc:
    print(exc)
"""


def _pairs() -> list[tuple[str, str]]:
    return [
        (context.format(blank), context.format(nearest))
        for context in CONTEXTS
        for blank, nearest in LIST_PAIRS + BLOCK_PAIRS
    ]


def _in_child(cases: list[tuple[str, str]], render: bool = False) -> list[dict[str, object]]:
    """Parse (and render) each (profile, text) in a child process under a time limit."""
    try:
        result = subprocess.run(
            [sys.executable, "-c", _CHILD],
            input=json.dumps({"cases": cases, "render": render}),
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=ROOT,
            timeout=TIME_LIMIT,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        out = exc.stdout.decode("utf-8") if isinstance(exc.stdout, bytes) else exc.stdout or ""
        done = len(out.splitlines())
        pytest.fail(f"Parsing did not finish on {cases[done]!r} ({done} cases finished first)")
    assert result.returncode == 0, result.stderr[-500:]
    return [json.loads(line) for line in result.stdout.splitlines()]


def _values(cases: list[tuple[str, str]], render: bool = False) -> list[object]:
    return [entry["value"] for entry in _in_child(cases, render)]


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_blank_list_items_and_anchor_led_blocks_read_as_their_nearest_input(
    profile: str,
) -> None:
    pairs = _pairs()
    cases = [(profile, _front(profile) + text) for pair in pairs for text in pair]
    values = _values(cases)
    assert len(values) == len(cases)
    for index, (blank, nearest) in enumerate(pairs):
        assert values[2 * index] == values[2 * index + 1], (blank, nearest)


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_anchor_led_blocks_render_as_their_nearest_input(profile: str) -> None:
    cases = [(profile, _front(profile) + text) for pair in BLOCK_PAIRS for text in pair]
    values = _values(cases, render=True)
    for index, (blank, nearest) in enumerate(BLOCK_PAIRS):
        assert values[2 * index] == values[2 * index + 1], (blank, nearest)
        assert "RENDER EXC" not in str(values[2 * index + 1][-1])[:20], nearest


def test_code_after_a_leading_anchor_keeps_its_anchor_text() -> None:
    values = _values([("plain", ANCHOR + "| " + KEEP + " |"), ("plain", ANCHOR + "> " + KEEP)])
    assert '<a id=\\"keep\\"></a>' not in json.dumps(values[0][0][2])
    assert "TextRun(text='<a id=\"keep\"></a>'" in values[0][0][1]
    assert values[1][0][2] == KEEP


@pytest.mark.parametrize("lead", ["| ", "> ", ANCHOR + "| ", ANCHOR + "> ", ">" + ANCHOR + " "])
def test_a_base64_image_never_takes_a_table_or_quote_line(lead: str) -> None:
    """Why a continued anchor-led line cannot have finished on c97841c.

    `MarkdownParser._anchor_led_line` relies on this: of the branches tried
    before tables and quotes, only a Base64 image could match such a line,
    and it needs `![` first, alone or wrapped across lines.
    """
    lines = [lead + IMAGE + " |", IMAGE[:40], IMAGE[40:] + " |"]
    assert Base64ImageParser.is_base64_image(lines[0])
    assert Base64ImageParser.parse_multiline(lines, 0)[0] is None
    assert Base64ImageParser.parse_multiline([lead + IMAGE[:40], IMAGE[40:]], 0)[0] is None
    assert Base64ImageParser.parse_multiline([IMAGE], 0)[0] is not None


def test_blank_list_items_are_text_not_items() -> None:
    values = _values(
        [("plain", "1.   "), ("plain", "-  "), ("plain", "- 가\n-   "), ("plain", "-   \n- 가")]
    )
    texts = [[(entry[0], entry[2]) for entry in value[:-1]] for value in values]
    assert texts == [
        [("PARAGRAPH", "1.")],
        [("PARAGRAPH", "-")],
        [("BULLET_LIST", "가 -")],
        [("PARAGRAPH", "-"), ("BULLET_LIST", "가")],
    ]


def test_a_pass_that_consumes_no_line_stops_the_parse() -> None:
    result = subprocess.run(
        [sys.executable, "-c", _GUARD_CHILD],
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=ROOT,
        timeout=TIME_LIMIT,
        check=False,
    )
    assert result.returncode == 0, result.stderr[-500:]
    assert result.stdout.strip() == "Markdown parser consumed no line at line 1"


# ═══════════════════════════════════════════════════════════════════════════════
# DIFFERENTIAL TESTS (fixed seeds)
# ═══════════════════════════════════════════════════════════════════════════════

_TEXT = ["가상", "text", "값 1,234", "(12.5)%", "**굵게**", "_기울임_", "\\|", "<br>", "@tbl-a"]
_EXTRA = [KEEP, '<a id="in"></a>', "`코드`", " ", "  ( 값 )"]
_PREFIXES = [ANCHOR, "<a id='y'></a> ", ' <A ID="z"></A>\t', ANCHOR + "  " + ANCHOR]
_MARKERS = ["-", "*", "+", "1.", "2."]


def _inline(rng: random.Random) -> str:
    """Text that always holds a visible word, with optional code, anchors and spaces."""
    parts = [rng.choice(_TEXT)] + [rng.choice(_TEXT + _EXTRA) for _ in range(rng.randint(0, 2))]
    rng.shuffle(parts)
    return " ".join(parts)


def _same(text: str) -> tuple[str, str]:
    return text, text


def _maybe(rng: random.Random, chance: float, text: str) -> tuple[str, str]:
    """The line as written, and with leading anchors at the given chance."""
    return text, (rng.choice(_PREFIXES) if rng.random() < chance else "") + text


def _document(rng: random.Random, finished_anchors: bool) -> tuple[str, str, str]:
    """A random document, the same with hanging lines written in, and a profile.

    The second document puts leading anchors before table and quote lines,
    anchors right after `>`, and blank list items (two or more spaces or
    anchors) where the first writes `- `. With `finished_anchors`, both
    documents also hold anchors where they never hung: before paragraphs,
    headings, list lines and Base64 image lines.
    """
    plain: list[str] = []
    hanging: list[str] = []
    lead = (lambda: rng.choice(_PREFIXES)) if finished_anchors else (lambda: "")
    for _ in range(rng.randint(1, 6)):
        kind = rng.choice(["paragraph", "heading", "table", "quote", "list", "code", "image"])
        lines: list[tuple[str, str]] = []
        if kind == "paragraph":
            lines.append(_same(lead() + _inline(rng)))
        elif kind == "heading":
            lines.append(_same(lead() + "## " + _inline(rng)))
        elif kind == "image":
            lines.append(_same(lead() + IMAGE))
        elif kind == "code":
            lines += [_same("```"), _same(_inline(rng)), _same("```")]
        elif kind == "table":
            columns = rng.randint(1, 3)
            rows = [
                "| " + " | ".join(_inline(rng) for _ in range(rng.randint(1, columns))) + " |"
                for _ in range(rng.randint(1, 4))
            ]
            rows.insert(1, "|" + "---|" * columns)
            lines += [_maybe(rng, 0.4, row) for row in rows]
            if rng.random() < 0.3:
                lines.append(_maybe(rng, 0.4, "| " + IMAGE + " |"))
        elif kind == "quote":
            body = ["> " + _inline(rng) for _ in range(rng.randint(1, 3))]
            if rng.random() < 0.3:
                body.insert(0, rng.choice(["> **요약**: ", "> [!NOTE]"]) + _inline(rng))
            for text in body:
                if rng.random() < 0.15:
                    lines.append(("> " + ANCHOR + " " + text[2:], ">" + ANCHOR + " " + text[2:]))
                else:
                    lines.append(_maybe(rng, 0.4, text))
            if rng.random() < 0.3:
                lines.append(_maybe(rng, 0.4, "> " + IMAGE))
        else:
            marker = rng.choice(_MARKERS)
            for _ in range(rng.randint(1, 3)):
                indent = rng.choice(["", "", "    "])
                if rng.random() < 0.3:
                    blank = rng.choice(
                        ["  ", "\t\t", " \u3000", "   ", " " + ANCHOR, "  " + ANCHOR]
                    )
                    lines.append((indent + marker + " ", indent + marker + blank))
                else:
                    prefix = lead() if rng.random() < 0.3 else ""
                    lines.append(_same(prefix + indent + marker + " " + _inline(rng)))
        gap = rng.choice(["\n", "\n\n"])
        plain.append("\n".join(first for first, _ in lines) + gap)
        hanging.append("\n".join(second for _, second in lines) + gap)
    profile = rng.choice(sorted(PROFILES))
    return _front(profile) + "".join(plain), _front(profile) + "".join(hanging), profile


def test_principle_a_documents_that_never_hung_take_no_fix_path() -> None:
    """Without a line that never finished, no new path runs: the code is c97841c's.

    The documents hold anchors before paragraphs, headings, list lines and
    Base64 image lines, anchors and code inside cells and quotes, blank-looking
    `- ` items and Base64 images in table rows and quotes. Measured against a
    c97841c copy as well (record): no parse differed.
    """
    rng = random.Random(20261114)
    cases = []
    for _ in range(600):
        plain, _hanging, profile = _document(rng, finished_anchors=True)
        cases.append((profile, plain))
    entries = _in_child(cases)
    assert len(entries) == len(cases)
    for (_, text), entry in zip(cases, entries, strict=True):
        assert entry["recorded"], "the fix paths are no longer recorded"
        assert entry["taken"] == [], text
        assert entry["value"][0] != "EXC", (text, entry["value"])


def test_principle_b_hanging_lines_read_as_the_document_without_them() -> None:
    rng = random.Random(1114)
    cases = []
    for _ in range(200):
        plain, hanging, profile = _document(rng, finished_anchors=False)
        cases += [(profile, plain), (profile, hanging)]
    entries = _in_child(cases, render=True)
    assert len(entries) == len(cases)
    differing = 0
    rendered = 0
    for index in range(0, len(cases), 2):
        plain_entry, hanging_entry = entries[index], entries[index + 1]
        assert plain_entry["taken"] == [], cases[index][1]
        assert hanging_entry["value"] == plain_entry["value"], (cases[index], cases[index + 1])
        rendered += "RENDER EXC" not in str(plain_entry["value"][-1])[:20]
        differing += cases[index] != cases[index + 1]
    # The generator does write hanging lines into most documents.
    assert differing > 100
    # Every profile renders the documents (required metadata is given).
    assert rendered == len(cases) // 2
