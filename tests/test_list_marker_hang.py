"""Lines the block parser used to take without consuming, so it never finished.

A list marker followed only by whitespace or HTML anchors (`1.   `, `-  `,
`- <a id="x"></a>`) and a table or quote line led by an HTML anchor
(`<a id="x"></a>| a |`) each kept `MarkdownParser._parse_elements` on one line
forever, in every profile. Each such input must now read exactly as its
nearest input that always finished: `1. `, `- `, `| a |`, `> 인용`.

Every parse runs in a child process under a time limit, so a regression fails
the test instead of stopping the run. Record:
docs/verification-list-marker-hang-20261114.md.
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

from document_profiles import PROFILES

ROOT = Path(__file__).resolve().parents[1]
TIME_LIMIT = 120
ANCHOR = '<a id="x"></a>'
TABLE = "| a | b |\n|---|---|\n| 1 | 2 |"

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
]
# Right after a table or quote, a line led by an anchor ends that block and
# starts a new one, while the same line without the anchor continues it; the
# anchor-led pairs are not compared there.
SAME_KIND = {"| 앞 |\n{}", "> 인용 앞\n{}"}
CONTEXTS += sorted(SAME_KIND)

FRONTMATTER = "---\ntitle: 가상 문서\n---\n\n"

_CHILD = r"""
import json
import sys

from md_parser import MarkdownParser

cases = json.load(sys.stdin)
for profile, text in cases:
    try:
        model = MarkdownParser(profile=profile).parse(text)
        value = []
        for element in model.elements:
            item = element.content[1] if isinstance(element.content, tuple) else element.content
            value.append([element.element_type.name, repr(element.content), getattr(item, "text", None)])
        value.append(model.warnings)
    except Exception as exc:
        value = ["EXC", type(exc).__name__, str(exc)]
    print(json.dumps(value, ensure_ascii=True), flush=True)
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
    pairs = []
    for context in CONTEXTS:
        for blank, nearest in LIST_PAIRS:
            pairs.append((context.format(blank), context.format(nearest)))
        if context not in SAME_KIND:
            for blank, nearest in BLOCK_PAIRS:
                pairs.append((context.format(blank), context.format(nearest)))
    return pairs


def _parse_in_child(cases: list[tuple[str, str]]) -> list[list[object]]:
    """Parse each (profile, text) in a child process; fail on the case that never ends."""
    try:
        result = subprocess.run(
            [sys.executable, "-c", _CHILD],
            input=json.dumps(cases),
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


@pytest.mark.parametrize("profile", sorted(PROFILES))
def test_blank_list_items_and_anchor_led_blocks_read_as_their_nearest_input(
    profile: str,
) -> None:
    pairs = _pairs()
    cases = [(profile, FRONTMATTER + text) for pair in pairs for text in pair]
    values = _parse_in_child(cases)
    assert len(values) == len(cases)
    for index, (blank, nearest) in enumerate(pairs):
        assert values[2 * index] == values[2 * index + 1], (blank, nearest)


def test_blank_list_items_are_text_not_items() -> None:
    values = _parse_in_child(
        [("plain", "1.   "), ("plain", "-  "), ("plain", "- 가\n-   "), ("plain", "-   \n- 가")]
    )
    texts = [[(kind, text) for kind, _, text in value[:-1]] for value in values]
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
