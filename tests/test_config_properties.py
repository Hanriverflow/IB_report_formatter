"""Strict outcome for YAML the document only points at, and for YAML aliases.

`tests/test_markdown_properties.py` generates the document's own frontmatter. This
module generates what its docstring lists as not generated: the content of a house
or theme file the document names, and frontmatter built with anchors and aliases
(shared, recursive and expanding structures). The property is the same one:
strict conversion either saves a readable document or ends in a documented
rejection, leaving no file. It also has to finish, which a recursive or
expanding structure used to prevent.

Bounded coverage, not a claim about all input: values come from the small
alphabet of `YAML_VALUE`, keys from the settings each file accepts, and alias
shapes from the fixed templates below.
"""

import json
import subprocess
import sys
import threading
from collections.abc import Callable
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest
import yaml
from hypothesis import given
from hypothesis import strategies as st

import input_safety
from input_safety import InputLimits, StrictValidationError, load_yaml
from md_parser import MarkdownParser

# Importing this module also registers the Hypothesis profile (fixed examples, no deadline).
from tests.test_markdown_properties import YAML_VALUE, _assert_strict_outcome

THEME_KEYS = (
    "body_font",
    "heading_font",
    "korean_font",
    "body_size",
    "primary_color",
    "margin_mm",
    "BODY_FONT",
    "NAVY",
    "wrong",
)
HOUSE_KEYS = ("prepared_by", "disclaimer", "confidential_label", "confirmation", "style", "wrong")
HOUSE_STYLE_KEYS = (
    "cover",
    "disclaimer",
    "page",
    "table_header",
    "table_sides",
    "header_rule",
    "footer_rule",
    "page_number_align",
    "page_number",
    "logo",
    "logo_width_mm",
    "label_color",
    "wrong",
)
TERM_SHEET = (
    "profile: term-sheet\ntitle: 예시 조건\nprepared_by: 가상 기관\ndisclaimer: 예시 문구\n"
)
BODY = "---\n| 예시 | 내용 |\n| --- | --- |\n| 가상 | 100 |"


def _config(keys: tuple[str, ...], nested: tuple[str, ...] = ()) -> st.SearchStrategy[object]:
    mapping = st.dictionaries(st.sampled_from(keys), YAML_VALUE, max_size=3)
    if nested:
        style = st.dictionaries(st.sampled_from(nested), YAML_VALUE, max_size=3)
        mapping = st.one_of(mapping, style.map(lambda value: {"style": value}))
    return st.one_of(mapping, YAML_VALUE)


def _finishes(action: Callable[[], None], seconds: float = 30.0) -> None:
    """Run in a thread so a structure that never finishes fails the test instead of hanging it.

    The thread cannot be stopped, so this only reports a stall. The module
    fixture below is what keeps a missing guard from exhausting memory here.
    """
    outcome: list[BaseException | None] = []

    def run() -> None:
        try:
            action()
        except BaseException as error:  # noqa: BLE001 - re-raised in the test thread
            outcome.append(error)
        else:
            outcome.append(None)

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    thread.join(seconds)
    assert not thread.is_alive(), "conversion did not finish"
    if outcome[0] is not None:
        raise outcome[0]


@pytest.mark.parametrize("route", ("text", "cli"))
@given(content=_config(THEME_KEYS))
def test_theme_file_content_strict_dichotomy(route: str, content: object) -> None:
    with TemporaryDirectory(prefix="ib-prop-theme-") as config_dir:
        theme = Path(config_dir) / "example-theme.yaml"
        theme.write_text(yaml.safe_dump(content, allow_unicode=True), encoding="utf-8")
        source = f"---\nprofile: plain\ntitle: 예시 문서\ntheme: '{theme.as_posix()}'\n{BODY}"
        with TemporaryDirectory(prefix="ib-prop-theme-out-") as temporary:
            _assert_strict_outcome(source, Path(temporary), "plain", route, InputLimits())


@pytest.mark.parametrize("route", ("text", "cli"))
@given(content=_config(HOUSE_KEYS, HOUSE_STYLE_KEYS))
def test_house_file_content_strict_dichotomy(route: str, content: object) -> None:
    with TemporaryDirectory(prefix="ib-prop-house-") as config_dir:
        house = Path(config_dir) / "example-house.yaml"
        house.write_text(yaml.safe_dump(content, allow_unicode=True), encoding="utf-8")
        source = f"---\n{TERM_SHEET}house: '{house.as_posix()}'\n{BODY}"
        with TemporaryDirectory(prefix="ib-prop-house-out-") as temporary:
            _assert_strict_outcome(source, Path(temporary), "term-sheet", route, InputLimits())


def _pyramid(depth: int, width: int) -> str:
    lines = ["a0: &a0 [가, 나, 다]"]
    for level in range(1, depth):
        lines.append(f"a{level}: &a{level} [" + ", ".join([f"*a{level - 1}"] * width) + "]")
    return "\n".join(lines)


def _repeated_text(field: str, length: int, count: int) -> str:
    return f"base: &t {'가' * length}\n{field}: [" + ", ".join(["*t"] * count) + "]"


def _merged_mapping(keys: int, copies: int) -> str:
    base = "base: &b {" + ", ".join(f"k{index}: 1" for index in range(keys)) + "}"
    return base + "\nrows:\n" + "\n".join(["- {<<: *b}"] * copies)


_GUARD_PROBE = """
import json
import sys

from input_safety import load_yaml

for text in json.load(sys.stdin):
    try:
        load_yaml(text)
    except ValueError:
        continue
    sys.exit(3)
"""


@pytest.fixture(scope="module", autouse=True)
def _alias_guard_is_in_place() -> None:
    """Fail the whole module at once if the guard is gone.

    Without it the cases below load structures that take gigabytes to walk, and
    a thread cannot be stopped. A child process can, so it goes first.
    """
    texts = ["loop: &x [*x]", _pyramid(9, 9), _merged_mapping(1_000, 1_000)]
    try:
        probe = subprocess.run(
            [sys.executable, "-c", _GUARD_PROBE],
            input=json.dumps(texts),
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=Path(__file__).resolve().parents[1],
            timeout=60,
            check=False,
        )
    except subprocess.TimeoutExpired:
        pytest.fail("load_yaml did not finish on alias structures; the module is not run")
    if probe.returncode != 0:
        pytest.fail(f"load_yaml accepted an alias structure it must refuse: {probe.stderr[-300:]}")


ALIAS_FRONTMATTER = (
    "base: &b 가상 값\ncopy: *b\nlist: &l [가, 나]\nagain: *l",
    "subtitle: &s 가상 부제\ncompany: *s",
    "defaults: &d {cover: false}\nlayout: *d",
    "defaults: &d {cover: false}\nlayout:\n  <<: *d\n  toc: false",
    "loop: &x [*x]",
    "loop: &x {self: *x}",
    "title: &x [*x]",
    "layout: &x {cover: *x}",
    "tables: &x [*x]",
    "sender: &x {organization: *x}",
    _pyramid(4, 3),
    _pyramid(6, 9),
    _pyramid(9, 9),
    "layout:\n  " + _pyramid(7, 9).replace("\n", "\n  "),
    _repeated_text("company", 100, 20),
    _repeated_text("company", 20_000, 2_000),
    _repeated_text("title", 20_000, 2_000),
    _repeated_text("attendees", 20_000, 2_000),
    _merged_mapping(5, 5),
    _merged_mapping(1_000, 1_000),
)


@pytest.mark.parametrize("profile", ("plain", "ib-report"))
@pytest.mark.parametrize("route", ("text", "stream", "cli"))
@pytest.mark.parametrize(
    "frontmatter",
    ALIAS_FRONTMATTER,
    ids=[f"alias{index}" for index in range(len(ALIAS_FRONTMATTER))],
)
def test_alias_frontmatter_strict_dichotomy(profile: str, route: str, frontmatter: str) -> None:
    source = f"---\nprofile: {profile}\ntitle: 예시 문서\n{frontmatter}\n{BODY}"

    def convert() -> None:
        with TemporaryDirectory(prefix="ib-prop-alias-") as temporary:
            _assert_strict_outcome(source, Path(temporary), profile, route, InputLimits())

    _finishes(convert)


@pytest.mark.parametrize("kind", ("house", "theme", "chart", "diagram"))
@pytest.mark.parametrize(
    "content", ("loop: &x [*x]", _pyramid(9, 9)), ids=("recursive", "expanding")
)
def test_alias_structures_in_named_files_and_fences_finish(kind: str, content: str) -> None:
    def convert() -> None:
        with TemporaryDirectory(prefix="ib-prop-alias-file-") as config_dir:
            if kind in {"house", "theme"}:
                config = Path(config_dir) / "example.yaml"
                config.write_text(content, encoding="utf-8")
                profile = "term-sheet" if kind == "house" else "plain"
                head = TERM_SHEET if kind == "house" else "profile: plain\ntitle: 예시 문서\n"
                source = f"---\n{head}{kind}: '{config.as_posix()}'\n{BODY}"
            else:
                profile = "plain"
                fence = "chart" if kind == "chart" else "diagram:flow"
                source = f"---\nprofile: plain\n---\n\n```{fence}\n{content}\n```\n"
            with TemporaryDirectory(prefix="ib-prop-alias-out-") as temporary:
                _assert_strict_outcome(source, Path(temporary), profile, "text", InputLimits())

    _finishes(convert)


@pytest.mark.parametrize(
    "text, message",
    [
        ("loop: &x [*x]", "YAML structure refers to itself"),
        ("loop: &x {self: *x}", "YAML structure refers to itself"),
        (_pyramid(6, 9), "YAML aliases expand to too much content"),
        (_pyramid(9, 9), "YAML aliases expand to too much content"),
        (_repeated_text("company", 20_000, 2_000), "YAML aliases expand to too much content"),
        (_repeated_text("anything", 100_000, 2_000), "YAML aliases expand to too much content"),
        (_merged_mapping(1_000, 1_000), "YAML aliases expand to too much content"),
        (_merged_mapping(3_000, 3_000), "YAML aliases expand to too much content"),
    ],
    ids=("list", "mapping", "pyramid6", "pyramid9", "text20k", "text100k", "merge1k", "merge3k"),
)
def test_recursive_and_expanding_yaml_is_refused(text: str, message: str) -> None:
    def load() -> None:
        with pytest.raises(ValueError, match=message):
            load_yaml(text)

    _finishes(load, 10.0)

    def parse() -> None:
        with pytest.raises(StrictValidationError, match="Invalid YAML frontmatter: " + message):
            MarkdownParser(profile="plain").parse(f"---\nprofile: plain\n{text}\n---\n가상 본문")

    _finishes(parse, 10.0)


@pytest.mark.parametrize(
    "text",
    [
        "base: &b 가상 값\ncopy: *b\nlist: &l [가, 나]\nagain: *l",
        "defaults: &d {a: 1}\nlayout:\n  <<: *d\n  b: 2",
        _pyramid(4, 3),
        "- &a [1, 2]\n- *a\n- *a\n- {k: *a}",
        _repeated_text("company", 100, 20),
        _repeated_text("disclaimer", 3_000, 30),
        _merged_mapping(20, 40),
        "",
    ],
    ids=("shared", "merge", "pyramid", "nested", "text100", "text3k", "merged", "empty"),
)
def test_ordinary_aliases_load_as_the_safe_loader_loads_them(text: str) -> None:
    assert load_yaml(text) == yaml.safe_load(text)


def test_expanding_yaml_is_refused_before_any_value_is_built(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A merge key copies its mapping while loading, so the check has to come first."""
    built: list[object] = []
    construct = yaml.SafeLoader.construct_document

    def recording(loader: yaml.SafeLoader, node: yaml.Node) -> object:
        built.append(node)
        return construct(loader, node)

    monkeypatch.setattr(yaml.SafeLoader, "construct_document", recording)
    for text in (_merged_mapping(3_000, 3_000), _pyramid(9, 9), "loop: &x [*x]"):
        with pytest.raises(ValueError, match="YAML (aliases expand|structure refers)"):
            load_yaml(text)
    assert built == []
    assert load_yaml(_merged_mapping(20, 40))["rows"][39] == {f"k{index}": 1 for index in range(20)}
    assert len(built) == 1


def test_expansion_is_counted_exactly() -> None:
    """One for each node and character, with an aliased value counted at every use."""
    node = yaml.compose("a: &t 가나다라마바사아자차\nb: [*t, *t]")
    # mapping 1 + key 1 + text 10 + key 1 + (list 1 + text 10 + text 10)
    input_safety._check_yaml_expansion(node, 34)
    with pytest.raises(ValueError, match="YAML aliases expand to too much content"):
        input_safety._check_yaml_expansion(node, 33)


@pytest.mark.parametrize(
    "length, accepted, refused",
    [(1_000, 95, 105), (50_000, 2, 4), (400_000, 2, 4)],
    ids=("fixed-part", "four-times-50k", "four-times-400k"),
)
def test_allowance_is_the_larger_of_the_fixed_part_and_four_times_the_length(
    length: int, accepted: int, refused: int
) -> None:
    """Short YAML may expand to 100,000; longer YAML to four times its own length."""
    assert len(load_yaml(_repeated_text("items", length, accepted))["items"]) == accepted
    with pytest.raises(ValueError, match="YAML aliases expand to too much content"):
        load_yaml(_repeated_text("items", length, refused))


def test_an_aliased_collection_is_opened_once(monkeypatch: pytest.MonkeyPatch) -> None:
    """Opening it at every use made the check itself quadratic (11 s for 148 KB)."""
    opened: list[object] = []
    children = input_safety._yaml_children

    def recording(node: yaml.Node) -> list[yaml.Node]:
        opened.append(node)
        return children(node)

    monkeypatch.setattr(input_safety, "_yaml_children", recording)
    shared = "x: &x {" + ",".join(f"k{index}" for index in range(300)) + "}"
    with pytest.raises(ValueError, match="YAML aliases expand to too much content"):
        load_yaml(shared + "\nr: [" + ",".join(["*x"] * 5_000) + "]")
    # Three collections (the document, the mapping, the list), each opened to
    # walk it and again to size it.
    assert len(opened) <= 6


def test_merge_keys_in_one_mapping_are_limited() -> None:
    """The loader removes merge keys one at a time, which is quadratic in one mapping."""
    assert load_yaml("{" + ",".join(["<<: {a: 1}"] * 100) + "}") == {"a": 1}
    with pytest.raises(ValueError, match="YAML mapping repeats the merge key too many times"):
        load_yaml("{" + ",".join(["<<: {a: 1}"] * 101) + "}")
    with pytest.raises(ValueError, match="YAML mapping repeats the merge key too many times"):
        load_yaml("{" + ",".join(["<<: {}"] * 60_000) + "}")
    # One merge key that lists many mappings, and many mappings with a merge key each, are fine.
    listed = "base: &b {a: 1}\nmerged:\n  <<: [" + ", ".join(["*b"] * 500) + "]"
    assert load_yaml(listed)["merged"] == {"a": 1}
    assert len(load_yaml(_merged_mapping(3, 500))["rows"]) == 500


def test_input_that_is_not_text_is_a_value_error() -> None:
    with pytest.raises(ValueError, match="YAML value cannot be constructed"):
        load_yaml(None)  # type: ignore[arg-type]


def test_yaml_without_aliases_is_never_refused_for_size() -> None:
    """A value and a character each take a character of YAML, so the allowance is out of reach."""
    wide = "values: [" + ", ".join(["1"] * 110_000) + "]"
    assert len(load_yaml(wide)["values"]) == 110_000
    nested = "\n".join(f"k{index}: [[1, 2], [3, 4], {{a: [5]}}]" for index in range(8_000))
    assert len(load_yaml(nested)) == 8_000
    long_text = "text: " + "가" * 500_000
    assert len(load_yaml(long_text)["text"]) == 500_000
    # The densest forms, a value for every one or two characters. Each is larger
    # than the fixed part of the allowance, so what admits it is the part that
    # grows with the length of the YAML.
    keys = ",".join(f"k{index}" for index in range(20_000))
    for dense, count in (
        ("{" + keys + "}", 20_000),
        ("[" + ",".join(["[]"] * 105_000) + "]", 105_000),
        ("[" + ",".join(["{a}"] * 35_000) + "]", 35_000),
        ("- ?\n" * 35_000, 35_000),
        ("!!set {" + keys + "}", 20_000),
        ("!!pairs [" + keys.replace(",", ": ,") + ": ]", 20_000),
    ):
        assert len(load_yaml(dense)) == count


class _NoAliasDumper(yaml.SafeDumper):
    def ignore_aliases(self, data: object) -> bool:
        return True


_PLAIN_YAML = st.recursive(
    st.one_of(
        st.none(),
        st.booleans(),
        st.integers(-(10**6), 10**6),
        st.text(max_size=12),
        st.text(alphabet="가나다 a\n", max_size=400),
    ),
    lambda children: st.one_of(
        st.lists(children, max_size=6),
        st.dictionaries(st.text(max_size=4), children, max_size=6),
    ),
    max_leaves=40,
)


@pytest.mark.parametrize("flow", (True, False))
@given(value=_PLAIN_YAML, copies=st.integers(1, 400))
def test_generated_yaml_without_aliases_loads_whatever_its_shape(
    flow: bool, value: object, copies: int
) -> None:
    text = yaml.dump(
        [value] * copies, Dumper=_NoAliasDumper, allow_unicode=True, default_flow_style=flow
    )
    assert "&id" not in text
    # Compared with the plain loader, which normalises some line-break characters itself.
    assert load_yaml(text) == yaml.safe_load(text)
