"""Term variables: `terms:` frontmatter values referenced as `{{key}}`.

Design: docs/term-sheet-design-20260929.md §6, with the decisions in
docs/term-sheet-plan-20260929.md §2. The parser substitutes validated values as
`TextRun` objects carrying `term_key`. Values are inserted only after inline
parsing, so they are never read as Markdown. Substituted values are tagged as
Word content controls and snapshotted in custom document properties so
`docx-audit` can report inconsistent edits made later in Word.

Changelog (term variables):
    - NEW: `validate_terms` schema for the `terms:` frontmatter mapping.
    - NEW: `TermResolver` reference diagnostics and private-use token helpers.
    - NEW: Render-scoped term-run collector and plain-text content controls.
"""

import re
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import replace
from typing import Any, Callable, Dict, Iterator, List, Mapping, Optional, Pattern, Set, Tuple

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from document_model import TextRun

# Word content-control tag and custom-property name prefixes shared by the
# renderer (writer) and docx_audit (reader).
TERM_TAG_PREFIX = "ibrep:term:"
TERM_PROPERTY_PREFIX = "ibrep.term."

TERM_KEY_RE = re.compile(r"[a-z][a-z0-9_]{0,39}")
# A reference candidate; the parser decides which text is literal context.
TERM_REFERENCE_RE = re.compile(r"\{\{([^{}\n]*)\}\}")

# Private-use token -> (key, value) for one substituted text field.
TokenMap = Dict[str, Tuple[str, str]]

_LINE_BREAK_RE = re.compile("[\n\r\x85\u2028\u2029]")
_CONTROL_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
_INVALID_HINT = "keys use lowercase letters, digits and underscores and start with a letter"


# ═══════════════════════════════════════════════════════════════════════════════
# SCHEMA
# ═══════════════════════════════════════════════════════════════════════════════


def validate_terms(value: Any) -> Dict[str, str]:
    """Validate the `terms:` frontmatter mapping.

    Args:
        value: The parsed YAML value of the `terms` key.

    Returns:
        A new key-to-text mapping in source order.

    Raises:
        ValueError: The value is not a mapping, a key is malformed, or a value is
            not a single-line string.
    """
    if not isinstance(value, dict):
        raise ValueError('terms must be a YAML mapping of key: "text" pairs (write terms: {} for none)')
    terms: Dict[str, str] = {}
    for key, text in value.items():
        if not isinstance(key, str) or not TERM_KEY_RE.fullmatch(key):
            raise ValueError(
                f"Invalid term key {key!r}: start with a lowercase letter and use at most 40 "
                "lowercase letters, digits or underscores"
            )
        terms[key] = _validate_value(key, text)
    return terms


def _validate_value(key: str, value: Any) -> str:
    """Accept exactly the text the author wrote; YAML scalars lose their written form.

    Args:
        key: Validated term key, for messages.
        value: Parsed YAML value.

    Returns:
        The value unchanged.

    Raises:
        ValueError: The value is missing, not a string, multi-line or not XML text.
    """
    if value is None:
        raise ValueError(f"terms.{key} has no value; write it as a quoted string")
    if isinstance(value, (dict, list, tuple, set)):
        kind = "mapping" if isinstance(value, dict) else "list"
        raise ValueError(f"terms.{key} must be a single string, not a YAML {kind}")
    if not isinstance(value, str):
        raise ValueError(
            f"terms.{key} must be text, but YAML read it as {type(value).__name__} {value!r}; "
            f'quote it to keep the written form, e.g. {key}: "1.10" (unquoted, 1.10 becomes 1.1)'
        )
    if _LINE_BREAK_RE.search(value):
        raise ValueError(
            f"terms.{key} must be a single line: a line break cannot be tracked as one Word control"
        )
    if _CONTROL_RE.search(value):
        raise ValueError(f"terms.{key} contains a control character that Word cannot store")
    return value


# ═══════════════════════════════════════════════════════════════════════════════
# REFERENCE RESOLUTION
# ═══════════════════════════════════════════════════════════════════════════════


class TermResolver:
    """Resolve `{{key}}` references for one parse and collect diagnostics.

    The parser decides which text is literal (code, math, escapes and link
    destinations) and asks the resolver for a private-use token for each other
    reference, following the `\\ue000CODE` collision approach of `TextParser`.
    Tokens pass through inline parsing as ordinary characters and become value
    runs afterwards, so a value is never parsed.
    """

    _TOKEN_PREFIX = "\ue000TERM"
    _TOKEN_END = "\ue001"

    def __init__(self, values: Mapping[str, str], corpus: str = "") -> None:
        """Create a resolver for validated values.

        Args:
            values: Validated key-to-text mapping (see `validate_terms`).
            corpus: Source text the token prefix must never occur in.
        """
        self.values: Dict[str, str] = dict(values)
        self.used: Set[str] = set()
        self.undefined: Set[str] = set()
        self.invalid: Set[str] = set()
        self._prefix = self._TOKEN_PREFIX
        self._count = 0
        self.reserve(corpus)

    def reserve(self, text: str) -> None:
        """Extend the token prefix until it does not occur in the given text."""
        while self._prefix in text:
            self._prefix += "X"

    def token(self, reference: str, key: str, tokens: TokenMap) -> Optional[str]:
        """Return a token for a defined key; record undefined or malformed references.

        Args:
            reference: The reference as written, for diagnostics.
            key: Reference content with escapes resolved and spaces stripped.
            tokens: Token map of the field being substituted; receives the token.

        Returns:
            A private-use token, or None when the reference stays literal.
        """
        if not TERM_KEY_RE.fullmatch(key):
            self.invalid.add(reference)
            return None
        if key not in self.values:
            self.undefined.add(key)
            return None
        self.used.add(key)
        token = f"{self._prefix}{self._count}{self._TOKEN_END}"
        self._count += 1
        tokens[token] = (key, self.values[key])
        return token

    def warnings(self) -> List[str]:
        """Return one sorted model warning per undefined key or malformed reference."""
        messages = ["Undefined term: {{" + key + "}}" for key in self.undefined]
        messages.extend(
            f"Invalid term reference: {reference} ({_INVALID_HINT})" for reference in self.invalid
        )
        return sorted(messages)

    def unused(self) -> List[str]:
        """Return defined keys that no substituted reference used."""
        return sorted(set(self.values) - self.used)


def split_term_runs(runs: List[TextRun], tokens: TokenMap) -> List[TextRun]:
    """Split token-bearing runs into before, value and after runs.

    Args:
        runs: Runs parsed from tokenized text.
        tokens: Tokens of that text.

    Returns:
        Runs where each token became a value run with `term_key`, inheriting the
        formatting (emphasis, colour, link) of the run that contained it.
    """
    if not tokens:
        return runs
    pattern = _token_pattern(tokens)
    result: List[TextRun] = []
    for run in runs:
        parts = pattern.split(run.text)
        if len(parts) == 1:
            result.append(run)
            continue
        for index, part in enumerate(parts):
            if index % 2:
                key, value = tokens[part]
                result.append(replace(run, text=value, term_key=key))
            elif part:
                result.append(replace(run, text=part))
    return result


def restore_terms(
    text: str, tokens: TokenMap, transform: Optional[Callable[[str], str]] = None,
) -> str:
    """Insert values into a plain string field in one pass (values are not rescanned).

    Args:
        text: Tokenized field text.
        tokens: Tokens of that text.
        transform: Optional conversion of each value, such as Markdown escaping.

    Returns:
        The field text with values in place of tokens.
    """
    if not tokens:
        return text

    def value(match: "re.Match[str]") -> str:
        text_value = tokens[match.group(0)][1]
        return transform(text_value) if transform else text_value

    return _token_pattern(tokens).sub(value, text)


def _token_pattern(tokens: TokenMap) -> Pattern[str]:
    """Match any token of one field, capturing it for `re.split`."""
    return re.compile("(" + "|".join(re.escape(token) for token in tokens) + ")")


# ═══════════════════════════════════════════════════════════════════════════════
# CONTENT CONTROLS
# ═══════════════════════════════════════════════════════════════════════════════


class TermRunCollector:
    """Value runs rendered in one request, tagged only after run post-processing.

    Renderer post-processing (negative colours, risk colours, base-case bold,
    native footnotes) reads `Paragraph.runs`, which does not see runs inside a
    content control, so runs are registered while rendering and wrapped last.
    """

    def __init__(self) -> None:
        """Start with no registered runs."""
        self._runs: Dict[Any, Tuple[str, str]] = {}

    def register(self, run: Any, key: str, value: str) -> None:
        """Remember a rendered `w:r` element and the term value it shows."""
        self._runs[run] = (key, value)

    def get(self, run: Any) -> Optional[Tuple[str, str]]:
        """Return the key and value registered for a `w:r` element, if any."""
        return self._runs.get(run)

    def __len__(self) -> int:
        """Number of registered runs."""
        return len(self._runs)


_TERM_RUNS: ContextVar[Optional[TermRunCollector]] = ContextVar("term_runs", default=None)
# Children of a rendered value run that a plain-text content control may hold.
_CONTROL_RUN_CHILDREN = frozenset(qn(tag) for tag in ("w:rPr", "w:t", "w:tab", "w:br", "w:cr"))


@contextmanager
def collect_term_runs(enabled: bool) -> Iterator[Optional[TermRunCollector]]:
    """Collect value runs for one render; nothing is collected when tags are off.

    Args:
        enabled: Resolved `term_tags` option of the render.

    Yields:
        The render's collector, or None when values stay plain text.
    """
    collector = TermRunCollector() if enabled else None
    token = _TERM_RUNS.set(collector)
    try:
        yield collector
    finally:
        _TERM_RUNS.reset(token)


def register_term_run(run: Any, key: str, value: str) -> None:
    """Register a rendered value run; a no-op outside a tagging render.

    Args:
        run: The `w:r` element that shows the value.
        key: Term key.
        value: Rendered value text.
    """
    collector = _TERM_RUNS.get()
    if collector is not None:
        collector.register(run, key, value)


def wrap_term_controls(document: Any, collector: TermRunCollector) -> Dict[str, str]:
    """Wrap registered runs still in the body in unlocked plain-text content controls.

    Each control holds one run, in `w:sdtPr` order alias, tag, id and text, with
    no lock or data binding, so Word edits it like ordinary text.

    Args:
        document: The rendered python-docx document.
        collector: Runs registered while rendering this document.

    Returns:
        Tagged key -> value in key order, for the generation snapshot.
    """
    if not len(collector):
        return {}
    root = document.element
    used_ids = {
        int(value) for value in root.xpath("//@w:id | //w:sdtPr/w:id/@w:val")
        if value.lstrip("-").isdigit()
    }
    control_id = 0
    tagged: Dict[str, str] = {}
    # Materialize first: wrapping moves runs later in document order.
    for run in list(root.body.iter(qn("w:r"))):
        entry = collector.get(run)
        if entry is None:
            continue
        parent = run.getparent()
        if parent is None or parent.tag != qn("w:p"):
            continue
        if any(child.tag not in _CONTROL_RUN_CHILDREN for child in run):
            continue  # e.g. converted into a native footnote reference
        control_id += 1
        while control_id in used_ids:
            control_id += 1
        key, value = entry
        control = _plain_text_control(key, control_id)
        run.addprevious(control)
        control[-1].append(run)
        tagged.setdefault(key, value)
    return dict(sorted(tagged.items()))


def _plain_text_control(key: str, control_id: int) -> Any:
    """Build an empty run-level `w:sdt` tagged for one term key."""
    control = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    for tag, value in (
        ("w:alias", key), ("w:tag", TERM_TAG_PREFIX + key), ("w:id", str(control_id)),
    ):
        element = OxmlElement(tag)
        element.set(qn("w:val"), value)
        properties.append(element)
    properties.append(OxmlElement("w:text"))
    control.append(properties)
    control.append(OxmlElement("w:sdtContent"))
    return control
