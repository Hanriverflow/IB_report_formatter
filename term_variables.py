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
"""

import re
from dataclasses import replace
from typing import Any, Callable, Dict, List, Mapping, Optional, Pattern, Set, Tuple

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

_LINE_BREAK_RE = re.compile("[\n\r\x85  ]")
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

    _TOKEN_PREFIX = "TERM"
    _TOKEN_END = ""

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
