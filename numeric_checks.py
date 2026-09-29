"""
Numeric consistency checks for term sheets: declared relations and schedules.

Values are read as displayed: Korean money units (`10억원`, `3,330만원`,
`5억 3,000만원`, `1,076,175,000원`), percentages (`4.78%`, `1.10%p`), basis
points and plain numbers (optionally `개월` or `년`). Brackets that mark
indicative figures (`[4.90]%`) and the words 약, 총, 내외, 수준 and 정도 are
ignored. A comparison passes within a tolerance: explicit, or by default half
of the display step of the stated value, so `12.66억원` agrees with any amount
that rounds to it. The checks only report disagreements; they never change a
value and do not establish financial correctness.

Changelog:
    - NEW: `checks:` relations between term values (`total = a + b * 2`).
    - NEW: Repayment schedule arithmetic for the table spec `schedule`.
"""

import json
import re
from dataclasses import dataclass
from decimal import Decimal
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

# ═══════════════════════════════════════════════════════════════════════════════
# READING DISPLAYED VALUES
# ═══════════════════════════════════════════════════════════════════════════════

_MONEY_UNITS: Dict[str, Decimal] = {
    "조": Decimal(10) ** 12, "억": Decimal(10) ** 8, "천만": Decimal(10) ** 7,
    "백만": Decimal(10) ** 6, "만": Decimal(10) ** 4, "천": Decimal(1000),
}
_SUFFIX_DIMENSIONS = {
    "%p": "percent", "%": "percent", "bps": "bp", "bp": "bp", "개월": "months", "년": "years",
}
_NUMBER = r"\d[\d,]*(?:\.\d+)?"
_MONEY_PART_RE = re.compile(r"(" + _NUMBER + r")\s*(조|억|천만|백만|만|천)?\s*")
_SUFFIX_RE = re.compile(r"^(" + _NUMBER + r")\s*(%p|%|bps|bp|개월|년)?$", re.IGNORECASE)
_NOISE_RE = re.compile(r"[\[\]]|^\s*(?:약|총)\s*|\s*(?:내외|수준|정도)\s*$")
_NEGATIVE_PREFIXES = ("-", "−", "△", "▲")
_DASHES = frozenset("-‒–—―－")


@dataclass(frozen=True)
class Quantity:
    """A value in its base unit (원, %, bp, months, years or plain) and its display step."""

    value: Decimal
    dimension: str
    step: Decimal = Decimal(0)


def _decimal(text: str) -> Decimal:
    return Decimal(text.replace(",", ""))


def _step(number: str) -> Decimal:
    """Smallest step a written number shows: `12.66` -> 0.01, `1,000` -> 1."""
    _, _, fraction = number.partition(".")
    return Decimal(1).scaleb(-len(fraction))


def read_quantity(text: str) -> Quantity:
    """Read a displayed amount, percentage or number.

    Args:
        text: Value as written, for example `[4.90]%`, `3,330만원` or `36개월`.

    Returns:
        The quantity; money is in 원.

    Raises:
        ValueError: The text is not a single readable value.
    """
    cleaned = _NOISE_RE.sub("", text).strip()
    while True:
        stripped = _NOISE_RE.sub("", cleaned).strip()
        if stripped == cleaned:
            break
        cleaned = stripped
    negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        negative, cleaned = True, cleaned[1:-1].strip()
    if cleaned[:1] in _NEGATIVE_PREFIXES:
        negative, cleaned = True, cleaned[1:].strip()
    quantity = _read_money(cleaned) or _read_suffixed(cleaned)
    if quantity is None:
        raise ValueError(f"cannot read {text!r} as a number")
    if negative:
        return Quantity(-quantity.value, quantity.dimension, quantity.step)
    return quantity


def _read_money(text: str) -> Optional[Quantity]:
    """Read `5억 3,000만원`-style money; None when the text is not money."""
    body = text[:-1].rstrip() if text.endswith("원") else text
    parts = list(_MONEY_PART_RE.finditer(body))
    if not parts or "".join(part.group(0) for part in parts) != body:
        return None
    units = [part.group(2) for part in parts]
    if not text.endswith("원") and not any(unit in ("조", "억", "만") for unit in units):
        return None
    if any(unit is None for unit in units[:-1]):
        return None  # only the last part may lack a unit (`5억 3,000만원`, not `5 3억원`)
    value = sum(
        (_decimal(part.group(1)) * _MONEY_UNITS.get(part.group(2) or "", Decimal(1)) for part in parts),
        Decimal(0),
    )
    last = parts[-1]
    step = _step(last.group(1)) * _MONEY_UNITS.get(last.group(2) or "", Decimal(1))
    return Quantity(value, "money", step)


def _read_suffixed(text: str) -> Optional[Quantity]:
    match = _SUFFIX_RE.match(text)
    if match is None:
        return None
    suffix = (match.group(2) or "").lower()
    return Quantity(_decimal(match.group(1)), _SUFFIX_DIMENSIONS.get(suffix, "plain"), _step(match.group(1)))


def read_cell_number(text: str) -> Optional[Tuple[Decimal, Decimal]]:
    """Read a plain table number such as `1,234.5`, `(12)` or `△3`.

    Args:
        text: Visible cell text.

    Returns:
        (value, display step), or None for a blank, a dash placeholder or text
        without digits (such as `대출실행`).

    Raises:
        ValueError: The text has digits but is not one plain number.
    """
    cleaned = text.strip()
    if not cleaned or set(cleaned) <= _DASHES or not any(char.isdigit() for char in cleaned):
        return None
    quantity = read_quantity(cleaned)
    if quantity.dimension != "plain":
        raise ValueError(f"cannot read {text!r} as a plain number")
    return quantity.value, quantity.step


def format_quantity(quantity: Quantity) -> str:
    """Display a quantity for a diagnostic message (money in 원 with separators)."""
    value = quantity.value.normalize()
    if value == value.to_integral():
        value = value.quantize(Decimal(1))
    shown = format(value, ",f")
    suffix = {"money": "원", "percent": "%", "bp": "bp", "months": "개월", "years": "년"}
    return shown + suffix.get(quantity.dimension, "")


def money_unit(unit_text: str) -> Optional[Decimal]:
    """The 원 multiplier of a table unit such as `억원`, `금액 백만원` or `원`."""
    match = re.search(r"(조|억|천만|백만|만|천)?원\s*$", unit_text.strip())
    if match is None:
        return None
    return _MONEY_UNITS.get(match.group(1) or "", Decimal(1))


# ═══════════════════════════════════════════════════════════════════════════════
# RELATIONS BETWEEN TERM VALUES (`checks:`)
# ═══════════════════════════════════════════════════════════════════════════════

# Custom document property holding the checks as JSON, for `docx-audit`.
CHECKS_PROPERTY = "ibrep.checks"

_TOKEN_RE = re.compile(r"\s*(?:(?P<number>\d+(?:\.\d+)?)|(?P<name>[a-z][a-z0-9_]*)|(?P<op>[-+*/()=]))")

Node = Tuple[Any, ...]  # ("number", Decimal), ("name", key), ("negate", node) or (op, left, right)


class CheckError(ValueError):
    """A check that cannot be evaluated (undefined term, unreadable value, units)."""


@dataclass(frozen=True)
class Check:
    """One declared relation `left = right`, with an optional explicit tolerance."""

    source: str
    left: Node
    right: Node
    tolerance: Optional[str] = None


def _tokens(source: str) -> List[Tuple[str, str]]:
    tokens: List[Tuple[str, str]] = []
    position = 0
    source = source.rstrip()
    while position < len(source):
        match = _TOKEN_RE.match(source, position)
        if match is None or match.end() == position:
            raise ValueError(f"checks: unexpected text in {source!r} at {source[position:]!r}")
        kind = match.lastgroup or ""
        tokens.append((kind, match.group(kind)))
        position = match.end()
    return tokens


class _Parser:
    """Recursive-descent parser for `+ - * /`, parentheses, numbers and term keys."""

    def __init__(self, source: str) -> None:
        self.source = source
        self.tokens = _tokens(source)
        self.index = 0

    def peek(self) -> Tuple[str, str]:
        return self.tokens[self.index] if self.index < len(self.tokens) else ("end", "")

    def take(self) -> Tuple[str, str]:
        token = self.peek()
        self.index += 1
        return token

    def fail(self) -> ValueError:
        return ValueError(f"checks: cannot read {self.source!r}")

    def expression(self) -> Node:
        node = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            node = (self.take()[1], node, self.term())
        return node

    def term(self) -> Node:
        node = self.factor()
        while self.peek() in (("op", "*"), ("op", "/")):
            node = (self.take()[1], node, self.factor())
        return node

    def factor(self) -> Node:
        kind, text = self.take()
        if kind == "number":
            return ("number", Decimal(text))
        if kind == "name":
            return ("name", text)
        if (kind, text) == ("op", "-"):
            return ("negate", self.factor())
        if (kind, text) == ("op", "("):
            node = self.expression()
            if self.take() != ("op", ")"):
                raise self.fail()
            return node
        raise self.fail()


def parse_checks(raw: Any) -> List[Check]:
    """Validate the `checks:` frontmatter list.

    Args:
        raw: A list of `"left = right"` strings or `{check, tolerance}` mappings.

    Returns:
        The parsed checks.

    Raises:
        ValueError: The list, an item, an expression or a tolerance is malformed.
    """
    if not isinstance(raw, list):
        raise ValueError("checks must be a list of `left = right` relations")
    checks: List[Check] = []
    for item in raw:
        tolerance: Optional[str] = None
        if isinstance(item, dict):
            if set(item) - {"check", "tolerance"} or "check" not in item:
                raise ValueError("checks: a mapping needs `check` and may have `tolerance`")
            source, raw_tolerance = item["check"], item.get("tolerance")
            if raw_tolerance is not None:
                if isinstance(raw_tolerance, bool) or not isinstance(
                    raw_tolerance, (str, int, float)
                ):
                    raise ValueError("checks: tolerance must be a value such as `0.01억원`")
                tolerance = str(raw_tolerance)
                read_quantity(tolerance)
        else:
            source = item
        if not isinstance(source, str) or source.count("=") != 1:
            raise ValueError("checks: each relation needs exactly one `=`")
        left_text, right_text = source.split("=")
        left, right = _parse_side(left_text, source), _parse_side(right_text, source)
        checks.append(Check(source.strip(), left, right, tolerance))
    return checks


def _parse_side(text: str, source: str) -> Node:
    parser = _Parser(text)
    if not parser.tokens:
        raise ValueError(f"checks: {source!r} has an empty side")
    node = parser.expression()
    if parser.peek()[0] != "end":
        raise parser.fail()
    return node


def _evaluate(node: Node, values: Mapping[str, str]) -> Quantity:
    kind = node[0]
    if kind == "number":
        return Quantity(node[1], "plain", Decimal(0))
    if kind == "name":
        name = node[1]
        if name not in values:
            raise CheckError(f"undefined term {{{{{name}}}}}")
        try:
            return read_quantity(values[name])
        except ValueError:
            raise CheckError(f"cannot read {{{{{name}}}}} = {values[name]!r} as a number") from None
    if kind == "negate":
        inner = _evaluate(node[1], values)
        return Quantity(-inner.value, inner.dimension)
    left, right = _evaluate(node[1], values), _evaluate(node[2], values)
    if kind in "+-":
        if left.dimension != right.dimension:
            raise CheckError(f"cannot add {left.dimension} and {right.dimension}")
        value = left.value + right.value if kind == "+" else left.value - right.value
        return Quantity(value, left.dimension)
    if kind == "*":
        if "plain" not in (left.dimension, right.dimension):
            raise CheckError(f"cannot multiply {left.dimension} by {right.dimension}")
        dimension = right.dimension if left.dimension == "plain" else left.dimension
        return Quantity(left.value * right.value, dimension)
    if right.value == 0:
        raise CheckError("division by zero")
    if right.dimension == "plain":
        return Quantity(left.value / right.value, left.dimension)
    if right.dimension == left.dimension:
        return Quantity(left.value / right.value, "plain")
    raise CheckError(f"cannot divide {left.dimension} by {right.dimension}")


def evaluate_check(check: Check, values: Mapping[str, str]) -> Optional[str]:
    """Evaluate one relation against term values.

    Args:
        check: Parsed relation.
        values: Term key -> value as displayed.

    Returns:
        None when both sides agree within the tolerance, else a diagnostic.
    """
    try:
        left, right = _evaluate(check.left, values), _evaluate(check.right, values)
        if left.dimension != right.dimension:
            raise CheckError(f"the sides are {left.dimension} and {right.dimension}")
        if check.tolerance is not None:
            tolerance = read_quantity(check.tolerance)
            if tolerance.dimension not in (left.dimension, "plain"):
                raise CheckError(f"the tolerance is {tolerance.dimension}, not {left.dimension}")
            allowed = abs(tolerance.value)
        else:
            allowed = left.step / 2
    except CheckError as error:
        return f"Check cannot be evaluated: {check.source} ({error})"
    difference = left.value - right.value
    if abs(difference) <= allowed:
        return None
    return (
        f"Check failed: {check.source} (left {format_quantity(left)}, right {format_quantity(right)}, "
        f"difference {format_quantity(Quantity(difference, left.dimension))})"
    )


def evaluate_checks(checks: Sequence[Check], values: Mapping[str, str]) -> List[str]:
    """Diagnostics for every relation that fails or cannot be evaluated."""
    return [message for check in checks if (message := evaluate_check(check, values)) is not None]


def check_names(checks: Sequence[Check]) -> List[str]:
    """Term keys the checks refer to, in first-use order."""
    names: List[str] = []

    def visit(node: Node) -> None:
        if node[0] == "name":
            if node[1] not in names:
                names.append(node[1])
        elif node[0] != "number":
            for child in node[1:]:
                visit(child)

    for check in checks:
        visit(check.left)
        visit(check.right)
    return names


def checks_to_json(checks: Sequence[Check], values: Mapping[str, str]) -> str:
    """Serialize checks for the `ibrep.checks` document property.

    The generated values of every referenced term are stored too, so the
    checks can be evaluated after editing even for terms without a content
    control in the body.

    Args:
        checks: Parsed relations.
        values: Term values at generation.

    Returns:
        JSON with `checks` and `values`.
    """
    return json.dumps(
        {
            "checks": [{"check": check.source, "tolerance": check.tolerance} for check in checks],
            "values": {name: values[name] for name in check_names(checks) if name in values},
        },
        ensure_ascii=False,
    )


def checks_from_json(text: str) -> Tuple[List[Check], Dict[str, str]]:
    """Read checks and generated values stored by `checks_to_json`.

    Raises:
        ValueError: The property is not a valid check list.
    """
    data = json.loads(text)
    if not isinstance(data, dict) or not isinstance(data.get("values"), dict):
        raise ValueError("the stored checks are not a checks/values object")
    items = data.get("checks")
    checks = parse_checks([
        {key: value for key, value in item.items() if value is not None} if isinstance(item, dict) else item
        for item in items
    ] if isinstance(items, list) else items)
    return checks, {str(key): str(value) for key, value in data["values"].items()}


# ═══════════════════════════════════════════════════════════════════════════════
# REPAYMENT SCHEDULES (table spec `schedule`)
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class ScheduleSpec:
    """Zero-based columns and optional stated figures of a repayment schedule."""

    repayment: int
    balance: int
    months: Optional[int] = None
    principal: Optional[Decimal] = None  # opening principal in table units
    total: bool = False  # the last body row is a totals row, not a period
    average_life: Optional[Quantity] = None  # stated weighted average life in years
    tolerance: Optional[Decimal] = None


def check_schedule(rows: Sequence[Sequence[str]], spec: ScheduleSpec) -> List[str]:
    """Check a repayment schedule's arithmetic.

    Each period's balance must equal the previous balance less that period's
    repayment (the first period starts from the stated principal, or else from
    its own balance plus repayment); the last balance must be zero, the
    repayments must add up to the principal, a totals row must show their sum,
    and a stated weighted average life (in years, from a months column) must
    agree. Blank cells, dash placeholders and text without digits (such as a
    drawdown label) count as no repayment. Tolerances are half of the display
    step of the stated value unless `tolerance` is given.

    Args:
        rows: Visible text of each body row's cells, in table order.
        spec: Columns and stated figures.

    Returns:
        One diagnostic per disagreement (row numbers count body rows from 1).
    """
    problems: List[str] = []
    periods = list(rows[:-1]) if spec.total and rows else list(rows)
    if not periods:
        return ["the schedule has no period rows"]

    def number(
        row_index: int, row: Sequence[str], column: Optional[int], what: str,
    ) -> Optional[Tuple[Decimal, Decimal]]:
        if column is None or column >= len(row):
            return None
        try:
            return read_cell_number(row[column])
        except ValueError:
            problems.append(f"row {row_index + 1}: cannot read the {what} {row[column]!r}")
            return None

    def allowed(step: Decimal) -> Decimal:
        return spec.tolerance if spec.tolerance is not None else step / 2

    repayments: List[Tuple[Decimal, Optional[Decimal]]] = []
    previous: Optional[Decimal] = spec.principal
    opening: Optional[Decimal] = spec.principal
    last_balance: Optional[Tuple[Decimal, Decimal]] = None
    for index, row in enumerate(periods):
        repaid = number(index, row, spec.repayment, "repayment")
        balance = number(index, row, spec.balance, "balance")
        months = number(index, row, spec.months, "months")
        amount = repaid[0] if repaid else Decimal(0)
        repayments.append((amount, months[0] if months else None))
        if balance is None:
            if repaid is not None or index == len(periods) - 1:
                balance = (Decimal(0), Decimal(1))  # a dash balance means nothing is left
            else:
                continue
        if previous is None:
            opening = balance[0] + amount
        elif abs(previous - amount - balance[0]) > allowed(balance[1]):
            expected = previous - amount
            problems.append(
                f"row {index + 1}: balance {_plain(balance[0])} should be "
                f"{_plain(previous)} − {_plain(amount)} = {_plain(expected)}"
            )
        previous = balance[0]
        last_balance = balance
    if last_balance is not None and abs(last_balance[0]) > allowed(last_balance[1]):
        problems.append(f"the last balance is {_plain(last_balance[0])}, not 0")
    repaid_total = sum((amount for amount, _ in repayments), Decimal(0))
    if opening is not None and abs(repaid_total - opening) > allowed(Decimal(1)):
        problems.append(
            f"repayments add up to {_plain(repaid_total)}, not the principal {_plain(opening)}"
        )
    if spec.total and rows:
        stated = number(len(rows) - 1, rows[-1], spec.repayment, "total repayment")
        if stated is not None and abs(stated[0] - repaid_total) > allowed(stated[1]):
            problems.append(
                f"the totals row shows {_plain(stated[0])}, "
                f"but the repayments add up to {_plain(repaid_total)}"
            )
    if spec.average_life is not None:
        if spec.months is None:
            problems.append("average_life needs the months column")
        elif repaid_total:
            if any(months is None for amount, months in repayments if amount):
                problems.append("average_life needs the months of every repayment")
            else:
                weighted = sum((amount * (months or 0) for amount, months in repayments), Decimal(0))
                years = weighted / repaid_total / 12
                stated_life = spec.average_life
                tolerance = spec.tolerance if spec.tolerance is not None else stated_life.step / 2
                if abs(years - stated_life.value) > tolerance:
                    problems.append(
                        f"the weighted average life is {years.quantize(Decimal('0.001'))}년, "
                        f"not the stated {format_quantity(stated_life)}"
                    )
    return problems


def _plain(value: Decimal) -> str:
    return format_quantity(Quantity(value, "plain"))
