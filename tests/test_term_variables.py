"""Term variables: `terms:` frontmatter substituted at parse time as tagged-value runs."""

import datetime
import json
import logging
from typing import Dict, List

import pytest

from document_model import (
    Blockquote,
    CodeBlock,
    DocumentModel,
    ElementType,
    Heading,
    ListItem,
    Paragraph,
    Table,
    TableType,
    TextRun,
)
from document_profiles import RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, TextParser
from term_variables import validate_terms

VALUES = {"amount": "500억원", "tenor": "3년", "rate": "CD(3개월) + [1.10]%"}
# Every inline syntax the parser knows must survive verbatim inside a value.
SPECIAL = (
    '**x** $1 and $2 <br> | \\ a  b ( y ) [^1] `c` _u_ ~s~ ^p^ '
    '<span style="color:#FF0000">r</span> [l](https://example.com) {{amount}}'
)


def front(values: Dict[str, str], extra: str = "") -> str:
    """Build frontmatter whose values are exact YAML strings (JSON is valid YAML)."""
    lines = "".join(
        f"  {key}: {json.dumps(value, ensure_ascii=False)}\n" for key, value in values.items()
    )
    return "---\n" + extra + "terms:\n" + lines + "---\n"


def parse(markdown: str, profile: str = "plain") -> DocumentModel:
    """Parse with the production parser."""
    return MarkdownParser(profile=profile).parse(markdown)


def first(model: DocumentModel, element_type: ElementType):
    """Return the first element content of a type."""
    return next(e.content for e in model.elements if e.element_type == element_type)


def texts(runs: List[TextRun]) -> List[str]:
    """Visible text of each run."""
    return [run.text for run in runs]


def keyed(runs: List[TextRun]) -> List[tuple]:
    """Text and term key of each run."""
    return [(run.text, run.term_key) for run in runs]


# ─────────────────────────────────────────────────────────────────────────────
# Schema
# ─────────────────────────────────────────────────────────────────────────────


def test_validate_terms_accepts_single_line_strings_in_order() -> None:
    values = {"amount": "500억원", "a" * 40: "", "rate_1": "  [1.10]%\t"}
    assert validate_terms(values) == values
    assert list(validate_terms(values)) == list(values)


@pytest.mark.parametrize("value", [None, [], ["amount"], "amount: 1", 3])
def test_validate_terms_requires_a_mapping(value) -> None:
    with pytest.raises(ValueError, match="mapping"):
        validate_terms(value)


@pytest.mark.parametrize("key", ["Amount", "1st", "a-b", "a b", "", "_a", "a" * 41, "금액", 1, True, None])
def test_validate_terms_rejects_malformed_keys(key) -> None:
    with pytest.raises(ValueError, match="term key"):
        validate_terms({key: "value"})


@pytest.mark.parametrize("value", [1.1, 500, True, datetime.date(2026, 9, 29)])
def test_validate_terms_rejects_scalars_with_a_quoting_hint(value) -> None:
    with pytest.raises(ValueError, match="quote"):
        validate_terms({"rate": value})


@pytest.mark.parametrize("value, message", [
    (None, "no value"),
    (["a"], "single string"),
    ({"a": "b"}, "single string"),
    ("a\nb", "single line"),
    ("a\r\nb", "single line"),
    ("a\u2028b", "single line"),
    ("a\x07b", "control character"),
])
def test_validate_terms_rejects_non_text_values(value, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        validate_terms({"rate": value})


@pytest.mark.parametrize("yaml_terms", ["terms:\n  rate: 1.10\n", "terms:\n", "terms: [a]\n"])
def test_invalid_frontmatter_terms_stop_parsing(yaml_terms: str) -> None:
    with pytest.raises(ValueError):
        parse("---\n" + yaml_terms + "---\nBody {{rate}}.")


def test_numeric_yaml_value_error_tells_the_author_to_quote() -> None:
    with pytest.raises(ValueError, match=r"quote.*1\.10"):
        parse("---\nterms:\n  rate: 1.10\n---\nBody {{rate}}.")


# ─────────────────────────────────────────────────────────────────────────────
# Tokenizer
# ─────────────────────────────────────────────────────────────────────────────


def test_paragraph_reference_becomes_a_value_run() -> None:
    paragraph = first(parse(front(VALUES) + "Deal {{amount}} closes."), ElementType.PARAGRAPH)
    assert isinstance(paragraph, Paragraph)
    assert paragraph.runs == [
        TextRun("Deal "), TextRun("500억원", term_key="amount"), TextRun(" closes."),
    ]
    assert paragraph.text == "Deal 500억원 closes."


@pytest.mark.parametrize("markdown, expected", [
    ("**{{amount}}**", TextRun("500억원", bold=True, term_key="amount")),
    ("*{{amount}}*", TextRun("500억원", italic=True, term_key="amount")),
    ("_{{amount}}_", TextRun("500억원", italic=True, term_key="amount")),
    ("^{{amount}}^", TextRun("500억원", superscript=True, term_key="amount")),
    ('<span style="color:#C00000">{{amount}}</span>',
     TextRun("500억원", color_hex="#C00000", term_key="amount")),
    ("[{{amount}}](https://example.com/deal)",
     TextRun("500억원", hyperlink="https://example.com/deal", term_key="amount")),
])
def test_value_run_inherits_surrounding_formatting(markdown: str, expected: TextRun) -> None:
    paragraph = first(parse(front(VALUES) + markdown), ElementType.PARAGRAPH)
    assert paragraph.runs == [expected]


def test_emphasis_around_a_reference_splits_into_three_formatted_runs() -> None:
    paragraph = first(parse(front(VALUES) + "**A{{amount}}B** {{ tenor }}{{rate}}"), ElementType.PARAGRAPH)
    assert paragraph.runs == [
        TextRun("A", bold=True), TextRun("500억원", bold=True, term_key="amount"),
        TextRun("B", bold=True), TextRun(" "), TextRun("3년", term_key="tenor"),
        TextRun("CD(3개월) + [1.10]%", term_key="rate"),
    ]


@pytest.mark.parametrize("template", [
    "Start {{raw}} end.",
    "- Start {{raw}} end.",
    "1. Start {{raw}} end.",
    "| Label |\n|---|\n| Start {{raw}} end. |",
    "> Start {{raw}} end.",
    "## Start {{raw}} end.",
])
def test_values_are_never_reparsed(template: str) -> None:
    model = parse(front({"raw": SPECIAL, "amount": "500억원"}) + template)
    content = model.elements[0].content
    if isinstance(content, tuple):
        content = content[1]
    if isinstance(content, Table):
        runs = content.rows[1].cells[0].runs
    else:
        runs = content.runs
    assert keyed(runs) == [("Start ", None), (SPECIAL, "raw"), (" end.", None)]
    value = runs[1]
    assert not any([value.bold, value.italic, value.superscript, value.subscript,
                    value.color_hex, value.hyperlink, value.footnote_id, value.is_latex])
    assert model.warnings == []


def test_escaped_reference_stays_literal_without_diagnostics() -> None:
    model = parse(front(VALUES) + r"Literal \{{amount}} and \{{nope}} but \\{{tenor}}.")
    paragraph = first(model, ElementType.PARAGRAPH)
    assert keyed(paragraph.runs) == [
        ("Literal {{amount}} and {{nope}} but \\", None), ("3년", "tenor"), (".", None),
    ]
    assert model.warnings == []


def test_escaped_underscore_inside_a_reference_still_names_the_key() -> None:
    paragraph = first(parse(front({"sign_date": "2026. 10. 1."}) + r"On {{sign\_date}}."), ElementType.PARAGRAPH)
    assert keyed(paragraph.runs) == [("On ", None), ("2026. 10. 1.", "sign_date"), (".", None)]


def test_code_math_fences_and_link_destinations_are_not_substituted() -> None:
    model = parse(
        front(VALUES)
        + "Use `{{amount}}` or `{{nope}}`, math $x^{{2}}$ and "
        "[{{amount}}](https://example.com/{{nope}}).\n\n"
        "```\n{{amount}} {{nope}}\n```\n\n"
        "```chart\ntype: bar\ntitle: \"{{nope}}\"\n```\n"
    )
    paragraph = first(model, ElementType.PARAGRAPH)
    assert [run.text for run in paragraph.runs if run.term_key] == ["500억원"]
    assert "`{{amount}}`" in "".join(texts(paragraph.runs))
    assert [run.text for run in paragraph.runs if run.is_latex] == ["x^{{2}}"]
    assert [run.hyperlink for run in paragraph.runs if run.hyperlink] == ["https://example.com/{{nope}}"]
    code = first(model, ElementType.CODE_BLOCK)
    assert isinstance(code, CodeBlock) and code.code == "{{amount}} {{nope}}"
    assert "{{nope}}" in first(model, ElementType.CHART).code
    assert model.warnings == []


MATH_LINK = "[X](https://example.com/{{amount}}?q=$foo$)"


def test_link_destination_with_math_stays_literal_and_unchecked() -> None:
    assert parse("---\nterms: {}\n---\n" + MATH_LINK).warnings == []
    model = parse(front({"amount": "500"}) + MATH_LINK)
    assert first(model, ElementType.PARAGRAPH).runs == TextParser.parse_runs(MATH_LINK)
    assert model.warnings == []


def test_escaped_link_syntax_is_prose_so_its_reference_is_substituted() -> None:
    paragraph = first(
        parse(front({"amount": "500"}) + r"\[X](https://example.com/{{amount}})"), ElementType.PARAGRAPH,
    )
    assert keyed(paragraph.runs) == [
        ("[X](https://example.com/", None), ("500", "amount"), (")", None),
    ]


def test_label_value_beside_a_math_destination_is_the_only_substitution() -> None:
    model = parse(front({"amount": "500"}) + "[{{amount}}](https://example.com/{{amount}}?q=$foo$)")
    runs = first(model, ElementType.PARAGRAPH).runs
    assert [run.text for run in runs if run.term_key] == ["500"]
    assert "https://example.com/{{amount}}?q=" in "".join(texts(runs))
    assert model.warnings == []


def test_value_containing_reference_syntax_is_not_substituted_again() -> None:
    model = parse(front({"a": "{{b}}", "b": "X"}) + "# Head {{a}}\n\nBody {{a}} and {{b}}.")
    heading = first(model, ElementType.HEADING_1)
    assert heading.text == "Head {{b}}"
    assert model.metadata.title == "Head {{b}}"
    assert keyed(model.metadata.display_runs["title"]) == [("Head ", None), ("{{b}}", "a")]
    paragraph = first(model, ElementType.PARAGRAPH)
    assert keyed(paragraph.runs) == [
        ("Body ", None), ("{{b}}", "a"), (" and ", None), ("X", "b"), (".", None),
    ]
    assert model.warnings == []


# ─────────────────────────────────────────────────────────────────────────────
# Application locations
# ─────────────────────────────────────────────────────────────────────────────


def test_lists_and_table_cells_receive_value_runs_but_cells_keep_raw_content() -> None:
    model = parse(
        front({**VALUES, "marker": "^^"})
        + "- Bullet {{amount}}\n\n1. Number {{tenor}}\n\n"
        "| {{tenor}} | Value |\n|---|---|\n| Amount | {{amount}} |\n| Marker | {{marker}} |"
    )
    bullet, numbered = (e.content for e in model.elements[:2])
    assert isinstance(bullet, ListItem) and bullet.text == "Bullet 500억원"
    assert keyed(bullet.runs) == [("Bullet ", None), ("500억원", "amount")]
    assert numbered[0] == "1" and keyed(numbered[1].runs) == [("Number ", None), ("3년", "tenor")]
    table = first(model, ElementType.TABLE)
    assert keyed(table.rows[0].cells[0].runs) == [("3년", "tenor")]
    assert table.rows[1].cells[1].content == "{{amount}}"
    assert keyed(table.rows[1].cells[1].runs) == [("500억원", "amount")]
    marker = table.rows[2].cells[1]
    assert marker.content == "{{marker}}" and keyed(marker.runs) == [("^^", "marker")]
    assert marker.merge is None


def test_table_semantics_follow_values_in_header_cells() -> None:
    risk = "| Risk | {{h}} |\n|---|---|\n| Rate | High |"
    detected = first(parse(front({"h": "Impact"}) + risk, profile="ib-memo"), ElementType.TABLE)
    assert detected.table_type == TableType.RISK_MATRIX
    assert [cell.risk_level for cell in detected.rows[1].cells] == [None, "high"]
    specified = first(
        parse(front({"h": "Impact"}, extra="tables:\n  - type: risk\n") + risk), ElementType.TABLE,
    )
    assert [cell.risk_level for cell in specified.rows[1].cells] == [None, "high"]
    year = first(
        parse(front({"year": "2026"}) + "| Item | {{year}} |\n|---|---|\n| Sales | 1 |", profile="ib-memo"),
        ElementType.TABLE,
    )
    assert year.table_type == TableType.FINANCIAL
    assert year.rows[0].cells[1].content == "{{year}}"


@pytest.mark.parametrize("markdown, element_type", [
    ("# Tenor {{tenor}}", ElementType.HEADING_1),
    ("## Tenor {{tenor}}", ElementType.HEADING_2),
    ("### Tenor {{tenor}}", ElementType.HEADING_3),
    ("#### Tenor {{tenor}}", ElementType.HEADING_4),
    ("Tenor {{tenor}}\n===", ElementType.HEADING_1),
])
def test_headings_carry_value_runs_and_substituted_text(markdown: str, element_type: ElementType) -> None:
    heading = first(parse(front(VALUES) + markdown), element_type)
    assert isinstance(heading, Heading)
    assert heading.text == "Tenor 3년"
    assert heading.runs == [TextRun("Tenor "), TextRun("3년", term_key="tenor")]


def test_headings_and_quotes_without_values_keep_empty_carried_runs() -> None:
    model = parse(front(VALUES) + "## Plain heading\n\n## Escaped \\{{tenor}}\n\n> Plain quote")
    headings = [e.content for e in model.elements if e.element_type == ElementType.HEADING_2]
    assert [h.runs for h in headings] == [[], []]
    assert headings[1].text == "Escaped \\{{tenor}}"
    assert first(model, ElementType.BLOCKQUOTE).runs == []


def test_ib_numbered_heading_substitutes_from_the_source_line() -> None:
    model = parse(front(VALUES) + "**1. 개요 {{amount}}** \\{{tenor}}", profile="ib-report")
    heading = first(model, ElementType.NUMBERED_HEADING)
    assert heading.text == "**1. 개요 500억원** {{tenor}}"
    assert keyed(heading.runs) == [("1. 개요 ", None), ("500억원", "amount"), (" {{tenor}}", None)]


def test_blockquote_carries_value_runs_after_its_label() -> None:
    quote = first(
        parse(front(VALUES) + "> [참고] Amount {{amount}}\n> for {{tenor}}"), ElementType.BLOCKQUOTE,
    )
    assert isinstance(quote, Blockquote)
    assert quote.title == "참고"
    assert quote.text == "Amount 500억원 for 3년"
    assert keyed(quote.runs) == [
        ("Amount ", None), ("500억원", "amount"), (" for ", None), ("3년", "tenor"),
    ]


def test_title_subtitle_and_date_get_display_runs_and_substituted_strings() -> None:
    model = parse(front(
        VALUES,
        extra='title: "가나다 {{amount}}"\nsubtitle: "Tenor {{tenor}}"\ndate: "{{sign}}"\n',
    ).replace("terms:\n", "terms:\n  sign: \"2026. 09.\"\n"))
    metadata = model.metadata
    assert (metadata.title, metadata.subtitle, metadata.extra["date"]) == (
        "가나다 500억원", "Tenor 3년", "2026. 09.",
    )
    assert metadata.display_runs == {
        "title": [TextRun("가나다 "), TextRun("500억원", term_key="amount")],
        "subtitle": [TextRun("Tenor "), TextRun("3년", term_key="tenor")],
        "date": [TextRun("2026. 09.", term_key="sign")],
    }


def test_metadata_without_references_keeps_no_display_runs() -> None:
    model = parse(front(VALUES, extra='title: "Plain title"\ndate: "2026. 09."\n'))
    assert model.metadata.display_runs == {}
    assert model.metadata.title == "Plain title"


def test_table_caption_unit_source_and_as_of_are_substituted_as_text() -> None:
    model = parse(
        front({**VALUES, "raw": SPECIAL}, extra=(
            "tables:\n  - caption: \"표 {{amount}}\"\n    unit: \"{{tenor}}\"\n"
            "    source: \"{{raw}}\"\n    as_of: \"\\\\{{tenor}}\"\n"
        ))
        + "| A | B |\n|---|---|\n| 1 | 2 |"
    )
    table = first(model, ElementType.TABLE)
    assert (table.caption, table.unit) == ("표 500억원", "3년")
    assert table.as_of == "\\{{tenor}}"
    # Rendering re-reads these fields as inline Markdown, so the value must display literally.
    assert texts(TextParser.parse_runs(table.source)) == [SPECIAL]
    assert model.warnings == []


def test_rendered_caption_shows_special_value_literally() -> None:
    model = parse(
        front({"raw": SPECIAL, "amount": "500억원"}, extra="tables:\n  - caption: \"표 {{raw}}\"\n")
        + "| A | B |\n|---|---|\n| 1 | 2 |"
    )
    doc = IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model)
    assert doc.paragraphs[0].text == "표 " + SPECIAL


def test_title_deduplication_uses_substituted_text() -> None:
    markdown = front(VALUES, extra='title: "가나다 {{amount}}"\n') + "# 가나다 {{amount}}\n\nBody."
    model = parse(markdown, profile="business-report")
    doc = IBDocumentRenderer(options=RenderOptions(profile="business-report", strict=True)).render(model)
    # Content-control text is invisible to Paragraph.text, so compare the XML text.
    visible = ["".join(p._p.xpath(".//w:t/text()")) for p in doc.paragraphs]
    assert visible.count("가나다 500억원") == 1
    assert not [p for p in doc.paragraphs if p.style.name == "Heading 1"]


# ─────────────────────────────────────────────────────────────────────────────
# Diagnostics and activation
# ─────────────────────────────────────────────────────────────────────────────


def test_undefined_and_invalid_references_warn_once_and_stay_literal() -> None:
    model = parse(
        front(VALUES) + "A {{nope}} {{nope}} {{Bad-Key}} {{}}\n\n- {{nope}}\n\n## {{Bad-Key}}"
    )
    assert model.warnings == [
        "Invalid term reference: {{Bad-Key}} (keys use lowercase letters, digits and "
        "underscores and start with a letter)",
        "Invalid term reference: {{}} (keys use lowercase letters, digits and "
        "underscores and start with a letter)",
        "Undefined term: {{nope}}",
    ]
    assert texts(first(model, ElementType.PARAGRAPH).runs) == ["A {{nope}} {{nope}} {{Bad-Key}} {{}}"]
    assert first(model, ElementType.HEADING_2).text == "{{Bad-Key}}"


def test_strict_render_rejects_undefined_terms_before_saving() -> None:
    model = parse(front(VALUES) + "Missing {{nope}}.")
    with pytest.raises(ValueError, match="Undefined term"):
        IBDocumentRenderer(options=RenderOptions(profile="plain", strict=True)).render(model)
    doc = IBDocumentRenderer(options=RenderOptions(profile="plain")).render(model)
    assert "Missing {{nope}}." in [p.text for p in doc.paragraphs]


def test_unused_terms_are_logged_not_warned(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.INFO, logger="md_parser"):
        model = parse(front({**VALUES, "spare": "x"}) + "Uses {{amount}}.")
    assert model.warnings == []
    assert any(
        "Unused terms" in record.getMessage() and "spare" in record.getMessage()
        and "tenor" in record.getMessage() and record.levelno == logging.INFO
        for record in caplog.records
    )


def test_without_terms_key_references_are_ordinary_text() -> None:
    source = "Text {{amount}} and \\{{x}}.\n\n## Head {{amount}}\n\n> Quote {{amount}}"
    model = parse("---\ntitle: T {{amount}}\n---\n" + source)
    paragraph = first(model, ElementType.PARAGRAPH)
    assert paragraph.runs == TextParser.parse_runs("Text {{amount}} and \\{{x}}.")
    assert first(model, ElementType.HEADING_2).runs == []
    assert first(model, ElementType.BLOCKQUOTE).runs == []
    assert model.metadata.title == "T {{amount}}"
    assert model.metadata.display_runs == {}
    assert model.warnings == []


def test_empty_terms_mapping_turns_checking_on() -> None:
    model = parse("---\nterms: {}\n---\nText {{amount}}.")
    assert model.warnings == ["Undefined term: {{amount}}"]
