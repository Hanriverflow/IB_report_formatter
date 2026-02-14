"""Tests for DeepResearch marker cleaner."""

from deep_md_cleaner import (
    CleanerConfig,
    DeepResearchCleaner,
    MARKER_END,
    MARKER_SEPARATOR,
    MARKER_START,
    clean_deepresearch_markdown,
)


def _marker(tag: str, payload: str) -> str:
    return "{start}{tag}{sep}{payload}{end}".format(
        start=MARKER_START,
        tag=tag,
        sep=MARKER_SEPARATOR,
        payload=payload,
        end=MARKER_END,
    )


def test_off_mode_keeps_text_unchanged():
    text = "A {cite} B".format(cite=_marker("cite", "turn1search0"))
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="off"))

    assert cleaned == text
    assert report.applied is False
    assert report.was_modified is False


def test_auto_mode_skips_plain_text():
    text = "Plain markdown without deepresearch markers."
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="auto"))

    assert cleaned == text
    assert report.markers_detected is False
    assert report.applied is False


def test_auto_mode_cleans_when_marker_exists():
    text = "Reference{cite}.".format(cite=_marker("cite", "turn1search0"))
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="auto"))

    assert "[^1]" in cleaned
    assert "## Citations" in cleaned
    assert "[^1]: turn1search0" in cleaned
    assert report.markers_detected is True
    assert report.applied is True
    assert report.cite_markers == 1


def test_cite_mode_inline():
    payload = "turn5search0{sep}turn1search8".format(sep=MARKER_SEPARATOR)
    text = "Ref{cite}".format(cite=_marker("cite", payload))
    cleaned, report = clean_deepresearch_markdown(
        text,
        CleanerConfig(activation_mode="on", cite_mode="inline"),
    )

    assert "(sources: turn5search0, turn1search8)" in cleaned
    assert "## Citations" not in cleaned
    assert report.footnotes_emitted == 0


def test_cite_mode_strip():
    text = "Before{cite}After".format(cite=_marker("cite", "turn3view3"))
    cleaned, report = clean_deepresearch_markdown(
        text,
        CleanerConfig(activation_mode="on", cite_mode="strip"),
    )

    assert cleaned == "BeforeAfter"
    assert report.cite_markers == 1
    assert report.footnotes_emitted == 0


def test_entity_keeps_display_name():
    payload = '["company", "신세계프라퍼티", "korean company"]'
    text = "파트너: {entity}".format(entity=_marker("entity", payload))
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="on"))

    assert cleaned == "파트너: 신세계프라퍼티"
    assert report.entity_markers == 1


def test_image_group_converts_to_comment():
    payload = '{"query":["a","b"],"layout":"carousel"}'
    text = "{img}\n본문".format(img=_marker("image_group", payload))
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="on"))

    assert "<!-- image_group: a, b -->" in cleaned
    assert "본문" in cleaned
    assert report.image_group_markers == 1


def test_unknown_marker_preserved_as_comment_by_default():
    text = "x{unknown}y".format(unknown=_marker("widget", "foo bar baz"))
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="on"))

    assert "<!-- widget: foo bar baz -->" in cleaned
    assert report.unknown_markers == 1


def test_unknown_marker_dropped_when_configured():
    text = "x{unknown}y".format(unknown=_marker("widget", "foo bar baz"))
    cleaned, report = clean_deepresearch_markdown(
        text,
        CleanerConfig(activation_mode="on", drop_unknown_markers=True),
    )

    assert cleaned == "xy"
    assert report.unknown_markers == 1


def test_leftover_pua_removed_after_processing():
    leftover = "\ue20a"
    text = "A{cite}{leftover}B".format(cite=_marker("cite", "turn1search0"), leftover=leftover)
    cleaned, report = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="on"))

    assert leftover not in cleaned
    assert report.leftover_pua_removed == 1


def test_citation_numbers_are_deduplicated_by_source_id():
    first = _marker("cite", "turn1search0")
    second = _marker("cite", "turn1search0{sep}turn2search1".format(sep=MARKER_SEPARATOR))
    text = "{a} then {b}".format(a=first, b=second)
    cleaned, _ = clean_deepresearch_markdown(text, CleanerConfig(activation_mode="on"))

    assert "[^1] then [^1][^2]" in cleaned
    assert "[^1]: turn1search0" in cleaned
    assert "[^2]: turn2search1" in cleaned


def test_class_api_detect():
    cleaner = DeepResearchCleaner(CleanerConfig(activation_mode="auto"))
    assert cleaner.detect(_marker("cite", "turn1search0")) is True
    assert cleaner.detect("plain text") is False
