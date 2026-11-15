"""Regressions for the 2026-11-15 audit follow-up, group D.

The converter registry's profile, formatter placeholders and label spacing,
save permissions, portable name reservation and the contents of the shipped
archives. Every case goes through the real CLI, formatter, registry or save path.
"""

import importlib.util
import logging
import os
import re
import tomllib
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier

import pytest
from docx import Document

import cli_utils
import md_to_word
from converters import get_default_registry
from document_profiles import RenderOptions
from md_formatter import ColonLabelDetector, format_markdown
from md_parser import MarkdownParser
from tools import portable_launcher

ROOT = Path(__file__).resolve().parents[1]
_PRIVATE_USE_RE = re.compile(f"[{chr(0xE000)}-{chr(0xF8FF)}]")


# ═══════════════════════════════════════════════════════════════════════════════
# SINGLE-FILE --format (the formatter's default output name still applies)
# ═══════════════════════════════════════════════════════════════════════════════


def test_single_file_format_still_replaces_its_own_formatted_copy(tmp_path):
    source = tmp_path / "a.md"
    source.write_text("New report\n", encoding="utf-8")
    formatted = tmp_path / "a_formatted.md"
    formatted.write_text("Earlier formatted copy\n", encoding="utf-8")
    args = md_to_word.build_parser().parse_args(
        [str(source), str(tmp_path / "a.docx"), "--format", "--profile", "plain"]
    )

    assert md_to_word.run_conversion(source, args) == 0
    assert formatted.read_text(encoding="utf-8") == "New report"


# ═══════════════════════════════════════════════════════════════════════════════
# D3: REGISTRY PROFILE FROM RENDER OPTIONS
# ═══════════════════════════════════════════════════════════════════════════════


def test_registry_parses_with_the_render_options_profile(tmp_path):
    source = tmp_path / "input.md"
    source.write_text("Hello\n", encoding="utf-8")
    output = tmp_path / "out.docx"
    options = RenderOptions(profile="plain", strict=True)
    registry = get_default_registry()

    model = registry.convert(str(source), render_options=options)
    assert model.parsed_profile == "plain"
    registry.convert(model, output_format="docx", render_options=options, output_path=output)
    assert "Hello" in [p.text for p in Document(output).paragraphs]


def test_registry_explicit_profile_argument_wins_over_render_options(tmp_path):
    source = tmp_path / "input.md"
    source.write_text("Hello\n", encoding="utf-8")
    model = get_default_registry().convert(
        str(source), profile="business-report", render_options=RenderOptions(profile="plain")
    )
    assert model.parsed_profile == "business-report"


# ═══════════════════════════════════════════════════════════════════════════════
# D2: FORMATTER PLACEHOLDERS
# ═══════════════════════════════════════════════════════════════════════════════


def test_bold_in_an_extracted_title_is_restored_without_markers():
    formatted = format_markdown("[보고서] **Alpha**작성일: 2026-11-15")
    model = MarkdownParser().parse(formatted)
    assert model.metadata.title == "Alpha"
    assert 'date: "2026-11-15"' in formatted
    assert chr(0xE000) not in formatted


def test_equation_in_an_extracted_title_is_restored_as_written():
    formatted = format_markdown("[보고서] $x^2$ 분석작성일: 2026-11-15")
    assert MarkdownParser().parse(formatted).metadata.title == "$x^2$ 분석"


@pytest.mark.parametrize(
    "text",
    [
        "__BOLD_0__ **Alpha**",
        "__LATEX_INLINE_0__ $x$ 값",
        "__LATEX_BLOCK_0__ $$y$$ 값",
        chr(0xE000) + "BOLD0" + chr(0xE001) + " **Alpha**",
    ],
)
def test_text_that_looks_like_a_placeholder_is_kept(text):
    formatted = format_markdown(text)
    literal = text.split(" ", 1)[0]
    assert literal in formatted
    assert formatted.count(literal) == 1
    protected = re.search(r"\*\*Alpha\*\*|\$\$y\$\$|\$x\$", text)
    assert protected is not None and protected.group(0) in formatted


@pytest.mark.parametrize(
    "text, expected",
    [
        # A false inline match ($5 ... $6) holds the block equation's placeholder.
        ("가격은 $5 이고 $$y=1$$ 이며 $6 이다.", ["$$y=1$$", "$5 이고", "이며 $6"]),
        # Bold inside an equation and an equation inside bold.
        ("식 $**a**+b$ 와 **굵게 $x$ 표시** 끝", ["$**a**+b$", "**굵게 $x$ 표시**"]),
        ("식 $$**a**$$ 와 **굵게 $$y$$ 표시** 끝", ["$$**a**$$", "$$y$$"]),
    ],
)
def test_nested_protected_text_is_restored(text, expected, caplog):
    with caplog.at_level(logging.ERROR, logger="md_formatter"):
        formatted = format_markdown(text)
    assert not _PRIVATE_USE_RE.search(formatted)
    for piece in expected:
        assert piece in formatted
    assert not caplog.records


@pytest.mark.parametrize(
    "text, title",
    [
        ("[보고서] **굵게 $x$ 값**작성일: 2026-11-15", "굵게 $x$ 값"),
        (
            "[보고서] 가격 $5 와 $$y=1$$ 및 $6 분석작성일: 2026-11-15",
            "가격 $5 와 $$y=1$$ 및 $6 분석",
        ),
    ],
)
def test_nested_protected_text_is_restored_in_metadata(text, title):
    formatted = format_markdown(text)
    assert not _PRIVATE_USE_RE.search(formatted)
    assert MarkdownParser().parse(formatted).metadata.title == title


def test_equation_text_holding_the_bold_marker_name_is_not_reported(caplog):
    text = f"식 $a{chr(0xE000)}BOLD$ 와 **굵게** 끝"
    with caplog.at_level(logging.ERROR, logger="md_formatter"):
        formatted = format_markdown(text)
    assert f"$a{chr(0xE000)}BOLD$" in formatted and "**굵게**" in formatted
    assert not caplog.records


@pytest.mark.parametrize("lines", [1, 30])
def test_a_placeholder_left_in_the_output_is_logged_not_stripped(lines, caplog, monkeypatch):
    import md_formatter

    if lines == 1:
        monkeypatch.setattr(md_formatter.LaTeXProtector, "restore", lambda self, text, **_: text)
        text = "식 $x$ 끝"
    else:
        # A rule that copies the text leaves a second copy of the literal block's
        # placeholder, which nothing restores.
        monkeypatch.setattr(
            md_formatter.ColonLabelDetector,
            "normalize_colon_spacing",
            staticmethod(lambda value: value + "\n" + value),
        )
        text = "```\ncode\n```\n" + "본문\n" * lines
    with caplog.at_level(logging.ERROR, logger="md_formatter"):
        formatted = format_markdown(text)
    assert _PRIVATE_USE_RE.search(formatted)
    assert any("1 protected-text placeholder" in r.getMessage() for r in caplog.records)


# ═══════════════════════════════════════════════════════════════════════════════
# D7: LABEL SPACING STAYS ON ITS LINE
# ═══════════════════════════════════════════════════════════════════════════════


def test_label_spacing_does_not_join_a_value_on_the_next_line():
    assert ColonLabelDetector.normalize_colon_spacing("작성일:\n2026-11-15") == (
        "작성일:\n2026-11-15"
    )
    assert ColonLabelDetector.normalize_colon_spacing("가정:\t \n값") == "가정:\t \n값"
    assert ColonLabelDetector.normalize_colon_spacing("작성일:\t 2026 수신:A") == (
        "작성일: 2026 수신: A"
    )


def test_light_format_keeps_a_label_and_its_next_line_apart():
    source = "작성일:\n2026-11-15\n\n가정:\n연 매출\n" + "본문\n" * 22
    formatted = format_markdown(source)
    assert formatted.startswith("작성일:\n2026-11-15\n\n가정:\n연 매출\n")


# ═══════════════════════════════════════════════════════════════════════════════
# D5: SAVE PERMISSIONS WITHOUT TOUCHING THE PROCESS UMASK
# ═══════════════════════════════════════════════════════════════════════════════


def _document(text: str):
    return md_to_word.IBDocumentRenderer(options=RenderOptions(profile="plain")).render(
        MarkdownParser(profile="plain").parse(text)
    )


def test_saving_a_new_file_does_not_change_the_process_umask(tmp_path, monkeypatch):
    def forbidden_umask(mask):
        raise AssertionError("the process umask must not be changed while saving")

    monkeypatch.setattr(cli_utils.os, "umask", forbidden_umask)
    output = tmp_path / "new.docx"
    assert md_to_word.safe_save(_document("New file"), output) == output
    assert "New file" in [p.text for p in Document(output).paragraphs]
    assert [path.name for path in tmp_path.iterdir()] == ["new.docx"]


@pytest.mark.skipif(os.name == "nt", reason="POSIX permission bits")
def test_new_file_mode_follows_the_umask_set_elsewhere(tmp_path):
    previous = os.umask(0o027)
    try:
        output = tmp_path / "new.docx"
        md_to_word.safe_save(_document("Mode"), output)
    finally:
        os.umask(previous)
    assert output.stat().st_mode & 0o777 == 0o640


# ═══════════════════════════════════════════════════════════════════════════════
# D4: PORTABLE CONVERTER RESERVES ITS DESTINATION
# ═══════════════════════════════════════════════════════════════════════════════


def test_two_portable_conversions_into_one_folder_keep_both_results(tmp_path, monkeypatch):
    sources = []
    for index, text in enumerate(("First result", "Second result")):
        folder = tmp_path / f"job{index}"
        folder.mkdir()
        source = folder / "report.md"
        source.write_text(f"---\nprofile: plain\n---\n{text}\n", encoding="utf-8")
        sources.append(source)
    output_dir = tmp_path / "out"
    ready = Barrier(2)
    original_convert = md_to_word.IBReportConverter.convert

    def convert_together(self):
        # Both windows have chosen their destination before either saves.
        ready.wait(timeout=10)
        return original_convert(self)

    monkeypatch.setattr(md_to_word.IBReportConverter, "convert", convert_together)
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [
            pool.submit(portable_launcher.convert_document, str(source), str(output_dir))
            for source in sources
        ]
        saved = [future.result(timeout=60) for future in futures]

    assert len(set(saved)) == 2
    assert output_dir / "report.docx" in saved
    texts = sorted(Document(path).paragraphs[0].text for path in saved)
    assert texts == ["First result", "Second result"]
    assert sorted(output_dir.iterdir()) == sorted(saved)


def test_failed_portable_conversion_removes_its_reservation(tmp_path):
    source = tmp_path / "missing image.md"
    source.write_text("---\nprofile: plain\n---\n![x](absent.png)\n", encoding="utf-8")
    output_dir = tmp_path / "out"
    existing = output_dir / "missing image.docx"
    output_dir.mkdir()
    existing.write_bytes(b"earlier result")
    with pytest.raises(RuntimeError):
        portable_launcher.convert_document(str(source), str(output_dir))
    assert list(output_dir.iterdir()) == [existing]
    assert existing.read_bytes() == b"earlier result"


# ═══════════════════════════════════════════════════════════════════════════════
# D6: SHIPPED READMES POINT AT SHIPPED FILES
# ═══════════════════════════════════════════════════════════════════════════════

_LINK_RE = re.compile(r"\]\(([^)\s#]+)(?:#[^)]*)?\)")
_PATH_RE = re.compile(r"(?<![\w/.-])((?:docs|samples/qa)/[\w./\-가-힣]+)")
# Linked from README.md but deliberately not shipped: the archived planning
# files (`docs/archive/`, repository only) and the test suite.
_NOT_SHIPPED = {
    "docs/archive",
    "docs/archive/plan.md",
    "docs/archive/plan_upgrade.md",
    "docs/archive/next_step.md",
    "docs/archive/roadmap.md",
    "docs/archive/git_checklist.md",
    "tests",
}


_FENCED_RE = re.compile(r"^ {0,3}(`{3,}|~{3,}).*?^ {0,3}\1[^\n]*$", re.MULTILINE | re.DOTALL)
_CODE_SPAN_RE = re.compile(r"(`+)(?!`).+?(?<!`)\1(?!`)", re.DOTALL)


def _readme_targets() -> set[str]:
    """Every relative file a shipped README links to or names under `docs/`.

    Links inside code (examples such as `![alt](path)`) are not links. A
    relative link that resolves to nothing fails the test instead of being
    skipped; absolute URLs and pure `#anchor` links are not files.
    """
    targets: set[str] = set()
    broken: list[str] = []
    for name in ("README.md", "README.ko.md"):
        text = (ROOT / name).read_text(encoding="utf-8")
        prose = _CODE_SPAN_RE.sub("", _FENCED_RE.sub("", text))
        for match in _LINK_RE.finditer(prose):
            raw = match.group(1)
            if re.match(r"[A-Za-z][A-Za-z0-9+.-]*:", raw):
                continue  # absolute URL (https:, mailto:, ...)
            target = urllib.parse.unquote(raw.split("?", 1)[0]).rstrip("/")
            if not target:
                continue
            if (ROOT / target).exists():
                targets.add(target)
            else:
                broken.append(f"{name}: {raw}")
        for match in _PATH_RE.finditer(text):
            target = match.group(1).rstrip("/.,")
            if (ROOT / target).exists():
                targets.add(target)
    assert broken == [], "README links that resolve to nothing"
    return targets - _NOT_SHIPPED


def _covered(target: str, files: set[str]) -> bool:
    return target in files or any(name.startswith(target + "/") for name in files)


def test_codex_bundle_ships_every_file_its_readmes_link():
    spec = importlib.util.spec_from_file_location(
        "build_codex_bundle", ROOT / "scripts/build_codex_bundle.py"
    )
    assert spec is not None and spec.loader is not None
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    _, paths = builder.source_files(ROOT)
    files = set(paths)
    missing = sorted(t for t in _readme_targets() if not _covered(t, files))
    assert missing == []
    assert {"samples/qa/charts-dcm.md", "samples/qa/office-letter-appendix.md"} <= files


def test_sdist_ships_every_document_its_readmes_link():
    config = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    include = config["tool"]["hatch"]["build"]["targets"]["sdist"]["include"]
    exact = {pattern.lstrip("/") for pattern in include if not pattern.endswith("/**")}
    trees = [pattern.lstrip("/")[: -len("/**")] for pattern in include if pattern.endswith("/**")]
    documents = {t for t in _readme_targets() if t.startswith("docs/")}
    documents.add("scripts/word_schema_validation.md")
    missing = sorted(
        t
        for t in documents
        if t not in exact and not any(t.startswith(f"{tree}/") for tree in trees)
    )
    assert missing == []
    # Never the archived planning files, the tests or anything private.
    assert not any(
        pattern.lstrip("/").split("/")[0] in {"tests", "private", "output", "dist"}
        or pattern.lstrip("/").startswith("docs/archive")
        for pattern in include
    )
