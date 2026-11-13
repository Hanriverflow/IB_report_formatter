"""Module layout guards for splitting ib_renderer.py and md_parser.py.

The split moves top-level definitions of the two large modules into internal
files while ``ib_renderer`` and ``md_parser`` keep exporting every name that
callers use today. These checks pin that contract:

1. Every listed name is importable from the original module.
2. Every name that repository code reaches through the two modules exists
   there. The scan reads ``from module import X``, ``module.X``, a literal
   ``"module.X"`` and calls that pass the module and the name separately,
   such as ``setattr(module, "X", value)``.
3. A module-level name that tests replace on an original module is read only by
   definitions that stay in that module; a replacement cannot reach a copy of
   the name in another file.
4. Each moved definition is in its split file and is the same object as the
   name exported by the original module.
5. The wheel, sdist and first-party lists in pyproject.toml match the engine
   files, and the split files on disk are exactly the listed ones.
6. Split files do not import their original module at module level, and
   term_sheet.py does not import ib_renderer.
7. Every split file and both original modules import on their own in a fresh
   interpreter, so no import order hides a cycle.
8. A split file's ``logger`` is the original module's logger.
"""

import ast
import fnmatch
import importlib
import os
import re
import subprocess
import sys
import tomllib
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

import pytest

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
SCANNED_DIRECTORIES = ("tests", "scripts", "tools")
SKIPPED_DIRECTORY_NAMES = frozenset(
    {".venv", ".git", "reports", "temp", "tmp", "private", "__snapshots__", "__pycache__"}
)
# Legacy unreadable pytest temp dirs may exist in local checkouts (AGENTS.md: do not read).
LEGACY_TEMP_DIRECTORY_PATTERN = "Users*AppDataLocalTemp*"
ORIGINAL_MODULES = ("ib_renderer", "md_parser")
SPLIT_FILE_PATTERNS = {"ib_renderer": "ib_render_*.py", "md_parser": "md_parse_*.py"}

# Split files, in move order, with their original module and moved definitions.
SPLIT_LAYOUT: dict[str, tuple[str, tuple[str, ...]]] = {
    "ib_render_fonts": ("ib_renderer", ("FontStyler", "FontPolicy")),
    "ib_render_document_style": (
        "ib_renderer",
        ("GeneratorSignatureWriter", "DocumentStyler", "TableStyler"),
    ),
    "ib_render_footnotes": (
        "ib_renderer",
        ("FootnoteRenderer", "NativeFootnotesPart", "rebase_local_links"),
    ),
    "ib_render_media": ("ib_renderer", ("ImageRenderer", "ChartRenderer", "LaTeXRenderer")),
    "ib_render_text": ("ib_renderer", ("TextRenderer",)),
    "ib_render_blocks": (
        "ib_renderer",
        ("TOCRenderer", "HeadingRenderer", "ParagraphRenderer", "ListRenderer", "CalloutRenderer"),
    ),
    "ib_render_tables": ("ib_renderer", ("TableRenderer", "_BLACK")),
    "md_parse_text": (
        "md_parser",
        (
            "TextParser",
            "file_uri",
            "rebase_link_target",
            "_rebase_link_target",
            "_parse_term_runs",
            "_substitute_text",
            "_document_int",
            "_DRIVE_PATH_RE",
            "_URI_DRIVE_PATH_RE",
            "_URI_SCHEME_RE",
        ),
    ),
    "md_parse_frontmatter": (
        "md_parser",
        (
            "FrontmatterParser",
            "_infer_metadata_from_elements",
            "_extract_leading_metadata_pair",
            "_infer_company_from_title",
        ),
    ),
    "md_parse_tables": ("md_parser", ("TableParser", "TableSpanResolver")),
    "md_parse_html_tables": (
        "md_parser",
        ("_HtmlCell", "_html_image", "_HtmlTableReader", "HtmlTableParser"),
    ),
    "md_parse_blocks": (
        "md_parser",
        ("FenceScanner", "FootnoteParser", "LaTeXParser", "Base64ImageParser"),
    ),
}

# ═══════════════════════════════════════════════════════════════════════════════
# NAMES THAT MUST STAY IMPORTABLE
# ═══════════════════════════════════════════════════════════════════════════════

IB_RENDERER_DEFINED_NAMES = (
    "logger",
    "_BLACK",
    "GeneratorSignatureWriter",
    "FontStyler",
    "FontPolicy",
    "DocumentStyler",
    "TableStyler",
    "TextRenderer",
    "CoverRenderer",
    "TOCRenderer",
    "HeadingRenderer",
    "ParagraphRenderer",
    "ListRenderer",
    "TableRenderer",
    "CalloutRenderer",
    "ImageRenderer",
    "FootnoteRenderer",
    "NativeFootnotesPart",
    "DisclaimerRenderer",
    "rebase_local_links",
    "ChartRenderer",
    "_metadata_text",
    "IBDocumentRenderer",
    "LaTeXRenderer",
)
# Imported by ib_renderer and used through it by repository code.
IB_RENDERER_IMPORTED_NAMES = ("STYLE", "IBStyle", "uuid4", "Document", "time", "platform")

MD_PARSER_DEFINED_NAMES = (
    "logger",
    "_document_int",
    "FrontmatterParser",
    "_DRIVE_PATH_RE",
    "_URI_DRIVE_PATH_RE",
    "_URI_SCHEME_RE",
    "file_uri",
    "rebase_link_target",
    "_rebase_link_target",
    "TextParser",
    "_parse_term_runs",
    "_substitute_text",
    "FenceScanner",
    "FootnoteParser",
    "TableParser",
    "TableSpanResolver",
    "_HtmlCell",
    "_html_image",
    "_HtmlTableReader",
    "HtmlTableParser",
    "LaTeXParser",
    "Base64ImageParser",
    "MarkdownParser",
    "_read_with_encoding",
    "_decode_bytes",
    "parse_markdown_file",
    "_infer_metadata_from_elements",
    "_extract_leading_metadata_pair",
    "_infer_company_from_title",
)
# Re-exported by md_parser from document_model as ``Name as Name``.
MD_PARSER_MODEL_NAMES = (
    "Blockquote",
    "BulletList",
    "Chart",
    "CodeBlock",
    "ConfirmationBlock",
    "Diagram",
    "DiagramArrow",
    "DiagramBox",
    "DocumentMetadata",
    "DocumentModel",
    "Element",
    "ElementContent",
    "ElementType",
    "Footnote",
    "Heading",
    "Image",
    "LaTeXEquation",
    "ListItem",
    "NumberedList",
    "Paragraph",
    "Section",
    "Table",
    "TableCell",
    "TableRow",
    "TableType",
    "TextRun",
)

# The imported objects the original modules must keep exporting unchanged.
IMPORTED_OBJECT_SOURCES = {
    ("ib_renderer", "STYLE"): ("render_styles", "STYLE"),
    ("ib_renderer", "IBStyle"): ("render_styles", "IBStyle"),
    ("ib_renderer", "uuid4"): ("uuid", "uuid4"),
    ("ib_renderer", "Document"): ("docx", "Document"),
    ("ib_renderer", "time"): ("time", None),
    ("ib_renderer", "platform"): ("platform", None),
    **{("md_parser", name): ("document_model", name) for name in MD_PARSER_MODEL_NAMES},
}

# References the scan must find; they prove each reference form is read.
KNOWN_REFERENCES = {
    ("ib_renderer", "uuid4", "string"),
    ("ib_renderer", "LaTeXRenderer", "string"),
    ("ib_renderer", "Document", "call"),
    ("ib_renderer", "time", "call"),
    ("ib_renderer", "platform", "attribute"),
    ("ib_renderer", "IBDocumentRenderer", "import"),
    ("md_parser", "TextParser", "import"),
}
# Module-level names that tests replace on an original module today.
KNOWN_PATCHED_NAMES = {
    ("ib_renderer", "uuid4"),
    ("ib_renderer", "Document"),
    ("ib_renderer", "time"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# REPOSITORY SCAN
# ═══════════════════════════════════════════════════════════════════════════════

_DOTTED_TARGET_RE = re.compile(r"(ib_renderer|md_parser)((?:\.[A-Za-z_][A-Za-z0-9_]*)+)")
_PATCHING_CALLS = frozenset({"setattr", "delattr", "patch", "object"})


@dataclass(frozen=True)
class ModuleReference:
    """A name that repository code reaches through an original module."""

    module: str
    name: str
    form: str
    location: str
    patched: bool


def _scanned_files() -> Iterator[Path]:
    """Yield engine files and every Python file under tests, scripts and tools."""
    yield from sorted(REPOSITORY_ROOT.glob("*.py"))

    def fail(error: OSError) -> None:
        raise error

    for directory in SCANNED_DIRECTORIES:
        for current, subdirectories, files in os.walk(REPOSITORY_ROOT / directory, onerror=fail):
            subdirectories[:] = sorted(
                name
                for name in subdirectories
                if name not in SKIPPED_DIRECTORY_NAMES
                and not fnmatch.fnmatch(name, LEGACY_TEMP_DIRECTORY_PATTERN)
            )
            yield from (Path(current) / name for name in sorted(files) if name.endswith(".py"))


def _call_name(call: ast.Call) -> str | None:
    if isinstance(call.func, ast.Name):
        return call.func.id
    if isinstance(call.func, ast.Attribute):
        return call.func.attr
    return None


def _references_in(path: Path) -> list[ModuleReference]:
    """Collect every reference form in one file."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    relative = path.relative_to(REPOSITORY_ROOT).as_posix()
    aliases: dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name in ORIGINAL_MODULES:
                    aliases[alias.asname or alias.name] = alias.name

    patch_targets: set[int] = set()
    found: list[ModuleReference] = []

    def add(module: str, name: str, form: str, node: ast.AST, patched: bool = False) -> None:
        location = f"{relative}:{getattr(node, 'lineno', 0)}"
        found.append(ModuleReference(module, name, form, location, patched))

    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.level == 0:
            if node.module in ORIGINAL_MODULES:
                for alias in node.names:
                    add(node.module, alias.name, "import", node)
        elif isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
            if node.value.id in aliases:
                add(aliases[node.value.id], node.attr, "attribute", node)
        elif isinstance(node, ast.Call) and node.args:
            patching = _call_name(node) in _PATCHING_CALLS
            first = node.args[0]
            if isinstance(first, ast.Constant) and isinstance(first.value, str) and patching:
                patch_targets.add(id(first))
            if (
                len(node.args) >= 2
                and isinstance(first, ast.Name)
                and first.id in aliases
                and isinstance(node.args[1], ast.Constant)
                and isinstance(node.args[1].value, str)
            ):
                add(aliases[first.id], node.args[1].value, "call", node, patched=patching)

    for node in ast.walk(tree):
        if not (isinstance(node, ast.Constant) and isinstance(node.value, str)):
            continue
        match = _DOTTED_TARGET_RE.fullmatch(node.value)
        if match is None or node.value.endswith(".py"):
            continue
        parts = match.group(2)[1:].split(".")
        patched = len(parts) == 1 and id(node) in patch_targets
        add(match.group(1), parts[0], "string", node, patched=patched)
    return found


@pytest.fixture(scope="module")
def repository_references() -> list[ModuleReference]:
    references: list[ModuleReference] = []
    for path in _scanned_files():
        references.extend(_references_in(path))
    return references


def _top_level_definitions(path: Path) -> Iterator[tuple[str, ast.stmt]]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in tree.body:
        if isinstance(node, ast.ClassDef | ast.FunctionDef | ast.AsyncFunctionDef):
            yield node.name, node
        elif isinstance(node, ast.Assign | ast.AnnAssign):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                if isinstance(target, ast.Name):
                    yield target.id, node


def _readers_of(module: str, name: str) -> dict[str, list[str]]:
    """Map the original file and its split files to definitions that load ``name``."""
    paths = [REPOSITORY_ROOT / f"{module}.py"]
    paths.extend(sorted(REPOSITORY_ROOT.glob(SPLIT_FILE_PATTERNS[module])))
    readers: dict[str, list[str]] = {}
    for path in paths:
        for definition, node in _top_level_definitions(path):
            if any(
                isinstance(child, ast.Name) and child.id == name and isinstance(child.ctx, ast.Load)
                for child in ast.walk(node)
            ):
                readers.setdefault(path.name, []).append(definition)
    return readers


# ═══════════════════════════════════════════════════════════════════════════════
# 1. LISTED NAMES
# ═══════════════════════════════════════════════════════════════════════════════


def test_listed_names_are_complete_and_unique() -> None:
    assert len(IB_RENDERER_DEFINED_NAMES) == len(set(IB_RENDERER_DEFINED_NAMES)) == 24
    assert len(MD_PARSER_DEFINED_NAMES) == len(set(MD_PARSER_DEFINED_NAMES)) == 29
    assert len(MD_PARSER_MODEL_NAMES) == len(set(MD_PARSER_MODEL_NAMES))


@pytest.mark.parametrize(
    "module_name, names",
    [
        ("ib_renderer", IB_RENDERER_DEFINED_NAMES + IB_RENDERER_IMPORTED_NAMES),
        ("md_parser", MD_PARSER_DEFINED_NAMES + MD_PARSER_MODEL_NAMES),
    ],
)
def test_listed_names_are_importable_from_original_module(
    module_name: str, names: tuple[str, ...]
) -> None:
    module = importlib.import_module(module_name)
    assert [name for name in names if not hasattr(module, name)] == []


def test_imported_names_are_the_same_objects_as_their_sources() -> None:
    different = []
    for (module_name, name), (source_name, attribute) in IMPORTED_OBJECT_SOURCES.items():
        source = importlib.import_module(source_name)
        expected = source if attribute is None else getattr(source, attribute)
        if getattr(importlib.import_module(module_name), name) is not expected:
            different.append(f"{module_name}.{name}")
    assert different == []


# ═══════════════════════════════════════════════════════════════════════════════
# 2. NAMES THAT REPOSITORY CODE REACHES THROUGH THE ORIGINAL MODULES
# ═══════════════════════════════════════════════════════════════════════════════


def test_scan_reads_every_reference_form(repository_references: list[ModuleReference]) -> None:
    found = {(ref.module, ref.name, ref.form) for ref in repository_references}
    assert KNOWN_REFERENCES - found == set()


def test_every_referenced_name_exists_in_original_module(
    repository_references: list[ModuleReference],
) -> None:
    missing = sorted(
        f"{ref.location}: {ref.module}.{ref.name} ({ref.form})"
        for ref in repository_references
        if ref.name == "*" or not hasattr(importlib.import_module(ref.module), ref.name)
    )
    assert missing == []


# ═══════════════════════════════════════════════════════════════════════════════
# 3. REPLACED MODULE-LEVEL NAMES ARE READ ONLY IN THE ORIGINAL FILE
# ═══════════════════════════════════════════════════════════════════════════════


def test_replaced_names_are_read_only_by_definitions_in_original_file(
    repository_references: list[ModuleReference],
) -> None:
    patched = {(ref.module, ref.name) for ref in repository_references if ref.patched}
    assert KNOWN_PATCHED_NAMES - patched == set()
    problems = []
    for module, name in sorted(patched):
        readers = _readers_of(module, name)
        if not readers.pop(f"{module}.py", None):
            problems.append(f"{module}.{name}: no definition in {module}.py reads it")
        problems.extend(
            f"{module}.{name}: read by {', '.join(names)} in {file_name}"
            for file_name, names in sorted(readers.items())
        )
    assert problems == []


# ═══════════════════════════════════════════════════════════════════════════════
# 4. MOVED DEFINITIONS
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("split_module", list(SPLIT_LAYOUT))
def test_moved_definitions_live_in_split_file_and_are_reexported(split_module: str) -> None:
    original_module, names = SPLIT_LAYOUT[split_module]
    split_defined = dict(_top_level_definitions(REPOSITORY_ROOT / f"{split_module}.py"))
    original_defined = dict(_top_level_definitions(REPOSITORY_ROOT / f"{original_module}.py"))
    split = importlib.import_module(split_module)
    original = importlib.import_module(original_module)
    problems = [
        f"{name}: not defined in {split_module}.py" for name in names if name not in split_defined
    ]
    problems += [
        f"{name}: still defined in {original_module}.py"
        for name in names
        if name in original_defined
    ]
    problems += [
        f"{name}: {original_module}.{name} is not {split_module}.{name}"
        for name in names
        if getattr(original, name, None) is not getattr(split, name, None)
    ]
    assert problems == []


# ═══════════════════════════════════════════════════════════════════════════════
# 5. PACKAGING LISTS AND FILES ON DISK
# ═══════════════════════════════════════════════════════════════════════════════


def test_packaging_lists_match_engine_files() -> None:
    config = tomllib.loads((REPOSITORY_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    targets = config["tool"]["hatch"]["build"]["targets"]
    wheel = {entry for entry in targets["wheel"]["only-include"] if entry.endswith(".py")}
    sdist = {entry.lstrip("/") for entry in targets["sdist"]["include"] if entry.endswith(".py")}
    engine = {path.name for path in REPOSITORY_ROOT.glob("*.py")}
    assert wheel == sdist == engine
    first_party = set(config["tool"]["ruff"]["lint"]["isort"]["known-first-party"])
    assert set(SPLIT_LAYOUT) - first_party == set()


def test_split_files_on_disk_are_the_listed_ones() -> None:
    on_disk = {
        path.stem
        for pattern in SPLIT_FILE_PATTERNS.values()
        for path in REPOSITORY_ROOT.glob(pattern)
    }
    assert on_disk == set(SPLIT_LAYOUT)


def test_split_layout_is_complete() -> None:
    """Twelve split files hold every moved definition; the rest stays defined in place."""
    assert len(SPLIT_LAYOUT) == 12
    defined_names = {
        "ib_renderer": IB_RENDERER_DEFINED_NAMES,
        "md_parser": MD_PARSER_DEFINED_NAMES,
    }
    staying = {
        "ib_renderer": {
            "IBDocumentRenderer",
            "_metadata_text",
            "CoverRenderer",
            "DisclaimerRenderer",
            "logger",
        },
        "md_parser": {
            "MarkdownParser",
            "_read_with_encoding",
            "_decode_bytes",
            "parse_markdown_file",
            "logger",
        },
    }
    for module, names in defined_names.items():
        moved = [
            name
            for original, split_names in SPLIT_LAYOUT.values()
            if original == module
            for name in split_names
        ]
        assert len(moved) == len(set(moved)) == len(names) - len(staying[module])
        assert set(moved) | staying[module] == set(names)
        in_place = {name for name, _ in _top_level_definitions(REPOSITORY_ROOT / f"{module}.py")}
        assert in_place == staying[module]


# ═══════════════════════════════════════════════════════════════════════════════
# 6. IMPORT DIRECTION
# ═══════════════════════════════════════════════════════════════════════════════


def _imported_modules(tree: ast.AST, module_level_only: bool) -> set[str]:
    """Collect imported module names, optionally skipping function bodies."""
    found: set[str] = set()
    pending: list[ast.AST] = [tree]
    while pending:
        node = pending.pop()
        if isinstance(node, ast.Import):
            found.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            found.add(node.module)
        for child in ast.iter_child_nodes(node):
            if module_level_only and isinstance(child, ast.FunctionDef | ast.AsyncFunctionDef):
                continue
            pending.append(child)
    return found


@pytest.mark.parametrize("split_module", list(SPLIT_LAYOUT))
def test_split_file_does_not_import_its_original_module(split_module: str) -> None:
    original_module, _ = SPLIT_LAYOUT[split_module]
    path = REPOSITORY_ROOT / f"{split_module}.py"
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    assert original_module not in _imported_modules(tree, module_level_only=True)


def test_term_sheet_does_not_import_ib_renderer() -> None:
    path = REPOSITORY_ROOT / "term_sheet.py"
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    assert "ib_renderer" not in _imported_modules(tree, module_level_only=False)


# ═══════════════════════════════════════════════════════════════════════════════
# 7. EACH MODULE IMPORTS ON ITS OWN
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("module_name", [*SPLIT_LAYOUT, *ORIGINAL_MODULES])
def test_module_imports_alone_in_fresh_interpreter(module_name: str) -> None:
    result = subprocess.run(
        [sys.executable, "-c", f"import {module_name}"],
        cwd=REPOSITORY_ROOT,
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    assert result.returncode == 0, result.stderr


# ═══════════════════════════════════════════════════════════════════════════════
# 8. ONE LOGGER PER ORIGINAL MODULE
# ═══════════════════════════════════════════════════════════════════════════════


@pytest.mark.parametrize("split_module", list(SPLIT_LAYOUT))
def test_split_file_logger_is_original_logger(split_module: str) -> None:
    original_module, _ = SPLIT_LAYOUT[split_module]
    split = importlib.import_module(split_module)
    if hasattr(split, "logger"):
        assert split.logger is importlib.import_module(original_module).logger
