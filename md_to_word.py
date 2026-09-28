"""
MD to IB Style Word Report Converter
Main entry point for converting Markdown files to professional IB-style Word documents.

Usage:
    uv run md_to_word.py input.md [output.docx]
    uv run md_to_word.py --list                    # Show available md files
    uv run md_to_word.py --list --interactive       # Select and convert interactively
    uv run md_to_word.py 파일명.md                  # Convert file from parent folder
    uv run md_to_word.py 파일명.md --format         # Auto-format before converting
    uv run md_to_word.py 파일명.md --no-cover       # Skip cover page

Changelog (v2):
    - Python 3.8+ compatible (removed Path.with_stem dependency)
    - Integrated md_formatter auto-format option (--format)
    - Replaced print-based output with logging module
    - Added --no-cover, --no-toc, --no-disclaimer flags
    - Added --interactive mode for --list
    - Added conversion timing information
    - Improved error handling with per-stage reporting
    - Unified output path resolution logic

Changelog (hardening):
    - Preflight batch destinations and reject all sources sharing an output path.
"""

import argparse
import logging
import os
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set

from cli_utils import (
    generate_output_path as build_output_path,
)
from cli_utils import (
    interactive_select as prompt_for_selection,
)
from cli_utils import (
    list_files,
)
from cli_utils import (
    resolve_input_path as resolve_project_input_path,
)
from cli_utils import (
    safe_save as safe_save_with_fallback,
)
from cli_utils import (
    setup_logging as configure_logging,
)
from document_profiles import PRESETS, PROFILES
from document_profiles import RenderOptions as RenderOptions
from ib_renderer import IBDocumentRenderer
from md_parser import DocumentModel, parse_markdown_file

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Parent folder path (where md files are typically located)
PARENT_DIR = Path(__file__).resolve().parent.parent

# Output suffix
OUTPUT_SUFFIX = "_Report_Pro.docx"

# Logger
logger = logging.getLogger("md_to_word")


# ═══════════════════════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════════════════════
# FORMATTER INTEGRATION
# ═══════════════════════════════════════════════════════════════════════════════


def auto_format_markdown(
    input_path: Path,
    cleaner_mode: str = "off",
    cite_mode: str = "footnote",
    drop_unknown_markers: bool = False,
    cleaner_report: bool = False,
) -> Path:
    """
    Auto-format a markdown file using md_formatter before conversion.

    Args:
        input_path: Path to the raw markdown file

    Returns:
        Path to the formatted markdown file

    Raises:
        ImportError: If md_formatter module is not available
        Exception: If formatting fails
    """
    try:
        from md_formatter import format_file_with_options
    except ImportError as err:
        raise ImportError(
            "md_formatter.py not found. Ensure it is in the same directory as md_to_word.py."
        ) from err

    logger.info("Auto-formatting: %s", input_path.name)

    formatted_path_str = format_file_with_options(
        input_path=str(input_path),
        cleaner_mode=cleaner_mode,
        cite_mode=cite_mode,
        drop_unknown_markers=drop_unknown_markers,
        cleaner_report=cleaner_report,
    )
    formatted_path = Path(formatted_path_str)

    logger.info("Formatted output: %s", formatted_path.name)
    return formatted_path


def _read_with_encoding(file_path: Path) -> str:
    """Read markdown text with encoding fallback."""
    encodings = ["utf-8", "utf-8-sig", "euc-kr", "cp949"]
    for enc in encodings:
        try:
            return file_path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise UnicodeDecodeError(
        "multiple",
        b"",
        0,
        1,
        f"Failed to decode {file_path} with encodings: {encodings}",
    )


def clean_deepresearch_markdown_file(
    input_path: Path,
    cleaner_mode: str,
    cite_mode: str,
    drop_unknown_markers: bool,
    cleaner_report: bool,
) -> Path:
    """Conditionally clean DeepResearch markers and return path to cleaned file."""
    if cleaner_mode == "off":
        return input_path

    try:
        from deep_md_cleaner import CleanerConfig, clean_deepresearch_markdown
    except ImportError as err:
        raise ImportError(
            "deep_md_cleaner.py not found. Ensure it is in the same directory as md_to_word.py."
        ) from err

    content = _read_with_encoding(input_path)
    cleaner_config = CleanerConfig(
        activation_mode=cleaner_mode,
        cite_mode=cite_mode,
        drop_unknown_markers=drop_unknown_markers,
    )
    cleaned, report = clean_deepresearch_markdown(content, cleaner_config)

    if cleaner_report and (report.applied or report.markers_detected):
        logger.info("DeepResearch cleaner: %s", report.summary())

    if not report.was_modified:
        return input_path

    cleaned_path = input_path.with_name(input_path.stem + "_cleaned.md")
    cleaned_path.write_text(cleaned, encoding="utf-8")
    logger.info("Cleaner output: %s", cleaned_path.name)
    return cleaned_path


# ═══════════════════════════════════════════════════════════════════════════════
# PATH RESOLUTION
# ═══════════════════════════════════════════════════════════════════════════════


def resolve_input_path(input_file: str) -> Path:
    """Resolve the user-provided input path."""
    return resolve_project_input_path(input_file, parent_dir=PARENT_DIR, script_path=Path(__file__))


def generate_output_path(input_path: Path, output_path: Optional[str] = None) -> Path:
    """Generate the Word output path."""
    return build_output_path(
        input_path,
        output_path,
        ".docx",
        default_name=lambda path: f"{path.stem}{OUTPUT_SUFFIX}".replace(" ", "_"),
    )


# ═══════════════════════════════════════════════════════════════════════════════
# SAFE SAVE
# ═══════════════════════════════════════════════════════════════════════════════


def safe_save(doc, output_path: Path) -> Path:
    """Save the generated Word document with permission-error fallback."""
    return safe_save_with_fallback(
        output_path=output_path,
        save_action=lambda path: doc.save(str(path)),
        logger=logger,
        lock_message=(
            "%s is locked (possibly open in another program). Saving with timestamp suffix."
        ),
    )


# ═══════════════════════════════════════════════════════════════════════════════
# CONVERTER
# ═══════════════════════════════════════════════════════════════════════════════


class IBReportConverter:
    """
    Main converter class that orchestrates MD parsing and Word rendering.
    """

    def __init__(
        self,
        md_file_path: str,
        output_path: Optional[str] = None,
        render_options: Optional[RenderOptions] = None,
    ):
        """
        Initialize the converter.

        Args:
            md_file_path: Path to the input markdown file
            output_path: Optional path for output docx file
            render_options: Options controlling which sections to render

        Raises:
            FileNotFoundError: If input file does not exist
        """
        self.md_file_path = Path(md_file_path).resolve()
        self.output_path = output_path
        self.render_options = render_options or RenderOptions()
        self._validate_input()

    def _validate_input(self):
        """Validate that the input file exists and is readable"""
        if not self.md_file_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.md_file_path}")

        if not self.md_file_path.is_file():
            raise FileNotFoundError(f"Not a file: {self.md_file_path}")

        if self.md_file_path.suffix.lower() != ".md":
            logger.warning(
                "Input file does not have .md extension: %s",
                self.md_file_path.name,
            )

        # Check file is not empty
        if self.md_file_path.stat().st_size == 0:
            raise ValueError(f"Input file is empty: {self.md_file_path}")

    def convert(self) -> str:
        """
        Execute the full conversion pipeline: Parse → Render → Save.

        Returns:
            Absolute path to the generated document as string

        Raises:
            Exception: If any stage of the pipeline fails
        """
        start_time = time.perf_counter()

        # ── Stage 1: Parse ──────────────────────────────────────────────────
        logger.info("Parsing: %s", self.md_file_path.name)
        try:
            model = parse_markdown_file(str(self.md_file_path), profile=self.render_options.profile)
        except UnicodeDecodeError as e:
            raise RuntimeError(
                f"Failed to decode file (encoding issue): {e}\nTry saving the file as UTF-8."
            ) from e

        logger.info("Parsed %d elements", len(model.elements))
        logger.info("Title: %s", model.metadata.title)
        logger.debug("Subtitle: %s", model.metadata.subtitle)
        logger.debug("Company: %s", model.metadata.company)
        logger.info("Footnotes: %d", len(model.footnotes))

        # ── Stage 2: Render ─────────────────────────────────────────────────
        logger.info("Rendering Word document... (%s)", self.render_options)
        try:
            doc = self._render(model)
        except Exception as e:
            raise RuntimeError(f"Rendering failed: {e}") from e

        # ── Stage 3: Save ───────────────────────────────────────────────────
        output_path = generate_output_path(self.md_file_path, self.output_path)
        saved_path = safe_save(doc, output_path)

        # ── Done ────────────────────────────────────────────────────────────
        elapsed = time.perf_counter() - start_time
        logger.info("Conversion completed in %.2fs", elapsed)

        return str(saved_path)

    def _render(self, model: DocumentModel):
        """Use the same composition service as the Python API and registry."""
        renderer = IBDocumentRenderer(options=self.render_options)
        document = renderer.render(model)
        for issue in renderer.errors:
            logger.warning("%s", issue)
        return document


# ═══════════════════════════════════════════════════════════════════════════════
# LIST & INTERACTIVE MODE
# ═══════════════════════════════════════════════════════════════════════════════


def list_md_files() -> List[Path]:
    """List all markdown files in the parent directory."""
    return list_files(PARENT_DIR, "*.md", ".md", "MD")


def interactive_select(md_files: List[Path]) -> Optional[Path]:
    """Prompt user to select a markdown file by number."""
    return prompt_for_selection(md_files, "Enter file number to convert (or 'q' to quit): ")


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════


def build_parser() -> argparse.ArgumentParser:
    """Build CLI argument parser"""
    parser = argparse.ArgumentParser(
        description="Convert Markdown to IB reports and Korean business Word documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    uv run md_to_word.py --list                        # List available md files
    uv run md_to_word.py --list -i                     # Interactive selection
    uv run md_to_word.py 보고서.md                      # Convert from parent folder
    uv run md_to_word.py 보고서.md --format             # Auto-format then convert
    uv run md_to_word.py 보고서.md --no-cover           # Skip cover page
    uv run md_to_word.py 보고서.md --no-toc --no-disc   # Minimal output
    uv run md_to_word.py reports/ --batch              # Convert all md files in a directory
    uv run md_to_word.py C:/path/to/report.md          # Absolute path
        """,
    )

    parser.add_argument(
        "input_file",
        nargs="?",
        help="Input Markdown file (filename or path)",
    )

    parser.add_argument(
        "output_file",
        nargs="?",
        help="Output Word document path (optional, auto-generated if omitted)",
    )

    # Mode flags
    mode_group = parser.add_argument_group("mode options")
    mode_group.add_argument(
        "-l",
        "--list",
        action="store_true",
        help="List available markdown files in parent folder",
    )
    mode_group.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Interactive file selection (use with --list)",
    )
    mode_group.add_argument(
        "--batch",
        action="store_true",
        help="Convert all .md files in the given input directory",
    )

    # Pre-processing
    preprocess_group = parser.add_argument_group("pre-processing")
    preprocess_group.add_argument(
        "-f",
        "--format",
        action="store_true",
        help="Auto-format markdown (single-line → structured) before conversion",
    )
    preprocess_group.add_argument(
        "--deepresearch-cleaner",
        choices=["off", "auto", "on"],
        default="off",
        help="Apply OpenAI DeepResearch marker cleaner",
    )
    preprocess_group.add_argument(
        "--cite-mode",
        choices=["footnote", "inline", "strip"],
        default="footnote",
        help="How to transform cite markers when cleaner is enabled",
    )
    preprocess_group.add_argument(
        "--drop-unknown-markers",
        action="store_true",
        help="Drop unknown DeepResearch marker blocks instead of comment-preserving",
    )
    preprocess_group.add_argument(
        "--cleaner-report",
        action="store_true",
        help="Print DeepResearch cleaner summary",
    )

    parser.add_argument(
        "--profile", choices=list(PROFILES), help="Document type (overrides frontmatter)"
    )
    parser.add_argument("--theme", help="default, mono, or a YAML theme file")
    parser.add_argument("--preset", choices=list(PRESETS), help="Named section bundle (overrides frontmatter)")
    parser.add_argument("--list-presets", action="store_true", help="List section presets")
    chart_group = parser.add_mutually_exclusive_group()
    chart_group.add_argument("--charts", action="store_true", default=None,
                             help="Render chart fences as images (default: code panels)")
    chart_group.add_argument("--no-charts", action="store_false", dest="charts",
                             help="Render chart fences as code panels, overriding frontmatter")
    parser.add_argument("--list-profiles", action="store_true", help="List document profiles")
    parser.add_argument(
        "--strict", action="store_true", default=None, help="Fail on input loss or rendering errors"
    )
    parser.add_argument(
        "--no-confidential",
        action="store_false",
        dest="confidential",
        default=None,
        help="Omit confidentiality label",
    )

    # Section toggles
    section_group = parser.add_argument_group("section options")
    section_group.add_argument(
        "--no-cover",
        action="store_true",
        help="Skip cover page",
    )
    section_group.add_argument(
        "--no-toc",
        action="store_true",
        help="Skip table of contents",
    )
    section_group.add_argument(
        "--no-disclaimer",
        "--no-disc",
        action="store_true",
        dest="no_disclaimer",
        help="Skip disclaimer page",
    )
    section_group.add_argument(
        "--separator-mode",
        choices=["auto", "rule", "page-break"],
        default=None,
        help="Render separators as horizontal rules, page breaks, or auto (`## ---` => page break)",
    )

    # Verbosity
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose (debug) output",
    )

    return parser


def run_conversion(input_path: Path, args) -> int:
    """
    Execute conversion for a single file.

    Args:
        input_path: Resolved input file path
        args: Parsed CLI arguments

    Returns:
        Exit code (0 = success, 1 = error)
    """
    # Build render options
    render_options = RenderOptions(
        include_cover=False if args.no_cover else None,
        include_toc=False if args.no_toc else None,
        include_disclaimer=False if args.no_disclaimer else None,
        separator_mode=getattr(args, "separator_mode", None),
        profile=getattr(args, "profile", None),
        theme=getattr(args, "theme", None),
        strict=getattr(args, "strict", None),
        confidential=getattr(args, "confidential", None),
        charts=getattr(args, "charts", None),
        preset=getattr(args, "preset", None),
    )

    # Auto-format / cleaner if requested
    actual_input = input_path
    if args.format:
        try:
            actual_input = auto_format_markdown(
                input_path,
                cleaner_mode=args.deepresearch_cleaner,
                cite_mode=args.cite_mode,
                drop_unknown_markers=args.drop_unknown_markers,
                cleaner_report=args.cleaner_report,
            )
        except ImportError as e:
            logger.error("%s", e)
            return 1
        except Exception as e:
            logger.error("Auto-format failed: %s", e)
            if args.verbose:
                import traceback

                traceback.print_exc()
            return 1
    elif args.deepresearch_cleaner != "off":
        try:
            actual_input = clean_deepresearch_markdown_file(
                input_path=input_path,
                cleaner_mode=args.deepresearch_cleaner,
                cite_mode=args.cite_mode,
                drop_unknown_markers=args.drop_unknown_markers,
                cleaner_report=args.cleaner_report,
            )
        except ImportError as e:
            logger.error("%s", e)
            return 1
        except Exception as e:
            logger.error("DeepResearch cleaner failed: %s", e)
            if args.verbose:
                import traceback

                traceback.print_exc()
            return 1

    # Convert
    try:
        converter = IBReportConverter(
            md_file_path=str(actual_input),
            output_path=args.output_file,
            render_options=render_options,
        )
        output_path = converter.convert()

        print(f"\n{'=' * 65}")
        print("  Conversion complete!")
        print(f"  Input:  {actual_input}")
        print(f"  Output: {output_path}")
        print(f"{'=' * 65}\n")
        return 0

    except FileNotFoundError as e:
        logger.error("%s", e)
        print("\nTip: Use --list to see available files")
        return 1

    except ValueError as e:
        logger.error("%s", e)
        return 1

    except RuntimeError as e:
        logger.error("%s", e)
        if args.verbose:
            import traceback

            traceback.print_exc()
        return 1

    except Exception as e:
        logger.error("Unexpected error: %s", e)
        if args.verbose:
            import traceback

            traceback.print_exc()
        return 1


def run_batch_conversion(input_dir: Path, args: argparse.Namespace) -> int:
    """Convert a directory after rejecting inputs with colliding destinations.

    Args:
        input_dir: Directory containing Markdown inputs.
        args: Parsed CLI arguments; output_file is an optional output directory.

    Returns:
        Zero if every input succeeds, otherwise one.
    """
    input_files = sorted(input_dir.glob("*.md"))
    if not input_files:
        logger.error("No .md files found in: %s", input_dir)
        return 1

    output_dir = Path(args.output_file) if args.output_file else input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    planned_outputs = {
        source: output_dir / generate_output_path(source).name for source in input_files
    }
    output_sources: Dict[str, List[Path]] = {}
    for source, destination in planned_outputs.items():
        key = os.path.normcase(str(destination.resolve()))
        output_sources.setdefault(key, []).append(source)

    colliding_inputs: Set[Path] = set()
    for sources in output_sources.values():
        if len(sources) > 1:
            colliding_inputs.update(sources)
            logger.error(
                "Batch output collision at %s; inputs will not be converted: %s",
                planned_outputs[sources[0]],
                ", ".join(str(source) for source in sources),
            )

    success_count = 0
    failure_count = len(colliding_inputs)

    for input_file in input_files:
        if input_file in colliding_inputs:
            continue
        batch_args = argparse.Namespace(**vars(args))
        batch_args.output_file = str(planned_outputs[input_file])
        exit_code = run_conversion(input_file, batch_args)
        if exit_code == 0:
            success_count += 1
        else:
            failure_count += 1

    logger.info("[BATCH] Completed: %d succeeded, %d failed", success_count, failure_count)
    return 0 if failure_count == 0 else 1


def main():
    """Main entry point with CLI argument parsing"""
    parser = build_parser()
    args = parser.parse_args()

    # Setup logging
    configure_logging(verbose=args.verbose)

    if args.list_profiles:
        logger.info("Available profiles: %s", ", ".join(PROFILES))
        return

    if args.list_presets:
        logger.info("Available presets: %s", ", ".join(PRESETS))
        return

    # ── List mode ───────────────────────────────────────────────────────────
    if args.list:
        md_files = list_md_files()

        if args.interactive and md_files:
            selected = interactive_select(md_files)
            if selected:
                sys.exit(run_conversion(selected, args))
            else:
                print("No file selected.")
                sys.exit(0)

        if not args.input_file:
            print("Usage: uv run md_to_word.py <filename.md>")
            print("       uv run md_to_word.py --list -i   (interactive)")
            sys.exit(0)

    # ── No input file → show list and exit ──────────────────────────────────
    if args.input_file is None:
        md_files = list_md_files()

        if args.interactive and md_files:
            selected = interactive_select(md_files)
            if selected:
                sys.exit(run_conversion(selected, args))

        print("Usage: uv run md_to_word.py <filename.md>")
        sys.exit(0)

    # ── Resolve input path ──────────────────────────────────────────────────
    input_path = resolve_input_path(args.input_file)
    if not input_path.exists():
        logger.error("File not found: %s", input_path)
        sys.exit(1)

    if input_path.is_dir() or args.batch:
        if not input_path.is_dir():
            logger.error("Batch mode requires a directory input: %s", input_path)
            sys.exit(1)
        sys.exit(run_batch_conversion(input_path, args))

    sys.exit(run_conversion(input_path, args))


if __name__ == "__main__":
    main()
