"""
Word to Markdown Converter
Main entry point for converting Word documents to Markdown.

Usage:
    uv run word_to_md.py input.docx [output.md]
    uv run word_to_md.py input.docx --strip           # Remove formatting for LLM
    uv run word_to_md.py input.docx --no-frontmatter  # Skip YAML metadata
    uv run word_to_md.py input.docx --extract-images  # Save images to folder
"""

import argparse
import logging
import sys
import time
from pathlib import Path
from typing import List, Optional

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
from md_renderer import render_to_markdown
from word_parser import parse_word_file

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Parent folder path (where files are typically located)
PARENT_DIR = Path(__file__).resolve().parent.parent

# Output suffix
OUTPUT_SUFFIX = ".md"

# Logger
logger = logging.getLogger("word_to_md")


# ═══════════════════════════════════════════════════════════════════════════════
def resolve_input_path(input_file: str) -> Path:
    """Resolve the user-provided input file path."""
    return resolve_project_input_path(input_file, parent_dir=PARENT_DIR, script_path=Path(__file__))


def generate_output_path(input_path: Path, output_path: Optional[str] = None) -> Path:
    """Generate the markdown output path."""
    return build_output_path(input_path, output_path, OUTPUT_SUFFIX)


def safe_save(content: str, output_path: Path) -> Path:
    """Save markdown content, falling back to a timestamped path if needed."""
    def write_content(path: Path) -> None:
        path.write_text(content, encoding="utf-8")

    return safe_save_with_fallback(
        output_path=output_path,
        save_action=write_content,
        logger=logger,
        lock_message="%s is locked. Saving with timestamp suffix.",
    )


# ═══════════════════════════════════════════════════════════════════════════════
# CONVERTER
# ═══════════════════════════════════════════════════════════════════════════════


class WordToMarkdownConverter:
    """Main converter class orchestrating Word to Markdown conversion."""

    def __init__(
        self,
        docx_file_path: str,
        output_path: Optional[str] = None,
        strip_formatting: bool = False,
        include_frontmatter: bool = True,
        extract_images: bool = False,
        image_output_dir: Optional[str] = None,
    ):
        self.docx_file_path = Path(docx_file_path).resolve()
        self.output_path = output_path
        self.strip_formatting = strip_formatting
        self.include_frontmatter = include_frontmatter
        self.extract_images = extract_images
        self.image_output_dir = image_output_dir
        self._validate_input()

    def _validate_input(self):
        """Validate that the input file exists and is readable."""
        if not self.docx_file_path.exists():
            raise FileNotFoundError(f"Input file not found: {self.docx_file_path}")
        if not self.docx_file_path.is_file():
            raise FileNotFoundError(f"Not a file: {self.docx_file_path}")
        if self.docx_file_path.suffix.lower() != ".docx":
            logger.warning("Input file may not be a Word document: %s", self.docx_file_path.name)

    def convert(self) -> str:
        """Execute the full conversion pipeline."""
        start_time = time.perf_counter()

        # Stage 1: Parse Word document
        logger.info("Parsing: %s", self.docx_file_path.name)

        image_dir = None
        if self.extract_images:
            image_dir = self.image_output_dir or f"{self.docx_file_path.stem}_images"

        model = parse_word_file(
            str(self.docx_file_path),
            extract_images=self.extract_images,
            image_output_dir=image_dir,
        )

        logger.info("Parsed %d elements", len(model.elements))
        logger.info("Title: %s", model.metadata.title)

        # Stage 2: Render to Markdown
        logger.info("Rendering Markdown...")

        markdown = render_to_markdown(
            model,
            include_frontmatter=self.include_frontmatter,
            strip_formatting=self.strip_formatting,
            image_path_prefix=image_dir if image_dir else "",
        )

        # Stage 3: Save output
        output_path = generate_output_path(self.docx_file_path, self.output_path)
        saved_path = safe_save(markdown, output_path)

        elapsed = time.perf_counter() - start_time
        logger.info("Conversion completed in %.2fs", elapsed)

        return str(saved_path)


def run_conversion(input_path: Path, args) -> int:
    """Execute conversion for a single file."""
    try:
        converter = WordToMarkdownConverter(
            docx_file_path=str(input_path),
            output_path=args.output_file,
            strip_formatting=args.strip,
            include_frontmatter=not args.no_frontmatter,
            extract_images=args.extract_images,
            image_output_dir=args.image_dir,
        )
        output_path = converter.convert()

        print(f"\n{'=' * 65}")
        print("  Conversion complete!")
        print(f"  Input:  {input_path}")
        print(f"  Output: {output_path}")
        print(f"{'=' * 65}\n")
        return 0

    except FileNotFoundError as e:
        logger.error("%s", e)
        print("\nTip: Use --list to see available files")
        return 1
    except Exception as e:
        logger.error("Conversion failed: %s", e)
        if args.verbose:
            import traceback

            traceback.print_exc()
        return 1


# ═══════════════════════════════════════════════════════════════════════════════
# LIST DOCX FILES
# ═══════════════════════════════════════════════════════════════════════════════


def list_docx_files() -> List[Path]:
    """List all Word documents in the parent directory."""
    return list_files(PARENT_DIR, "*.docx", ".docx", "DOCX")


def interactive_select(docx_files: List[Path]) -> Optional[Path]:
    """Prompt user to select a file by number."""
    return prompt_for_selection(docx_files, "Enter file number to convert (or 'q' to quit): ")


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════


def build_parser() -> argparse.ArgumentParser:
    """Build CLI argument parser"""
    parser = argparse.ArgumentParser(
        description="Convert Word documents to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    uv run word_to_md.py --list                        # List available docx files
    uv run word_to_md.py --list -i                     # Interactive selection
    uv run word_to_md.py report.docx                   # Convert to report.md
    uv run word_to_md.py report.docx --strip           # Plain text for LLM
    uv run word_to_md.py report.docx --extract-images  # Save images
        """,
    )

    parser.add_argument(
        "input_file",
        nargs="?",
        help="Input Word file (filename or path)",
    )

    parser.add_argument(
        "output_file",
        nargs="?",
        help="Output Markdown path (optional, auto-generated if omitted)",
    )

    # Mode flags
    mode_group = parser.add_argument_group("mode options")
    mode_group.add_argument(
        "-l",
        "--list",
        action="store_true",
        help="List available docx files in parent folder",
    )
    mode_group.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Interactive file selection (use with --list)",
    )

    # Conversion options
    conv_group = parser.add_argument_group("conversion options")
    conv_group.add_argument(
        "-s",
        "--strip",
        action="store_true",
        help="Strip formatting (bold, italic, etc.) for clean LLM input",
    )
    conv_group.add_argument(
        "--no-frontmatter",
        action="store_true",
        help="Do not include YAML frontmatter in output",
    )
    conv_group.add_argument(
        "--extract-images",
        action="store_true",
        help="Extract and save images to local directory",
    )
    conv_group.add_argument(
        "--image-dir",
        type=str,
        help="Directory to save extracted images (default: ./images)",
    )

    # Verbosity
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose (debug) output",
    )

    return parser


def main():
    """Main entry point with CLI argument parsing"""
    parser = build_parser()
    args = parser.parse_args()

    # Setup logging
    configure_logging(verbose=args.verbose)

    # List mode
    if args.list:
        docx_files = list_docx_files()
        if args.interactive and docx_files:
            selected = interactive_select(docx_files)
            if selected:
                sys.exit(run_conversion(selected, args))
        sys.exit(0)

    # Resolve paths
    if args.input_file is None:
        list_docx_files()
        print("Usage: uv run word_to_md.py <filename.docx>")
        sys.exit(0)

    input_path = resolve_input_path(args.input_file)
    if not input_path.exists():
        logger.error("File not found: %s", input_path)
        sys.exit(1)

    # Execute conversion
    sys.exit(run_conversion(input_path, args))


if __name__ == "__main__":
    main()
