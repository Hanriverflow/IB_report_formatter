"""
Shared CLI helpers for converter entry points.
"""

import logging
import sys
import time
from pathlib import Path
from typing import Callable, List, Optional, Sequence


class LogFormatter(logging.Formatter):
    """Format CLI logs with a compact level prefix."""

    PREFIXES = {
        logging.DEBUG: "[DEBUG]",
        logging.INFO: "[INFO]",
        logging.WARNING: "[WARNING]",
        logging.ERROR: "[ERROR]",
        logging.CRITICAL: "[CRITICAL]",
    }

    def format(self, record: logging.LogRecord) -> str:
        prefix = self.PREFIXES.get(record.levelno, "[LOG]")
        return f"{prefix} {record.getMessage()}"


def setup_logging(verbose: bool = False) -> None:
    """
    Configure root logging for CLI tools.

    Using the root logger lets imported helper modules emit logs without
    every caller wiring handlers manually.
    """

    level = logging.DEBUG if verbose else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(LogFormatter())

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.handlers.clear()
    root_logger.addHandler(handler)


def resolve_input_path(input_file: str, parent_dir: Path, script_path: Path) -> Path:
    """Resolve an input path using the project's standard search order."""
    input_path = Path(input_file)

    if input_path.is_absolute():
        return input_path

    parent_path = parent_dir / input_path.name
    if parent_path.exists():
        return parent_path

    cwd_path = Path.cwd() / input_file
    if cwd_path.exists():
        return cwd_path

    script_dir = script_path.resolve().parent
    script_candidate = script_dir / input_file
    if script_candidate.exists():
        return script_candidate

    return parent_path


def generate_output_path(
    input_path: Path,
    output_path: Optional[str],
    suffix: str,
    default_name: Optional[Callable[[Path], str]] = None,
) -> Path:
    """Generate an output path while preserving Python 3.8 compatibility."""
    if output_path:
        out = Path(output_path)
        if out.suffix.lower() != suffix.lower():
            out = out.with_suffix(suffix)
        return out

    if default_name is not None:
        return input_path.with_name(default_name(input_path))

    return input_path.with_suffix(suffix)


def safe_save(
    output_path: Path,
    save_action: Callable[[Path], None],
    logger: logging.Logger,
    lock_message: str,
) -> Path:
    """Save output, falling back to a timestamped filename on permission errors."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        save_action(output_path)
        logger.info("Saved: %s", output_path)
        return output_path
    except PermissionError:
        logger.warning(lock_message, output_path.name)
        timestamp = int(time.time())
        new_name = f"{output_path.stem}_{timestamp}{output_path.suffix}"
        new_path = output_path.with_name(new_name)
        save_action(new_path)
        logger.info("Saved: %s", new_path)
        return new_path


def list_files(parent_dir: Path, pattern: str, file_label: str, title_label: str) -> List[Path]:
    """Print a project-local file listing and return the discovered files."""
    files = sorted(parent_dir.glob(pattern))

    if not files:
        print(f"No {file_label} files found in parent folder.")
        print(f"  Searched: {parent_dir}")
        return []

    print(f"\n{'=' * 65}")
    print(f"  Available {title_label} files in: {parent_dir}")
    print(f"{'=' * 65}")

    for index, file_path in enumerate(files, 1):
        size_kb = file_path.stat().st_size / 1024
        print(f"  [{index:2d}]  {file_path.name:<40s}  ({size_kb:>7.1f} KB)")

    print(f"{'=' * 65}")
    print(f"  Total: {len(files)} file(s)")
    print(f"{'=' * 65}")
    print()

    return files


def interactive_select(files: Sequence[Path], prompt: str) -> Optional[Path]:
    """Prompt the user to select a file by number."""
    if not files:
        return None

    print(prompt, end="")

    try:
        user_input = input().strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return None

    if user_input.lower() in ("q", "quit", "exit", ""):
        return None

    try:
        index = int(user_input)
    except ValueError:
        print(f"Invalid input: '{user_input}'. Enter a number.")
        return None

    if 1 <= index <= len(files):
        return files[index - 1]

    print(f"Invalid number. Choose between 1 and {len(files)}.")
    return None
