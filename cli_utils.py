"""
Shared CLI helpers for converter entry points.

Changelog (hardening):
    - Save through sibling temporary files to preserve existing reports on failure.
    - Reserve locked-file fallback names exclusively, including concurrent saves.
    - Limit project-directory input fallback to bare filenames.
"""

import logging
import os
import stat
import sys
import tempfile
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


def _configure_text_stream(stream) -> None:
    """Prefer UTF-8 console output when the current stream supports reconfiguration."""
    reconfigure = getattr(stream, "reconfigure", None)
    if not callable(reconfigure):
        return

    try:
        reconfigure(encoding="utf-8", errors="replace")
    except (ValueError, OSError):
        # Some redirected streams reject runtime reconfiguration. Best effort only.
        return


def setup_logging(verbose: bool = False) -> None:
    """
    Configure root logging for CLI tools.

    Using the root logger lets imported helper modules emit logs without
    every caller wiring handlers manually.
    """

    _configure_text_stream(sys.stdout)
    _configure_text_stream(sys.stderr)

    level = logging.DEBUG if verbose else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(LogFormatter())

    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.handlers.clear()
    root_logger.addHandler(handler)


def resolve_input_path(input_file: str, parent_dir: Path, script_path: Path) -> Path:
    """Resolve explicit paths as given and search project locations for bare names.

    Args:
        input_file: User-provided filename or path.
        parent_dir: Fallback directory for bare filenames.
        script_path: Entry point whose directory is the final bare-name fallback.

    Returns:
        Resolved candidate, which may not exist and must be checked by the caller.
    """
    input_path = Path(input_file)

    if input_path.is_absolute():
        return input_path

    cwd_path = Path.cwd() / input_file
    if cwd_path.exists() or input_file != input_path.name:
        return cwd_path

    parent_path = parent_dir / input_path.name
    if parent_path.exists():
        return parent_path

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


def _target_mode(output_path: Path) -> int:
    """Return the existing destination mode, or the umask default for a new file.

    Args:
        output_path: Destination that may or may not exist.

    Returns:
        Permission bits to apply before replacing the destination.
    """
    try:
        return stat.S_IMODE(output_path.stat().st_mode)
    except FileNotFoundError:
        umask = os.umask(0)
        os.umask(umask)
        return 0o666 & ~umask


def _atomic_save(output_path: Path, save_action: Callable[[Path], None]) -> None:
    """Serialize beside the destination and replace it only after success.

    Args:
        output_path: Destination in an existing directory.
        save_action: Callback that serializes the complete output to a path.
    """
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_", suffix=output_path.suffix, dir=str(output_path.parent)
    )
    temp_path = Path(temp_name)
    try:
        os.close(fd)
        save_action(temp_path)
        # mkstemp creates 0600 files; keep the mode a direct save would have produced.
        os.chmod(temp_path, _target_mode(output_path))
        os.replace(temp_path, output_path)
    finally:
        temp_path.unlink(missing_ok=True)


def _reserve_fallback_path(output_path: Path) -> Path:
    """Exclusively reserve an unused timestamped sibling destination.

    Args:
        output_path: Original destination whose stem and extension are retained.

    Returns:
        Path to an empty file owned by this save operation.
    """
    timestamp = int(time.time())
    counter = 0
    while True:
        suffix = f"_{timestamp}" if counter == 0 else f"_{timestamp}_{counter}"
        candidate = output_path.with_name(f"{output_path.stem}{suffix}{output_path.suffix}")
        try:
            with candidate.open("xb"):
                return candidate
        except FileExistsError:
            counter += 1


def safe_save(
    output_path: Path,
    save_action: Callable[[Path], None],
    logger: logging.Logger,
    lock_message: str,
) -> Path:
    """Save atomically, falling back to a timestamped path on permission errors.

    Args:
        output_path: Requested destination.
        save_action: Callback that serializes the complete output to a path.
        logger: Logger for save and lock messages.
        lock_message: Warning format with one filename placeholder.

    Returns:
        Path containing the successfully saved output.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        _atomic_save(output_path, save_action)
        logger.info("Saved: %s", output_path)
        return output_path
    except PermissionError:
        logger.warning(lock_message, output_path.name)
        new_path = _reserve_fallback_path(output_path)
        try:
            _atomic_save(new_path, save_action)
        except BaseException:
            new_path.unlink(missing_ok=True)
            raise
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
