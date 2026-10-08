"""Run the source plugin with dependencies kept in an external user cache.

Changelog (plugin distribution):
    - Add explicit setup and Codex rendering without modifying the plugin tree.
"""

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional, Sequence

CACHE_ENV = "IB_REPORT_FORMATTER_CACHE_DIR"
LOCK_TIMEOUT_SECONDS = 120.0


def plugin_root() -> Path:
    """Return the source plugin root independently of the caller's directory."""
    return Path(__file__).resolve().parents[1]


def cache_directory(override: Optional[str] = None) -> Path:
    """Resolve the explicitly selected or platform-default external user cache."""
    selected = override or os.environ.get(CACHE_ENV)
    if selected:
        return Path(selected).expanduser().resolve()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local")))
    elif sys.platform == "darwin":
        base = Path.home() / "Library/Caches"
    else:
        base = Path(os.environ.get("XDG_CACHE_HOME", str(Path.home() / ".cache")))
    return (base / "ib-report-formatter").resolve()


def runtime_key(root: Path) -> str:
    """Fingerprint dependency declarations, OS, architecture and interpreter."""
    digest = hashlib.sha256()
    for name in ("pyproject.toml", "uv.lock"):
        digest.update(name.encode("utf-8") + b"\0" + (root / name).read_bytes() + b"\0")
    identity = (
        sys.platform, platform.machine(), sys.implementation.name,
        sys.implementation.cache_tag, platform.python_version(), str(Path(sys.executable).resolve()),
    )
    digest.update(json.dumps(identity).encode("utf-8"))
    return digest.hexdigest()[:24]


def runtime_python(environment: Path) -> Path:
    """Return the virtual environment Python path for the current platform."""
    return environment / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")


@contextmanager
def runtime_lock(directory: Path) -> Iterator[None]:
    """Serialize setup and use; time out with actionable guidance after a stale lock."""
    lock = directory / ".runtime.lock"
    deadline = time.monotonic() + LOCK_TIMEOUT_SECONDS
    while True:
        try:
            lock.mkdir()
            break
        except FileExistsError:
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"Runtime is busy: {lock}. Retry after the other job finishes. "
                    "If its process has stopped, remove this empty lock directory and retry."
                ) from None
            time.sleep(0.2)
    try:
        yield
    finally:
        lock.rmdir()


def runtime_environment(cache: Path, environment: Path) -> dict[str, str]:
    """Route dependency, bytecode and plotting caches away from plugin files."""
    env = os.environ.copy()
    env.update({
        "UV_PROJECT_ENVIRONMENT": str(environment),
        "UV_CACHE_DIR": str(cache / "uv-cache"),
        "UV_PYTHON_INSTALL_DIR": str(cache / "python"),
        "PYTHONDONTWRITEBYTECODE": "1",
        "PYTHONPYCACHEPREFIX": str(cache / "pycache"),
        "MPLCONFIGDIR": str(cache / "matplotlib"),
    })
    return env


def execute(arguments: argparse.Namespace) -> int:
    """Ensure locked dependencies and optionally delegate a strict document render.

    Args:
        arguments: Parsed setup/render options.

    Returns:
        Zero for successful setup, or the rendering process's exit status.
    """
    root = plugin_root()
    cache = cache_directory(arguments.cache_dir)
    if cache.is_relative_to(root):
        raise ValueError("Cache directory must be outside the plugin installation")
    if arguments.command == "render" and Path(arguments.output_dir).resolve().is_relative_to(root):
        raise ValueError("Output directory must be outside the plugin installation")
    uv = shutil.which("uv")
    if uv is None:
        raise RuntimeError(
            "uv was not found on PATH. Install uv using https://docs.astral.sh/uv/getting-started/installation/ "
            "and rerun setup. This launcher does not install uv system-wide."
        )
    cache.mkdir(parents=True, exist_ok=True)
    runtime = cache / "runtimes" / runtime_key(root)
    runtime.mkdir(parents=True, exist_ok=True)
    environment = runtime / "venv"
    env = runtime_environment(cache, environment)
    # No ready stamp is trusted: uv checks the lock each time. Hold the lock through
    # rendering so another sync cannot change an environment that is being used.
    with runtime_lock(runtime):
        synced = subprocess.run(
            [uv, "sync", "--locked", "--no-dev", "--extra", "full", "--no-install-project",
             "--project", str(root), "--python", sys.executable],
            cwd=cache, env=env, stdout=sys.stderr, stderr=sys.stderr, check=False,
        )
        if synced.returncode:
            raise RuntimeError(f"Dependency setup failed (uv exit {synced.returncode}); rerun setup after correcting the error above")
        python = runtime_python(environment)
        if not python.is_file():
            raise RuntimeError(f"Dependency setup did not create the runtime Python: {python}")
        if arguments.command == "setup":
            print(json.dumps({
                "ok": True, "status": "ready", "cache_dir": str(cache),
                "runtime": {"path": str(environment), "python": str(python)},
            }, ensure_ascii=True))
            return 0
        command = [str(python), "-B", str(root / "scripts/agent_render.py"), arguments.source,
                   "--output-dir", arguments.output_dir]
        if arguments.profile is not None:
            command.extend(["--profile", arguments.profile])
        if arguments.expected_terms is not None:
            command.extend(["--expected-terms", arguments.expected_terms])
        # Deliberately inherit the user's cwd: relative input, expected-terms and
        # output paths must retain the same meaning as the original invocation.
        return subprocess.run(command, env=env, check=False).returncode


class _JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValueError(message)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Parse setup/render arguments and return actionable JSON on runtime failure."""
    parser = _JsonArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    setup = commands.add_parser("setup", help="Install locked dependencies into an external cache")
    render = commands.add_parser("render", help="Ensure dependencies and render using the shared engine")
    for command in (setup, render):
        command.add_argument("--cache-dir", help=f"External writable cache (or {CACHE_ENV})")
    render.add_argument("source")
    render.add_argument("--output-dir", required=True)
    render.add_argument("--profile")
    render.add_argument("--expected-terms")
    try:
        arguments = parser.parse_args(argv)
        if sys.version_info < (3, 12):  # noqa: UP036 - bootstrap may use an older interpreter
            raise RuntimeError("Python 3.12 or newer is required; run with uv run --no-project --python 3.12 python")
        return execute(arguments)
    except Exception as exc:
        print(json.dumps({
            "ok": False, "status": "failed", "stage": "runtime",
            "diagnostics": [{"code": type(exc).__name__, "message": str(exc)}],
        }, ensure_ascii=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
