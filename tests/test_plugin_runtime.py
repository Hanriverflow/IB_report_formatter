"""External-cache and process-boundary tests for the installable plugin launcher."""

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from scripts import plugin_runtime as launcher


@pytest.fixture
def plugin(tmp_path, monkeypatch):
    root = tmp_path / "installed plugin with spaces"
    root.mkdir()
    (root / "pyproject.toml").write_text('[project]\nname = "test"\n', encoding="utf-8")
    (root / "uv.lock").write_text("version = 1\n", encoding="utf-8")
    scripts = root / "scripts"
    scripts.mkdir()
    (scripts / "agent_render.py").write_text("# delegated entry point\n", encoding="utf-8")
    monkeypatch.setattr(launcher, "plugin_root", lambda: root)
    monkeypatch.setattr(launcher.shutil, "which", lambda _: "/tools with spaces/uv")
    monkeypatch.delenv(launcher.CACHE_ENV, raising=False)
    return root


def snapshot(root):
    return {str(path.relative_to(root)): path.read_bytes() for path in root.rglob("*") if path.is_file()}


def fake_success(monkeypatch, render_code=0):
    calls = []

    def run(command, **kwargs):
        calls.append((command, kwargs))
        if "sync" in command:
            environment = Path(kwargs["env"]["UV_PROJECT_ENVIRONMENT"])
            python = launcher.runtime_python(environment)
            python.parent.mkdir(parents=True, exist_ok=True)
            python.touch()
            return SimpleNamespace(returncode=0)
        return SimpleNamespace(returncode=render_code)

    monkeypatch.setattr(launcher.subprocess, "run", run)
    return calls


def test_setup_syncs_locked_dependencies_outside_plugin(plugin, tmp_path, monkeypatch, capsys):
    before = snapshot(plugin)
    calls = fake_success(monkeypatch)
    cache = tmp_path / "external cache"
    assert launcher.main(["setup", "--cache-dir", str(cache)]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["ok"] and result["status"] == "ready"
    command, options = calls[0]
    assert command == [
        "/tools with spaces/uv", "sync", "--locked", "--no-dev", "--extra", "full",
        "--no-install-project", "--project", str(plugin), "--python", launcher.sys.executable,
    ]
    assert options["cwd"] == cache
    assert options.get("shell", False) is False
    env = options["env"]
    for name in ("UV_PROJECT_ENVIRONMENT", "UV_CACHE_DIR", "UV_PYTHON_INSTALL_DIR",
                 "PYTHONPYCACHEPREFIX", "MPLCONFIGDIR"):
        assert Path(env[name]).is_relative_to(cache)
        assert not Path(env[name]).is_relative_to(plugin)
    assert env["PYTHONDONTWRITEBYTECODE"] == "1"
    assert snapshot(plugin) == before
    assert not (plugin / ".venv").exists()
    assert not list(cache.rglob(".runtime.lock"))


def test_render_preserves_caller_cwd_and_exact_arguments(plugin, tmp_path, monkeypatch):
    caller = tmp_path / "사용자 작업 폴더"
    caller.mkdir()
    monkeypatch.chdir(caller)
    calls = fake_success(monkeypatch, render_code=7)
    result = launcher.main([
        "render", "입력 자료.md", "--output-dir", "새 결과", "--profile", "term-sheet",
        "--expected-terms", "조건 정보.json", "--cache-dir", str(tmp_path / "cache"),
    ])
    assert result == 7
    command, options = calls[1]
    assert command[1:] == [
        "-B", str(plugin / "scripts/agent_render.py"), "입력 자료.md", "--output-dir", "새 결과",
        "--profile", "term-sheet", "--expected-terms", "조건 정보.json",
    ]
    assert "cwd" not in options
    assert Path.cwd() == caller
    assert options.get("shell", False) is False


def test_environment_override_and_explicit_precedence(plugin, tmp_path, monkeypatch):
    monkeypatch.setenv(launcher.CACHE_ENV, str(tmp_path / "env cache"))
    assert launcher.cache_directory() == tmp_path / "env cache"
    assert launcher.cache_directory(str(tmp_path / "explicit")) == tmp_path / "explicit"


def test_runtime_fingerprint_tracks_lock_project_and_interpreter(plugin, monkeypatch):
    initial = launcher.runtime_key(plugin)
    (plugin / "uv.lock").write_text("version = 2\n", encoding="utf-8")
    changed_lock = launcher.runtime_key(plugin)
    assert changed_lock != initial
    (plugin / "pyproject.toml").write_text("# another project\n", encoding="utf-8")
    changed_project = launcher.runtime_key(plugin)
    assert changed_project != changed_lock
    monkeypatch.setattr(launcher.platform, "machine", lambda: "other-architecture")
    assert launcher.runtime_key(plugin) != changed_project


def test_missing_uv_is_actionable_without_install_attempt(plugin, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(launcher.shutil, "which", lambda _: None)
    calls = fake_success(monkeypatch)
    assert launcher.main(["setup", "--cache-dir", str(tmp_path / "cache")]) == 1
    result = json.loads(capsys.readouterr().out)
    assert "Install uv" in result["diagnostics"][0]["message"]
    assert result["stage"] == "runtime"
    assert calls == []


def test_install_failure_leaves_no_ready_stamp_and_releases_lock(plugin, tmp_path, monkeypatch, capsys):
    cache = tmp_path / "cache"
    monkeypatch.setattr(launcher.subprocess, "run", lambda *args, **kwargs: SimpleNamespace(returncode=8))
    assert launcher.main(["setup", "--cache-dir", str(cache)]) == 1
    assert "uv exit 8" in json.loads(capsys.readouterr().out)["diagnostics"][0]["message"]
    assert not list(cache.rglob("*ready*"))
    assert not list(cache.rglob(".runtime.lock"))


def test_setup_never_trusts_an_old_stamp(plugin, tmp_path, monkeypatch, capsys):
    calls = fake_success(monkeypatch)
    args = ["setup", "--cache-dir", str(tmp_path / "cache")]
    assert launcher.main(args) == 0
    capsys.readouterr()
    assert launcher.main(args) == 0
    assert len(calls) == 2


def test_runtime_busy_timeout_preserves_other_jobs_lock(plugin, tmp_path, monkeypatch, capsys):
    cache = tmp_path / "cache"
    runtime = cache / "runtimes" / launcher.runtime_key(plugin)
    lock = runtime / ".runtime.lock"
    lock.mkdir(parents=True)
    monkeypatch.setattr(launcher, "LOCK_TIMEOUT_SECONDS", 0)
    calls = fake_success(monkeypatch)
    assert launcher.main(["setup", "--cache-dir", str(cache)]) == 1
    result = json.loads(capsys.readouterr().out)
    assert result["diagnostics"][0]["code"] == "TimeoutError"
    assert lock.is_dir()
    assert calls == []


@pytest.mark.parametrize("relative", [".", "cache"])
def test_cache_inside_plugin_is_rejected(plugin, monkeypatch, capsys, relative):
    before = snapshot(plugin)
    calls = fake_success(monkeypatch)
    assert launcher.main(["setup", "--cache-dir", str(plugin / relative)]) == 1
    assert "outside the plugin" in json.loads(capsys.readouterr().out)["diagnostics"][0]["message"]
    assert snapshot(plugin) == before
    assert calls == []


def test_invalid_arguments_are_structured(capsys):
    assert launcher.main(["render", "draft.md"]) == 1
    assert json.loads(capsys.readouterr().out)["stage"] == "runtime"


@pytest.mark.parametrize("output", [".", "new output"])
def test_output_inside_plugin_is_rejected_before_sync(plugin, tmp_path, monkeypatch, capsys, output):
    monkeypatch.chdir(plugin)
    before = snapshot(plugin)
    calls = fake_success(monkeypatch)
    result = launcher.main([
        "render", "draft.md", "--output-dir", output, "--cache-dir", str(tmp_path / "cache"),
    ])
    assert result == 1
    assert "Output directory must be outside" in json.loads(capsys.readouterr().out)["diagnostics"][0]["message"]
    assert snapshot(plugin) == before
    assert calls == []


def test_success_without_python_is_not_reported_ready(plugin, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(launcher.subprocess, "run", lambda *args, **kwargs: SimpleNamespace(returncode=0))
    assert launcher.main(["setup", "--cache-dir", str(tmp_path / "cache")]) == 1
    assert "did not create" in json.loads(capsys.readouterr().out)["diagnostics"][0]["message"]


def test_lock_released_if_process_launch_raises(plugin, tmp_path, monkeypatch, capsys):
    def fail(*args, **kwargs):
        raise OSError("simulated process failure")

    monkeypatch.setattr(launcher.subprocess, "run", fail)
    cache = tmp_path / "cache"
    assert launcher.main(["setup", "--cache-dir", str(cache)]) == 1
    assert json.loads(capsys.readouterr().out)["diagnostics"][0]["code"] == "OSError"
    assert not list(cache.rglob(".runtime.lock"))


def test_platform_specific_python_layout(monkeypatch, tmp_path):
    monkeypatch.setattr(launcher.sys, "platform", "linux")
    assert launcher.runtime_python(tmp_path) == tmp_path / "bin/python"
    monkeypatch.setattr(launcher.sys, "platform", "win32")
    assert launcher.runtime_python(tmp_path) == tmp_path / "Scripts/python.exe"
