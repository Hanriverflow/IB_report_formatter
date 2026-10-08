---
name: setup
description: Prepare or diagnose the external Python runtime for the IB Report Formatter plugin in Claude Code or Codex, without modifying the installed plugin or requiring a project-specific model API key.
---

# Set up IB Report Formatter

Resolve the plugin root from this loaded `SKILL.md`'s actual absolute path: the directory containing `skills` is the plugin root (`Path(skill_file).resolve().parents[2]`). Confirm its `scripts/plugin_runtime.py`, `pyproject.toml` and `uv.lock` exist. The user's working directory is not necessarily the plugin root. Use the verified path on both hosts; do not require `CLAUDE_PLUGIN_ROOT`.

Check `uv --version`. If unavailable, report that concrete prerequisite and use the host's available approved installation method when authorized; otherwise provide [uv's official installation instructions](https://docs.astral.sh/uv/getting-started/installation/). Do not claim setup succeeded or silently replace the runtime mechanism. The launcher reports a missing uv executable; it does not install uv system-wide.

Run, replacing the placeholder with the real path and quoting for the active shell:

```sh
uv run --no-project --python 3.12 python "/absolute/plugin/scripts/plugin_runtime.py" setup
```

`uv` supplies Python 3.12 when necessary. Initial downloads need network access. The launcher prepares the locked engine dependencies in an external cache and returns JSON. It keeps dependency/Python caches and matplotlib configuration outside the plugin. Do not run `uv sync` in the plugin, edit the user's project dependencies, or write real documents into the installation.

For a user-requested custom cache, use an absolute directory outside the plugin:

```sh
uv run --no-project --python 3.12 python "/absolute/plugin/scripts/plugin_runtime.py" setup --cache-dir "/absolute/runtime-cache"
```

`IB_REPORT_FORMATTER_CACHE_DIR` can select that cache as well. Pass the same cache choice on later render commands; `--cache-dir` belongs after the `setup` or `render` subcommand. Never clear an unrelated cache or alter global settings to repair a local error.

Check the exit status and JSON. Success has `status: "ready"` and a runtime path; failure has `ok: false`, `stage: "runtime"` and diagnostics. Report the actual cache/runtime location and actionable failure details. A successful setup verifies environment preparation, not document rendering or host plugin discovery. Do not claim the plugin has been installed in either application merely because this command ran.

After successful setup, continue the user's requested document work using [ib-document](../ib-document/SKILL.md). For a setup-only request, stop after reporting the result. The current host account supplies AI; no separate project API key is needed. Do not promise support or verified execution in Claude web or native Cowork.
