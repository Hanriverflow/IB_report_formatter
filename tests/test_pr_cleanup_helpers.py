"""Regression coverage for the narrow helper cleanup during PR promotion."""

from pathlib import Path
from typing import get_args, get_type_hints

import pytest

import diagram_renderer
from stream_utils import normalize_source


def test_normalize_source_annotations_can_be_resolved():
    assert Path in get_args(get_type_hints(normalize_source)["source"])
    assert normalize_source(Path("report.md")) == "report.md"


@pytest.mark.parametrize("available", [True, False])
def test_matplotlib_availability_probe_retains_optional_dependency_behavior(monkeypatch, available):
    requested = []

    def probe(name):
        requested.append(name)
        if not available:
            raise ImportError("optional dependency unavailable")
        return object()

    monkeypatch.setattr(diagram_renderer.importlib, "import_module", probe)
    assert diagram_renderer._matplotlib_available() is available
    assert requested == ["matplotlib"]
