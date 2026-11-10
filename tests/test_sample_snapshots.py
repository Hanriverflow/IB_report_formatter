"""B-6 snapshots of the shared strict corpus: complete models and all XML parts.

Normal runs NEVER create missing snapshots. Explicit update command (PowerShell):

    uv run --no-sync pytest tests/test_sample_snapshots.py --snapshot-update -q -p no:cacheprovider --basetemp "$env:LOCALAPPDATA/Temp/ib-t8-update-$([guid]::NewGuid().ToString('N'))"

On Ubuntu use a fresh absolute OS temporary directory for --basetemp instead.
Review every changed model/XML file after updating. No binary DOCX parts are
snapshotted. A separate inventory snapshot catches deleted/renamed XML parts.

The session fixture reuses scripts.sample_corpus's definition AND strict render
path; it renders each case once for all snapshot assertions in this module.
"""

import json

import pytest
from syrupy.assertion import SnapshotAssertion

from document_profiles import resolve_options
from md_parser import parse_markdown_file
from scripts.sample_corpus import PROJECT_ROOT, SampleCase, render_corpus, sample_cases
from tests.snapshot_helpers import (
    JSONTextSnapshotExtension,
    XMLTextSnapshotExtension,
    model_json,
    xml_parts,
)

CASES = sample_cases()


@pytest.fixture(scope="session")
def corpus_snapshots(
    tmp_path_factory: pytest.TempPathFactory,
) -> dict[SampleCase, tuple[str, dict[str, str]]]:
    """Cache source models and one strict rendering per shared corpus case."""
    models = {}
    copyright_companies = {}
    for case in CASES:
        model = parse_markdown_file(
            str(PROJECT_ROOT / case.source),
            profile=case.options.profile,
            asset_roots=case.options.asset_roots,
        )
        models[case] = model_json(model, (PROJECT_ROOT / case.source).parent)
        resolved = resolve_options(model.metadata, case.options)
        copyright_companies[case] = model.metadata.company if resolved.disclaimer else None
    outputs = render_corpus(tmp_path_factory.mktemp("sample-snapshots") / "rendered")
    return {
        case: (models[case], xml_parts(output, copyright_company=copyright_companies[case]))
        for case, output in zip(CASES, outputs, strict=True)
    }


@pytest.mark.parametrize(
    "case",
    CASES,
    ids=lambda case: case.output_relative.with_suffix("").as_posix().replace("/", "--"),
)
def test_sample_model(
    case: SampleCase,
    corpus_snapshots: dict[SampleCase, tuple[str, dict[str, str]]],
    snapshot: SnapshotAssertion,
) -> None:
    """Keep every parsed field, in order, without absolute checkout paths."""
    model, _ = corpus_snapshots[case]
    assert model == snapshot(extension_class=JSONTextSnapshotExtension)


@pytest.mark.parametrize(
    "case",
    CASES,
    ids=lambda case: case.output_relative.with_suffix("").as_posix().replace("/", "--"),
)
def test_sample_xml(
    case: SampleCase,
    corpus_snapshots: dict[SampleCase, tuple[str, dict[str, str]]],
    snapshot: SnapshotAssertion,
) -> None:
    """Detect part addition/removal as well as changes inside any XML part."""
    _, parts = corpus_snapshots[case]
    assert json.dumps(list(parts), indent=2) + "\n" == snapshot(
        extension_class=JSONTextSnapshotExtension, name="parts"
    )
    for part, xml in parts.items():
        assert xml == snapshot(
            extension_class=XMLTextSnapshotExtension, name=part.replace("/", "--")
        ), part
