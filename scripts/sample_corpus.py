"""Shared strict sample corpus for schema validation and future snapshots."""

from dataclasses import dataclass
from pathlib import Path

from document_profiles import RenderOptions
from md_to_word import IBReportConverter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DIRECTORIES = ("profiles", "qa", "distribution")
CHARTS_SAMPLE = Path("samples/qa/charts.md")


@dataclass(frozen=True)
class SampleCase:
    """A repository-relative source and its explicit render override."""

    source: Path
    charts: bool | None = None

    @property
    def options(self) -> RenderOptions:
        """Preserve source settings except strict mode and the extra charts run."""
        return RenderOptions(strict=True, charts=self.charts)

    @property
    def output_relative(self) -> Path:
        """Keep default and charts-on outputs in separate namespaces."""
        variant = "charts-on" if self.charts else "default"
        return Path(variant) / self.source.relative_to("samples").with_suffix(".docx")


def sample_cases(root: Path = PROJECT_ROOT) -> tuple[SampleCase, ...]:
    """Discover every Markdown sample, then add the charts-on variant.

    Args:
        root: Source checkout containing the three sample directories.

    Returns:
        Stable, relative-path-ordered cases with no profile overrides.
    """
    sources: list[Path] = []
    for directory in SAMPLE_DIRECTORIES:
        found = sorted((root / "samples" / directory).rglob("*.md"))
        if not found:
            raise ValueError(f"No Markdown samples found in samples/{directory}")
        sources.extend(path.relative_to(root) for path in found)
    if CHARTS_SAMPLE not in sources:
        raise ValueError(f"Missing charts sample: {CHARTS_SAMPLE}")
    return tuple(SampleCase(path) for path in sorted(sources)) + (
        SampleCase(CHARTS_SAMPLE, charts=True),
    )


def render_corpus(output_dir: Path, root: Path = PROJECT_ROOT) -> tuple[Path, ...]:
    """Render the shared corpus through the normal strict conversion path.

    Args:
        output_dir: New directory reserved exclusively for this run.
        root: Source checkout; source-relative house files remain resolvable.

    Returns:
        Absolute paths to every saved document, in case order.
    """
    cases = sample_cases(root)
    destination = output_dir.resolve()
    destination.mkdir(parents=True, exist_ok=False)
    outputs: list[Path] = []
    for case in cases:
        output = destination / case.output_relative
        output.parent.mkdir(parents=True, exist_ok=True)
        converter = IBReportConverter(
            str(root / case.source), str(output), render_options=case.options
        )
        outputs.append(Path(converter.convert()).resolve())
    return tuple(outputs)
