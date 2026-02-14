"""
DeepResearch markdown cleaner.

This module handles OpenAI DeepResearch markers embedded in markdown text,
for example:
    \ue200cite\ue202turn5search0\ue202turn1search8\ue201
    \ue200entity\ue202["company", "신세계", "..."]\ue201
    \ue200image_group\ue202{"query": [...]}\ue201
"""

import json
import logging
import re
from collections import OrderedDict
from dataclasses import dataclass
from typing import Optional, Tuple


logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# MARKER CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Use escaped code points (no literal glyphs in source).
MARKER_START = "\ue200"
MARKER_END = "\ue201"
MARKER_SEPARATOR = "\ue202"

# Any leftover private-use markers in this block should not leak to output.
PUA_RANGE_RE = re.compile(r"[\ue200-\ue20f]")


# ═══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═══════════════════════════════════════════════════════════════════════════════


@dataclass
class CleanerConfig:
    """Configuration for DeepResearch cleanup."""

    activation_mode: str = "off"  # off | auto | on
    cite_mode: str = "footnote"  # footnote | inline | strip
    drop_unknown_markers: bool = False


@dataclass
class CleanerReport:
    """Structured report for cleaner execution."""

    activation_mode: str = "off"
    markers_detected: bool = False
    applied: bool = False
    was_modified: bool = False
    original_size: int = 0
    cleaned_size: int = 0
    cite_markers: int = 0
    entity_markers: int = 0
    image_group_markers: int = 0
    unknown_markers: int = 0
    footnotes_emitted: int = 0
    leftover_pua_removed: int = 0

    def summary(self) -> str:
        """Return concise summary for logging."""
        return (
            "mode={mode} detected={detected} applied={applied} modified={modified} "
            "cite={cite} entity={entity} image_group={image_group} "
            "unknown={unknown} footnotes={footnotes} pua_removed={pua_removed}"
        ).format(
            mode=self.activation_mode,
            detected=self.markers_detected,
            applied=self.applied,
            modified=self.was_modified,
            cite=self.cite_markers,
            entity=self.entity_markers,
            image_group=self.image_group_markers,
            unknown=self.unknown_markers,
            footnotes=self.footnotes_emitted,
            pua_removed=self.leftover_pua_removed,
        )


# ═══════════════════════════════════════════════════════════════════════════════
# CLEANER
# ═══════════════════════════════════════════════════════════════════════════════


class DeepResearchCleaner:
    """Cleaner for DeepResearch-specific marker blocks."""

    _MARKER_BLOCK_RE = re.compile(
        re.escape(MARKER_START)
        + r"([a-zA-Z_][a-zA-Z0-9_]*)"
        + re.escape(MARKER_SEPARATOR)
        + r"(.*?)"
        + re.escape(MARKER_END),
        re.DOTALL,
    )

    def __init__(self, config: Optional[CleanerConfig] = None):
        self.config = config or CleanerConfig()

    def detect(self, text: str) -> bool:
        """Detect whether DeepResearch marker blocks are present."""
        return bool(self._MARKER_BLOCK_RE.search(text))

    def clean(self, text: str) -> Tuple[str, CleanerReport]:
        """Clean marker blocks according to config and return report."""
        report = CleanerReport(
            activation_mode=self.config.activation_mode,
            original_size=len(text),
            cleaned_size=len(text),
        )

        mode = self.config.activation_mode
        if mode not in ("off", "auto", "on"):
            raise ValueError("activation_mode must be one of: off, auto, on")

        if self.config.cite_mode not in ("footnote", "inline", "strip"):
            raise ValueError("cite_mode must be one of: footnote, inline, strip")

        report.markers_detected = self.detect(text)

        if mode == "off":
            return text, report

        if mode == "auto" and not report.markers_detected:
            return text, report

        report.applied = True

        footnotes = OrderedDict()

        def replace_block(match: re.Match) -> str:
            tag = match.group(1).strip().lower()
            payload = match.group(2)

            if tag == "cite":
                report.cite_markers += 1
                return self._handle_cite(payload, footnotes)

            if tag == "entity":
                report.entity_markers += 1
                return self._handle_entity(payload)

            if tag == "image_group":
                report.image_group_markers += 1
                return self._handle_image_group(payload)

            report.unknown_markers += 1
            return self._handle_unknown(tag, payload)

        cleaned = self._MARKER_BLOCK_RE.sub(replace_block, text)

        if self.config.cite_mode == "footnote" and footnotes:
            report.footnotes_emitted = len(footnotes)
            cleaned = self._append_citations(cleaned, footnotes)

        leftover = PUA_RANGE_RE.findall(cleaned)
        if leftover:
            report.leftover_pua_removed = len(leftover)
            cleaned = PUA_RANGE_RE.sub("", cleaned)

        report.cleaned_size = len(cleaned)
        report.was_modified = cleaned != text
        return cleaned, report

    def _handle_cite(self, payload: str, footnotes: "OrderedDict[str, int]") -> str:
        ids = [token.strip() for token in payload.split(MARKER_SEPARATOR) if token.strip()]
        if not ids and payload.strip():
            ids = [payload.strip()]

        if self.config.cite_mode == "strip":
            return ""

        if self.config.cite_mode == "inline":
            if not ids:
                return ""
            return " (sources: {sources})".format(sources=", ".join(ids))

        refs = []
        for sid in ids:
            if sid not in footnotes:
                footnotes[sid] = len(footnotes) + 1
            refs.append("[^{num}]".format(num=footnotes[sid]))
        return "".join(refs)

    @staticmethod
    def _handle_entity(payload: str) -> str:
        try:
            arr = json.loads(payload)
            if isinstance(arr, list) and len(arr) >= 2:
                return str(arr[1])
        except Exception:
            pass
        return " ".join(payload.split())

    @staticmethod
    def _handle_image_group(payload: str) -> str:
        try:
            obj = json.loads(payload)
            queries = obj.get("query") or obj.get("queries") or []
            if isinstance(queries, list) and queries:
                joined = ", ".join(str(item) for item in queries)
                return "\n<!-- image_group: {joined} -->\n".format(joined=joined)
        except Exception:
            pass
        return "\n<!-- image_group removed -->\n"

    def _handle_unknown(self, tag: str, payload: str) -> str:
        if self.config.drop_unknown_markers:
            return ""
        compact = " ".join(payload.split())
        if len(compact) > 240:
            compact = compact[:240] + "..."
        return "<!-- {tag}: {payload} -->".format(tag=tag, payload=compact)

    @staticmethod
    def _append_citations(text: str, footnotes: "OrderedDict[str, int]") -> str:
        lines = ["## Citations"]
        for source_id, num in footnotes.items():
            lines.append("[^{num}]: {source}".format(num=num, source=source_id))

        block = "\n".join(lines)
        stripped = text.rstrip()
        if not stripped:
            return block + "\n"
        return stripped + "\n\n" + block + "\n"


# ═══════════════════════════════════════════════════════════════════════════════
# CONVENIENCE
# ═══════════════════════════════════════════════════════════════════════════════


def clean_deepresearch_markdown(
    text: str,
    config: Optional[CleanerConfig] = None,
) -> Tuple[str, CleanerReport]:
    """Clean text with DeepResearchCleaner and return output/report."""
    cleaner = DeepResearchCleaner(config=config)
    return cleaner.clean(text)
