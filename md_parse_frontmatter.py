"""YAML frontmatter parsing and metadata inferred from the document body."""

import logging
import re

import yaml

from document_model import (
    CrossReference,
    DefaultTitle,
    DocumentMetadata,
    DocumentModel,
    ElementType,
    Heading,
    InferredText,
    Paragraph,
)
from document_profiles import default_metadata
from input_safety import StrictValidationError, document_validation, load_yaml

logger = logging.getLogger("md_parser")


class FrontmatterParser:
    """Parses YAML frontmatter from markdown"""

    _MARKDOWN_HEADING_RE = re.compile(r"^#{1,6}\s+")
    _BOLD_LABEL_RE = re.compile(r"\*\*[^*]+:\*\*")
    _SIMPLE_KEY_VALUE_RE = re.compile(r"^[A-Za-z0-9_.-]+\s*:\s*.*$")
    _DECLARED_RE = re.compile(
        r"^(profile|layout|tables|sender|charts|preset|house|terms|confirmation):", re.IGNORECASE
    )

    @staticmethod
    def parse(lines: list[str], profile: str | None = None) -> tuple[DocumentMetadata, list[str]]:
        """
        Parse YAML frontmatter and return metadata + remaining lines.

        Args:
            lines: All lines from the markdown file

        Returns:
            Tuple of (DocumentMetadata, remaining_lines)
        """
        metadata = default_metadata(profile or "ib-report")

        if not lines or lines[0].strip() != "---":
            return metadata, lines

        # Find end of frontmatter
        content_start_idx = 0
        frontmatter_lines: list[str] = []

        for i, line in enumerate(lines[1:], 1):
            if line.strip() == "---":
                content_start_idx = i + 1
                break
            frontmatter_lines.append(line)

        if content_start_idx == 0:
            return metadata, lines

        declared = any(FrontmatterParser._DECLARED_RE.match(line) for line in frontmatter_lines)
        if not declared and not FrontmatterParser._is_valid_frontmatter(frontmatter_lines):
            logger.debug("Frontmatter markers found, but content is not YAML frontmatter")
            return metadata, lines

        # Parse YAML content
        yaml_content = "\n".join(frontmatter_lines)
        try:
            parsed_data = load_yaml(yaml_content) or {}
        except (yaml.YAMLError, ValueError) as exc:
            if declared:
                raise StrictValidationError("Invalid YAML frontmatter: ", [str(exc)]) from exc
            # Fallback to simple key: value parsing only for simple YAML-like blocks
            if not FrontmatterParser._is_simple_key_value_block(frontmatter_lines):
                logger.debug(
                    "Frontmatter YAML parse failed and fallback is not safe; "
                    "treating block as document content"
                )
                return metadata, lines
            parsed_data = FrontmatterParser._parse_simple_key_values(frontmatter_lines)

        if not isinstance(parsed_data, dict):
            logger.debug(
                "Frontmatter parsed to %s (not mapping); treating as document content",
                type(parsed_data).__name__,
            )
            return metadata, lines

        data = {
            str(key).strip().lower(): value for key, value in parsed_data.items() if key is not None
        }

        with document_validation(profile is None):
            metadata = default_metadata(profile or str(data.get("profile", "ib-report")))

        # Map to metadata fields
        if metadata.profile == "term-sheet":
            # Preserve types so the profile contract can reject malformed YAML.
            metadata.title = data.get("title", metadata.title)
            metadata.subtitle = data.get("subtitle", metadata.subtitle)
        else:
            if "title" in data:
                # Without a title the profile's stand-in (DefaultTitle) stays marked.
                metadata.title = str(data["title"])
            metadata.subtitle = str(data.get("subtitle", metadata.subtitle))
        metadata.company = str(data.get("company", metadata.company))
        metadata.ticker = str(data.get("ticker", metadata.ticker))
        metadata.sector = str(data.get("sector", metadata.sector))
        metadata.analyst = str(data.get("analyst", metadata.analyst))

        # Store extra fields
        known_keys = {"title", "subtitle", "company", "ticker", "sector", "analyst", "profile"}
        structured = {
            "layout",
            "tables",
            "sender",
            "recipients",
            "cc",
            "attachments",
            "attendees",
            "letter",
            "charts",
            "preset",
            "house",
            "terms",
            "confirmation",
            "checks",
        }
        if metadata.profile == "term-sheet":
            structured.update({"prepared_by", "disclaimer", "confidential_label", "style"})
        metadata.extra = {
            k: v if k in structured else str(v) for k, v in data.items() if k not in known_keys
        }
        if "version" in data and data["version"] is None:
            # An empty `version:` means no version, not the text "None".
            metadata.extra["version"] = ""

        return metadata, lines[content_start_idx:]

    @staticmethod
    def _is_valid_frontmatter(frontmatter_lines: list[str]) -> bool:
        """
        Validate that a frontmatter block looks like YAML metadata, not markdown content.

        This prevents accidental content loss when documents begin with horizontal rules.
        """
        if not frontmatter_lines:
            return False

        non_empty_count = 0
        key_value_count = 0
        empty_streak = 0

        for line in frontmatter_lines:
            stripped = line.strip()

            if not stripped:
                empty_streak += 1
                if empty_streak >= 2:
                    return False
                continue

            empty_streak = 0
            non_empty_count += 1

            # Markdown content indicators
            if FrontmatterParser._MARKDOWN_HEADING_RE.match(stripped):
                return False
            if FrontmatterParser._BOLD_LABEL_RE.search(stripped):
                return False
            if "**" in stripped:
                return False

            if FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                _, value = stripped.split(":", 1)
                if len(value.strip()) > 120:
                    return False
                key_value_count += 1
                continue

            # Long prose line without key-value shape is likely document content
            if len(stripped) > 120 and ":" not in stripped:
                return False

        return non_empty_count > 0 and key_value_count > 0

    @staticmethod
    def _is_simple_key_value_block(frontmatter_lines: list[str]) -> bool:
        """Check if all non-empty lines look like simple key: value pairs."""
        has_key_value = False

        for line in frontmatter_lines:
            stripped = line.strip()
            if not stripped:
                continue
            if not FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                return False
            has_key_value = True

        return has_key_value

    @staticmethod
    def _parse_simple_key_values(frontmatter_lines: list[str]) -> dict[str, str]:
        """Parse simple key: value lines when YAML parsing fails."""
        data: dict[str, str] = {}

        for line in frontmatter_lines:
            stripped = line.strip()
            if not stripped or not FrontmatterParser._SIMPLE_KEY_VALUE_RE.match(stripped):
                continue
            key, value = stripped.split(":", 1)
            data[key.strip().lower()] = value.strip().strip('"').strip("'")

        return data


def _infer_metadata_from_elements(
    model: DocumentModel,
    allow_company_inference: bool = True,
) -> Heading | None:
    """Backfill metadata from document content when frontmatter is absent or partial.

    The parser calls this for IB profiles inside its reference scope and
    judges the title it returns. Every other value read here is an
    `InferredText`: a label value carries the references that resolve in its
    paragraph, and a company read from the title carries none, because the
    title it is part of is judged.

    Returns:
        The heading that became the title, or None.
    """
    metadata = model.metadata

    # Only the profile's stand-in gives way to the first H1. A written title, even
    # "IB Report" or one made by a term reference, is a plain string and stays.
    title_heading: Heading | None = None
    if isinstance(metadata.title, DefaultTitle):
        title_heading = next(
            (
                element.content
                for element in model.elements
                if element.element_type == ElementType.HEADING_1
                and isinstance(element.content, Heading)
                and element.content.text.strip()
            ),
            None,
        )
        if title_heading is not None:
            metadata.title = title_heading.text.strip()

    if allow_company_inference and metadata.company == "Korea Development Bank":
        inferred_company = _infer_company_from_title(metadata.title)
        if inferred_company:
            metadata.extra.setdefault("subject_company", InferredText.read(inferred_company))

    label_map = {
        "기준일": ("extra", "date"),
        "as of": ("extra", "date"),
        "작성일": ("extra", "date"),
        "report date": ("extra", "date"),
        "date": ("extra", "date"),
        "분석 대상 기간": ("extra", "analysis_period"),
        "analysis period": ("extra", "analysis_period"),
        "분석 기준": ("extra", "analysis_basis"),
        "analysis basis": ("extra", "analysis_basis"),
        "prepared by": ("analyst", None),
        "작성자": ("analyst", None),
        "analyst": ("analyst", None),
        "institution": ("company", None),
        "company": ("company", None),
        "기관": ("company", None),
        "sector": ("sector", None),
        "업종": ("sector", None),
        "prepared for": ("extra", "recipient"),
        "recipient": ("extra", "recipient"),
        "수신": ("extra", "recipient"),
    }

    for element in model.elements:
        if element.element_type not in {ElementType.PARAGRAPH, ElementType.HEADING_2}:
            continue
        if element.element_type == ElementType.HEADING_2:
            break
        if not isinstance(element.content, Paragraph):
            continue

        extracted = _extract_leading_metadata_pair(element.content)
        if extracted is None:
            continue

        label, text = extracted
        mapped = label_map.get(label.lower())
        if not mapped or not text:
            continue
        # The label run never holds a mapped value's reference: a label that is
        # a reference maps to nothing.
        value = InferredText.read(
            text,
            tuple(run.text for run in element.content.runs[1:] if isinstance(run, CrossReference)),
        )

        field_name, extra_key = mapped
        if field_name == "extra" and extra_key:
            metadata.extra.setdefault(extra_key, value)
        elif field_name and getattr(metadata, field_name, "") in {
            "",
            "IB Report",
            "SECTOR",
            "DCM Team 1",
            "Korea Development Bank",
        }:
            setattr(metadata, field_name, value)
    return title_heading


def _extract_leading_metadata_pair(paragraph: Paragraph) -> tuple[str, str] | None:
    """Extract a bold-label metadata pair like '**작성일:** 2026-03-20' from a paragraph."""
    if not paragraph.runs:
        return None

    first_run = paragraph.runs[0]
    if not first_run.bold:
        return None

    label = first_run.text.strip().rstrip(":").rstrip("：").strip()
    if not label:
        return None

    value = "".join(run.text for run in paragraph.runs[1:]).strip()
    if not value and ":" in first_run.text:
        raw_label, raw_value = first_run.text.split(":", 1)
        label = raw_label.strip().rstrip("：").strip()
        value = raw_value.strip()
    if not value:
        return None

    return label, value


def _infer_company_from_title(title: str) -> str:
    """Infer the company/subject name from a report-style title."""
    if not title or title == "IB Report":
        return ""

    for delimiter in (" — ", " - ", " – ", ": "):
        if delimiter in title:
            candidate = title.split(delimiter, 1)[0].strip()
            if candidate:
                return candidate

    suffix_patterns = [
        r"\s+수익성\s+변화\s+분석\s+보고서$",
        r"\s+수익성\s+분석\s+보고서$",
        r"\s+분석\s+보고서$",
        r"\s+보고서$",
        r"\s+리포트$",
    ]

    for pattern in suffix_patterns:
        candidate = re.sub(pattern, "", title).strip()
        if candidate and candidate != title:
            return candidate

    return ""
