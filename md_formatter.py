"""
MD Formatter - Preprocessor for single-line markdown files
Converts Gemini Deep Research clipboard output to properly formatted markdown.

This module handles the specific case where Deep Research output is copied
as a single continuous line with no newlines, containing:
    - Korean IB report structure (numbered headings, metadata, bullets)
    - LaTeX equations ($...$ and $$...$$)
    - Bold markers (**...**)
    - Callout labels ([시사점], [참고], etc.)

Usage:
    uv run md_formatter.py input.md [output.md]
    uv run md_formatter.py --check input.md        # Check if formatting needed

Changelog (v2):
    - Complete rewrite for single-line Deep Research clipboard handling
    - LaTeX equation preservation (block $$ and inline $)
    - Context-aware Korean sentence boundary detection
    - Numbered heading vs. numbered list discrimination
    - Metadata extraction (작성일, 수신, 주제, 작성자)
    - Bullet pattern restoration (Korean bullet markers)
    - Callout box detection ([시사점], [참고], etc.)
    - YAML frontmatter generation
"""

import argparse
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Match

from deep_md_cleaner import CleanerConfig, clean_deepresearch_markdown

# Parent folder path
PARENT_DIR = Path(__file__).resolve().parent.parent


# ═══════════════════════════════════════════════════════════════════════════════
# LATEX PROTECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class LaTeXProtector:
    """
    Protects LaTeX equations from being mangled during text reformatting.

    Replaces $$ ... $$ and $ ... $ with unique placeholders,
    then restores them after formatting is complete.
    """

    BLOCK_PLACEHOLDER = "__LATEX_BLOCK_{idx}__"
    INLINE_PLACEHOLDER = "__LATEX_INLINE_{idx}__"

    def __init__(self):
        self._block_store: List[str] = []
        self._inline_store: List[str] = []

    def protect(self, text: str) -> str:
        """Replace all LaTeX with placeholders"""
        # Block equations first ($$...$$) — greedy but not across other $$
        text = self._protect_block(text)
        # Inline equations ($...$) — non-greedy
        text = self._protect_inline(text)
        return text

    def restore(self, text: str) -> str:
        """Restore all LaTeX from placeholders"""
        # Restore block equations
        for idx, latex in enumerate(self._block_store):
            placeholder = self.BLOCK_PLACEHOLDER.format(idx=idx)
            text = text.replace(placeholder, f"\n\n{latex}\n\n")

        # Restore inline equations
        for idx, latex in enumerate(self._inline_store):
            placeholder = self.INLINE_PLACEHOLDER.format(idx=idx)
            text = text.replace(placeholder, latex)

        return text

    def _protect_block(self, text: str) -> str:
        """Protect block equations $$...$$"""

        def replacer(m):
            idx = len(self._block_store)
            self._block_store.append(m.group(0))
            return self.BLOCK_PLACEHOLDER.format(idx=idx)

        return re.sub(r"\$\$[^$]+?\$\$", replacer, text)

    def _protect_inline(self, text: str) -> str:
        """Protect inline equations $...$"""

        def replacer(m):
            idx = len(self._inline_store)
            self._inline_store.append(m.group(0))
            return self.INLINE_PLACEHOLDER.format(idx=idx)

        # Match $...$ but not $$...$$ (already removed)
        return re.sub(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)", replacer, text)


# ═══════════════════════════════════════════════════════════════════════════════
# BOLD PROTECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class BoldProtector:
    """
    Protects **bold** text from being split by sentence boundary detection.

    Korean sentence endings inside bold markers should not trigger line breaks.
    """

    PLACEHOLDER = "__BOLD_{idx}__"

    def __init__(self):
        self._store: List[str] = []

    def protect(self, text: str) -> str:
        def replacer(m):
            idx = len(self._store)
            self._store.append(m.group(0))
            return self.PLACEHOLDER.format(idx=idx)

        return re.sub(r"\*\*[^*]+?\*\*", replacer, text)

    def restore(self, text: str) -> str:
        for idx, bold_text in enumerate(self._store):
            placeholder = self.PLACEHOLDER.format(idx=idx)
            text = text.replace(placeholder, bold_text)
        return text


# ═══════════════════════════════════════════════════════════════════════════════
# METADATA EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════


class MetadataExtractor:
    """
    Extracts report metadata from the beginning of single-line text.

    Handles patterns like:
        [보고서] 제목작성일: ...수신: ...주제: ...
    """

    # Metadata field patterns (order matters — checked sequentially)
    _FIELDS = [
        ("보고서 작성일", re.compile(r"보고서\s*작성일:\s*")),
        ("작성일", re.compile(r"작성일:\s*")),
        ("수신", re.compile(r"수신:\s*")),
        ("주제", re.compile(r"주제:\s*")),
        ("작성자", re.compile(r"작성자:\s*")),
    ]

    # Report title pattern
    _TITLE_RE = re.compile(r"^\[보고서\]\s*(.+?)(?=작성일:|수신:|주제:|작성자:|\d+\.\s)")

    # Generic label boundary (e.g., "가정:", "트리거 요건:") to avoid
    # over-consuming metadata values when additional label blocks follow.
    _GENERIC_LABEL_BOUNDARY_RE = re.compile(r"[가-힣A-Za-z][가-힣A-Za-z0-9\s]{0,20}:\s*")

    @classmethod
    def extract(cls, text: str) -> Tuple[Dict[str, str], str]:
        """
        Extract metadata from the start of text.

        Returns:
            Tuple of (metadata_dict, remaining_text)
        """
        metadata: Dict[str, str] = {}
        remaining = text

        # Extract title
        title_match = cls._TITLE_RE.match(remaining)
        if title_match:
            metadata["title"] = title_match.group(1).strip()
            remaining = remaining[title_match.end() :]

        # Extract metadata fields
        for field_name, pattern in cls._FIELDS:
            match = pattern.search(remaining)
            if match:
                # Find the value: everything until the next field or heading
                start = match.end()
                # Look for end boundary
                end = len(remaining)
                for _, next_pattern in cls._FIELDS:
                    next_match = next_pattern.search(remaining[start:])
                    if next_match:
                        end = min(end, start + next_match.start())

                # Also check for heading patterns as boundary
                heading_match = re.search(r"\d+\.\s+[가-힣A-Z]", remaining[start:])
                if heading_match:
                    end = min(end, start + heading_match.start())

                # Also stop at generic label:value boundaries (e.g. "가정:")
                label_match = cls._GENERIC_LABEL_BOUNDARY_RE.search(remaining[start:])
                if label_match:
                    end = min(end, start + label_match.start())

                value = remaining[start:end].strip()
                metadata[field_name] = value

                # Remove this field from remaining text
                remaining = remaining[: match.start()] + remaining[end:]

        return metadata, remaining.strip()

    @classmethod
    def to_frontmatter(cls, metadata: Dict[str, str]) -> str:
        """Convert metadata dict to YAML frontmatter"""
        if not metadata:
            return ""

        lines = ["---"]

        if "title" in metadata:
            # Escape quotes in title
            title = metadata["title"].replace('"', '\\"')
            lines.append(f'title: "{title}"')

        field_mapping = {
            "주제": "subtitle",
            "작성일": "date",
            "보고서 작성일": "date",
            "수신": "recipient",
            "작성자": "analyst",
        }

        for kr_key, en_key in field_mapping.items():
            if kr_key in metadata:
                value = metadata[kr_key].replace('"', '\\"')
                lines.append(f'{en_key}: "{value}"')

        lines.append("---")
        return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
# STRUCTURE DETECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class StructureDetector:
    """
    Detects structural elements in a single-line markdown string.

    Inserts line breaks before structural markers to split
    the single line into proper markdown structure.
    """

    # ── Heading patterns ────────────────────────────────────────────────────
    # Major sections: "1. 제목" "2. Title" (1-2 digit number)
    MAJOR_HEADING_RE = re.compile(
        r"(\d{1,2})\.\s+"  # Section number
        r"([가-힣A-Z][가-힣a-zA-Z\s()\-:]{1,50})"  # Title (Korean or English)
    )

    # Sub-sections: "2.1. 제목" "3.2. Title"
    SUB_HEADING_RE = re.compile(
        r"(?<!\d\.)"
        r"(\d{1,2}\.\d{1,2})\.\s+"
        r'([가-힣A-Z"][가-힣a-zA-Z\s()\-:"\'"]{1,60})'
    )

    # Sub-sub-sections: "3.1.1. 제목"
    SUBSUB_HEADING_RE = re.compile(
        r"(\d{1,2}\.\d{1,2}\.\d{1,2})\.\s+" r"([가-힣A-Z][가-힣a-zA-Z\s()\-:]{1,60})"
    )

    # ── Korean sentence ending patterns ─────────────────────────────────────
    # These endings indicate a sentence boundary when followed by Korean text
    SENTENCE_END_RE = re.compile(
        r"(다|요|음|함|임|됨|것|수|점|니다|입니다|습니다|됩니다|합니다|입니까)"
        r"\."
        r"(?=[가-힣A-Z\[])"  # Followed by Korean char, uppercase, or bracket
    )

    # ── Callout patterns ────────────────────────────────────────────────────
    CALLOUT_RE = re.compile(r"\[(시사점|참고|주의|결론|요약|핵심|NOTE|WARNING)\]")

    # ── Bullet-like patterns ────────────────────────────────────────────────
    # Korean bullet markers that appear mid-line
    BULLET_MARKERS_RE = re.compile(r"([ㆍ•※])\s*")

    # ── Paragraph starters (typically begin new logical blocks) ──────────────
    PARA_STARTERS = [
        "결론적으로,",
        "따라서,",
        "그러나,",
        "다만,",
        "이에,",
        "특히,",
        "또한,",
        "한편,",
        "즉,",
        "예를 들어,",
        "구체적으로,",
        "본 보고서는",
        "본 건은",
        "위 알고리즘",
    ]

    @classmethod
    def insert_structure(cls, text: str) -> str:
        """
        Insert line breaks at structural boundaries in single-line text.

        Processing order matters — more specific patterns first.
        """
        result = text

        # 1. Sub-sections
        result = cls.SUB_HEADING_RE.sub(r"\n\n### \1. \2", result)

        # 2. Sub-sub-sections
        result = cls.SUBSUB_HEADING_RE.sub(r"\n\n#### \1. \2", result)

        # 3. Major sections (need lookbehind, so handle differently)
        result = cls._insert_major_headings(result)

        # 4. Callout boxes
        result = cls.CALLOUT_RE.sub(r"\n\n> [\1]", result)

        # 5. Bullet markers
        result = cls.BULLET_MARKERS_RE.sub(r"\n\n- ", result)

        # 6. Korean sentence boundaries (paragraph breaks)
        result = cls.SENTENCE_END_RE.sub(r"\1.\n\n", result)

        # 7. Paragraph starters
        for starter in cls.PARA_STARTERS:
            # Do not force a paragraph break when starter immediately follows ": "
            # (e.g., "Disclaimer: 본 보고서는 ...")
            pattern = rf"(?<!: ){re.escape(starter)}"
            result = re.sub(pattern, f"\n\n{starter}", result)

        return result

    @classmethod
    def _insert_major_headings(cls, text: str) -> str:
        """
        Insert ## markers before major section headings.

        Uses context to distinguish section headings from mid-sentence numbers.
        """
        # Find all potential major heading positions
        matches = list(cls.MAJOR_HEADING_RE.finditer(text))

        if not matches:
            return text

        # Process in reverse order to preserve positions
        result = text
        for m in reversed(matches):
            number = m.group(1)
            title = m.group(2)

            # Reject if this "N." is part of sub/sub-sub numbering (e.g. 2.1. / 3.1.1.)
            prefix = result[: m.start()]
            if re.search(r"\d\.\d\.$", prefix) or re.search(r"\d\.$", prefix):
                continue

            # Require boundary before major heading (start, whitespace, or sentence punctuation)
            if m.start() > 0:
                prev_char = result[m.start() - 1]
                if prev_char not in (" ", "\n", "\t", ".", "]", ")"):
                    continue

            # Heuristic: reject if title looks like a sentence continuation
            title_stripped = title.strip()
            if len(title_stripped) > 45:
                continue
            if re.search(r"[다요음함임됨]\.$", title_stripped):
                continue

            # Insert heading marker
            heading_text = f"\n\n## {number}. {title}"
            result = result[: m.start() + 1] + heading_text + result[m.end() :]

        return result


# ═══════════════════════════════════════════════════════════════════════════════
# COLON-LABEL DETECTOR
# ═══════════════════════════════════════════════════════════════════════════════


class ColonLabelDetector:
    """
    Detects label: value patterns that should be on separate lines.

    Handles patterns like:
        가정: 연 매출 2조 원...
        일 평균 매출: ...
        트리거 요건: ...
    """

    # Labels that should start new lines (typically followed by content)
    _LABELS = [
        "가정",
        "일 평균 매출",
        "보유 채권 잔액",
        "최대 조달 가능액",
        "트리거 요건",
        "발동 효과",
        "확정일자부 채권양도 승낙",
        "계좌 변경 금지 확약",
        "계좌 분리",
        "인출 통제",
        "API 기반 모니터링",
        "동적 적립률 조정",
        "희석화 준비금",
        "기타 준비금",
        "총 필요 신용보강률",
        "구조",
        "적정성",
        "실행 전략",
    ]

    # Metadata/label variants where ':' should be followed by one space
    _SPACING_LABELS = [
        "작성일",
        "보고서 작성일",
        "수신",
        "주제",
        "작성자",
    ] + _LABELS

    _BOLD_SPAN_RE = re.compile(r"\*\*[^*]+?\*\*")

    @staticmethod
    def _is_inside_spans(idx: int, spans: List[Tuple[int, int]]) -> bool:
        """Return True if index is inside one of the given [start, end) spans."""
        for start, end in spans:
            if start <= idx < end:
                return True
        return False

    @classmethod
    def insert_breaks(cls, text: str) -> str:
        """Insert line breaks before known label: patterns"""
        result = text
        for label in cls._LABELS:
            # Only break if label appears mid-sentence (preceded by content)
            pattern = f"(?<=[가-힣a-zA-Z0-9.%)\\]])({re.escape(label)}:)"
            result = re.sub(pattern, r"\n\n- \1", result)
        return result

    @classmethod
    def normalize_colon_spacing(cls, text: str) -> str:
        """Normalize label-value separator to ': ' for known labels."""
        result = text
        for label in cls._SPACING_LABELS:
            bold_spans = [(m.start(), m.end()) for m in cls._BOLD_SPAN_RE.finditer(result)]
            pattern = rf"({re.escape(label)}):\s*"

            def replacer(match: Match[str]) -> str:
                if cls._is_inside_spans(match.start(), bold_spans):
                    return match.group(0)
                return f"{match.group(1)}: "

            result = re.sub(pattern, replacer, result)
        return result

    @classmethod
    def normalize_colon_spacing_global(cls, text: str) -> str:
        """Normalize ':' spacing globally to ': ' with safe exceptions.

        Exceptions (preserved as-is):
        - URL schemes (`http://`, `https://`, `ftp://`)
        - Windows drive paths (`C:\\`)
        - Digit-to-digit time/ratio forms (`10:30`)
        - Double-colon tokens (`::`)
        - Bold markdown spans (`**...**`)
        """

        bold_spans = [(m.start(), m.end()) for m in cls._BOLD_SPAN_RE.finditer(text)]

        def replacer(match: Match[str]) -> str:
            idx = match.start()
            prev_char = text[idx - 1] if idx > 0 else ""
            next_char = text[idx + 1] if idx + 1 < len(text) else ""

            # Keep colons inside bold markdown unchanged
            if cls._is_inside_spans(idx, bold_spans):
                return ":"

            # URL schemes and Windows drive paths
            if next_char in ("/", "\\"):
                return ":"

            # Keep numeric time/ratio like 10:30
            if prev_char.isdigit() and next_char.isdigit():
                return ":"

            # Keep double-colon tokens
            if prev_char == ":" or next_char == ":":
                return ":"

            return ": "

        return re.sub(r":(?=\S)", replacer, text)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN FORMATTER
# ═══════════════════════════════════════════════════════════════════════════════


def format_markdown(
    text: str,
    cleaner_config: Optional[CleanerConfig] = None,
    cleaner_report: bool = False,
) -> str:
    """
    Format single-line markdown text into properly structured markdown.

    Pipeline:
        1. Detect if already formatted (multi-line)
        2. Protect LaTeX equations
        3. Protect bold markers
        4. Extract metadata → generate frontmatter
        5. Insert structural breaks (headings, sections)
        6. Insert sentence breaks (paragraph splitting)
        7. Insert label breaks (bullet-like items)
        8. Restore bold markers
        9. Restore LaTeX equations
        10. Clean up excess whitespace

    Args:
        text: Raw markdown text (possibly single-line)

    Returns:
        Properly formatted markdown with structure
    """
    result = text

    # Optional DeepResearch cleanup (off/auto/on via config)
    if cleaner_config is not None:
        result, report = clean_deepresearch_markdown(result, cleaner_config)
        if cleaner_report and (report.applied or report.markers_detected):
            print("[INFO] DeepResearch cleaner: {summary}".format(summary=report.summary()))

    # Quick check: if already has many lines, minimal formatting needed
    line_count = result.count("\n")
    if line_count > 20:
        print(f"[INFO] Text already has {line_count} lines - applying light formatting only")
        return _light_format(result)

    # ── Step 1: Protect LaTeX ───────────────────────────────────────────────
    latex_protector = LaTeXProtector()
    result = latex_protector.protect(result)

    # ── Step 2: Protect bold markers ────────────────────────────────────────
    bold_protector = BoldProtector()
    result = bold_protector.protect(result)

    # ── Step 3: Normalize colon spacing (global + label-specific) ─────────
    result = ColonLabelDetector.normalize_colon_spacing_global(result)
    result = ColonLabelDetector.normalize_colon_spacing(result)

    # ── Step 4: Extract metadata ────────────────────────────────────────────
    metadata, result = MetadataExtractor.extract(result)
    frontmatter = MetadataExtractor.to_frontmatter(metadata)

    # ── Step 5: Insert structural breaks ────────────────────────────────────
    result = StructureDetector.insert_structure(result)

    # ── Step 6: Insert label breaks ─────────────────────────────────────────
    result = ColonLabelDetector.insert_breaks(result)

    # ── Step 7: Restore bold markers ────────────────────────────────────────
    result = bold_protector.restore(result)

    # ── Step 8: Restore LaTeX ───────────────────────────────────────────────
    result = latex_protector.restore(result)

    # ── Step 9: Clean up ────────────────────────────────────────────────────
    result = _cleanup(result)

    # ── Step 10: Prepend frontmatter ────────────────────────────────────────
    if frontmatter:
        result = frontmatter + "\n\n" + result

    return result


def _light_format(text: str) -> str:
    """Apply minimal formatting to already-structured text"""
    result = text

    # Normalize colon spacing even in light-format mode
    result = ColonLabelDetector.normalize_colon_spacing_global(result)
    result = ColonLabelDetector.normalize_colon_spacing(result)

    # Ensure callout boxes have proper formatting
    result = re.sub(r"\[(시사점|참고|주의|결론|요약|핵심)\]", r"\n\n> [\1]", result)

    # Clean up excess whitespace
    result = re.sub(r"\n{3,}", "\n\n", result)

    return result.strip()


def _cleanup(text: str) -> str:
    """
    Final cleanup pass.

    - Remove excess blank lines (max 2 consecutive)
    - Ensure headings have blank line before them
    - Trim trailing whitespace per line
    """
    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Remove excess blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Ensure blank line before headings
    text = re.sub(r"([^\n])\n(#{1,4}\s)", r"\1\n\n\2", text)

    # Ensure blank line before block equations
    text = re.sub(r"([^\n])\n(\$\$)", r"\1\n\n\2", text)

    # Ensure blank line after block equations
    text = re.sub(r"(\$\$)\n([^\n])", r"\1\n\n\2", text)

    # Ensure blank line before blockquotes
    text = re.sub(r"([^\n])\n(>\s)", r"\1\n\n\2", text)

    # Ensure blank line before bullet items
    text = re.sub(r"([^\n])\n(-\s)", r"\1\n\n\2", text)

    # Trim trailing whitespace per line
    lines = text.split("\n")
    lines = [line.rstrip() for line in lines]
    text = "\n".join(lines)

    return text.strip()


# ═══════════════════════════════════════════════════════════════════════════════
# FILE I/O
# ═══════════════════════════════════════════════════════════════════════════════


def format_file(input_path: str, output_path: Optional[str] = None) -> str:
    """
    Format a markdown file.

    Args:
        input_path: Path to input markdown file
        output_path: Optional path for output (default: input_formatted.md)

    Returns:
        Path to the formatted file
    """
    return format_file_with_options(input_path=input_path, output_path=output_path)


def _build_cleaner_config(
    cleaner_mode: str,
    cite_mode: str,
    drop_unknown_markers: bool,
) -> CleanerConfig:
    """Build DeepResearch cleaner config."""
    return CleanerConfig(
        activation_mode=cleaner_mode,
        cite_mode=cite_mode,
        drop_unknown_markers=drop_unknown_markers,
    )


def format_file_with_options(
    input_path: str,
    output_path: Optional[str] = None,
    cleaner_mode: str = "off",
    cite_mode: str = "footnote",
    drop_unknown_markers: bool = False,
    cleaner_report: bool = False,
) -> str:
    """Format markdown file with optional DeepResearch cleaner controls."""
    input_file = Path(input_path)

    # Resolve path
    if not input_file.is_absolute():
        parent_path = PARENT_DIR / input_file.name
        if parent_path.exists():
            input_file = parent_path

    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    content = _read_with_encoding(input_file)
    line_count = content.count("\n")
    print(f"[INFO] Input: {input_file.name} ({line_count} lines, {len(content)} chars)")

    cleaner_config = _build_cleaner_config(
        cleaner_mode=cleaner_mode,
        cite_mode=cite_mode,
        drop_unknown_markers=drop_unknown_markers,
    )

    formatted = format_markdown(
        content,
        cleaner_config=cleaner_config,
        cleaner_report=cleaner_report,
    )

    if output_path:
        output_file = Path(output_path)
    else:
        output_file = input_file.with_name(input_file.stem + "_formatted.md")

    output_file.write_text(formatted, encoding="utf-8")

    new_line_count = formatted.count("\n")
    print(f"[OK] Formatted: {output_file}")
    print(f"     Lines: {line_count} -> {new_line_count}")
    print(f"     Size: {len(content)} -> {len(formatted)} chars")

    return str(output_file)


def _read_with_encoding(file_path: Path) -> str:
    """Read file with encoding detection fallback"""
    encodings = ["utf-8", "utf-8-sig", "euc-kr", "cp949"]
    for enc in encodings:
        try:
            return file_path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise UnicodeDecodeError(
        "multiple", b"", 0, 1, f"Failed to read {file_path} with encodings: {encodings}"
    )


def check_needs_formatting(input_path: str) -> bool:
    """
    Check if a file needs formatting.

    Returns True if the file appears to be single-line or minimally structured.
    """
    content = _read_with_encoding(Path(input_path))
    line_count = content.count("\n")
    char_count = len(content)

    if line_count == 0 and char_count > 100:
        return True
    if line_count < 5 and char_count > 500:
        return True

    return False


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════


def build_parser() -> argparse.ArgumentParser:
    """Build CLI parser for formatter."""
    parser = argparse.ArgumentParser(
        description="Format single-line markdown into structured markdown",
    )
    parser.add_argument("input_file", nargs="?", help="Input markdown file path")
    parser.add_argument("output_file", nargs="?", help="Output markdown file path")
    parser.add_argument("--check", action="store_true", help="Check if formatting is needed")

    parser.add_argument(
        "--deepresearch-cleaner",
        choices=["off", "auto", "on"],
        default="off",
        help="Apply OpenAI DeepResearch marker cleaner",
    )
    parser.add_argument(
        "--cite-mode",
        choices=["footnote", "inline", "strip"],
        default="footnote",
        help="How to transform cite markers when cleaner is enabled",
    )
    parser.add_argument(
        "--drop-unknown-markers",
        action="store_true",
        help="Drop unknown DeepResearch marker blocks instead of comment-preserving",
    )
    parser.add_argument(
        "--cleaner-report",
        action="store_true",
        help="Print DeepResearch cleaner summary",
    )
    return parser


def main():
    """Main entry point"""
    parser = build_parser()
    args = parser.parse_args()

    if args.input_file is None:
        parser.print_help()
        sys.exit(1)

    if args.check:
        needs_fmt = check_needs_formatting(args.input_file)
        if needs_fmt:
            print(f"[NEEDS FORMATTING] {args.input_file}")
        else:
            print(f"[OK] {args.input_file} appears already formatted")
        sys.exit(0 if not needs_fmt else 1)

    try:
        result = format_file_with_options(
            input_path=args.input_file,
            output_path=args.output_file,
            cleaner_mode=args.deepresearch_cleaner,
            cite_mode=args.cite_mode,
            drop_unknown_markers=args.drop_unknown_markers,
            cleaner_report=args.cleaner_report,
        )
        print(f"\nFormatted file: {result}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
