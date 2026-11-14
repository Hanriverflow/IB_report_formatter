"""Charts with the alt-text warnings (E9) and repeated footnote references.

Pins how the E10 chart paths meet two changes that reached main beside them:
chart pictures always carry a description (their title, else `Chart N`), so
neither the renderer nor `docx-audit` warns for them; and a footnote mark in a
chart's own texts is plain text, so it never adds a footnote reference.
"""

import re
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document

from document_profiles import RenderOptions
from docx_audit import audit_file, audit_to_dict
from ib_renderer import IBDocumentRenderer
from md_parser import MarkdownParser, parse_markdown_file

ROOT = Path(__file__).resolve().parents[1]
ALT_RENDER = "Image without a description"
ALT_AUDIT = "Picture without a description"


def _render(
    model, tmp_path: Path, name: str, profile: str, **options
) -> tuple[IBDocumentRenderer, Path]:
    renderer = IBDocumentRenderer(
        options=RenderOptions(profile=profile, charts=True, strict=True, **options)
    )
    doc = renderer.render(model)
    out = tmp_path / f"{name}.docx"
    doc.save(str(out))
    return renderer, out


def _descriptions(path: Path) -> list[str]:
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    return re.findall(r'<wp:docPr [^>]*descr="([^"]*)"', xml)


@pytest.mark.parametrize(
    ("sample", "chart_style"),
    [
        ("samples/qa/charts.md", "standard"),
        ("samples/qa/charts.md", "classic"),
        ("samples/qa/charts-dcm.md", "standard"),
    ],
)
def test_chart_samples_add_no_alt_text_warning(tmp_path, sample, chart_style):
    model = parse_markdown_file(str(ROOT / sample))
    renderer, out = _render(
        model, tmp_path, "sample", model.metadata.profile or "ib-report", chart_style=chart_style
    )

    assert [w for w in renderer.warnings if ALT_RENDER in w] == []
    descriptions = _descriptions(out)
    assert descriptions and all(d.strip() for d in descriptions)
    audit = audit_to_dict(audit_file(str(out)))
    assert audit["issues"] == []
    assert [w for w in audit["warnings"] if ALT_AUDIT in w] == []


UNTITLED = """```chart
type: bar
labels: [A, B]
series:
  - name: 매출
    values: [1, 2]
```

```chart
type: stacked_bar
id: fig-mix
title: 구성
labels: [A, B]
series:
  - name: 갑
    values: [1, 2]
  - name: 을
    values: [3, 4]
```

@fig-mix 참조.
"""


@pytest.mark.parametrize("chart_style", ["standard", "classic"])
def test_untitled_and_numbered_charts_are_described(tmp_path, chart_style):
    model = MarkdownParser(profile="plain").parse(UNTITLED)
    renderer, out = _render(model, tmp_path, "untitled", "plain", chart_style=chart_style)

    assert renderer.warnings == []
    assert _descriptions(out) == ["Chart 1", "구성"]
    audit = audit_to_dict(audit_file(str(out)))
    assert audit["issues"] == []
    assert [w for w in audit["warnings"] if ALT_AUDIT in w] == []


FOOTNOTED = """---
title: 각주 시험
charts: true
---

# 본문

본문 문장입니다.[^1]{again}

```chart
type: bar
title: 매출[^1]
subtitle: 백만원[^1]
note: 주석[^1]
source: 가상 자료[^1]
labels: [A, B]
series:
  - name: 매출
    values: [1, 2]
```

[^1]: 각주 내용.
"""


@pytest.mark.parametrize("again", ["", " 다시 한 번.[^1]"])
def test_footnote_marks_in_chart_texts_stay_text(tmp_path, again):
    model = MarkdownParser(profile="plain").parse(FOOTNOTED.format(again=again))
    renderer, out = _render(model, tmp_path, "footnoted", "plain")

    assert renderer.errors == [] and renderer.warnings == []
    xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8")
    # Only the body's marks reach Word: one native reference, a later one a NOTEREF field.
    assert xml.count("<w:footnoteReference") == 1
    assert xml.count("NOTEREF") == (1 if again else 0)
    texts = [p.text for p in Document(BytesIO(out.read_bytes())).paragraphs]
    for shown in ("매출[^1]", "백만원[^1]", "주: 주석[^1]", "자료: 가상 자료[^1]"):
        assert shown in texts
    assert _descriptions(out) == ["매출[^1]"]
    audit = audit_to_dict(audit_file(str(out)))
    assert audit["issues"] == []
