# IB Report Formatter — 프로젝트 현황 & 다음 단계

> **최종 목표**: MD → 사람이 읽기에 좋은 Word / Word → LLM이 처리하기 좋은 MD
> **Branch**: `codex/mac-parser-stability`
> **Last updated**: 2026-03-21
> **Tests**: 250/250 passing
> **Python**: 3.8+ / `uv` package manager

---

## 아키텍처 개요

```
┌─────────────┐     ┌──────────────┐     ┌──────────────┐
│  Word (.docx)│────▶│ DocumentModel│────▶│ Markdown (.md)│  ← LLM용
│             │     │   (허브)      │     │              │
│  사람용      │◀────│              │◀────│              │
└─────────────┘     └──────────────┘     └──────────────┘
     ↑                    ↑                    ↑
 ib_renderer.py     md_parser.py         md_renderer.py
 word_parser.py     converters.py        word_to_md.py (CLI)
```

**핵심 패턴**: `DocumentModel`이 허브 — 모든 파서가 생산, 모든 렌더러가 소비.

---

## 완료된 작업 (Step 1~5)

### Step 1. OMML→LaTeX 변환 ✅

Word 수식(OMML XML)을 LaTeX로 자동 변환.

| 파일 | 역할 |
|------|------|
| `omml_latex.py` (신규) | OMML→LaTeX 변환 엔진 |
| `word_parser.py` (수정) | DOCX 열 때 자동 pre-processing |
| `tests/test_omml_latex.py` (신규) | 20개 테스트 |

- WordParser가 DOCX ZIP 내부 XML에서 `<m:oMath>` 태그를 `$...$` / `$$...$$`로 변환
- 대상: document.xml, footnotes.xml, endnotes.xml
- 분수, 상/하첨자, 근호, 행렬, 적분 등 지원
- 실패 시 graceful fallback (원본 그대로 열림)

---

### Step 2. 플러그인 아키텍처 ✅

새 포맷을 클래스 + `register()`만으로 추가할 수 있는 구조.

| 파일 | 역할 |
|------|------|
| `converters.py` (신규) | BaseConverter, InputConverter, OutputConverter, ConverterRegistry |
| `tests/test_converters.py` (신규) | 31개 테스트 |

```python
from converters import get_default_registry

registry = get_default_registry()
model = registry.convert("report.md")           # 입력
registry.convert(model, output_format="docx", output_path="out.docx")  # 출력
```

빌트인: MarkdownInput, DocxInput, DocxOutput, MarkdownOutput

---

### Step 3. Stream 기반 처리 + MIME 감지 ✅

파일 경로 외에 `BinaryIO` 스트림 입력 + CLI 파이프 지원.

| 파일 | 역할 |
|------|------|
| `stream_utils.py` (신규) | `detect_format()`, `ensure_seekable()`, `is_stream()` |
| `word_parser.py` (수정) | `parse()` / `_open_document()` → `Union[str, BinaryIO]` |
| `md_parser.py` (수정) | `parse_markdown_file()` → `Union[str, BinaryIO]`, `_decode_bytes()` 추가 |
| `converters.py` (수정) | InputConverter에 `supported_format` + 스트림 자동 감지 |
| `word_to_md.py` (수정) | `cat file.docx \| word_to_md.py -` 파이프 지원 |
| `tests/test_stream.py` (신규) | 24개 테스트 |

MIME 감지 전략 (경량, 외부 의존성 없음):
1. Extension hint → 2. ZIP 시그니처(docx) → 3. UTF-8 heuristic(md) → 4. unknown

---

### Step 4. MD 출력 정규화 (LLM 최적화) ✅

markitdown 출력 정규화 패턴을 반영. 모든 `render_to_markdown()` 출력에 자동 적용.

| 파일 | 역할 |
|------|------|
| `md_renderer.py` (수정) | `_normalize_markdown()` 후처리 파이프라인 |
| `tests/test_md_renderer.py` (수정) | 15개 테스트 추가 (4→19) |

정규화 규칙:
1. 줄 끝 공백 제거 → LLM 토큰 절약
2. 3줄+ 연속 빈줄 → 1줄로 압축
3. 블록 요소(테이블, `$$수식$$`, `` ```코드``` ``) 전후 빈줄 보장
4. 최종 단일 개행

---

### Step 5. Word 출력 품질 개선 (MD→Word) ✅

13개 품질 격차 진단 후 MD→Word 방향 핵심 4개 수정.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | TextRenderer 확장, TableRenderer runs 활용, SEPARATOR, 콜아웃 서식 |
| `tests/test_ib_renderer.py` (수정) | 15개 테스트 추가 (21→36) |

| 수정 항목 | Before | After |
|-----------|--------|-------|
| 테이블 데이터 셀 | `cell.runs` 무시, plain text | `render_runs()` 우선 (bold/italic/super 보존) |
| 테이블 헤더 셀 | `**` 마커 단순 제거 | runs 구조적 렌더링 |
| SEPARATOR (`---`) | 완전히 무시됨 | Word 수평선 (gray bottom border) |
| 콜아웃 본문 | plain text 단일 run | `**bold**` / `*italic*` / `^super^` 파싱 |

신규 메서드: `TextRenderer.render_text_with_formatting()` (기존 `render_text_with_bold()`는 위임)

---

### Step 6. Word→MD 입력 품질 개선 ✅

Word에서 MD로 변환할 때 손실되던 인라인 정보와 이미지 메타데이터를 보존.

| 파일 | 역할 |
|------|------|
| `word_parser.py` (수정) | 실제 이미지 alt text 추출, Base64 임베딩 옵션, subscript 추출 |
| `md_renderer.py` (수정) | `~subscript~`, 테이블 셀 runs, Base64 이미지 렌더링 |
| `word_to_md.py` (수정) | `--embed-images-base64` CLI 옵션 추가 |
| `md_parser.py` (수정) | `~text~` subscript 마커 파싱 |
| `ib_renderer.py` (수정) | MD→Word 방향 subscript 렌더링 |
| `tests/test_word_parser.py` / `tests/test_md_renderer.py` / `tests/test_stream.py` / `tests/test_word_to_md_cli.py` / `tests/test_md_parser.py` / `tests/test_ib_renderer.py` (수정) | 11개 회귀 테스트 추가 |

개선 항목:
- 이미지 alt text: Word XML `docPr descr/title` 우선 추출
- 이미지 임베딩: `--embed-images-base64`로 data URI 출력 지원
- subscript 캡처: Word subscript → `~text~`로 보존
- 테이블 셀 서식: `runs` 기반으로 `**bold**`, `*italic*`, `^super^`, `~sub~` 출력
- 스트림 입력: `parse_word_file(BytesIO, embed_images_base64=True)` 경로도 검증

---

### Step 7. Word→MD 구조/메타데이터 보존 강화 ✅

Word에서 구조와 문서 속성 메타데이터를 더 안정적으로 복구.

| 파일 | 역할 |
|------|------|
| `word_parser.py` (수정) | 커스텀 heading style 상속 인식, outline level 인식, `docProps/custom.xml` 추출 |
| `tests/test_word_parser.py` (수정) | 커스텀 style heading / custom props 회귀 테스트 추가 |
| `tests/test_stream.py` (수정) | BinaryIO 스트림에서도 custom props 추출 회귀 테스트 추가 |

개선 항목:
- 커스텀 heading style: `Heading N`을 직접 쓰지 않아도 base style 체인과 outline level을 따라 heading level 복구
- 문서 커스텀 메타데이터: `docProps/custom.xml`에서 known field는 `DocumentMetadata`로 매핑
- 추가 커스텀 속성: 표준 필드가 아닌 값도 `metadata.extra`에 snake_case 키로 보존
- 스트림 입력: 파일 경로뿐 아니라 `BytesIO` 입력에서도 custom metadata 추출 검증

---

### Step 8. Word→MD 색상 정보 보존 ✅

Word에서 추출한 인라인 텍스트 색상을 Markdown과 Word 왕복 경로 모두에서 보존.

| 파일 | 역할 |
|------|------|
| `word_parser.py` (수정) | Word run의 RGB 색상을 `TextRun.color_hex`로 추출 |
| `md_renderer.py` (수정) | 색상 run을 `<span style="color:#RRGGBB">...</span>`으로 렌더링 |
| `md_parser.py` (수정) | color span을 다시 `TextRun.color_hex`로 파싱 |
| `ib_renderer.py` (수정) | 색상 run / color span을 Word run 색상으로 복원 |
| `tests/test_md_parser.py` / `tests/test_md_renderer.py` / `tests/test_word_parser.py` / `tests/test_ib_renderer.py` (수정) | 색상 보존 회귀 테스트 추가 |

개선 항목:
- Word→MD: run RGB 색상을 `#RRGGBB`로 추출 후 HTML color span으로 출력
- MD→Word: color span과 colored run을 다시 Word run 색상으로 렌더링
- Strip mode: LLM용 `--strip` 경로에서는 색상 마크업 없이 순수 텍스트 유지
- 공통 모델: `TextRun.color_hex` 추가로 파서/렌더러 간 공유 표현 통일

---

### Step 9. MD→Word 이미지 alt text 삽입 ✅

Word에 삽입되는 이미지가 보이는 caption뿐 아니라 Word 자체 이미지 메타데이터에도 alt text를 갖도록 개선.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | `run.add_picture()` 결과의 `docPr descr/title`에 alt text 삽입 |
| `word_parser.py` (수정) | TOC skip 로직이 텍스트 없는 이미지 문단을 잘못 건너뛰지 않도록 보정 |
| `tests/test_ib_renderer.py` (수정) | alt text가 Word 메타데이터와 Word→MD roundtrip에 모두 남는지 검증 |

개선 항목:
- MD→Word: 이미지 삽입 시 Word `docPr descr/title`에 alt text 기록
- Word→MD: IB 문서의 TOC skip 과정에서 이미지 문단이 blank line로 오인되는 버그 수정
- Roundtrip: MD 이미지 alt text → Word alt text → 다시 Markdown alt text로 복구 확인

---

### Step 10. MD→Word 네이티브 footnote 지원 ✅

Markdown 인용 표기를 Word의 실제 footnote XML로 렌더링하고, 다시 읽을 때도 복구 가능하도록 개선.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | native `footnotes.xml` part 생성, superscript 숫자 run을 `w:footnoteReference`로 치환 |
| `word_parser.py` (수정) | native Word footnotes part와 inline `footnoteReference`를 다시 파싱 |
| `tests/test_word_parser.py` (수정) | markdown → docx → markdown native footnote roundtrip 검증 |

개선 항목:
- MD→Word: superscript citation marker가 있으면 ENDNOTES 텍스트 리스트 대신 native Word footnote 생성
- Fallback: inline marker가 없으면 기존 ENDNOTES 섹션 fallback 유지
- Word→MD: native footnotes part가 있으면 footnote 본문과 inline reference 숫자를 함께 복구
- Roundtrip: `^1^` + `## Citations`가 실제 Word footnote XML로 저장되고 다시 파싱 가능

---

### Step 11. MD→Word 인라인 LaTeX 이미지 렌더링 ✅

문단 안 `$...$` 수식이 더 이상 텍스트 fallback으로만 남지 않고, inline 이미지로 렌더링되도록 개선.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | `TextRun.is_latex`를 inline 이미지 삽입으로 렌더링 |
| `md_parser.py` (수정) | balanced inline LaTeX가 있는 table cell은 LaTeX-aware runs로 보존 |
| `tests/test_ib_renderer.py` (수정) | inline LaTeX run/paragraph가 inline picture로 렌더되는지 검증 |
| `tests/test_md_parser.py` (수정) | inline LaTeX run / table cell LaTeX run 파싱 검증 |

개선 항목:
- 문단/리스트: `TextRun.is_latex`를 inline PNG로 렌더링
- 테이블 셀: balanced `$...$`가 있으면 currency-safe plain parser 대신 LaTeX-aware parser 사용
- Fallback: 이미지 렌더링 실패 시 기존 `[expr]` 텍스트 fallback 유지
- 테스트: matplotlib 없이도 monkeypatch된 image path로 inline picture 삽입 경로 검증

---

### Step 12. MD→Word 깊은 리스트 중첩 안정화 ✅

깊은 nested list가 Word 페이지 폭을 선형으로 잠식하지 않도록 들여쓰기 공식을 조정.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | deep list indent 압축 계산 + 최대 indent 상한 적용 |
| `tests/test_ib_renderer.py` (수정) | deep bullet/numbered list indent 회귀 테스트 추가 |

개선 항목:
- 0~3레벨: 기존과 동일한 `0.25in` step 유지
- 4레벨 이후: 더 작은 step으로 압축해 hierarchy는 유지하고 폭 증가는 완화
- 최대 indent 상한: 과도한 nesting에서도 left indent가 일정 폭 이상 커지지 않도록 제한
- 회귀 테스트: deep bullet과 numbered list 모두 선형 폭증이 사라졌는지 확인

---

### Step 13. MD→Word separator page break 모드 ✅

separator를 수평선으로만 렌더하던 동작을 확장해서, 명시적 page break와 CLI override를 지원.

| 파일 | 역할 |
|------|------|
| `ib_renderer.py` (수정) | separator를 `rule` / `page-break` / `auto` 모드로 렌더 |
| `md_to_word.py` (수정) | `--separator-mode` CLI 옵션 추가 |
| `tests/test_ib_renderer.py` (수정) | `## ---` auto page break / 강제 page-break 모드 회귀 테스트 추가 |
| `tests/test_cli_batch.py` (수정) | 배치 CLI 하위호환 인자 보강 |

개선 항목:
- 기본값 `auto`: plain `---`는 기존처럼 수평선 유지
- 명시적 page break: `## ---`는 Word page break로 렌더
- 강제 모드: CLI에서 `--separator-mode page-break`를 주면 모든 separator를 page break로 렌더
- 호환성: 기존 문서와 테스트는 기본 동작을 유지

---

### Step 14. Word→MD theme/tint 색상 복원 ✅

직접 RGB가 아닌 Word theme color와 tint/shade도 실제 `#RRGGBB` 값으로 복원되도록 개선.

| 파일 | 역할 |
|------|------|
| `word_parser.py` (수정) | `theme1.xml` 파싱, theme/tint/shade 색상 복원 |
| `tests/test_word_parser.py` (수정) | theme color / tinted theme color 회귀 테스트 추가 |

개선 항목:
- `word/theme/theme1.xml`에서 Office color scheme 추출
- `w:themeColor`만 있는 run도 실제 RGB로 복원
- `w:themeTint` / `w:themeShade`가 있으면 밝기 조정까지 반영
- 기존 direct RGB parsing과 호환 유지

---

### Step 15. 실제 MD→Word 사례 평가 및 usability 개선 ✅

`tests/일동제약_수익성분석.md` 실제 사례를 기준으로 MD→Word 변환물을 점검하고, 사용자가 Word에서 직접 손보던 부분을 줄이도록 개선.

| 파일 | 역할 |
|------|------|
| `md_parser.py` (수정) | frontmatter가 없어도 첫 H1과 선행 bold 메타 문단에서 title/date 등 metadata 보강 |
| `ib_renderer.py` (수정) | TOC field와 함께 즉시 보이는 static TOC preview 생성 |
| `md_to_word.py` (수정) | TOC renderer에 parsed model 전달 |
| `tests/test_md_parser.py` (수정) | H1/date metadata 추론 회귀 테스트 추가 |
| `tests/test_ib_renderer.py` (수정) | TOC preview heading 항목 가시성 회귀 테스트 추가 |

실제 사례 평가:
- 입력: `tests/일동제약_수익성분석.md`
- 기존 문제:
  - cover title이 frontmatter 부재 시 `IB Report`로 출력됨
  - TOC는 Word field 업데이트 전 비어 보여 사용자가 목차를 직접 채운 것처럼 느껴짐
- 개선 후:
  - 첫 H1(`일동제약 주식회사 수익성 변화 분석 보고서`)가 cover title로 자동 승격
  - 선행 `**작성일:** ...` 문단이 metadata로 흡수되어 cover metadata 품질 개선
  - TOC 페이지에 heading 기반 preview가 바로 보여 field 업데이트 전에도 outline 확인 가능
  - `분석 대상 기간`, `분석 기준`이 cover metadata panel로 승격
  - 제목 기반으로 `Institution`을 `일동제약 주식회사`로 추론
  - placeholder였던 `PREPARED BY = DCM Team 1`, `SECTOR = SECTOR` 행은 자동 숨김 처리되어 수동 정리 부담 감소
  - 표 셀 내부는 vertical alignment와 paragraph spacing을 정규화해 숫자/텍스트 정렬이 더 안정적으로 보이도록 polish
  - cover 상단에서 회사명과 제목이 중복될 때 `회사명 / 보고서명`으로 분리해 첫인상 개선
  - cover 상단의 `SECTOR` placeholder도 숨김 처리되어 템플릿 느낌 감소

추가로 실제 사례에서 발견되어 함께 수정한 버그:
- `1월~12월` 같은 범위 표기에서 `~...~` subscript 파서가 멀리 떨어진 두 `~`를 한 쌍으로 잡아 분석 기간 문자열을 훼손하던 문제 수정
- subscript 문법을 `H~2~O` 같은 짧은 토큰 중심으로 보수화하여 일반 한국어 범위 표기와 충돌하지 않도록 조정

주의:
- Word field를 업데이트하면 page number가 포함된 정식 TOC가 생성됨
- 현재 preview는 가시성 보강용이며, page number 계산 자체는 Word field update에 의존

---

### Step 16. 두 번째 실제 사례(웅진 계열사) 기반 polish ✅

`tests/웅진_계열사.md` 실제 사례를 기준으로 cover 메타데이터 추론과 구조 도식 표현을 추가 개선.

| 파일 | 역할 |
|------|------|
| `md_parser.py` (수정) | `기준일:` 메타 추론, em dash 제목에서 company/title 추론, fenced code block 파싱 |
| `ib_renderer.py` (수정) | code block을 monospaced shaded block으로 렌더, default report type 숨김, cover title split polish |
| `md_renderer.py` (수정) | code block element를 fenced block으로 다시 렌더 |
| `tests/test_md_parser.py` / `tests/test_ib_renderer.py` (수정) | code block / cover split / metadata inference 회귀 테스트 추가 |

실제 사례 평가:
- 입력: `tests/웅진_계열사.md`
- 기존 문제:
  - `**기준일: ...**` 문단이 cover metadata에 반영되지 않음
  - `주식회사 웅진 — 자회사 보유구조 및 영업관계 분석`이 cover에서 한 덩어리로 보여 첫인상이 무거움
  - `지배구조 도식 (간략)` fenced block이 Word에서 백틱 문단으로 깨짐
  - default `DCM RESEARCH` 문구가 불필요하게 노출됨
- 개선 후:
  - `REPORT DATE`에 `기준일`이 반영됨
  - cover가 `주식회사 웅진` / `자회사 보유구조 및 영업관계 분석`으로 분리되어 가독성 개선
  - 구조 도식이 monospaced shaded block으로 렌더되어 훨씬 읽기 쉬워짐
  - default report type는 명시되지 않은 경우 숨김 처리

---

## 현재 프로젝트 전체 파일 구조

```
신규 파일 (이번 프로젝트에서 추가):
  omml_latex.py                 — OMML→LaTeX 변환기
  converters.py                 — 플러그인 아키텍처 (BaseConverter/Registry)
  stream_utils.py               — 스트림 유틸리티 + MIME 감지
  tests/test_omml_latex.py      — 20개 테스트
  tests/test_converters.py      — 31개 테스트
  tests/test_stream.py          — 24개 테스트

수정 파일:
  word_parser.py                — OMML 통합 + BinaryIO 스트림 지원
  md_parser.py                  — BinaryIO 스트림 지원 + _decode_bytes()
  md_renderer.py                — _normalize_markdown() 출력 정규화
  ib_renderer.py                — 테이블 셀 서식, SEPARATOR, 콜아웃 서식, TextRenderer 확장
  word_to_md.py                 — CLI stdin 파이프 지원 + Base64 이미지 임베딩 옵션
  converters.py                 — InputConverter 스트림 분기 + supported_format
  pyproject.toml                — isort known-first-party 갱신
  tests/test_ib_renderer.py     — 36개 테스트 (기존 21 + 신규 15)
  tests/test_md_renderer.py     — 19개 테스트 (기존 4 + 신규 15)
```

---

## 다음 단계 후보

### 후보 A. Word→MD 입력 품질 (난이도: 중)
완료됨. Step 6 참고.

### 후보 B. MCP 서버 (난이도: 중) -> 제외 할 것, 외부 mcp 제공계획 현재없음

Claude Desktop/Claude Code에서 직접 변환 호출 가능하게.

```
사용자: "이 Word 파일을 마크다운으로 변환해줘"
Claude: [MCP tool 호출] → 변환 결과 반환
```

### 후보 C. 3rd-party 플러그인 (난이도: 중) -> 제외 할 것, 외부 플러그인 제공계획 없음

`pyproject.toml`의 `entry_points` 기반으로 외부 패키지가 컨버터를 등록.

```toml
[project.entry-points."ib_report_formatter.plugin"]
pdf = "my_pdf_plugin:PdfConverter"
```

---

## 품질 격차 진단 결과 (참고)

2026-03-21 분석에서 발견된 13개 격차 중 4개 해결 완료. 나머지:

### Word→MD (LLM용) 미해결
- 명시적 미해결 주요 갭 없음 (추가 개선은 edge-case / polish 영역)

### MD→Word (사람용) 미해결
- 명시적 미해결 주요 갭 없음 (추가 개선은 polish/heuristic 영역)
