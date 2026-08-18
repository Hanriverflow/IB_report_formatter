<p align="right">
  <a href="./README.md"><img alt="lang English" src="https://img.shields.io/badge/lang-English-blue"></a>
  <a href="./README.ko.md"><img alt="lang 한국어" src="https://img.shields.io/badge/lang-한국어-orange"></a>
</p>

# IB Report Formatter

Markdown → Word 변환기로, IB(투자은행) 스타일의 전문 Word 보고서(`.docx`)를 생성합니다.

이 프로젝트는 리서치/내부 메모 형태의 markdown을 구조화된 제목, 표 스타일링, 콜아웃 박스, 이미지, 수식, 헤더/푸터가 포함된 보고서 형태로 출력합니다.

## 주요 기능

- **Markdown → Word** 변환 (IB 스타일 문서 생성)
- Markdown 입력 폴더 batch 변환 지원
- 단일 라인(클립보드) markdown 자동 구조화 포맷팅
- OpenAI DeepResearch 마커 정리기(선택 적용: `off`/`auto`/`on`)
- frontmatter가 없는 실제 보고서에서도 제목/날짜/분석기간/분석기준 메타데이터 자동 추론
- Word TOC field + 즉시 보이는 preview 목차 동시 생성
- cover / TOC에 `Malgun Gothic` 기반 한글 친화 타이포그래피 적용
- YAML frontmatter 파싱 (`title`, `date`, `recipient`, `analyst` 등)
- 금융 표 렌더링(천 단위 콤마, 의미 기반 정렬, 내용 기반 열 폭 조정)
- 구조 도식/코드블록을 monospaced shaded panel로 렌더링
- 콜아웃 박스 렌더링 (`[요약]`, `[시사점]`, `[주의]`, `[참고]` 등)
- 이미지 삽입(파일 경로, Base64 `data:image/...`)
- LaTeX 수식 지원 (`$inline$`, `$$block$$`, 기본 설치에서 Word 이미지로 렌더링)
- inline citation marker가 있을 때 markdown 인용을 Word 네이티브 footnote로 렌더링
- 헤더/푸터 구성(회사명, `CONFIDENTIAL`, 페이지 번호)

## 프로젝트 구조

```text
IB_report_formatter/
├── md_to_word.py      # Markdown → Word 변환 CLI
├── md_parser.py       # Markdown/frontmatter/요소 파서
├── md_formatter.py    # 단일 라인 markdown 전처리기
├── ib_renderer.py     # Word 렌더러 및 스타일 시스템
├── tests/             # Pytest 테스트
└── pyproject.toml     # 의존성/도구 설정
```

## 요구 사항

- Python 3.8+
- [uv](https://docs.astral.sh/uv/) (권장 패키지 매니저)

## 다른 PC에서 설치하기

아래 순서대로 진행하면 어느 컴퓨터에서든 프로젝트를 실행할 수 있습니다.

### 1. Python 설치

[python.org](https://www.python.org/downloads/)에서 Python 3.8 이상을 다운로드하여 설치합니다.

설치 확인:

```bash
python --version
```

### 2. uv 설치 (패키지 매니저)

**Windows (PowerShell):**

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

**macOS / Linux:**

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

설치 확인:

```bash
uv --version
```

### 3. 프로젝트 복사

`IB_report_formatter` 폴더 전체를 대상 PC로 복사하거나, 저장소에서 클론합니다:

```bash
git clone <저장소-주소> IB_report_formatter
cd IB_report_formatter
```

### 4. 의존성 설치

프로젝트 폴더로 이동 후 실행:

```bash
uv sync
```

가상환경 생성과 필수 패키지 설치가 자동으로 완료됩니다.

**선택:** 한글 파일 인코딩 보강 설치:

```bash
uv sync --extra full
```

**선택:** 개발/테스트 도구 설치:

```bash
uv sync --extra dev
```

### 5. 설치 확인

```bash
uv run md_to_word.py --list
```

성공하면 상위 폴더의 markdown 파일 목록이 표시됩니다.

## GitHub에 올려야 할 파일

다른 PC에서 동일하게 실행하려면, 아래 실행 필수 파일만 올리면 됩니다.

포함 권장:

- `md_to_word.py`
- `md_parser.py`
- `md_formatter.py`
- `ib_renderer.py`
- `tests/`
- `pyproject.toml`
- `uv.lock`
- `README.md`
- `README.ko.md`
- `AGENTS.md` (선택, 협업 가이드용)
- `docs/` (선택, 내부/민감 내용 제거 후)

제외 필수:

- `.venv/`, `__pycache__/`, `.pytest_cache/`, `.mypy_cache/`, `.ruff_cache/`
- 결과물 `*.docx`
- 로컬 도구 상태 파일 (`.claude/`, `.sisyphus/`)
- 사내 민감정보가 들어간 원본 markdown 파일

현재 루트 `.gitignore`에 위 제외 항목이 기본 반영되어 있습니다.

## 문제 해결

| 증상 | 해결 방법 |
|------|----------|
| `uv: command not found` | 터미널 재시작 또는 uv를 PATH에 추가 |
| `python: command not found` | Python 설치 후 PATH에 추가되었는지 확인 |
| Windows 권한 오류 | PowerShell을 관리자 권한으로 실행 |
| 한글 파일 인코딩 오류 | `uv sync --extra full`로 인코딩 지원 강화 |

## 빠른 시작

markdown -> Word 변환:

```bash
uv run md_to_word.py input.md
```

스크립트 엔트리포인트 사용:

```bash
uv run ib-report input.md
```

출력 파일 경로 지정:

```bash
uv run md_to_word.py input.md output.docx
```

사전 포맷팅 후 변환:

```bash
uv run md_to_word.py input.md --format
```

OpenAI DeepResearch 마커를 감지될 때만 정리 후 변환:

```bash
uv run md_to_word.py input.md --deepresearch-cleaner auto --cite-mode footnote --cleaner-report
```

폴더 단위 batch 변환:

```bash
uv run md_to_word.py reports/ --batch
```

## 추천 사용 흐름

### 1. 일반 보고서 markdown을 바로 Word로 변환

이미 제목/문단/표 구조가 잘 잡혀 있는 보고서라면:

```bash
uv run md_to_word.py tests/웅진_계열사.md
```

출력 경로를 명시적으로 고정하려면:

```bash
uv run md_to_word.py tests/웅진_계열사.md tests/웅진_계열사_Report_Pro.docx
```

### 2. 복붙된 난삽한 markdown을 정리 후 Word 변환

문서 구조가 깨져 있거나 한 줄로 뭉친 원문이라면:

```bash
uv run md_to_word.py raw_report.md --format
```

OpenAI DeepResearch 마커가 섞였을 가능성이 있으면:

```bash
uv run md_to_word.py raw_report.md --format --deepresearch-cleaner auto --cite-mode footnote --cleaner-report
```

추천 순서:

1. `uv run md_formatter.py --check input.md`
2. 구조가 뭉쳐 있으면 `--format`
3. 마커 정리가 필요할 수 있으면 `--deepresearch-cleaner auto`
4. 최종 `.docx` 생성

`--format`(사전 포맷팅)은 무엇을 하나요?

- 한 줄로 뭉친 markdown을 문서 구조로 자동 복원합니다.
- 내부적으로 `md_formatter.py`를 먼저 실행해 `input_formatted.md` 형태의 중간 파일을 만든 뒤, 그 파일로 Word 변환을 진행합니다.
- 특히 아래 같은 클립보드 원문(Deep Research 복붙)에 효과적입니다.
  - 제목/소제목 경계가 없는 긴 문장
  - 콜아웃 라벨(`[시사점]`, `[요약]` 등)이 본문에 붙어 있는 경우
  - 수식(`$...$`, `$$...$$`)과 볼드(`**...**`)가 섞여 줄바꿈이 깨진 경우

사전 포맷팅 시 주요 정리 항목:

- 제목/소제목 패턴 감지 후 줄바꿈 삽입
- 문장 경계 기준 문단 분리
- 콜아웃/불릿 라인 정리
- LaTeX/볼드 토큰 보호 후 복원
- 메타데이터를 YAML frontmatter로 정리

언제 쓰면 좋나요?

- 원문이 거의 1~5줄 내외로 붙어 있을 때
- Word 변환 결과에서 문단/제목이 비정상적으로 이어질 때

언제 생략해도 되나요?

- 이미 markdown 구조가 잘 잡혀 있고(제목/문단/표가 정상), 바로 변환해도 결과가 괜찮을 때

미리 확인만 하고 싶다면:

```bash
uv run md_formatter.py --check input.md
```

## 변환기 CLI (`md_to_word.py`)

```bash
uv run md_to_word.py [input_file] [output_file] [options]
```

옵션:

- `-l, --list`: 상위 폴더의 markdown 파일 목록 표시
- `-i, --interactive`: 목록에서 대화형 선택 (`--list`와 함께 사용)
- `--batch`: 지정한 디렉터리의 `.md` 파일 전체 변환
- `-f, --format`: 변환 전에 formatter 실행
- `--deepresearch-cleaner {off,auto,on}`: DeepResearch 마커 정리기 적용 (`off` 기본)
- `--cite-mode {footnote,inline,strip}`: 인용 마커 변환 방식
- `--drop-unknown-markers`: 알 수 없는 DeepResearch 마커 블록 제거 (기본은 주석 보존)
- `--cleaner-report`: 정리기 실행 요약 출력
- `--no-cover`: 표지 생략
- `--no-toc`: 목차 생략
- `--no-disclaimer` / `--no-disc`: 디스클레이머 생략
- `--separator-mode {auto,rule,page-break}`: separator를 수평선 또는 페이지 나누기로 렌더링
- `--theme <name|path>`: `themes/<name>.yaml`의 스타일 프로필 또는 지정한 YAML 경로 적용
- `--preset <name|path>`: 문서 유형 프리셋(`ib-report`, `termsheet`, `legal-memo`, `lecture-note` 또는 YAML 경로) 적용; 명시한 CLI 옵션이 프리셋 값보다 우선
- `--charts`: ` ```chart ` fenced YAML 블록을 차트 이미지로 렌더링 (기본값: 끔, 코드 패널로 렌더링)
- `-v, --verbose`: 디버그 로그 출력

예시:

```bash
uv run md_to_word.py --list
uv run md_to_word.py --list -i
uv run md_to_word.py "네페스_기업분석2026.md"
uv run md_to_word.py report.md --format --no-toc
uv run md_to_word.py report.md --deepresearch-cleaner auto --cite-mode strip --cleaner-report
uv run md_to_word.py report.md --theme default
uv run md_to_word.py report.md --preset termsheet
uv run md_to_word.py reports/ --batch
```

```chart
chart_type: bar
labels: ["2025", "2026"]
series:
  - name: 매출
    values: [100, 120]
```

실무 팁:

- `--separator-mode auto`에서는 plain `---`는 수평선으로 유지되고, `## ---`는 페이지 나누기로 렌더됩니다.
- frontmatter가 없어도 첫 H1과 선행 bold 메타 문단에서 제목/날짜/분석 메타를 추론합니다.
- cover의 `INSTITUTION`은 분석대상 회사를 반영할 수 있지만, disclaimer/header/footer 등 문서 브랜딩은 house company identity를 유지합니다.
- 테마는 `themes/*.yaml`에 두며, 기본 제공 `default.yaml`은 코드에 내장된 기본 스타일과 같습니다. 사용자 테마에서는 공개 `IBStyle` 필드 중 필요한 항목만 덮어쓸 수 있고, 알 수 없는 키는 허용되지 않습니다.
- 프리셋은 `presets/*.yaml`에 있으며 문서 유형별 표지, 목차, 디스클레이머, separator mode, 테마의 기본값을 지정합니다. `--no-cover`, `--separator-mode`, `--theme`처럼 CLI에서 직접 지정한 옵션은 언제나 프리셋 값보다 우선합니다.

## 포맷터 CLI (`md_formatter.py`)

파일 포맷팅:

```bash
uv run md_formatter.py input.md
```

출력 파일 지정:

```bash
uv run md_formatter.py input.md output_formatted.md
```

포맷 필요 여부 확인:

```bash
uv run md_formatter.py --check input.md
```

DeepResearch 정리 옵션과 함께 포맷팅:

```bash
uv run md_formatter.py input.md output_formatted.md --deepresearch-cleaner auto --cite-mode inline --cleaner-report
uv run md_formatter.py input.md --deepresearch-cleaner on --cite-mode strip --drop-unknown-markers
```

스크립트 엔트리포인트 사용:

```bash
uv run md-format --check input.md
```

## Word → Markdown 추출(지원 범위 외)

Word → Markdown 추출은 v2.0.0부터 의도적으로 지원 범위에서 제외됩니다. 이 기능이 필요하면 HWP/HWPX/PDF/DOCX를 Markdown으로 변환하고 CLI와 MCP 서버로 제공되는 [kordoc](https://github.com/chrisryugj/kordoc) 같은 전용 파서나 유사 도구를 사용하세요.
제거된 Word-to-Markdown 구현은 참고용으로 `archive/word-to-md-final` 브랜치에 보관되어 있습니다.

## 지원 Markdown 패턴

- 제목: `#`, `##`, `###`, `####`
- 번호형 제목/구조 라인
- 문단, 리스트
- 표(일반/금융/리스크/민감도 패턴)
- 인용구 기반 콜아웃
- 이미지:
  - `![alt](path/to/image.png)`
  - Base64: `![alt](data:image/png;base64,...)`
- LaTeX:
  - 인라인: `$E=mc^2$`
  - 블록: `$$\\int_a^b f(x)dx$$`

## 권장 워크플로우

1. 단일 라인 markdown이면 먼저 구조화:
   `uv run md_formatter.py raw.md`
2. 포맷된 markdown을 Word로 변환:
   `uv run md_to_word.py raw_formatted.md`
3. Word에서 목차 필드(TOC) 업데이트

## 테스트 및 점검

테스트 실행:

```bash
uv run pytest tests/ -v
```

타입 체크:

```bash
uv run mypy ib_renderer.py md_formatter.py md_parser.py md_to_word.py
```

## 참고 사항

- 출력 파일이 Word에서 열려 잠겨 있으면 타임스탬프가 붙은 파일명으로 자동 저장됩니다.
- LaTeX 렌더링은 기본 설치에 포함된 `matplotlib`를 사용합니다. 런타임에서 사용할 수 없으면 fallback 처리됩니다.
- 한글 문서 안정성을 위해 `utf-8`, `utf-8-sig`, `euc-kr`, `cp949` 인코딩 fallback을 사용합니다.

## Markdown 문단 정규화 정책

`md_parser.py`의 최근 동작 변경 사항:

- 문단 내부의 soft wrap 줄바꿈은 공백으로 병합되어 하나의 문단으로 처리됩니다.
- 명시적 hard break(`<br>` 또는 줄 끝 `\\`)만 줄바꿈으로 유지됩니다.
- 줄 끝 공백 2칸 기반 hard break는 기본 비활성화이며, `MarkdownParser(preserve_trailing_double_space_break=True)`로만 활성화됩니다. (복붙/OCR 원문에서 우발적 `↵` 생성을 줄이기 위함)
- 문단 텍스트는 과도한 연속 공백/탭을 정규화하고, 괄호 내부 불필요 공백을 정리합니다. (예: `( PFV ) -> (PFV)`)

검증용 테스트는 `tests/test_md_parser.py`에 추가되어 soft wrap 병합, hard break 정책, legacy 옵션 동작, 공백 정규화를 확인합니다.
