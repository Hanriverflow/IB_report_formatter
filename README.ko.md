<p align="right">
  <a href="./README.md"><img alt="lang English" src="https://img.shields.io/badge/lang-English-blue"></a>
  <a href="./README.ko.md"><img alt="lang 한국어" src="https://img.shields.io/badge/lang-한국어-orange"></a>
</p>

# IB Report Formatter

Markdown ↔ Word 양방향 변환기로, IB(투자은행) 스타일의 전문 Word 보고서(`.docx`)를 생성합니다.

이 프로젝트는 리서치/내부 메모 형태의 markdown을 구조화된 제목, 표 스타일링, 콜아웃 박스, 이미지, 수식, 헤더/푸터가 포함된 보고서 형태로 출력합니다. 또한 역방향으로 Word 문서에서 깔끔한 Markdown을 추출하여 LLM에 활용할 수 있습니다.

## 주요 기능

- **Markdown → Word** 변환 (IB 스타일 문서 생성)
- **Word → Markdown** 변환 (LLM 활용용, 신규!)
- 단일 라인(클립보드) markdown 자동 구조화 포맷팅
- YAML frontmatter 파싱 (`title`, `date`, `recipient`, `analyst` 등)
- 금융 표 렌더링(천 단위 콤마, 조건부 스타일)
- 콜아웃 박스 렌더링 (`[요약]`, `[시사점]`, `[주의]`, `[참고]` 등)
- 이미지 삽입(파일 경로, Base64 `data:image/...`)
- LaTeX 수식 지원 (`$inline$`, `$$block$$`, matplotlib 사용 시 이미지 렌더링)
- 헤더/푸터 구성(회사명, `CONFIDENTIAL`, 페이지 번호)

## 프로젝트 구조

```text
IB_report_formatter/
├── md_to_word.py      # Markdown → Word 변환 CLI
├── md_parser.py       # Markdown/frontmatter/요소 파서
├── md_formatter.py    # 단일 라인 markdown 전처리기
├── ib_renderer.py     # Word 렌더러 및 스타일 시스템
├── word_to_md.py      # Word → Markdown 변환 CLI (신규!)
├── word_parser.py     # Word 문서 파서
├── md_renderer.py     # Markdown 텍스트 렌더러
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

**선택:** 전체 기능 설치 (LaTeX 렌더링 + 인코딩 보강):

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
- `word_to_md.py`
- `word_parser.py`
- `md_renderer.py`
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
- `-f, --format`: 변환 전에 formatter 실행
- `--no-cover`: 표지 생략
- `--no-toc`: 목차 생략
- `--no-disclaimer` / `--no-disc`: 디스클레이머 생략
- `-v, --verbose`: 디버그 로그 출력

예시:

```bash
uv run md_to_word.py --list
uv run md_to_word.py --list -i
uv run md_to_word.py "네페스_기업분석2026.md"
uv run md_to_word.py report.md --format --no-toc
```

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

스크립트 엔트리포인트 사용:

```bash
uv run md-format --check input.md
```

## Word → Markdown 변환기 CLI (`word_to_md.py`)

Word 문서를 LLM 활용에 적합한 깔끔한 Markdown으로 변환합니다:

```bash
uv run word_to_md.py [input_file] [output_file] [options]
```

옵션:

- `-l, --list`: 상위 폴더의 Word 파일 목록 표시
- `-i, --interactive`: 목록에서 대화형 선택 (`--list`와 함께 사용)
- `-s, --strip`: 서식 제거 (볼드/이탤릭 없음) - LLM 최적화
- `--no-frontmatter`: YAML 메타데이터 헤더 생략
- `--extract-images`: 포함된 이미지를 폴더로 추출
- `-v, --verbose`: 디버그 로그 출력

예시:

```bash
uv run word_to_md.py --list
uv run word_to_md.py --list -i
uv run word_to_md.py report.docx
uv run word_to_md.py report.docx output.md
uv run word_to_md.py report.docx --strip              # LLM 최적화 출력
uv run word_to_md.py report.docx --strip --no-frontmatter
uv run word_to_md.py report.docx --extract-images     # 이미지 폴더로 저장
```

`--strip` 옵션은 언제 쓰나요?

- 볼드/이탤릭 마커가 필요 없는 LLM에 넣을 때
- 더 깔끔하고 간결한 텍스트가 필요할 때
- RAG/임베딩 파이프라인에서 서식이 노이즈일 때

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
- LaTeX 렌더링은 `matplotlib`가 필요합니다 (`uv sync --extra full`). 없으면 fallback 처리됩니다.
- 한글 문서 안정성을 위해 `utf-8`, `utf-8-sig`, `euc-kr`, `cp949` 인코딩 fallback을 사용합니다.
