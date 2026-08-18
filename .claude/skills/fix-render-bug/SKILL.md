---
name: fix-render-bug
description: Use when 변환 결과가 이상할 때 — 표·수식·콜아웃·목록이 Word에서 깨져 보임, 한국어가 mojibake로 깨짐, 요소가 사라지거나 다른 타입으로 나옴, 사용자가 "보고서가 이상해/깨졌어"라며 .md나 .docx를 가져올 때.
---

# Fix Render Bug — 변환 버그의 층 확정 → 회귀 테스트 → 최소 수정

## 개요

이 파이프라인은 4개 층을 지난다: **정형화(md_formatter/deep_md_cleaner) → 파싱(md_parser) → DocumentModel → 렌더링(ib_renderer)**.
증상은 항상 마지막(출력물)에서 보이지만 원인은 대개 앞 층에 있다.
**층을 확정하기 전에는 코드를 한 줄도 고치지 않는다.** 렌더러에서 파서 버그를 후처리로 덮는 것(CLAUDE.md M8)이 이 저장소 최악의 수술이다.

## 절차

### 1. 재현 입력 확보 — 민감정보 격리부터

사용자가 준 파일이 `reports/`의 실제 업무 문서라면:
- 그 파일로 **로컬 재현만** 한다. 내용을 테스트 픽스처·커밋·보고서에 복사하지 않는다.
- 증상을 재현하는 **합성 최소 입력**을 스크래치패드에 새로 만든다 (회사명은 "샘플기업", 숫자는 가짜로).

### 2. 최소화

증상이 유지되는 한 입력을 반씩 잘라낸다. 목표: **10줄 이하**의 재현 파일.
최소화가 끝나면 어떤 마크다운 구성요소(표? 볼드 경계? 수식? 콜아웃 라벨?)가 방아쇠인지 이미 반쯤 안 것이다.

### 3. 층 확정 — DocumentModel 덤프

**MD → Word 방향** (검증된 스니펫 — `reconfigure` 줄을 빼면 한국어가 깨진다, M5):

```bash
uv run python -c "
import sys
sys.stdout.reconfigure(encoding='utf-8')
from md_parser import parse_markdown_file
m = parse_markdown_file('repro.md')
for i, e in enumerate(m.elements):
    print(i, e.element_type.name, '|', str(e.content)[:80].replace(chr(10),' '))
"
```

판정표:

| 관찰 | 원인 층 | 고칠 파일 | 테스트 파일 |
|------|---------|-----------|-------------|
| 모델부터 틀림 (타입 오분류, 내용 소실) | 파싱 | md_parser.py | test_md_parser.py |
| 모델은 맞는데 .docx가 틀림 | 렌더링 | ib_renderer.py | test_ib_renderer.py |
| `--format` 거치면 재현, raw는 정상 | 정형화 | md_formatter.py / deep_md_cleaner.py | test_md_formatter.py / test_deep_md_cleaner.py |

입력이 원래 `--format` 대상(한 줄 붙여넣기)인지 먼저 `uv run md_formatter.py --check repro.md`로 확인하라 —
정형화가 필요한 입력을 raw로 변환해 놓고 파서 버그로 오진하는 경우가 흔하다.

### 4. 실패하는 회귀 테스트 먼저

판정표의 테스트 파일에, 기존 스타일(파일 상단 docstring 목록, `═══` 섹션 배너)을 따라 작성한다.

- 렌더러 버그면 텍스트 비교가 아니라 **XML 레벨 assertion** (기존 헬퍼 `_semantic_bookmark_names`, `doc.element.xpath(...)` 패턴 모방).
- 증상에 한국어가 관여하면 테스트 문자열도 한국어로.
- 작성 직후 실행해 **빨간불 출력을 확보**한다. 이 출력이 보고서의 증거다:

```bash
uv run pytest tests/test_md_parser.py -q -k "새테스트이름"
```

### 5. 최소 수정

- 3단계에서 확정한 층 **하나에서만** 고친다.
- 컴파일된 클래스 레벨 정규식(`_BOLD_SPLIT_RE` 등)을 수정할 때는 그 패턴의 다른 사용처를 `grep`으로 전수 확인.
- 엣지 케이스 수정이 공통 경로의 기본 출력을 바꾸면 안 된다(철칙 1). 바꿔야만 한다면 멈추고 에스컬레이션.

### 6. 검증

1. 새 테스트 그린 + `/preflight` 전체 통과 (docx 게이트 3단계 포함).
2. 사용자가 실제 파일을 줬다면 그 파일을 다시 변환해 증상 해소를 확인한다 (출력물은 커밋하지 않는다).

### 7. 보고 양식

```
증상: (한 줄)
원인: <파일>:<줄> — (메커니즘 한두 문장)
층: 정형화/파싱/렌더링
수정: (무엇을 바꿨나)
증거: 회귀 테스트 <이름> — 수정 전 FAIL 출력 확보, 수정 후 PASS. preflight 표 첨부.
```

## 흔한 실수

| 실수 | 대신 이렇게 |
|------|-------------|
| 렌더러에서 문자열 후처리로 파서 버그 은폐 | 3단계 덤프로 층 확정 후 원인 층 수정 |
| ASCII 재현만 만들고 통과 선언 | 한국어 포함 케이스 필수 (M5) |
| 실제 보고서 내용을 테스트에 복사 | 합성 최소 입력 신규 작성 (철칙 4) |
| 기존 테스트 기대값을 고쳐서 통과 | 철칙 7 — 멈추고 에스컬레이션 |
| `python -c` 출력이 깨지자 "인코딩 버그" 오진 | 콘솔 cp949 문제다. 스니펫의 reconfigure 줄 확인 |
| 수정하며 주변 코드 정리 | M10/M12 — 목적 외 줄 불가침 |
