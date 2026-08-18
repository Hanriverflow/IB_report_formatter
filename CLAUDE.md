# CLAUDE.md — IB Report Formatter 운영 매뉴얼

이 문서는 이 저장소에서 작업하는 모든 AI 에이전트의 운영 매뉴얼이다.
목표: 어떤 모델이 오든 같은 품질로 일하게 만드는 것. 판단이 서지 않으면 §6(에스컬레이션)을 따른다.
코드 스타일 상세(예시 포함)는 [AGENTS.md](AGENTS.md) 참조. 두 문서가 충돌하면 이 문서가 우선한다.

## 0. 이 프로젝트는 무엇인가

- Markdown → Word(.docx) 변환기. 한국어 IB/DCM 스타일 보고서를 생성한다.
- **실제 업무 도구다.** `reports/`의 파일들은 실제 딜 자료다(민감정보 — §2-M9). 장난감 리포처럼 다루지 말 것.
- 아키텍처 허브는 `DocumentModel`(md_parser.py에 정의). 모든 파서가 이것을 생산하고, 모든 렌더러가 소비한다.

| 방향 | CLI | 파서 | 렌더러 |
|------|-----|------|--------|
| MD → Word | `md_to_word.py` | `md_parser.py` | `ib_renderer.py` |

보조 모듈: `md_formatter.py`(한 줄 붙여넣기 정형화), `deep_md_cleaner.py`(DeepResearch 마커 정리),
`converters.py`(플러그인 레지스트리, MD 입력 → Word 출력 전용), `diagram_renderer.py`,
`cli_utils.py`, `stream_utils.py`.

## 1. 철칙 (Iron Rules)

이 중 하나라도 어기게 되는 상황이면 **작업을 멈추고 사용자에게 물어라** (§6의 질문 형식 사용).

1. **기본 동작 보존.** 플래그 없이 실행한 기존 입력의 출력이 달라지는 변경은 금지.
   새 동작은 opt-in 플래그로 추가하고 기본값은 off (전례: `--deepresearch-cleaner`의 off/auto/on 설계).
2. **페어 계약.** 요소의 의미(파싱·렌더링 방식, ElementType)를 바꾸면
   짝(파서 `md_parser.py` ↔ 렌더러 `ib_renderer.py`)을 같은 작업에서 맞춘다. 한쪽만 고치고 끝내지 않는다.
   `_ibrep_` 시맨틱 북마크는 출력물에 남아있는 안정적 공개 계약이다 — 외부 파서가 이를 소비하므로
   임의로 형식을 바꾸지 않는다.
3. **폴백은 기능이다.** 인코딩 사다리(utf-8→euc-kr→cp949), matplotlib 부재 폴백, PermissionError
   타임스탬프 저장, 요소 단위 try/except — "죽은 코드처럼 보여도" 삭제 금지. 다른 머신에서 살아있는 코드다.
4. **민감정보 격리.** `reports/`, `GPT_deep/`, `*.docx` 산출물은 절대 커밋하지 않고,
   내용을 테스트·커밋 메시지·문서에 인용하지 않는다.
5. **Python 3.8 하한.** 3.9+ 문법·API 금지 (§2-M3 목록 참조).
6. **실행은 `uv run`으로만.** `python foo.py` 직접 실행 금지 (시스템 파이썬은 의존성이 다르다).
7. **기존 테스트의 기대값을 바꿔서 통과시키지 않는다.** 기존 테스트가 빨간불이면 그 테스트가 보호하던
   계약이 깨졌다는 뜻이다. 기대값 수정은 사용자 승인 사항.

## 2. 명명된 실수들 (Named Failure Modes)

각 실수는 이 저장소에서 실제로 발생했거나 구조적으로 유발되기 쉬운 것들이다. 이름 → 증거 → 규칙 순.

### M1. NUL 파일 사고 (Windows 셸 혼동)
- **증거**: 저장소 루트의 `nul` 파일. 내용이 `del: command not found` — 과거 에이전트가 Git Bash에서
  Windows 명령을 실행하고 `> nul` 리다이렉트로 파일까지 만들었다.
- **규칙**: Bash 도구 = Git Bash = POSIX만 (`rm`, `/dev/null`, 슬래시 경로). `del`, `> nul`, `%VAR%` 금지.
  PowerShell 5.1에는 `&&`/`||`가 없다 → `if ($?) { }`. 출력 버리기는 `> /dev/null`(Bash) 또는 `| Out-Null`(PS).
  한국어 파일명 인자는 항상 따옴표로 감싼다.
- **검출**: 작업 후 `git status`에 `nul`/`NUL`이 보이면 당신이 만든 것이다. 삭제하고 명령을 고쳐라.

### M2. 게이트 우회 (Gate Bypass)
- **증거**: `_ibrep_` 시맨틱 북마크는 파서(`md_parser.py`)와 렌더러(`ib_renderer.py`)가 짝으로 의존하는
  공개 계약이다. `tests/test_docx_gates.py`가 이 계약을 스냅샷 + XML 게이트로 고정한다.
- **규칙**: 요소 렌더링/파싱을 바꾸면 `uv run pytest tests/test_docx_gates.py -q`가 그린이어야 한다.
  정당한 시맨틱 변경 때문에 내장 스냅샷을 갱신해야 한다면, 그것은 테스트 기대값 변경(철칙 7) —
  사용자 승인 후 같은 커밋에서 갱신한다. 새 요소 타입은 AGENTS.md의 5단계(enum→dataclass→파싱→렌더링→등록)를 전부 밟는다.

### M3. 3.9 침입 (py38 위반)
- **규칙**: 금지 목록 — `Path.with_stem()`, `str.removeprefix/removesuffix`, `dict | dict`,
  `match` 문, 내장 제네릭 주석(`list[str]` — `from typing import List` 스타일 유지), `functools.cache`.
- **검출**: `uv run mypy`(python_version=3.8)와 ruff(target py38)가 잡는다. §3-A 게이트 통과 필수.

### M4. 몰래 기본값 변경 (Default Drift)
- **증거**: roadmap.md의 설계 원칙 "기본 동작 보존: 옵션 미사용 시 기존 결과와 동일해야 함".
- **규칙**: 철칙 1. 엣지 케이스를 고치려고 공통 경로의 동작을 바꾸지 않는다.
- **검출**: 수정 전후로 `uv run md_to_word.py tests/웅진_계열사.md`가 성공하고, 의도하지 않은
  요소 변화가 `tests/test_docx_gates.py`에 나타나지 않아야 한다.

### M5. 인코딩 러시안룰렛 (cp949 Mojibake)
- **증거**: Windows 콘솔은 cp949다. `python -c`로 한국어를 print하면 그대로 깨진다(이 매뉴얼 작성 중에도 재현됨).
  CLI들이 `cli_utils._configure_text_stream`으로 UTF-8을 강제하는 이유다.
- **규칙**: 파일 I/O는 항상 `encoding="utf-8"` 명시. 디버그 출력 스크립트는 첫 줄에
  `sys.stdout.reconfigure(encoding="utf-8")`. 텍스트 경로를 건드리는 모든 변경은 한국어 문자열 테스트 케이스를 포함한다.
  ASCII로만 테스트하고 통과라고 선언하지 않는다.

### M6. 문서 화석화 (Stale Docs)
- **증거**: AGENTS.md에 "267 tests", 커밋 본문에 279, 실제는 284 (2026-07-07 측정). next_step.md는 "250/250".
- **규칙**: 문서에 개수·상태를 손으로 쓰지 않는다. 이미 있는 숫자를 갱신할 때는 반드시 실측값
  (`uv run pytest -q | tail -1`)을 같은 커밋에서 반영한다. CLI 플래그·기능 변경 시 README.md와
  README.ko.md를 **같은 커밋에서** 같은 내용으로 갱신한다 (두 파일은 번역 쌍이다).
- **예외**: `plan.md`, `next_step.md`, `roadmap.md`, `word_to_md_plan.md`는 과거 계획 스냅샷이다.
  요청 없이 "친절하게" 갱신하거나 삭제하지 않는다.

### M7. 폴백 청소부 (Fallback Deletion)
- **증거**: `samples/demo_pitchbook_1775575235.docx` — PermissionError 타임스탬프 폴백이 실제로 발동한 산출물.
- **규칙**: 철칙 3. 커버리지 0%로 보이는 폴백도 삭제 제안만 하고 실행은 승인 후에.

### M8. 잘못된 층 수술 (Wrong-layer Fix)
- **증거 패턴**: 파서 버그를 렌더러에서 문자열 후처리로 덮는 수정. 증상은 사라지고 원인은 남는다.
- **규칙**: 변환 버그는 반드시 `/fix-render-bug` 절차로 DocumentModel을 덤프해 어느 층(정형화→파싱→렌더링)이
  틀렸는지 확정한 뒤, 그 층에서만 고친다.

### M9. git add . 사고 (Staging Blast Radius)
- **증거**: `reports/`(실제 딜 문서)가 현재 untracked 상태로 존재한다. `git add .` 한 번이면 유출이다.
- **규칙**: 스테이징은 항상 명시적 파일 경로로. `git add .`/`-A` 금지. 커밋 직전
  `git diff --cached --name-only` 출력에서 `reports/`, `*.docx`, `GPT_deep/`, `nul` 이 보이면 즉시 unstage.

### M10. 대수술 diff (Monolith Rewrite)
- **증거**: `ib_renderer.py` 3,000+줄, `md_parser.py` 2,100+줄. 넓은 리라이트는 리뷰 불가능한 diff를 만든다.
- **규칙**: 목적과 무관한 줄은 건드리지 않는다(포매팅·정렬·이름 변경 포함). 리팩터는 동작 불변을
  스위트로 증명하는 **별도 커밋**. `git diff --stat`에 의도하지 않은 파일이 나타나면 안 된다.

### M11. PUA 마커 훼손 (Invisible Marker Corruption)
- **증거**: deep_md_cleaner.py는 DeepResearch 마커(U+E200~U+E202)를 다룬다. 소스 주석에 명시된 원칙 —
  "no literal glyphs in source".
- **규칙**: 소스·테스트에 PUA 문자를 리터럴로 붙여넣지 않는다. 항상 `""` 이스케이프로 쓴다.
  픽스처에서 "보이지 않는 문자 정리"를 하지 않는다 — 그 문자가 테스트 대상이다.

### M12. 드라이브바이 린트 청소 (Baseline Noise)
- **증거**: 기존 코드에 ruff 43개, mypy 21개 에러가 있다(§3-D 기준선). 기능 커밋에서 이걸 건드리면
  diff가 오염되고 회귀 위험이 생긴다.
- **규칙**: 당신이 추가·수정한 줄의 에러만 0으로 만든다. 기존 에러 정리는 요청받았을 때 별도 `chore:` 커밋으로.

## 3. 산출물별 품질 기준

형용사 금지. 아래 명령과 기대 출력이 기준이다. **모든 기준은 실행 결과로 증명한 뒤에만 "완료"를 선언한다.**

### A. 모든 코드 변경 (공통 게이트) — `/preflight`로 실행
| 검사 | 명령 | 통과 기준 |
|------|------|-----------|
| 테스트 | `uv run pytest -q` | `N passed, 0 failed`, N ≥ 직전 커밋의 개수(감소 시 사유 보고) |
| 린트 | `uv run ruff check .` | 총 에러 ≤ 기준선(§3-D), 당신이 만든 줄의 에러 0 |
| 타입 | `uv run mypy ib_renderer.py md_formatter.py md_parser.py md_to_word.py` | 총 에러 ≤ 기준선, 신규 0 |
| 스모크 | `uv run md_to_word.py tests/웅진_계열사.md smoke_output_pf.docx` | exit 0, 파일 생성(끝나면 삭제) |
| diff 위생 | `git diff --stat` | 의도한 파일만, 무관한 줄 0 |

### B. 버그 수정
- A 전체 + **수정 전에는 실패하는 회귀 테스트**가 같은 커밋에 포함된다.
  (증명: 수정 코드를 잠시 되돌리고 테스트가 빨간불인 것을 확인하거나, 테스트를 먼저 작성하고 빨간불 출력을 보고한다.)
- 수정은 원인 층 하나에서만 이뤄진다 (M8).

### C. 파서/렌더러 변경 (요소 의미에 손대는 모든 변경)
- A + B 전체.
- `uv run pytest tests/test_docx_gates.py -q` 전부 통과 (스냅샷 + XML 게이트).
- Word 출력 검증은 텍스트 비교가 아니라 **XML 레벨 assertion**
  (기존 패턴: `tests/test_ib_renderer.py`의 `_semantic_bookmark_names`, xpath 사용).
- 한국어 텍스트 포함 케이스 ≥ 1개 (M5).

### D. 기준선 (2026-08-19, v2.1.0 리뷰 프리플라이트 측정)
| 항목 | 값 |
|------|-----|
| pytest | **230 passed** |
| ruff (전체) | 28 errors |
| mypy (위 4개 모듈) | 22 errors (mypy 1.14.1) |

수치를 **개선**했으면 같은 커밋에서 이 표와 스킬 속 수치를 갱신한다. 악화는 게이트 실패다.

### E. CLI 플래그 추가/변경
- 기본 동작 불변(철칙 1) + `--help`에 노출 + README.md·README.ko.md 양쪽에 예제 포함 +
  AGENTS.md Quick Reference 갱신. 네 파일이 같은 커밋에 있다.

### F. 문서 변경
- README.md ↔ README.ko.md는 항상 쌍으로. 검증: `grep -c '^## ' README.md README.ko.md` 의 두 값이 같다.
- 개수·버전 표기는 실측값만 (M6).

### G. 커밋
- 형식: `type: 소문자 영어 명령형 요약` (feat/fix/chore/docs/refactor/test).
- 본문: 파일별 무엇·왜. 테스트 개수가 변했으면 `Tests: N → M (+k)` 줄 포함 (전례: 715fb64).
- 트레일러: `Co-Authored-By: Claude <모델명> <noreply@anthropic.com>`.
- 스테이징 검증: `git diff --cached --name-only`에 민감 파일 없음 (M9).

### H. 릴리즈 — `/ship`으로만 실행
- pyproject.toml version == CHANGELOG.md 최신 `## [X.Y.Z]` == uv.lock의 패키지 version. 셋 중 하나라도 다르면 실패.
- 이 저장소는 git tag를 쓰지 않는다. 태그를 만들지 않는다.
- push는 사용자 승인 후에만.

## 4. 관례 (Conventions)

- **커밋 이력이 스펙이다.** 형식이 헷갈리면 `git log`의 최근 feat/fix 커밋을 모방한다.
- **테스트 스타일**: 모듈별 `tests/test_<module>.py`. 파일 상단 docstring에 커버 범위 나열,
  섹션은 `═══` 배너 주석. Word 검증은 xpath. 한국어 픽스처(`tests/웅진_계열사.md`, `tests/일동제약_수익성분석.md`)는 실데이터 기반 — 수정 금지.
- **브랜치**: 소규모는 main 직행, 큰 작업은 `codex/<topic>` 스타일 브랜치 + PR (전례: PR #1).
- **에러 처리 철학**: 요소 하나가 실패해도 문서 전체를 죽이지 않는다 — 로그 + 눈에 보이는
  `[Render Error: ...]` 마커 삽입 (AGENTS.md "Element-Level Resilience").
- **로깅**: print 금지, `logging` 모듈 + `cli_utils.LogFormatter` 프리픽스 체계.
- **의존성 최소주의**: 런타임 의존성은 python-docx, pyyaml, matplotlib 셋뿐. 추가는 사용자 승인 사항.

## 5. 세션 시작 리추얼 + 명령어

매 세션 첫 코드 편집 **전에**:

```bash
git status --short            # 기존 untracked(.python-version, reports/ 등)를 파악. 내 소행과 구분.
uv run pytest -q | tail -1    # 그린 확인 (~10초). 시작부터 빨간불이면 편집하지 말고 즉시 보고.
```

자주 쓰는 명령:

```bash
uv sync --extra dev                                  # 최초 1회 셋업
uv run pytest tests/test_md_parser.py -q             # 단일 파일 테스트
uv run md_formatter.py --check input.md              # 정형화 필요 여부 판단
uv run md_to_word.py input.md --format               # 정형화 + 변환
```

## 6. 불확실할 때 (Escalation)

### 묻지 않고 진행해도 되는 것
- 테스트 추가, 요청받은 버그의 원인 층 최소 수정(+회귀 테스트), README 쌍 동기화,
  요청 범위 내 신규 코드, 스크래치패드에서의 실험.

### 멈추고 물어야 하는 것 (정확한 트리거)
1. 기존 테스트의 기대값을 바꿔야만 통과하는 상황 (철칙 7).
2. 플래그 없이 기본 출력이 달라지는 변경 (철칙 1).
3. 의존성 추가·제거·버전 범위 변경.
4. `IBStyle` 상수(색·폰트·여백) 변경 — 시각 아이덴티티는 사용자 소관.
5. 폴백·에러 처리 경로 삭제 (철칙 3).
6. 파일 삭제, 대량 이동/리네임, `.gitignore`의 민감 경로 항목 수정.
7. `reports/`·`samples/` 내용물 조작.
8. push, PR 생성, 버전 범프, 태그.
9. 같은 실패에 대한 수정 시도 2회 초과.

### 질문 형식 (열린 질문 금지)
```
발견: md_parser.py:1240에서 X가 Y로 처리됨 (테스트 test_z가 이 동작을 고정).
옵션 A: ... / 옵션 B: ...
추천: A — 이유 한 줄.
```

### 검증 실패 프로토콜
수정 2회 시도 후에도 게이트(§3)가 실패하면: **추가 수정 중단**, 작업 트리는 그대로 두고
실패 출력 전문 + `git diff` 요약 + 원인 가설을 보고한다. 통과할 때까지 몰래 반복하지 않는다.
절대 하지 말 것: 실패를 숨기는 커밋, 테스트 스킵 마크, 기대값 완화.

## 7. 프로젝트 스킬

해당 상황이 오면 **반드시** 호출한다 (선택이 아니다):

| 스킬 | 언제 |
|------|------|
| `/preflight` | 커밋 직전, "완료" 선언 직전, 모든 코드 변경 후 |
| `/fix-render-bug` | 변환 결과가 이상할 때 (표·수식·콜아웃 깨짐, 문자 깨짐) |
| `/ship` | 릴리즈·버전 범프 요청 시 |
