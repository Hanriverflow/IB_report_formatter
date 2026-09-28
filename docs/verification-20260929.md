# 입력 손실·저장 안전성 강화 검증 (2026-09-29)

기준: `origin/main` `7b654b6`에서 만든 worktree 브랜치 `codex/quality-hardening-20260929`. 커밋·push·배포는 하지 않았다. Word→MD 폐기 원칙은 유지한다.

## 범위

2026-09-28 분석에서 재현한 결함을 수정했다. 각 항목은 실제 파싱→렌더링→DOCX 회귀 테스트를 먼저 작성해 실패를 확인한 뒤 고쳤다.

| 영역 | 항목 | 회귀 테스트 |
|---|---|---|
| 입력 손실(P1) | `\$` 금액 소실, `\*\*` 리터럴, References 뒤 본문·표 삭제, 선두 굵은 라벨 문단 삭제, `\| - \| - \|` 행 삭제 | `tests/test_parser_hardening.py` |
| 금액 변조(P1) | 숫자 셀 `1234[^1]` → `12,341`·각주 소실 | `tests/test_render_hardening.py` |
| strict 누락(P1) | 수식 렌더 실패가 strict에서 저장됨 | `tests/test_render_hardening.py` |
| 저장(P1) | 비원자적 저장, batch 출력명 충돌 덮어쓰기 | `tests/test_cli_hardening.py` |
| 파서(P2) | 펜스(틸드·4+·문단 직후), CRLF, 목록 연속줄, 이미지 제목, 코드 안 각주, setext H1, HTML 주석, 참조형 링크, 연속 메타데이터 줄 | `tests/test_parser_hardening.py` |
| 렌더·감사(P2~P3) | 중첩 번호 재시작, Matplotlib 전역 상태, 렌더 경로 감사 성능, 자리표시 문자열 오탐, numId 무결성 | `tests/test_render_hardening.py` |
| CLI(P2) | 잠금 대체 파일명 충돌, 한정 경로 fallback, 레지스트리 초기화 경합, 저장 파일 권한 | `tests/test_cli_hardening.py` |
| CI | Ubuntu/Windows ruff·mypy·pytest·build·wheel 내용·3.8 문법 | `.github/workflows/ci.yml` |

의도된 출력 변경: 선두 메타데이터 키 정규화(`date`, `analysis_period`, `analysis_basis`), CRLF 입력 목록 끝 불필요 줄바꿈 제거, 중첩 번호 목록 인스턴스 분리. 표지를 끈 IB 문서에서 추정 부제목(H1 직후 H2)은 이제 본문 제목으로 남는다(이전에는 소실). 표지를 켠 경우는 기존과 같다.

## 자동 검증

| 검사 | 결과 |
|---|---|
| 기준선(`7b654b6`) | 311 passed, ruff·mypy 통과 |
| Windows, Python 3.12.12, 새 OS 임시 basetemp | **453 passed, 1 skipped**(POSIX 권한 테스트) |
| WSL Ubuntu, Python 3.12, 격리 복사본 | **450 passed, 4 skipped**(Word/Windows 전용), ruff·mypy 통과 |
| Ruff / mypy(14개 모듈) | 통과 / 통과 |
| wheel | 14개 엔진 모듈만 포함 |
| 합성 샘플 9종 `document.xml` | 의도된 변경 외 동일(파서 후속 수정 전후 byte 동일) |
| 로컬 비공개 실보고서 7건(저장소 밖 보관, 결과만 기록) | 기준선 대비 모든 XML 파트 동일(북마크 ID·날짜만 정규화) |
| 성능(2,000요소·표 200개, 같은 PC) | 렌더 13.7초 → 5.5초 |

GitHub Actions 워크플로는 YAML 파싱과 단계 스크립트만 로컬에서 확인했다. GitHub에서 실제 실행은 하지 않았다. Python 3.8 실행은 검증하지 않았고 문법 검사만 했다.

## Word 페이지 검토

Microsoft Word 16.0, build 16.0.20326. `scripts/word_visual_qa.ps1`로 새 출력 폴더 `.dryforge/qa/hardening-20260929/`(git 제외)에 렌더했다. manifest SHA-256 `46D5E530B118D85643BC91E1AFCF93809A224A1D9CB08D49045DF8C1B1C398C4`. manifest의 `visualReview`는 `pending`으로 둔다. 육안 검토 결과는 이 절에 기록한다.

| 문서 | 쪽 | 확인 |
|---|---:|---|
| business-report | 1 | 작성정보·표·번호 목록, 목록 끝 빈 줄 없음 |
| ib-memo | 1 | 음수·기준 셀·코드 선행 0·각주 |
| ib-report | 4 | 표지·목차·본문·면책, Page n of 4 |
| landscape | 3 | 세로→가로→세로, 연속 쪽번호 |
| long-table | 2 | 001~065, 반복 머리행 |
| meeting-minutes | 1 | 참석정보·번호·글머리표·표 |
| office-letter | 1 | 1·2·3 번호 아래 가·나 하위 번호, 붙임 1·2, 끝 |
| office-letter-appendix | 2 | 본문 1~4, 붙임, 2쪽 별첨 표·셀 내 줄바꿈 |
| plain | 1 | 1. 아래 가. 하위 번호, 빈 셀·파이프·링크·각주 |

모든 16쪽에서 잘림·겹침·한글 누락·의도하지 않은 빈 페이지를 관찰하지 못했다. 이는 AI 육안 검토이며 사람·업무 승인이 아니다.

## 남은 과제

- PR #5(`codex/remove-word-to-md`)의 테마·프리셋·차트 이식 여부 결정.
- LICENSE 권리 확인(README 릴리스 게이트 [blocked]).
- 인라인 토크나이저 단일화와 `ib_renderer.py` 모듈 분리(이번 변경은 개별 결함 수정).
- 로컬 dirty 체크아웃 `codex/md-word-profiles` 정리 여부 결정.
