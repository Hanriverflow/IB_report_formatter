# 구현 및 검증 결과 — 2026-09-14

## 결론

계획의 코드 구현·문서·예제·패키징과 실제 Word 페이지 검증까지 완료했다. 자동 테스트 251개, 구조·정적 검사, 별도 환경 설치 검증을 통과했다. 6개 프로파일과 긴 표·가로 표의 **14페이지 전체를 확인**했으며, Word에서 목록·표를 편집하고 저장/재열기하는 검증도 통과했다. 이는 회사별 실제 양식 승인이나 모든 입력의 무조건적인 인쇄 품질 보장을 의미하지 않는다.

Word→MD는 사용자 결정에 따라 폐기했다. 삭제한 구현을 다른 이름으로 재개하지 않았으며 `docx-audit`는 생성된 DOCX의 제한적인 구조 검사만 수행한다.

## 변경 범위

- 프로파일 6종, 공통 모델·불변 설정, 요청별 스타일 격리, 단일 CLI/API/레지스트리 조립 경로.
- 일반 문서: A4, 중립 기본값, 회사 공문 필드·붙임·발신, 보고 작성정보, 회의 참석정보, 네이티브 다단계 번호.
- IB/공통: 표의 명시적 열 의미·단위·출처·기준일·가로 구역·반복 머리행·기준 셀, 외부 링크·숫자 각주.
- 빈 셀·이스케이프 파이프·전체 빈 행 보존, 굵은 금융 숫자 출력·음수 색상 수정, 본문의 ‘출처’가 이후 표를 숨기던 기존 결함 수정.
- 엄격 검사에서 알려진 입력 손실·각주 계약 오류·렌더 실패·이미지/도식 실패를 성공으로 처리하지 않음.
- 버전 2.0.0, README 두 언어, AGENTS, 변경이력, 구 계획의 폐기 안내, 6종 예제, 코드 한정 배포 설정.
- 완료 점검에서 긴/잘못된 YAML 처리, 파일 파싱의 프로파일 우선순위, UTF-8·BOM·EUC-KR·CP949 한글 인코딩, MD 기준 상대 이미지 경로를 보강했다.
- 제목·표 머리글의 명시적 각주, 일반 위첨자 보존, 갱신 시 중복되지 않는 목차 필드, 중립 제목·명시적 글꼴 우선순위, 짧은 IB 메모 간격과 구역 너비별 머리말 정렬을 수정했다.

## 자동 검증

환경: Windows, Python 3.12.12, pytest 9.0.2. 테스트 임시 경로는 매 실행마다 OS 임시 폴더 아래 새 절대 경로를 사용했다.

| 검사 | 실제 결과 |
|---|---|
| 변경 전 전체 테스트 | 284 passed |
| 결함 재현 테스트 3개 | 수정 전 모두 실패: 빈 셀 이동, escaped pipe 열 누락, 굵은 숫자 포맷 미적용 |
| 최종 전체 테스트 | **251 passed** (완료 점검 실행 7.33s) |
| 구성 | 유지/조정된 기존 183개 + 신규 프로파일/회귀 68개; 역변환 관련 기존 101개 제거 |
| Ruff | 변경한 Python·테스트 16개 파일에서 All checks passed |
| mypy | 핵심 9개 모듈에서 Success: no issues found |
| Python 3.8 문법 | wheel 내 14개 엔진 모듈을 `ast.parse(feature_version=(3, 8))`로 검사해 통과 |
| wheel/sdist 빌드 | `uv build` 성공 |
| 패키지 구성 | 엔진 14개 모듈, console script 4개; 폐기 모듈·개인 보고서·QA 산출물 없음 |
| 별도 환경 설치 | wheel과 선언된 의존성을 OS 임시 폴더의 별도 가상환경에 설치; 저장소 밖에서 실제 설치 모듈 경로 확인 |
| 설치된 CLI | `md-to-word`로 공문 생성 → `docx-audit` 오류 0, 종료 코드 0 |
| Word 페이지 | 8개 DOCX, 총 14페이지를 필드 갱신·재열기·PNG 출력 후 육안 확인 |
| Word 편집 | 목록 새 항목 `3.`, 표 코드 `000777` 저장/재열기, 링크 1개·각주 1개 유지 |

Ruff 대상: `document_model`, `document_profiles`, `render_styles`, `office_layout`, `docx_audit`, `md_parser`, `ib_renderer`, `md_to_word`, `converters` 및 변경한 테스트 7개.

mypy 대상은 위 엔진 9개이며 `--follow-imports=silent`를 사용했다. 전체 저장소가 타입 안전하다는 주장이나 미타입 함수의 완전한 검증은 아니다. Python 3.8은 문법 검사이며 **실제 3.8 런타임 테스트는 수행하지 않았다.**

## 시나리오 검증

테스트는 다음을 실제 파싱/렌더링 경로로 확인한다.

- 6개 프로파일, 구조화 YAML, CLI 우선순위, 상대 테마 경로, 잘못된 옵션 거부.
- CLI/API/레지스트리의 문서 XML과 머리말 일치, 표지·최종 면책 모두 끄기.
- 일반 문서의 IB 메타데이터 미삽입, References 미삭제·미중복, 코드 선행 0 보존.
- 편집 가능한 번호 정의와 가나다 수준, 공문 제목 중복 방지, A4 구역 크기.
- 네이티브 링크·각주, 정의 누락/중복/예약 ID 0의 오류 처리.
- 금융 수치의 실제 저장 후 표 출력, 명시한 열 의미·기준 셀·위험 등급, 가로 표 이후 원래 구역 복구.
- 입력 열 손실·요소 렌더 실패·비어 있는 도식에 대한 엄격 모드 실패와 저장 차단.
- 동시 요청 간 스타일 격리, 오류 후 스타일 복구, 전역 호환 프록시의 변경 차단.
- 한글 파일명 CLI, 폐기된 입력/출력 방향 거부, 배포 예제 6종의 엄격 렌더링.

## 생성된 예제 DOCX 구조 검사

예제의 내용과 금융 수치는 모두 가상 데이터다. 실행한 파일은 `samples/profiles/*.md`와 `samples/qa/*.md`이며, QA 산출물은 git에서 제외된 `.dryforge/qa/completion/`에 있다.

| 예제 | 본문 문단 | 표 | 네이티브 번호 문단 | 알려진 구조 오류 |
|---|---:|---:|---:|---:|
| business-report | 17 | 2 | 2 | 0 |
| ib-memo | 18 | 2 | 0 | 0 |
| ib-report | 42 | 3 | 0 | 0 |
| meeting-minutes | 17 | 1 | 5 | 0 |
| office-letter | 20 | 1 | 5 | 0 |
| plain | 10 | 1 | 4 | 0 |
| long-table | 7 | 1 | 0 | 0 |
| landscape | 11 | 1 | 0 | 0 |

표 개수에는 IB 표지 구성용 표도 포함된다. `docx-audit`는 모든 OOXML 규격·누락 콘텐츠·시각적 문제를 포괄하지 않으며, 위 숫자는 페이지 수가 아니다.

## 페이지 시각검증: 완료

`documents` 스킬의 render-and-verify 절차를 적용하려고 번들 Python과 `render_docx.py`를 실행했으나 `LibreOffice soffice.exe was not found on PATH`로 실패했다. 사용자 시스템에 새 LibreOffice를 설치하거나 기존 설치를 변경하지 않았다.

숨김 Word 자동화의 PDF 내보내기도 기본 한 줄 문서에서 완료되지 않았다. 이 문제를 특정 프로파일 결함으로 단정하지 않고 조사한 뒤, Word의 공식 `Page.EnhMetaFileBits` 경로로 실제 페이지를 출력했다. 따라서 이전의 ‘시각검증 미완료’ 상태는 해소됐다. PDF 내보내기 성공을 주장하는 것은 아니다.

재현 스크립트: `scripts/word_visual_qa.ps1`. Windows PowerShell 5.1과 Word 365 16.0.20326.20144에서 실행했다. 원본 read-only 열기 → 필드/목차 갱신 → QA 복사본 저장·재열기 → 전체 페이지 EMF 캐시 준비 → 바닥글 필드 갱신 → 각 페이지 머리말 캐시 갱신 → PNG 출력 순서다. 단순 `Repaginate`만 사용했을 때 쪽수가 `1 of 2 → 2 of 3`으로 증가하고 반복 머리말이 빠지는 EMF 캐시 문제를 위 순서로 해결했다. 문서의 필드를 정적 숫자로 바꾸거나 이미지를 합성하지 않았다.

| 문서 | 실제 페이지 | 확인 결과 |
|---|---:|---|
| business-report | 1 | 중립 제목, 작성정보, 두 표, 번호, 쪽번호 정상 |
| meeting-minutes | 1 | 참석정보, 번호·글머리표, 담당자·기한 표 정상 |
| office-letter | 1 | 수신·참조·시행일·붙임·발신인·담당 연락처, 가나다 중첩 번호 정상 |
| plain | 1 | 선행 0·빈 셀·파이프 문자, 링크, 페이지 하단 각주 정상 |
| ib-memo | 1 | 금융 숫자·음수·기준 셀·출처·각주 정상; 초기 2페이지의 마지막 항목 밀림은 메모 간격 조정 후 해소 |
| ib-report | 4 | 표지 1, 목차 2, 본문 3, 면책 4; 목차 중복 없음, 페이지 참조·머리말·Page n of 4 확인 |
| long-table | 2 | 001~065 순서 유지; 두 번째 페이지 머리행 반복, 행 분할 없음, 표 뒤 출처·본문 보존 |
| landscape | 3 | A4 세로→가로→세로, 전체 열·음수·출처 표시, 구역별 머리말 폭·연속 쪽번호 정상 |

최종 PNG와 페이지 목록은 `.dryforge/qa/completion/verified-pages/` 및 `manifest.json`, 원래 생성 DOCX는 `release-docx/`에 있다. 문단/표가 페이지 밖으로 잘리거나 서로 겹치는 문제, 한글 누락, 뜻하지 않은 빈 페이지는 확인되지 않았다. 페이지 기하 검증용 가로 표의 앞뒤 세로 페이지 여백은 의도된 테스트 구성이다.

별도 Word 편집 검증은 `word-edit-smoke.ps1`로 수행했다. 원본 plain 문서의 두 번째 최상위 목록 뒤에 항목을 삽입하고 표 코드를 수정한 QA 복사본을 저장·재열었다. 추가 항목은 네이티브 `3.`으로 이어지고, `000777`과 외부 링크 1개·각주 1개가 유지됐다. Word의 `ListParagraphs` 열거 순서는 문서 순서와 달라 번호 속성으로 대상 문단을 선택했다.

검증용으로 만든 Word/보조 프로세스만 종료했다. 사용자 문서는 열거나 저장하지 않았고 프린터·추가 기능·레지스트리 등 전역 설정은 변경하지 않았다. 이 스크립트는 독립 로컬 QA 프로세스용이며 장기 무인 Office 서버용이 아니다. 다른 OS·Word 버전·누락 글꼴 환경과 모든 신규 입력은 별도 확인이 필요하다. 실제 문서에 재무·법률적 정확성을 보증하지 않는다.

공식 API 근거: [Page.EnhMetaFileBits](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.page.enhmetafilebits?view=word-pia), [Document.Repaginate](https://learn.microsoft.com/en-us/office/vba/api/word.document.repaginate).

재현 명령(프로젝트 루트에서):

```powershell
$testRoot = Join-Path $env:TEMP ('ib-report-tests-' + [guid]::NewGuid().ToString('N'))
uv run pytest tests/ -q --basetemp $testRoot
uv run md_to_word.py samples/profiles .dryforge/qa/completion/release-docx --batch --strict
uv run md_to_word.py samples/qa .dryforge/qa/completion/release-docx --batch --strict
powershell.exe -NoProfile -File scripts/word_visual_qa.ps1 -InputDirectory .dryforge/qa/completion/release-docx -OutputDirectory .dryforge/qa/completion/verified-pages
uv build
```

## 변경 관리와 잔여 범위

- 구현 브랜치: `codex/md-word-profiles`. 커밋·push·PR·외부 배포는 수행하지 않았다.
- 제거 전 소스: `d819bbb`, 복구용 로컬 브랜치 `codex/archive-word-to-md-d819bbb`.
- 기존 미추적 보고서·개인 계획·사용자 이미지·접근 불가 임시 디렉터리는 변경/삭제하지 않았다.
- 임의 DOCX 양식 자동 이해, 신규 차트 엔진, 편집형 차트, 재무 계산, HWPX, AI 작성 스킬, 자동 PDF 배포는 이번 구현 범위 밖이다.
- 이번 확정 범위의 미완료 구현은 없다. 향후 회사 실제 참조 양식 기반 템플릿 정밀화나 별도 IB 표/차트 확장은 신규 범위이며, Word→MD는 포함하지 않는다.
