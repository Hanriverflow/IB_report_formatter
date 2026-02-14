# OpenAI DeepResearch 전용 Cleaner 구현 로드맵

## 1) 배경과 범위

- 이 이슈는 일반 Markdown 품질 문제가 아니라, **OpenAI DeepResearch export 마커**(`cite...`, `entity...`, `image_group...`)에 한정된 문제다.
- 따라서 기본 변환 파이프라인은 유지하고, **필요할 때만 Cleaner를 활성화**하는 방식으로 설계한다.
- 샘플 근거:
  - `GPT_deep/신세계프라퍼티_사업분석_20260214-research-report.md`
  - `GPT_deep/deep_md_issue.md`

## 2) 설계 원칙

- 기본 동작 보존: 옵션 미사용 시 기존 결과와 동일해야 함.
- DeepResearch 특화: 패턴이 명확한 토큰만 정리하고 일반 Markdown에는 관여하지 않음.
- 가시성 확보: 몇 건을 어떤 방식으로 정리했는지 통계 제공.
- 안전한 롤아웃: `off` 기본값 + 테스트 완료 후 `auto`를 선택적으로 권장.

## 3) 활성화 정책 (핵심)

Cleaner 모드를 3단계로 제공한다.

- `off` (기본): Cleaner 비활성화
- `auto`: DeepResearch 패턴이 감지될 때만 활성화
- `on`: 항상 Cleaner 적용

권장 기본값은 `off`로 유지한다. (사용자 의도와 동일)

## 4) CLI/설정 스펙

### `md_to_word.py`

- 신규 옵션
  - `--deepresearch-cleaner {off,auto,on}` (default: `off`)
  - `--cite-mode {footnote,inline,strip}` (default: `footnote`)
  - `--drop-unknown-markers` (default: false)
  - `--cleaner-report` (치환 통계 출력)

### `md_formatter.py`

- 동일 옵션 추가 (standalone formatter 사용 시 일관성 확보)

## 5) 모듈 구조

- 신규 파일: `deep_md_cleaner.py`
- 주요 구성
  - `CleanerConfig` dataclass
  - `CleanerReport` dataclass (치환 건수/실패 건수)
  - `DeepResearchCleaner` class
    - `detect(text) -> bool`
    - `clean(text, config) -> Tuple[str, CleanerReport]`

## 6) 토큰 처리 정책

- `cite`
  - `footnote`: 본문 `[^n]` + 문서 하단 `## Citations`
  - `inline`: `(sources: ...)`
  - `strip`: 완전 제거
- `entity`
  - JSON 배열 2번째 요소(표시명) 우선 사용
  - 파싱 실패 시 payload 축약 문자열 사용
- `image_group`
  - 기본: HTML comment 보존
  - 필요 시 제거
- unknown 블록
  - 기본: HTML comment 축약 보존
  - 옵션으로 제거

## 7) 파이프라인 통합 지점

### 경로 A: `md_to_word.py --format`

1. 입력 파일 로드
2. Cleaner 조건 평가(`off/auto/on`)
3. 필요 시 Cleaner 적용
4. 기존 `md_formatter` 수행
5. 기존 `parse_markdown_file` -> `IBDocumentRenderer`

### 경로 B: `md_to_word.py` (format 미사용)

1. 입력 파일 로드
2. Cleaner 조건 평가
3. 필요 시 Cleaner 적용
4. 바로 `parse_markdown_file` -> `IBDocumentRenderer`

### 경로 C: `md_formatter.py` 단독 실행

1. Cleaner 적용(조건부)
2. 기존 구조 복원 로직 실행

## 8) 구현 작업계획 (효율 중심)

### Phase 1. Cleaner 코어 구현

- [ ] `deep_md_cleaner.py` 추가
- [ ] detect/clean/report 구현
- [ ] 단위 테스트 초안 작성 (`tests/test_deep_md_cleaner.py`)

**완료 기준**: 샘플 텍스트에서 토큰 정리 + report 값 확인 가능

### Phase 2. CLI 연결

- [ ] `md_to_word.py` 옵션 추가 및 전처리 단계 연결
- [ ] `md_formatter.py` 옵션 추가 및 연결
- [ ] 도움말/usage 문구 업데이트

**완료 기준**: `off/auto/on` 각각 기대 동작 확인

### Phase 3. 회귀 방지 테스트

- [ ] `tests/test_md_formatter.py`에 light path(20줄 초과) 케이스 추가
- [ ] `tests/test_md_parser.py`에 잔여 특수 토큰 방어 테스트 추가
- [ ] 일반 markdown 입력 비변경성 테스트 추가

**완료 기준**: 신규/기존 테스트 모두 통과

### Phase 4. 운영 검증

- [ ] 실제 샘플 파일 변환 전후 비교
- [ ] DOCX 결과에서 인용구/본문 잔여 토큰 0건 확인
- [ ] `--cleaner-report` 통계 검토

**완료 기준**: 실사용 샘플에서 문제 재현 불가

## 9) 검증 매트릭스

- 케이스 1: 일반 markdown + `off` -> 기존 결과 동일
- 케이스 2: DeepResearch markdown + `off` -> 기존처럼 잔여 토큰 존재(의도된 통제군)
- 케이스 3: DeepResearch markdown + `auto` -> 토큰 제거/변환 성공
- 케이스 4: DeepResearch markdown + `on --cite-mode strip` -> 인용 완전 제거
- 케이스 5: 깨진 JSON entity 포함 -> 실패 없이 fallback 처리

## 10) 문서화/롤아웃

- `README.md`와 `AGENTS.md`에 옵션 설명 추가
- 초기 릴리스는 기본 `off`로 배포
- 안정화 후 사용자 가이드에서 DeepResearch 입력 시 `auto` 권장

## 11) 최종 완료 기준 (Definition of Done)

- DeepResearch 입력에서 ``, ``, `` 잔여 토큰이 Word 본문/인용구에 남지 않음
- 기본 모드(`off`)에서 기존 사용자 결과가 변하지 않음
- 옵션(`off/auto/on`, `cite-mode`)이 문서/테스트/실행 모두 일치
- 테스트 통과 + 샘플 수동 검증 완료
