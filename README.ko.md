# IB Report Formatter 2.0

IB 보고서와 회사 업무문서를 위한 **Markdown → 편집 가능한 Word 전용 엔진**입니다.

[구현계획](docs/implementation-plan-20260914.md) · [검증 결과](docs/verification-20260914.md) · [예제 6종](samples/profiles) · [English / API 상세](README.md)

## 제품 방향

Word→MD는 일시 중단이 아니라 **자체 개발 폐기**입니다. 이미 좋은 외부 프로젝트가 많으므로 이 프로젝트는 역변환을 다시 개발하지 않습니다. 버전 2에서 관련 CLI·파서·MD 역렌더러·OMML 역변환·왕복 검사와 전용 테스트를 제거했습니다.

## 바로 사용하기

```sh
uv sync
uv run md-to-word samples/profiles/office-letter.md 공문.docx --strict
uv run md-to-word samples/profiles/business-report.md 업무보고.docx --strict
uv run md-to-word samples/profiles/meeting-minutes.md 회의록.docx --strict
uv run md-to-word samples/profiles/ib-memo.md IB메모.docx --strict
uv run md-to-word input.md output.docx --profile plain
uv run docx-audit 공문.docx
```

기존 `uv run md_to_word.py ...`, `uv run ib-report ...` 명령도 유지됩니다. 생성에는 Word 설치가 필요하지 않습니다. Word에서 파일을 열어 잠긴 경우 다른 이름으로 저장하고 실제 경로를 알립니다.

## 문서 프로파일

| 프로파일 | 기본 구성 | 용지 |
|---|---|---|
| `ib-report` | 기존 IB 표지·목차·면책·기밀 표시 | Letter |
| `ib-memo` | 제목·작성정보·기밀 표시, 표지·목차·면책 없음 | A4 |
| `plain` | 본문만, 금융·메타데이터 자동 추정 없음 | A4 |
| `office-letter` | 회사명·문서번호·수신·참조·제목·붙임·발신 | A4 |
| `business-report` | 제목·작성일·부서·작성자·본문 | A4 |
| `meeting-minutes` | 제목·일시·장소·참석자·작성자·본문 | A4 |

일반 문서 4종에는 IB 브랜드·면책문구를 자동 삽입하지 않습니다. 짧은 번호 항목을 제목으로 오인하거나 숫자 코드를 금융 수치로 추정하지 않습니다. 목록은 Word에서 수정 가능한 다단계 번호 `1. → 가. → (1)`로 생성합니다. 하위 목록은 **4칸씩 들여쓰기**합니다. IB 프로파일의 기존 목록 방식은 유지됩니다.

## 공문 입력 예시

```markdown
---
profile: office-letter
title: 자료 제출 요청
document_no: 기획-2026-015
date: "2026-09-14"
sender:
  organization: 주식회사 예시
  department: 경영기획팀
  signatory: 주식회사 예시 대표이사
  contact: planning@example.com
recipients: [협력회사 담당부서장]
cc: [재무담당자]
attachments: [제출 양식 1부]
---
# 자료 제출 요청

1. 제출 자료를 확인해 주시기 바랍니다.
    1. 월별 실적
    2. 운영계획
2. 기한 내 회신해 주시기 바랍니다.
```

공문에는 `recipients`, `sender.organization`이 필수입니다. 직인·서명 이미지를 임의로 만들지 않습니다. 날짜·문서번호·숫자형 코드는 따옴표로 감싸 원문을 보존하십시오.

공문은 회사명, 수신·참조·제목, 본문, 편집 가능한 붙임 번호, 끝 표기, 발신인, 시행·연락처 순으로 구성합니다. 긴 수신처·제목은 값 아래로 줄맞춤하며 회사명을 머리말에 중복 출력하지 않습니다. `sender.signatory`에는 YAML 여러 줄 문자열을 사용할 수 있습니다.

### 본문 뒤에 별첨을 넣는 공문

아래 설정을 추가하면 지정한 H1 **앞에서** 붙임·발신인·연락처를 출력하여 공문을 마치고, 새 페이지에서 별첨을 시작합니다.

```yaml
letter:
  appendix_heading: 운영자료 확인 내역
  appendix_label: 붙임 1
```

`appendix_heading`은 본문에 정확히 한 번 나오는 H1의 텍스트여야 하며, 문서 제목과 달라야 하고 앞에 공문 본문이 있어야 합니다. 경계 누락·중복·빈 본문·잘못된 설정은 저장 전에 거부합니다. `appendix_label`은 선택 항목이며 경계 설정 없이 사용할 수 없습니다. 설정을 생략하면 MD 전체를 공문 본문으로 처리합니다.

붙임 목록은 사용자가 기재한 정보일 뿐, 증빙 파일의 실제 존재나 문서 내 삽입을 확인하는 기능은 아닙니다. [완전한 가상 공문 예제](samples/qa/office-letter-appendix.md)와 [실사용 개선계획](docs/improvement-plan-20260915.md)을 참고하십시오.

업무보고는 `analyst`, `date`, `sender.department`; 회의록은 추가로 `meeting_time`, `location`, `attendees`를 사용합니다. 보고 요지·결정사항·후속 조치는 Markdown 본문에 직접 작성합니다. 이미 프로파일에서 만든 제목과 같은 첫 H1은 중복 출력하지 않습니다.

## 설정과 회사 테마

우선순위: **CLI/API 명시값 > YAML > 프로파일 기본값**.

```yaml
layout:
  cover: false
  toc: false
  disclaimer: false
  confidential: false
  strict: true
  separator_mode: auto
```

CLI에서는 `--profile`, `--theme`, `--strict`, `--no-cover`, `--no-toc`, `--no-disclaimer`, `--no-confidential`, `--separator-mode auto|rule|page-break`를 사용할 수 있습니다. 표지·목차를 켜는 설정은 YAML 또는 API로 지정합니다.

테마는 `default`, `mono`, 또는 [회사 테마 YAML](samples/profiles/company-theme.yaml)을 지원합니다. 설정 항목은 본문·제목·한글 글꼴, 본문 크기, 주색, 여백입니다. YAML 안의 테마 경로는 입력 MD 폴더 기준, CLI 경로는 실행 폴더 기준입니다. 임의의 회사 DOCX 양식을 자동 해석하는 기능은 아닙니다.

## IB 금융표 고도화

```yaml
tables:
  - type: financial
    columns: [text, money, percent, code]
    caption: 실적 요약
    unit: 금액 백만원
    as_of: "2026-06-30"
    source: 회사 제공 자료
    landscape: false
  - type: sensitivity
    base_case: {row: 2, column: 3}
```

- `tables`는 본문 표 순서에 대응합니다. 건너뛸 표는 `{}`로 지정합니다.
- 표 종류: `generic`, `financial`, `sensitivity`, `risk`.
- 열 의미: `text`, `code`, `date`, `number`, `money`, `percent`, `bps`, `multiple`. 지정할 때는 모든 열 개수와 일치해야 합니다.
- 숫자에는 천 단위 구분을 적용합니다. `percent`의 `12.5`는 `12.5%`가 됩니다. **비율·단위를 환산하거나 재무 계산을 수행하지 않습니다.**
- `code`는 `001234` 같은 선행 0을 보존합니다. 빈 셀·이스케이프 파이프 `A\|B`도 원래 열에 남습니다.
- 표 제목·단위·기준일·출처, 반복 머리행과 행 분할 제어를 지원합니다. `landscape: true`는 해당 표 앞뒤에 구역을 나누어 가로 표를 배치합니다.
- 민감도표 기준 셀은 **명시한 좌표만 강조**합니다. 행 1은 첫 데이터 행, 열은 라벨 열을 포함한 1부터 시작합니다. 기존 중앙 셀 자동 추정은 제거했습니다.
- 위험도 표는 인식 가능한 등급 열과 높음/중간/낮음 등의 값을 강조합니다.

## 문법·검증의 범위

제목, 문단, 강조, 표, 목록, 인용/콜아웃, 파일·Base64 이미지, 코드블록, 기존 수식·도식 출력을 지원합니다. 외부 HTTP(S)/메일 링크와 숫자 각주 `[^1]` / `[^1]: 설명`는 Word 네이티브 요소로 만듭니다. 각주는 숫자 식별자·한 줄 정의를 사용하십시오.

일반 문서의 References는 본문으로 보존합니다. 문단의 자연스러운 줄바꿈은 공백으로 합치며, 강제 줄바꿈은 `<br>` 또는 줄 끝 역슬래시를 사용합니다.

`<br>`, `<br/>`, `<BR />`는 본문·강조·제목·목록·표 셀 중간에서도 실제 Word 줄바꿈으로 출력합니다. 이스케이프한 `\<br>`와 코드의 리터럴은 보존하며 링크 주소는 바꾸지 않습니다. 줄 끝 공백 2칸 방식은 파서의 명시적 opt-in 설정에서만 지원합니다.

`--strict`는 알려진 입력 손실·각주 오류·요소 렌더 실패·미해결 이미지 표시를 검사하고 실패하면 저장하지 않습니다. 기본 모드에서는 경고와 함께 부분 문서를 생성할 수 있습니다. `docx-audit`는 구조 검사이지 Word→MD나 시각 검사가 아닙니다.

검은 사각형이 보이면 실제 글머리표인지, Word의 **인쇄되지 않는 문단 페이지 제어 표시**인지 먼저 구별하십시오. `docx-audit`의 `pagination_marked_paragraphs`는 표 안의 문단과 스타일 상속까지 고려한 keep-lines/keep-next/page-break-before 문단 개수이며, 글머리표 개수가 아닙니다. `warnings`는 Normal의 과도한 페이지 묶음을 경고하고 `issues`는 구조 오류를 보고합니다. 필요한 제목 제어까지 일괄 제거하거나 사용자의 Word 표시 설정을 임의 변경하지 않습니다. `visual_review: not_performed`는 페이지를 눈으로 확인하지 않았다는 뜻입니다.

아래 사항은 별도 확인 또는 후속 개발 대상입니다.

- 완전한 CommonMark/GFM·복잡한 중첩 문법·임의 HTML 지원
- 회사 DOCX 원본 양식 자동 복제, HWP/HWPX 출력
- 신규 차트 엔진·편집형 차트·재무모델 계산·법적 적합성 판단
- 자동 PDF 배포, Word와 다른 렌더러 간 완전한 페이지 일치

수식·도식은 기존 방식대로 이미지가 될 수 있습니다. 목차·쪽번호는 Word에서 필드 업데이트가 필요할 수 있습니다. 매우 긴 표, 글꼴 대체, 실제 인쇄 페이지는 사용 전 눈으로 확인하십시오.

UTF-8/BOM·EUC-KR·CP949 한글 입력을 지원합니다. 파일 이미지의 상대 경로는 MD 파일 폴더 기준입니다. 일반 문서에서 `m^2^`와 같은 위첨자는 각주로 바꾸지 않으며 실제 각주는 `[^2]`처럼 명시합니다. 기존 IB 위첨자 참조 호환은 유지합니다.

## 실제 Word 페이지 검증

Windows와 설치된 Microsoft Word가 있다면 다음 명령으로 전체 페이지 PNG와 검증용 DOCX 복사본을 만들 수 있습니다. 원본은 읽기 전용으로 열고 시스템 프린터·추가 기능 설정은 바꾸지 않습니다. 이 스크립트는 로컬 QA용이며 무인 서버·PDF 배포 기능이 아닙니다.

```powershell
uv run md_to_word.py samples/profiles .dryforge/qa/docx --batch --strict
uv run md_to_word.py samples/qa .dryforge/qa/docx --batch --strict
powershell.exe -NoProfile -File scripts/word_visual_qa.ps1 -InputDirectory .dryforge/qa/docx -OutputDirectory .dryforge/qa/pages
```

Word의 목차·쪽번호 필드를 갱신하고 페이지 캐시를 준비한 뒤 출력합니다. **출력 폴더는 새 폴더이거나 비어 있어야 하며**, 이전 검증 이미지는 덮어쓰지 않습니다. `-ExpectedPages 2`를 추가하면 모든 입력 파일이 2페이지인지 검사합니다. 문서마다 예상 쪽수가 다르면 별도로 실행하거나 이 옵션을 생략하십시오. 예상 쪽수 불일치 또는 검증 중 원본 변경은 완료된 기록을 남긴 뒤 비정상 종료합니다.

`manifest.json`은 항상 배열이며 Word 버전·시각, 원본/갱신 DOCX/PNG의 SHA-256, 실제·기대 페이지 수, `visualReview: pending`을 기록합니다. 렌더 성공만으로 시각검증을 완료하지 않습니다. 모든 PNG를 확인하고 manifest 해시와 연결한 별도 검토 기록을 남기십시오. [최신 검증 기록](docs/verification-20260915.md)을 참고하십시오.

한글 Windows 경로는 실제 파일 경로를 따옴표로 감싸 전달하십시오. Markdown에서 복사한 `파일\_초안.md`와 실제 `파일_초안.md`는 다릅니다. `_초안.md`가 실제 하위 경로일 수도 있으므로 역슬래시를 일괄 제거하지 말고 실제 파일을 확인해야 합니다. CLI가 입력·출력 경로를 추측하여 고치지는 않습니다.

## 개발 및 이전

```sh
uv sync --extra dev
uv run pytest tests/
uv build
```

Python 3.8 문법 호환을 유지하며 이번 실제 테스트 환경은 Python 3.12입니다. 모델은 `document_model.py`로 분리했으며 기존 `md_parser` 모델 import는 유지합니다. CLI·API·레지스트리는 하나의 조립 경로를 사용합니다. API에서는 파싱 단계부터 원하는 프로파일을 전달하십시오. 파싱된 모델은 적용한 프로파일을 기록하며, 렌더링 시 다른 프로파일을 요청하면 strict 여부와 무관하게 `ValueError`를 발생시킵니다. 원하는 프로파일로 Markdown을 다시 파싱해야 합니다. 파싱 이력이 없는 직접 구성 모델은 렌더 프로파일을 선택할 수 있습니다.

역변환 모듈과 `roundtrip-audit` 명령은 제거되었습니다. 제거 전 코드는 `d819bbb` 및 로컬 `codex/archive-word-to-md-d819bbb` 브랜치에서 복구할 수 있습니다. 기본 IB 경로는 유지하되 빈칸·숫자 출력·머리말/쪽번호·면책 설정 오류를 고쳤으므로 해당 결과는 의도적으로 달라집니다.

과거 계획서의 양방향 변환 계획보다 이 문서와 2026-09-14 구현계획이 우선합니다.
