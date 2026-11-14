# IB Report Formatter 2.0

IB 보고서와 회사 업무문서를 위한 **Markdown → 편집 가능한 Word 전용 엔진**입니다.

[구현계획](docs/implementation-plan-20260914.md) · [검증 결과](docs/verification-20260914.md) · [2026-11 검증 기록](docs/verification-20261110.md) · [예제 7종](samples/profiles) · [English / API 상세](README.md)

## 어떤 방식으로 사용하시나요?

같은 Markdown → Word 엔진을 **사용 방식에 따라** 선택합니다. 브랜치나 문서 종류별로 다른 제품이 아닙니다.

| 하고 싶은 일 | 선택할 배포본 | 필요한 환경 |
|---|---|---|
| 완성된 Markdown을 편집 가능한 Word로 변환 | **문서 변환기 · Windows 실행형** | Windows x64, Python 포함, AI 계정 불필요 |
| 내 작업 폴더의 자료로 AI에게 작성·수정까지 요청 | **AI 문서 작성 도우미 · Claude Code / Codex 플러그인** | 지원 호스트와 계정, uv; 필요 시 Python 3.12 준비 |

**[시작하기: 배포본 선택과 설치](docs/start-here.md)** · [플러그인 설치](docs/plugin-guide.md) · [공개 다운로드](https://github.com/Hanriverflow/IB_report_formatter/releases)

<details>
<summary>고급 사용자: Codex 프로젝트 키트 (소스형)</summary>

압축을 푼 소스 폴더 자체를 Codex 프로젝트로 열어 예제를 배우거나 작업 흐름·코드를 수정합니다. Codex 계정, uv, Python 3.12 이상이 필요합니다. 플러그인보다 문서 기능이 많은 상위 제품은 아닙니다. [프로젝트 시작](docs/start-here.md#codex-project-kit) · [상세 로컬 HTML 안내](docs/codex-start.html)

</details>

**버전 구분:** 엔진 **2.0.0 (Beta)** / 플러그인 **0.1.0**. 서로 다른 버전 체계이며 신·구 제품 순서가 아닙니다. ZIP과 `.plugin`은 같은 플러그인이고, Claude 웹·Cowork 실행은 검증되지 않았습니다. 소스가 있다는 것이 다운로드 가능한 Release 게시를 뜻하지는 않습니다. [검증 범위](https://github.com/Hanriverflow/IB_report_formatter/blob/main/docs/verification-plugin-20261008.md)

**문서 종류는 그다음에 선택합니다:** 일반 업무(공문·보고서·회의록), IB 보고서·메모, 금융거래 텀시트. 기관별 문구·로고는 YAML 설정으로 구분하며 실제 기관 자료는 저장소 밖에 둡니다. 모든 배포본이 같은 엔진을 사용하므로 기능 브랜치를 별도 제품으로 내려받지 않습니다. [문서 프로필](#문서-프로파일) · [관리자용 빌드·배포 안내](https://github.com/Hanriverflow/IB_report_formatter/blob/main/docs/distribution-build.md)

## 제품 방향

Word→MD는 일시 중단이 아니라 **자체 개발 폐기**입니다. 이미 좋은 외부 프로젝트가 많으므로 이 프로젝트는 역변환을 다시 개발하지 않습니다. 버전 2에서 관련 CLI·파서·MD 역렌더러·OMML 역변환·왕복 검사와 전용 테스트를 제거했습니다.

## 바로 사용하기

### 고급 사용자: Codex 소스 프로젝트로 문서 작성하기

이 소스 폴더를 Codex 앱의 로컬 프로젝트로 열고 `$ib-document`에 문서 작성을 요청할 수 있습니다. Codex가 자료를 읽어 MD/YAML을 작성·수정하고, 기존 엔진이 엄격 검사 후 Word를 생성합니다. 별도 LLM 프로그램·프로젝트 API 키·EXE는 필요 없습니다. GitHub를 쓰지 않는 분은 소스 ZIP을 전체 압축 해제하면 됩니다.

**[시작하기: 설치부터 텀시트 작성까지](docs/codex-start.html)** · [Codex 작성 스킬](.agents/skills/ib-document/SKILL.md) · [가상 자료 연습](samples/harness/)

스킬이 표시되지 않으면 `.agents/skills/ib-document/SKILL.md`를 읽고 진행해 달라고 요청하십시오. Python/uv 최초 설치와 Codex 모델 사용에는 인터넷이 필요할 수 있습니다. 원자료 대조·문서 구조 검사·실제 페이지 검토는 각각 별도입니다.

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
| `term-sheet` | 거래 첫머리·기관 문구·병합 조건표·고객확인란 | A4 |

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

공문은 회사명, 수신·참조·제목, 본문, 편람 모양의 붙임(편집 가능한 번호), 끝 표기, 발신인, 시행·연락처 순으로 구성합니다. 긴 수신처·제목은 값 아래로 줄맞춤하며 회사명을 머리말에 중복 출력하지 않습니다. `sender.signatory`에는 YAML 여러 줄 문자열을 사용할 수 있습니다.

### 본문 뒤에 별첨을 넣는 공문

아래 설정을 추가하면 지정한 H1 **앞에서** 붙임·발신인·연락처를 출력하여 공문을 마치고, 새 페이지에서 별첨을 시작합니다.

```yaml
letter:
  appendix_heading: 운영자료 확인 내역
  appendix_label: 붙임 1
```

`appendix_heading`은 본문에 정확히 한 번 나오는 H1의 텍스트여야 하며, 문서 제목과 달라야 하고 앞에 공문 본문이 있어야 합니다. 경계 누락·중복·빈 본문·잘못된 설정은 저장 전에 거부합니다. `appendix_label`은 선택 항목이며 경계 설정 없이 사용할 수 없습니다. 설정을 생략하면 MD 전체를 공문 본문으로 처리합니다.

붙임 목록은 사용자가 기재한 정보일 뿐, 증빙 파일의 실제 존재나 문서 내 삽입을 확인하는 기능은 아닙니다. [완전한 가상 공문 예제](samples/qa/office-letter-appendix.md)와 [실사용 개선계획](docs/improvement-plan-20260915.md)을 참고하십시오.

### 붙임, 표로 끝나는 본문, 항목 기호

「행정업무의 운영 및 혁신에 관한 규정 시행규칙」 제2조①·제4조④⑤와 2025 행정업무운영 편람의 모양을 따릅니다. 이 규정은 행정기관에만 적용되므로 회사 공문에는 법적 구속력이 없고, 공공기관 상대방이 익숙한 형식을 맞추는 것입니다. 엔진은 법적 적합성을 보증하지 않습니다. 확인 내용은 [검증 기록](docs/verification-office-letter-20261114.md)에 있습니다.

- **붙임.** 1개면 번호 없이 한 줄 `붙임  발행조건 요약표 1부.  끝.`. 여러 개면 `붙임  1.` 줄에서 시작하고 다음 번호는 `1.` 아래에 맞춰집니다. 번호는 Word 자동 번호이므로 Word에서 고칠 수 있습니다. `끝.`은 마지막 붙임 뒤에 붙습니다.
- **표로 끝나는 본문.** 마지막 칸까지 채운 표는 표 아래 한 글자 들여 `끝.`을 씁니다. 마지막 표가 중간까지만 채워졌으면(빈 행이 남은 표) 직접 적습니다.

  ```yaml
  letter:
    table_end: partial
  ```

  그러면 마지막으로 작성된 행 다음의 첫 빈 행 첫 칸에 `이하 빈칸`을 쓰고 `끝.`은 쓰지 않습니다(붙임이 있으면 `끝.`은 붙임 뒤에 그대로). 받는 쪽이 채우라고 빈 행을 둔 회신 양식도 있으므로, 엔진이 빈 행만 보고 추측하지 않습니다. 값이 `partial`이 아니거나, 공문 본문이 표로 끝나지 않거나, 그 표의 마지막 작성 행 뒤에 빈 행이 없으면 저장 전에 거부합니다.
- **항목 기호.** 공문의 번호 목록은 네 칸 들여쓰기 단계마다 `1.` `가.` `1)` `가)` `(1)` `(가)` `①`(1~7단계)를 씁니다. 8단계 `㉮`는 Word 자동 번호 형식이 없어, strict 모드는 8단계 이상을 거부하고 일반 모드는 경고 뒤 `가`, `나`, `다`로 대신합니다. Word의 가나다 번호는 `하` 다음을 `가`로 되돌립니다(편람은 `거`). 다른 프로필의 번호(`1.` `가.` `(1)`)는 바뀌지 않습니다.

[이하 빈칸 가상 공문 예제](samples/qa/office-letter-partial-table.md)를 참고하십시오.

업무보고는 `analyst`, `date`, `sender.department`; 회의록은 추가로 `meeting_time`, `location`, `attendees`를 사용합니다. 보고 요지·결정사항·후속 조치는 Markdown 본문에 직접 작성합니다. 이미 프로파일에서 만든 제목과 같은 첫 H1은 중복 출력하지 않습니다.

## 버전 표시와 초안 워터마크

```yaml
version: "v1.2"   # 따옴표로 감쌉니다. 따옴표 없이 1.10이라 쓰면 YAML이 1.1로 읽습니다
status: draft     # draft 또는 final
```

- `version`을 적으면 파일 속성의 사용자 지정 항목 `Version`에 기록하고, 머리글 회사명 뒤에 `예시금융 | v1.2`처럼 보여 줍니다(회사명이 없으면 `v1.2`만). 머리글의 버전은 `DOCPROPERTY Version` 필드이며 지금 값을 담고 있어 파일을 열자마자 바르게 보입니다. 공문은 속성만 기록하고 머리글에 버전을 보여 주지 않습니다. 텀싯은 지금처럼 바닥글에 `subtitle version`을 보여 주며, 버전 부분이 같은 필드입니다.
- Word에서 버전을 바꾸려면 **파일 > 정보 > 속성 > 고급 속성 > 사용자 지정**에서 `Version` 값을 고친 뒤 필드를 갱신합니다. 머리글을 두 번 눌러 **Ctrl+A, F9**를 누르면 그 구역의 머리글이 바뀌고, **Ctrl+P**(인쇄 화면)나 인쇄 미리보기를 열면 모든 구역이 바뀝니다. 가로 표가 있는 문서는 구역이 여럿이므로 Ctrl+P가 간편합니다. 본문에서 Ctrl+A, F9를 눌러서는 머리글이 바뀌지 않습니다. 인쇄와 PDF 저장에는 새 값이 찍힙니다.
- `status: draft`를 적으면 표지와 가로 표 쪽을 포함한 모든 쪽에 DRAFT 워터마크가 찍힙니다. Word가 쓰는 것과 같은 워터마크이므로 **디자인 > 워터마크 > 워터마크 제거** 한 번으로 모든 쪽에서 지워집니다. `status: final`이거나 적지 않으면 워터마크가 없고, 다른 값은 오류로 거부합니다.
- 파일 > 정보에는 문서 제목이 나오고, 만든 이와 메모는 비어 있으며, 만든 날짜와 수정한 날짜는 변환한 시각입니다.

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

## 텀싯 (Term sheet)

국내 구조화금융 조건서는 `profile: term-sheet`로 작성합니다. [가상 ABCP 예제](samples/profiles/term-sheet.md)와 [가상 기관 문구 파일](samples/profiles/term-sheet-house.yaml)을 참고하십시오.

```yaml
---
profile: term-sheet
title: "가나다머티리얼즈㈜ ABCP 300억원"
date: "2026-10-16"
version: v1
house: term-sheet-house.yaml
tables:
  - {}
  - {columns: [text, date, date, number, number, number], unit: 억원}
---
```

`title`은 필수이며 `subtitle` 기본값은 `Term Sheet`입니다. 선택 항목인 `date`는 입력한 표기대로 표시하고 `version`은 바닥글에 씁니다. `prepared_by`와 `disclaimer`는 frontmatter나 house 중 한 곳에 공백이 아닌 문구로 있어야 합니다. House YAML은 `prepared_by`, `disclaimer`, `confidential_label`, `confirmation`(`intro`, `items` 목록, `signature`), `style`만 받습니다. 같은 키를 frontmatter에 지정하면 빈 값도 house보다 우선하며, `style`은 항목별로 덮어씁니다. 기관 문구에는 `( *주요내용* )` 같은 인라인 강조를 쓸 수 있습니다.

`style`은 기관 서식을 정합니다. 모든 항목은 선택이며 기본값은 표준 배치입니다.

```yaml
style:
  cover: page            # inline(기본) | page: 제목·로고·고지문을 별도 표지 쪽에 배치
  logo: assets/logo.png  # 이미지 경로(house 파일 기준, frontmatter면 Markdown 기준) 또는 data: URI, ""이면 house 로고 제거
  logo_width_mm: 40
  disclaimer: box        # rules(기본) | box
  header_rule: true      # 머리글 아래 강조선
  footer_rule: true      # 바닥글 위 강조선
  label_color: "#C00000" # 대외비 표시 색, ""이면 기본 회색
  page_number: "- {page} -"   # 기본 "{page} / {pages}"
  page_number_align: center  # right(기본) | center
  table_header: dark     # light(기본) | dark: 강조색 바탕에 흰 글씨
  table_sides: open      # closed(기본) | open: 표 바깥 좌우 세로선 생략
```

기본값을 쓰려면 항목을 생략합니다. null 값은 설정 오류로 거부합니다. 로고를 불러오지 못하면 렌더 진단으로 기록하며 strict에서 거부합니다. 실제 기관 로고는 실제 house 파일처럼 저장소 밖에 보관하십시오.

실제 딜 문서와 실제 기관 house 파일은 **저장소 밖**에 보관하십시오. Frontmatter의 `house` 상대 경로는 원본 Markdown 폴더 기준입니다. `--house /absolute/path/to/house.yaml`로 파일을 재지정할 수 있으며 CLI의 상대 경로는 실행 폴더 기준입니다. 원본 경로가 없는 문자열·스트림 입력은 절대 house 경로가 필요합니다. 저장소에 포함된 house는 새로 작성한 가상 문구입니다.

장 번호는 `## 1. 본건 개요`처럼 직접 쓰고 소제목은 `###`를 사용합니다. `tables`는 본문 표 순서에 맞추며 설정이 없는 표도 `{}`로 자리를 유지합니다. 위 코드는 표 두 개의 작성 예시이며, 전체 예제를 복사할 때는 예제 파일의 표 사양 순서를 사용하십시오.

- 셀 내용 전체가 `^^`이면 위로, `<<`이면 왼쪽으로 병합합니다. 병합 영역은 직사각형이어야 하며 머리행과 본문 사이를 넘을 수 없습니다. 빈칸은 병합하지 않습니다. 기호를 그대로 쓰려면 `\^^`, `\<<`로 이스케이프합니다. 잘못된 병합 묶음은 원문과 경고를 남기며 strict에서 거부합니다.
- 텀싯은 병합이 기본 활성화됩니다. `spans: false`로 끌 수 있고, 다른 프로파일은 `spans: true`로 명시해야 활성화됩니다.
- `label_columns`는 선행 라벨 열 수를 재지정합니다(0 이상, 전체 열 수 미만의 정수). 생략하면 병합 후 첫 머리 셀의 폭을 사용하되 전체 열 수−1을 넘지 않습니다. `| 구 분 | << | 내 용 |`은 2개, `| 구 분 | 내 용 |`은 1개입니다.
- 라벨 1~2개 뒤에 내용 열 하나가 오는 조건표는 첫 라벨 폭 33.5mm, 음영·굵은 글씨를 씁니다. 둘째 라벨은 30mm, 흰 바탕·보통 굵기·회색 글씨입니다. 그 밖의 격자 표는 각 열의 내용이 한 줄에 들어가는 폭을 먼저 주고 남는 폭을 열마다 똑같이 나눕니다. 내용이 나란히 들어가지 않으면 기존의 내용 기반 추정을 씁니다. 라벨에는 음영과 보통 굵기를 적용합니다. `label_columns: 0`이면 라벨 음영이 없습니다.
- 둘째 라벨 열에서 시작해 내용 열까지 병합한 셀은 내용 셀(왼쪽 정렬, 음영 없음)로 배치합니다. 첫 라벨 열에서 시작한 셀은 합계 행처럼 넓게 병합해도 라벨로 남습니다.
- `header_rows: 2`(기본 1)는 앞의 두 행을 머리행으로 그립니다. 음영과 병합을 적용하고 쪽마다 반복합니다. 둘째 머리행은 구분선 바로 다음 Markdown 행입니다. 머리행끼리는 병합할 수 있지만 본문으로 넘어갈 수는 없습니다. 이때 `base_case`의 행 번호는 첫 본문 행부터 셉니다.
- `columns`로 날짜·코드·숫자 역할을 정하고 `unit`으로 표 위 단위를 표시합니다. 선택 문자열 `note`는 표 아래 오른쪽에 8pt 회색으로 표시하며 모든 프로파일에서 지원합니다. 금융 의미를 추론하거나 상환스케줄을 계산하지는 않습니다. 선언한 `schedule`은 검사만 합니다(수치 검증 참고).

셀 안 줄바꿈은 `<br>`를 사용합니다. 텀싯 본문 문단의 `<br>`도 줄마다 별도 Word 문단으로 만듭니다. 줄 앞 `•`, `-`, `·`는 각각 0/3/6mm에서 시작하고 3mm 내어쓰기, `①`–`⑳`는 0mm 시작·4.5mm 내어쓰기, `※`는 0mm 시작·4mm 내어쓰기와 8pt 글씨를 적용합니다. 이어지는 줄은 기호 뒤 본문에 맞춰 정렬됩니다. 본문의 Markdown `-` 목록은 기존 목록 처리를 유지합니다.

고객확인란을 넣을 위치에 비어 있고 닫힌 펜스를 둡니다.

````markdown
```confirmation
```
````

House/frontmatter 문구로 안내·체크 항목 행과 음영·가운데 정렬 서명 행을 만들고 한 페이지에 모읍니다. 문구 누락, 내용이 있는 펜스, 닫히지 않은 펜스는 원문을 진단 코드 패널로 보존하며 strict에서 거부합니다. 다른 프로파일에서는 일반 코드 블록입니다.

모든 페이지 머리글에 `confidential_label`(기본 `Strictly Confidential`)을 표시합니다. 빈 라벨, `layout.confidential: false`, `--no-confidential`로 끌 수 있습니다. 바닥글 왼쪽은 `version`이 있을 때 `subtitle version`(버전은 `DOCPROPERTY Version` 필드, [버전 표시와 초안 워터마크](#버전-표시와-초안-워터마크) 참고), 오른쪽은 `PAGE / NUMPAGES`입니다. 첫머리 면책은 `--no-disclaimer`와 무관하게 필수입니다. 기존 **`termsheet` 프리셋**은 표지·목차·끝 면책만 전환하며, 텀싯 프로파일 선택·기관 문구 로딩·조건표 서식은 적용하지 않습니다.

텀싯 문서는 한글을 어절 단위로 줄바꿈하고(`w:wordWrap=1`), 한글과 영문·숫자 사이의 자동 간격을 끄므로(`w:autoSpaceDE/DN=0`) `300억원`, `SPC에`처럼 붙여 씁니다.

### 조건 변수와 불일치 검사

금액·금리·일정처럼 여러 곳에 반복되는 값은 frontmatter `terms:`에 한 번만 정의하고 본문에서 `{{키}}`로 참조합니다. 모든 프로파일에서 쓸 수 있습니다.

```yaml
terms:
  amount: "300억원"
  cap_spread: "[1.10]%p"
```

- 키는 영문 소문자로 시작하고 소문자·숫자·밑줄만 씁니다(최대 40자). 값은 한 줄 문자열만 받습니다. `1.10`처럼 숫자로 읽히는 값은 YAML이 `1.1`로 바꾸므로 반드시 따옴표로 감쌉니다.
- 제목·부제·날짜, 장 제목, 문단, 목록, 표 셀, 인용 상자의 참조를 치환합니다. 표 캡션·단위·출처·기준일은 태그 없이 글자로만 치환합니다. 인라인 코드, 코드 블록, 수식, 링크 주소 안은 치환하지 않습니다. `\{{`는 글자 그대로 나옵니다.
- 값은 Markdown으로 다시 해석하지 않고 쓴 그대로 들어가며, 표의 숫자 서식도 적용하지 않습니다. 둘레의 굵게·기울임은 이어받습니다.
- 정의되지 않은 키는 경고와 함께 `{{키}}`를 남기고 strict에서 거부합니다. `terms:`가 없으면 `{{…}}`는 일반 글자입니다. `terms: {}`는 기능을 켜고 모든 참조를 검사합니다.
- 치환된 값은 Word 콘텐츠 컨트롤(태그 `ibrep:term:<키>`)로 감싸고, 생성 시점 값을 문서 속성 `ibrep.term.<키>`에 기록합니다. 외부 송부용으로 컨트롤 없는 사본이 필요하면 `--no-term-tags` 또는 `layout: {term_tags: false}`를 씁니다.
- Word에서 수정한 문서에 `docx-audit`를 실행하면 JSON의 `terms` 항목이 다음을 보고합니다: `mismatched`(같은 키의 값이 서로 다름), `changed`(생성 후 바뀐 값, 원본 YAML에 반영할 목록), `missing`(컨트롤이 모두 사라진 키), `indicative`(대괄호가 남은 확정 전 조건). 이 검사는 경고이며 구조 오류(`issues`)와 종료 코드에는 영향을 주지 않습니다. 태그된 값끼리의 일관성만 보며 재무적 타당성은 판단하지 않습니다.

### 수치 검증

조건 값이 지켜야 할 산식을 `checks:`에 선언합니다. 모든 프로파일에서 쓸 수 있습니다.

```yaml
checks:
  - "all_in = issue_rate + credit_fee + running_cost"
  - "facility = amount * 1.05"
  - {check: "upfront = advisory + legal + audit", tolerance: "30만원"}
```

- 식은 `=`를 하나만 씁니다. 양변에는 조건 키와 숫자를 `+ - * /`와 괄호로 조합합니다.
- 값은 화면에 보이는 대로 읽습니다.
  - 원화 금액: `10억원`, `3,330만원`, `5억 3,000만원`, `1,076,175,000원`.
  - 비율과 포인트: `4.78%`, `1.10%p`, `29bp`.
  - `36개월`, `4.54년`, 일반 숫자.
  - 대괄호와 약·총·내외·수준·정도라는 말은 무시합니다.
- 금액과 비율은 섞거나 환산하지 않습니다.
- 기본 허용 오차는 좌변 값이 표시한 자릿수의 절반입니다. 그래서 `12.66억원`은 반올림해 그 값이 되는 금액과 일치로 봅니다. 더 넓게 허용하려면 `tolerance`를 줍니다.
- 틀리거나 읽을 수 없는 식은 경고이며 strict에서 거부합니다. 식은 생성 시점 값과 함께 DOCX 속성 `ibrep.checks`에 저장합니다. `docx-audit`는 Word에서 수정한 현재 값으로 다시 계산해 `terms.failed_checks`에 보고합니다.

상환스케줄 표는 열을 선언하면 산식을 검사합니다.

```yaml
tables:
  - unit: 억원
    schedule: {repayment: 5, balance: 6, months: 4, principal: "{{amount}}", total: true, average_life: "{{wal}}"}
```

- `repayment`와 `balance`는 1부터 세는 열 번호이고, `months`(경과 개월)는 선택입니다.
- `principal`은 표 단위의 숫자이거나, 금액으로 된 조건 키·값입니다. 금액이면 표의 `unit`으로 환산합니다. 없으면 첫 행의 잔액과 상환액을 더해 기초 원금으로 봅니다.
- `total: true`는 마지막 본문 행을 합계 행으로 봅니다.
- 검사 항목:
  - 각 행 잔액 = 직전 잔액 − 그 행 상환액;
  - 마지막 잔액 0;
  - 상환액 합계 = 원금;
  - 합계 행 = 상환액 합계;
  - `average_life`(년, 또는 개월) = 상환액 가중 경과개월 ÷ 12.
- 빈칸, 대시, `대출실행` 같은 글자는 상환 없음으로 봅니다. 잔액의 대시는 잔액 없음입니다.
- 불일치는 경고이며 strict에서 거부합니다. 값을 다시 계산해 바꾸지는 않습니다.

### 변환본 입력 정리 (HWP·Word)

외부 변환기로 HWP나 Word를 바꾼 Markdown은 텀싯 입력으로 쓰기 전에 정리가 필요합니다. 정리는 사용자가 직접 실행하는 MD→MD 단계입니다. 렌더링 전에 보고와 결과를 확인하십시오.

```sh
uv run md-format converted.md term-sheet.md --converted-term-sheet
```

- 한 줄짜리 단락 `1. 제목`(한두 자리 번호, `1\.`도 포함, 번호 뒤 40자 이내)이 굵거나 바로 다음 블록이 파이프·HTML 표이면 `## 1. 제목`으로 바꿉니다. 그 밖의 번호 문장은 그대로 두며, 파서가 스스로 제목으로 올리지도 않습니다. 그대로 둔 한 줄짜리 번호 단락은 장 제목인지 확인하도록 보고합니다.
- 머리행만 있고 본문 행이 없는 표(예: `| 별첨 | | 상환스케줄 |`)는 `## 별첨 | 상환스케줄` 제목으로 바꿉니다. 제목 글자에 인라인 문법으로 읽힐 문자(`*`, `_`, `$`, `^`, `~`, `` ` ``, 괄호 `[]`, `<`, `>`, 역슬래시, HTML 엔티티)가 있거나 HTML 표의 셀 밖에 무엇이든(예: caption, 주석) 있거나 셀 안에 `<p>` 하나와 굵게·기울임·밑줄·span 태그 말고 다른 태그(예: `<br>`, `<sup>`, 링크)가 있으면 표로 두고 보고합니다.
- 첫 장 제목(또는 표로 남긴 1행 표) 앞의 블록을 표지로 봅니다. 첫 대외비 줄은 `confidential_label`이 되고 같은 줄의 반복은 지웁니다. 맨 앞의 굵은 줄(또는 1단계 제목)은 `title`, 이어지는 굵은 줄은 `subtitle`, 날짜 줄은 `date`, 80자 이상 단락은 `disclaimer`가 됩니다(강제 줄바꿈 `\`·`<br>`은 줄바꿈으로, 이스케이프한 `\<br>`은 글자로 남깁니다). 표지 이미지는 본문에서 빼고 보고하므로, 로고라면 house `style.logo`로 지정하십시오.
- 최상위 블록만 바꿉니다. 들여쓴 코드, 목록 항목과 이어지는 줄, 인용, 코드 블록, HTML 표(중첩 표 포함, 위의 1행 띠 표는 예외)는 그대로 둡니다. `#` 제목 줄은 앞 단락을 끝내고, 2단계 이상 제목은 들여써도 표지 범위를 끝냅니다.
- 판단하지 못한 줄은 그 자리에 두고 보고합니다. `prepared_by`는 추론하지 않으니 house 파일이나 frontmatter에 넣으십시오. 기존 frontmatter의 줄과 키(파서처럼 대소문자 구분 없이 비교)는 그대로 두고 새 키는 닫는 `---` 앞에 붙이며, 이미 설정된 키에 해당하는 표지 줄은 본문에 남깁니다. YAML `...` 닫는 줄은 `---`로 바꾸고, 파서가 읽지 않는 흐름 형식(`{...}`) frontmatter는 블록 YAML로 다시 쓰며 둘 다 보고합니다.
- 옮기거나 바꾸거나 지우거나 남긴 줄을 입력 줄 번호(여러 줄이면 범위)와 함께 한 줄씩 출력합니다. 이 모드는 그 실행에서 Deep Research 정리를 대신하며 `--check`나 기본값이 아닌 DeepResearch 옵션과 함께 쓸 수 없습니다.

## 차트 사용 (기본 비활성)

```sh
md-to-word report.md report.docx --charts --strict
md-to-word samples/qa/charts.md charts.docx --strict
md-to-word report.md code-panels.docx --no-charts
```

7개 프로파일 모두 기본값은 비활성입니다. `--charts`, `RenderOptions(charts=True)` 또는 최상위 frontmatter `charts: true`로 활성화합니다. 명시한 CLI/API 값이 YAML보다 우선하며 `--no-charts` / `charts=False`로 코드 패널 출력을 강제할 수 있습니다. 파서는 원문을 차트 모델 요소에 보존하므로 같은 모델에서도 출력 방식을 선택할 수 있습니다.

````markdown
```chart
chart_type: bar
title: 가상 매출
labels: [상반기, 하반기]
series:
  - name: 매출
    values: [100, 120]
y_label: 백만원
source: 기능 검증용 가상 자료
number_format: ',.1f'
```
````

`bar`는 계열별 묶은 막대, `line`은 선, `waterfall`은 누적 증감 차트입니다. PR #5의 `chart_type`, `y_label`, `title`, `labels`, `series`, `source`, `total_label` 문법을 유지합니다. `type`도 `chart_type` 대신 사용할 수 있으며 두 값이 충돌하면 오류입니다. `unit`은 `y_label`이 없을 때 세로축 라벨로 사용합니다.

계열마다 문자열 `name`과 라벨 수에 맞는 유한 숫자 `values`가 필요합니다. 라벨은 문자열·숫자를 지원하며 불리언은 거부합니다. `number_format`은 `number`(기존 기본 형식), `percent`, `bps`, `multiple`, `',.2f'`나 `'.1f'` 같은 고정소수점 형식을 지원합니다. **수치를 환산하지 않으므로** `12.5`는 `percent` 형식에서 `12.5%`입니다.

Waterfall은 계열 하나만 지원하며 0부터 모든 값을 차례로 더합니다. `[100, -30, 20]`이면 마지막에 합계 `90` 막대를 추가합니다. 음수 합계도 지원하고 `total_label`로 합계 라벨을 지정합니다. 이미 계산한 합계를 증감값에 다시 넣지 마십시오.

잘못된 YAML·차트 사양·렌더 실패는 차트를 식별하는 진단을 남깁니다. Strict 모드는 새 파일 생성이나 기존 파일 교체 전에 거부하고, 일반 모드는 경고와 함께 원문 코드 패널을 표시합니다. 비활성 상태에서는 잘못된 YAML도 기존 코드 패널로 출력합니다. 차트는 본문 폭에 맞춘 PNG이며 요청별 색상과 한글 글꼴 정책을 사용합니다. 한글 출력에는 지정 글꼴이 설치되어 있어야 합니다(Windows 기본: 맑은 고딕). [가상 자료 예제](samples/qa/charts.md)는 세 차트 유형을 모두 포함합니다.

## 문서 구성 프리셋

```sh
md-to-word --list-presets
md-to-word report.md terms.docx --preset termsheet
md-to-word notes.md notes.docx --profile plain --preset lecture-note --no-cover
```

| 프리셋 | 표지 | 목차 | 면책 |
|---|---|---|---|
| `ib-report` | 상속 | 상속 | 상속 |
| `termsheet` | 끔 | 끔 | 끔 |
| `legal-memo` | 끔 | 켬 | 끔 |
| `lecture-note` | 켬 | 켬 | 끔 |

`RenderOptions(preset="termsheet")` 또는 최상위 frontmatter `preset: termsheet`로도 지정합니다. 항목별 우선순위는 **명시한 CLI/API 개별 값 > CLI/API 프리셋 값 > YAML `layout` 개별 값 > YAML 프리셋 값 > 프로파일 기본값**입니다. `ib-report` 프리셋은 값을 재정의하지 않습니다. 알 수 없는 이름은 저장 없이 오류로 종료합니다.

모든 프리셋을 7개 프로파일과 조합할 수 있습니다. 표지·목차·면책만 바꾸므로 일반 프로파일의 중립 메타데이터, 네이티브 번호, 일반 표 의미는 유지합니다. `termsheet`는 `ib-report`/`ib-memo`, `legal-memo`는 `ib-memo`/`plain`, `lecture-note`는 `plain`/`business-report`와 조합하면 유용합니다. 공문이나 짧은 회의록에는 보통 표지·목차가 필요하지 않습니다. 프리셋이 법률 문구나 문서별 메타데이터를 생성하지는 않습니다.

표지를 끄면(`--no-cover`, API `include_cover=False`, YAML `layout.cover: false`, `termsheet`/`legal-memo`) `ib-report`는 문서 맨 앞에 테마를 반영한 제목·부제와 `ib-memo`와 같은 작성일·작성자 행을 표시하고, 목차가 있으면 같은 페이지에서 이어 표시한 뒤 기존 페이지 나누기로 본문을 시작합니다. 제목과 일치하는 본문 H1 및 추출된 부제는 이 블록에만 한 번 표시하고 목차에서는 제외하며, 표지를 켠 출력과 다른 프로파일의 제목 처리는 유지합니다.

## 테마 확장 키

기존 소문자 키 6개 외에 다음 `IBStyle` 대문자 표현 필드(PR #5 표기)를 YAML 테마에서 사용할 수 있습니다.

| 종류 | 키 / 단위 |
|---|---|
| 색상 | `NAVY`, `DARK_GRAY`, `LIGHT_GRAY`, `ACCENT_BLUE`, `WHITE`, `RED`, `GREEN`, `ORANGE`, `MEDIUM_GRAY`, `CODE_BG`, `CHART_NEGATIVE_COLOR`, `TABLE_HEADER_COLOR` |
| OOXML 색상 | `NAVY_HEX`, `LIGHT_GRAY_HEX`, `ACCENT_BLUE_HEX`, `GRAY_BORDER_HEX`, `YELLOW_HEX`, `TABLE_HEADER_BG` |
| 글꼴 | `HEADING_FONT`, `BODY_FONT`, `KOREAN_FONT`, `COVER_FONT`, `TOC_FONT` |
| 크기(pt) | `H1_SIZE`–`H4_SIZE`, `BODY_SIZE`, `SMALL_SIZE`, `TABLE_HEADER_SIZE`, `TABLE_BODY_SIZE` |
| 간격(pt) | `H1_SPACE_BEFORE/AFTER`, `H2_SPACE_BEFORE/AFTER`, `H3_SPACE_BEFORE/AFTER`, `BODY_SPACE_AFTER`, `BULLET_SPACE_AFTER` |
| 길이(inch) | `TOP_MARGIN`, `BOTTOM_MARGIN`, `LEFT_MARGIN`, `RIGHT_MARGIN`, `BULLET_INDENT`, `DEEP_LIST_INDENT`, `MAX_LIST_INDENT` |
| 기타 | `BODY_LINE_SPACING`(양수 배수), `FULL_LIST_INDENT_LEVELS`(0–9 정수), `BULLET_CHAR`, `TABLE_ZEBRA`, `BODY_JUSTIFY`, `HEADING_BORDER`(불리언), `TOC_TITLE`, `PAGE_LABEL`, `PAGE_OF_LABEL`(문자열) |

```yaml
NAVY: "234567"
H1_SIZE: 18
BODY_FONT: Calibri
KOREAN_FONT: Malgun Gothic
BODY_SPACE_AFTER: 6
TOP_MARGIN: 0.8
CHART_NEGATIVE_COLOR: "995544"
TABLE_ZEBRA: false
```

색상은 따옴표로 감싼 6자리 16진 문자열입니다(`"#234567"`도 허용). 숫자 색상, 크기의 불리언 값, 비유한 숫자, 알 수 없는 키, 별칭 충돌은 저장 전에 거부합니다. 크기는 1–144pt, 간격은 0–144pt, inch 길이는 0–3, 줄 간격은 0 초과 5 이하입니다. 기존 `body_size`는 6–30pt, `margin_mm`는 5–60mm 범위를 유지합니다. 글꼴 이름은 비어 있을 수 없습니다. 내부 `STYLE_*` 식별자와 프로파일의 `NATIVE_NUMBERING` 정책은 테마 키가 아닙니다.

RGB/hex 쌍은 한쪽만 지정하면 함께 반영되며, 양쪽을 명시하면 각각 적용합니다. `NAVY`는 별도 `TABLE_HEADER_BG`가 없을 때 IB 표 머리 배경에도 적용합니다. `RED`는 별도 `CHART_NEGATIVE_COLOR`가 없을 때 waterfall 음수 색상에도 적용합니다. 일반/mono 차트의 기본 색상은 무채색입니다. 제목·콜아웃·코드 패널·차트는 요청별 불변 스타일을 읽으며 기존 `default`·`mono` 문서 출력은 유지합니다.

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
  - type: financial
    subtotal_rows: [3, 5]   # 본문 행 번호: 머리행 다음 첫 행이 1
    total_rows: [8]
    negative: paren         # paren (1,234), minus -1,234, triangle △1,234
```

- `tables`는 본문 표 순서에 대응합니다. 건너뛸 표는 `{}`로 지정합니다. `header_rows`(기본 1)로 반복 머리행 수를 정하며 모든 프로파일에서 지원합니다.
- 표 종류: `generic`, `financial`, `sensitivity`, `risk`.
- 열 의미: `text`, `code`, `date`, `number`, `money`, `percent`, `bps`, `multiple`. 지정할 때는 모든 열 개수와 일치해야 합니다.
- 열 의미를 지정하지 않으면 첫 본문 값 대부분이 숫자인 열을 숫자 열(오른쪽 정렬)로 봅니다. 빈칸과 대시(`-`, `–`) 자리표시는 판단에서 뺍니다.
- 숫자에는 천 단위 구분을 적용합니다. `percent`의 `12.5`는 `12.5%`가 됩니다. **비율·단위를 환산하거나 재무 계산을 수행하지 않습니다.**
- `code`는 `001234` 같은 선행 0을 보존합니다. 빈 셀·이스케이프 파이프 `A\|B`도 원래 열에 남습니다.
- 표 제목·단위·기준일·출처, 반복 머리행과 행 분할 제어를 지원합니다. `landscape: true`는 해당 표 앞뒤에 구역을 나누어 가로 표를 배치합니다.
- 민감도표 기준 셀은 **명시한 좌표만 강조**합니다. 행 1은 첫 데이터 행, 열은 라벨 열을 포함한 1부터 시작합니다. 기존 중앙 셀 자동 추정은 제거했습니다.
- 위험도 표는 인식 가능한 등급 열과 높음/중간/낮음 등의 값을 강조합니다.
- 합계·소계도 **명시한 행에만** 표시하며 모든 프로파일에서 씁니다. `subtotal_rows`에 적은 본문 행은 위에 단선, 글자는 굵게, `total_rows`에 적은 행은 여기에 아래 이중선이 더해집니다. 행 1은 머리행 다음 첫 행입니다(`header_rows: 2`이면 Markdown 셋째 행이 1). 선 색은 프로파일의 표 강조색(IB 남색, 일반 프로파일 짙은 회색, 텀싯 남색)이고, 합계 행이 표의 마지막 행이면 아래 이중선이 표 아래 바깥선을 대신합니다. `합계`라고 적힌 행이라도 목록에 없으면 아무것도 하지 않으며, `schedule.total`은 검사만 합니다. 목록에 든 행을 세로로 지나가는 병합 칸은 한 칸 그대로이고 Word는 그 안에 선을 긋지 않습니다. 정수가 아니거나 범위 밖이거나 같은 행이 두 번 나오거나 두 목록에 같은 행이 있으면 strict·일반 모드 모두 거부합니다.
- 금융표의 음수는 기본으로 적은 그대로(`(1,234)`, `-1,234`) 빨강으로 표시합니다. `type: financial` 사양에 `negative: paren`·`minus`·`triangle`을 주면 그 표의 음수를 `(1,234)`·`-1,234`·`△1,234`로 통일하고 자동 빨강을 끕니다.
  - 숫자 열에서 칸 전체가 음수 금액 하나일 때만 바꿉니다. `-…`, `−…`, `(…)`, `△…` 모두 읽으며, 단위는 괄호 안이든 밖이든(`(3.5%)`, `(3.5)%`) 같습니다. 같은 금액은 어떻게 적었든 같은 글자가 됩니다.
  - 숫자와 단위는 그대로이고 천 단위 쉼표만 더합니다. 단위는 괄호 안으로 들어갑니다(`(3.5%)`, `(1,200억원)`). percent·bps·multiple 열의 단위 없는 숫자에는 옵션이 없을 때처럼 그 열의 단위가 붙습니다.
  - 그 밖의 칸은 적은 그대로입니다: `1-2` 같은 범위, 한 칸의 여러 숫자나 글자, 음의 0, text·code·date 열, `▲`·`▼`, `NM` 같은 표시.
  - 앞자리 0 금액(정수부가 0으로 시작하는 금액: `(0001.20%)`, `−0001200`, `-01,200` — 코드일 수 있음)에는 새 음수 재표기를 적용하지 않습니다. 기존 숫자 서식 결과를 그대로 따르므로 일부 표기(`-0001200`, `(0001.20)%`, `-00.5`)에서는 옵션 없는 표와 마찬가지로 앞자리 0이 없어집니다. 전각 숫자는 옵션 없는 표가 읽는 대로 읽습니다(`paren`이면 `-１２３４`가 `(1,234)`).
  - 다음도 바꾸지 않고 옵션 없는 표와 같은 글자로 둡니다: 통화 기호가 붙은 수(`-$1,200`, `₩-1,200`), 부호·괄호와 숫자 사이에 공백이 있는 수(`- 1,200`, `( 1,200 )`), `-.5`, 지수 표기, 전각 괄호, `USD`·`%p` 접미사, 금액 앞의 각주 표시, 전체가 위첨자로 적힌 금액. 단위 `백만원`·`천원`·`bp`는 지원합니다.
  - 한 금액이 서식이 다른 조각으로 나뉜 칸(`-1200**.50**`, `-**1200**`, `-1200^.50^`)도 적은 그대로 두고, 변환 중 경고를 남깁니다. 이 경고로 strict 모드가 문서를 거부하지는 않습니다. 금액 앞뒤의 공백과 각주 표시(`&nbsp;`, 줄바꿈, `[^1]`)는 순서와 개수에 관계없이 금액에 들지 않아 서식이 달라도 나뉜 것으로 보지 않으며, 그 글자·서식·자리는 그대로 남습니다. 각주 표시 뒤에 오는 금액은 여전히 바꾸지 않지만, 서식이 나뉘었으면 경고합니다.
  - 옵션이 없는 표에서는 엔진의 자동 빨강이 문서에 직접 지정한 색을 덮습니다(기존 동작). 옵션은 자동 빨강만 끄므로 직접 지정한 색(`<span style="color:…">`)이 그대로 보입니다. 링크 색은 두 경우 모두 그대로입니다.
  - 정렬과 열 너비는 옵션이 없을 때와 같습니다.
  - `negative`가 없으면 `△…`로 적은 값도 적은 그대로 둡니다. 다른 값이나 같은 사양에 `type: financial`이 없는 `negative`는 거부합니다.

## 표·그림 번호와 상호 참조

번호를 붙일 표와 그림에 ID를 주면 엔진이 Word의 "캡션 삽입"과 같은 구조로 `표 1. 제목`, `그림 1. 설명`을 만듭니다. 번호는 Word 필드(`SEQ`)이고 엔진이 미리 계산한 값이 들어 있어, 받은 사람이 Word에서 표를 끼워 넣고 필드를 갱신하면 뒤 번호와 본문 참조가 함께 바뀝니다. ID가 없는 표와 그림은 지금과 똑같이 나옵니다.

```yaml
tables:
  - id: tbl-results        # tbl- 로 시작, 영문 소문자·숫자, 단어는 하이픈으로
    caption: 연도별 실적
    unit: 억원
```

```markdown
![거래 구조도](structure.png){#fig-structure}

@tbl-results에서 보듯 매출이 늘었다. 구조는 @fig-structure 참조.
```

- 표: 표 사양에 `id: tbl-…`를 줍니다. 제목(`caption`)이 있어야 합니다. 제목 줄이 `표 1. 연도별 실적`이 되고 단위 줄·출처·주석은 그대로입니다. 텀싯은 같은 줄 왼쪽 제목 앞에 번호가 붙고 오른쪽 `(단위 : …)`는 그대로입니다.
- 그림: 이미지 파일 한 줄의 끝에 `{#fig-…}`를 붙입니다. 설명(`![설명]`)이 있어야 합니다. 그림 아래 설명 줄이 `그림 1. 거래 구조도`가 됩니다. 문단·목록·표 칸 안의 그림 뒤에 쓴 `{#…}`, Base64로 넣은 그림의 ID는 번호를 붙이지 않고 보고합니다. 차트·도식에도 번호를 붙이지 않습니다.
- 번호는 문서 처음부터 표와 그림을 따로 셉니다(장 번호 `표 2-1`은 넣지 않습니다). 실제로 그려지는 표와 그림만 셉니다. 레이블은 모든 프로파일에서 `표`, `그림`입니다.
- 참조: 본문에 `@tbl-…`, `@fig-…`를 쓰면 Word 상호 참조(`REF` 필드)가 되어 `표 1`, `그림 1`로 보이고 클릭하면 그 표로 갑니다. 문단·목록·인용·제목·표 칸(HTML 표 포함)·표 제목/단위/출처/주석에서 쓸 수 있습니다.
- 참조로 읽는 규칙:
  - `@` 바로 앞이 줄 처음, 공백, 여는 괄호나 따옴표(`(`, `[`, `「`, `"` 등)여야 합니다. 앞에 글자·숫자나 `+`, `.`, `_`, `-`가 붙은 `@`(이메일 `user+@tbl-x.example`, `doc@tbl-x.kr`)와 겹쳐 쓴 `@@`는 참조가 아닙니다. 그 밖의 글자 바로 뒤에 붙여 쓴 것(`@tbl-a,@tbl-b`, `@tbl-a~@tbl-b`, `상세:@tbl-a`, `다음@tbl-a`)도 참조가 아니어서 글자로 남는데, 그 ID가 이 문서의 표·그림 ID이면 참조로 쓰려던 것으로 보고 오류로 보고합니다. `@` 앞에 공백을 넣거나(`@tbl-a, @tbl-b`), 글자로 두려면 `\@tbl-b`로 씁니다. 글자로만 쓰이는 frontmatter 값에서는 이 보고를 하지 않습니다.
  - `tbl-`, `fig-`는 소문자로 씁니다. `@Tbl-a`, `@FIG-a`처럼 대문자가 섞이면 모양이 틀린 참조로 보고합니다.
  - 차트·도식 블록 안의 글자는 코드 블록과 같아서 참조로 읽지 않습니다.
  - `@` 뒤의 ID는 영문·숫자·`-`·`_`가 이어지는 데까지 통째로 읽어 검사합니다. `@tbl-a-`, `@tbl-`, `@tbl-Results`처럼 모양이 틀리면 앞부분만 참조로 쓰지 않고 오류로 보고합니다. ID 뒤에는 한글 조사, 공백, 문장부호가 바로 와도 됩니다(`@tbl-results에서`, `(@tbl-results)`).
  - ID 뒤에 `.`과 영문·숫자가 이어지면(`@tbl-x.example` 같은 도메인) 참조가 아닙니다.
  - 코드(`` `@tbl-a` ``, 코드 블록), 수식(`$…$`), 링크 주소 안, 역슬래시로 막은 `\@tbl-…`는 어느 자리에서나 참조가 아닙니다.
  - 글자는 읽는 사람이 보는 대로 읽습니다. 굵게·기울임·색이 바뀌어도 이어 읽고, 내용이 빈 색 지정(`<span …></span>`)은 없는 것으로 봅니다. 서식이 ID를 가르면(`@tbl-a**b**`, `@tbl-a<span …>b</span>`) 참조로 쓰지 않고 오류로 보고합니다. 도메인 규칙도 보이는 ID 전체를 읽은 뒤 판정합니다(`@tbl-a**b**.com`은 참조가 아닙니다).
  - 제목은 본문과 같이 `**` 표시가 있는 원문으로 판정합니다. `## @tbl-**a**`는 오류로 보고합니다.
  - 조건 변수(`{{key}}`)가 넣은 값은 참조로 읽지 않습니다. 그 제목이 문서 제목이 되어도 같습니다.
  - 참조는 파싱한 뒤의 글자에서 읽습니다. `@**tbl-a**`, HTML `&#64;tbl-a`, frontmatter 값에서 YAML 이스케이프로 쓴 `@`도 다른 참조와 같이 읽고 보고합니다.
  - 참조로 읽지 않은 글자와 풀리지 않는 참조는 원문 그대로, 기능을 넣기 전과 같은 모양으로 나옵니다. 사용자 정의 영역 문자(U+E000 등)도 그대로입니다. 엔진은 풀리는 참조가 있는 자리만 나눕니다.
- 번호가 바뀌면 "표 1과"가 "표 2과"처럼 조사가 어긋날 수 있습니다. "표 1에서", "표 1의", "표 1 참조"처럼 받침과 관계없는 조사를 권합니다.
- Word에서 번호 갱신: 표나 그림을 Word에서 더 넣었다면 본문 전체를 선택(Ctrl+A)하고 F9를 누릅니다. 문서를 열기만 하거나 인쇄 미리보기를 하는 것으로는 갱신되지 않습니다(이 기기의 Word로 확인). 엔진은 열었을 때 바로 맞는 값이 보이도록 번호와 참조 글자를 미리 적어 둡니다.
- Word에서 그림을 새로 넣으며 번호를 이으려면: 한국어 Word에는 `그림` 레이블이 기본으로 없습니다. 참조 › 캡션 삽입 › 새 레이블에서 `그림`을 한 번 만든 뒤 그 레이블로 캡션을 넣으면 엔진이 만든 그림 번호에 이어집니다. 표는 기본 레이블 `표`를 그대로 쓰면 됩니다.
- 거부되는 입력: 위 규칙으로 참조인 자리에서 없는 ID를 가리키는 참조, 모양이 틀린 참조(`@tbl-a-`, `@tbl-`, `@tbl-Results`), 서식이 ID를 가른 참조(`@tbl-a**b**`), 링크 글자 안의 참조, 그림 설명(인라인 그림과 HTML `data:` 그림 포함)·각주·문서에 나오는 frontmatter 값(아래) 안의 참조(이 자리에는 Word 필드를 넣을 수 없습니다)는 strict 모드에서 저장하지 않습니다. 기본 모드는 글자를 원문 그대로 두고 경고만 남깁니다. 같은 ID를 두 번 쓴 그림, 번호를 붙일 수 없는 `{#…}`(모양 오류 `{#Fig-1}`, 표용 `{#tbl-…}`, 설명 없는 그림, Base64 그림, 문단·목록·표 칸 안의 그림)도 strict에서 거부하고 기본 모드에서는 지금처럼 글자로 둡니다. 표 사양 `id`의 모양 오류, 중복, 제목 없음은 다른 표 사양 오류처럼 두 모드 모두 거부합니다. 칸이 없어 그려지지 않는 표의 `id`는 strict 모드에서 거부하고, 기본 모드에서는 경고를 남기고 그 표에 번호를 주지 않습니다.
- 참조를 검사하는 frontmatter 값(글자로만 쓰입니다): 문서 제목과 부제는 언제나, 날짜는 `plain`을 뺀 모든 프로파일에서 조건 변수를 넣기 전 원문으로 검사합니다. 나머지 값은 문서에 나오는 자리에서만 검사합니다. IB 표지: `report_type`, `sector`, `ticker`, `subject_company`(없으면 `company`), `analyst`, `recipient`, `analysis_period`, `analysis_basis`. 공문 별지 첫머리: `letter.appendix_label`. 제목 블록: `sender.department`, `analyst`, `meeting_time`, `location`, `attendees`(공문은 `sender.organization`, `recipients`, `cc`). 공문 끝: `attachments`, `sender.signatory`(없으면 `sender.organization`), `document_no`, `sender.department`, `sender.contact`. 머리글: `sender.organization`(없으면 `company`)과 `version`(공문 제외). 면책 쪽: `company`. 텀싯: `version`과 `prepared_by`·`disclaimer`·기밀 표시 문구. 값은 보이는 모양대로 읽습니다(일반 글자, 텀싯 문구는 인라인 Markdown). `report_type`은 대문자로 나오므로 참조가 될 수 없습니다. IB 프로파일이 본문 앞쪽 문단(`**작성자:** …`)에서 읽은 값은 그 문단에서 판정합니다. 그 안의 조건 변수 값, 이스케이프한 `@`, 코드는 참조가 아니고, 그 문단에서 풀린 참조는 값이 보이는 자리에서 보고합니다.
- 풀린 참조는 자기 글자 말고는 아무것도 바꾸지 않습니다. 표 칸에서 참조 옆의 금액은 ID가 없을 때와 같은 모양을 유지하고, IB 목차 미리보기는 제목을 쓴 그대로 두고 참조만 `표 1`로 바꿉니다.
- `docx-audit`: 대상 북마크가 없는 참조(Word에서 대상 캡션을 지운 경우)는 issue, 갱신이 필요한 참조·번호(F9 전 상태)는 warning으로 보고합니다. 번호 순서는 스위치가 없거나 `\* ARABIC`만 있는 `SEQ` 열만 판정합니다. 본문만 검사합니다(머리글·바닥글·각주는 검사하지 않습니다).

## 자산 허용 폴더·입력 한도·거부 코드

**이미지·로고 경로 제한은 기존 동작을 바꾸는 변경입니다.** 이전에 읽혔던 로컬 이미지도 허용 폴더 밖이면 거부될 수 있습니다. Markdown 파일은 그 파일이 있는 폴더와 하위 폴더를 기본으로 허용합니다. 원본 파일 없이 넘긴 Markdown 글자·스트림과 직접 구성한 문서 모델은 기본 허용 폴더가 없습니다. 현재 작업 폴더를 자동으로 허용하지 않습니다.

운영자는 `--asset-root 폴더`를 여러 번 지정해 폴더를 추가할 수 있습니다. 대응 API 인자 `asset_roots`는 `parse_markdown_file`, `MarkdownParser`, `IBDocumentRenderer.render`, 레지스트리 변환 호출에서 받으며, 변환기 설정에 `RenderOptions(asset_roots=(... ,))`도 사용할 수 있습니다. 파싱된 모델은 허용 폴더를 별도 렌더·레지스트리 단계까지 전달합니다. 추가한 폴더는 하위 폴더도 포함하며 기본 폴더를 빼지는 못합니다. 허용 폴더와 한도는 운영자의 CLI/API 인자로만 바꿀 수 있습니다. 본문·앞머리·기관 파일·테마 파일로는 바꾸지 못합니다.

```sh
uv run md-to-word job/draft.md output.docx --strict --asset-root shared-assets
```

상대 경로, 심볼릭 링크와 정션을 따라 최종 위치를 해석한 뒤 허용 여부를 판단하며, 파일 존재 여부는 그다음에 확인합니다. 폴더 밖으로 나가는 `../`, 절대 경로·UNC 경로, 밖을 가리키는 링크는 거부합니다. 최종 위치가 허용 폴더 안이면 절대 경로와 UNC 경로도 허용하며, 네트워크 공유의 Markdown 폴더도 같은 규칙을 따릅니다. 퍼센트 인코딩 경로의 원문과 디코딩한 후보 각각에 같은 검사를 적용합니다.

| 자산·참조 | 경로 제한 |
|---|---|
| 본문·참조 정의·HTML·표 셀 이미지 | 적용 |
| 앞머리 `style.logo` | 적용 |
| 문서 앞머리가 지정한 기관 파일의 로고 | 적용. 기관 파일 자체가 다른 폴더에 있어도 동일 |
| 운영자가 `--house` 또는 API `RenderOptions(house=...)`로 직접 지정한 기관 파일의 로고 | 미적용. 앞머리에서 로고를 덮어쓰면 그 로고에는 적용 |
| `data:` / Base64 이미지 | 경로 제한 없음. 이미지 크기 한도는 적용 |
| 앞머리 `house:`, `theme:` 파일 경로 자체 | 경로 제한 없음 |
| 로컬 파일 하이퍼링크 | 경로 제한 없음. 파일 내용은 문서에 넣지 않음 |

기관 파일 안의 로고는 기관 파일 기준, 앞머리 로고는 Markdown 기준으로 경로를 해석합니다. 원본 파일이 없으면 상대 경로 기관 파일·로고는 계속 지원하지 않으며, 문서가 지정한 절대 경로 로고에는 명시적인 허용 폴더가 필요합니다. 원격 HTTP(S) 이미지는 계속 지원하지 않습니다. 폴더 밖 이미지는 읽지 않습니다. 일반 모드에서는 본문 이미지에 자리표시를 남기거나 로고를 생략하고 `asset_outside_root` 경고를 남기며, 엄격 모드에서는 저장을 거부합니다.

다음 한도를 넘으면 **일반 모드와 엄격 모드 모두** 변환을 중단하고 DOCX를 만들지 않습니다. MiB는 1,048,576바이트이며, CLI에는 단위를 붙이지 않은 정수 개수를 지정합니다.

| 코드 | 측정 대상 | 기본값 | CLI 옵션 / `InputLimits` 필드 |
|---|---|---|---|
| `limit_input_bytes` | 원본 Markdown 바이트 수 | 16 MiB (16,777,216) | `--max-input-bytes` / `max_input_bytes` |
| `limit_table_cells` | 표 하나의 칸 수. 머리행 포함 행 수 × 열 수, HTML 병합은 펼친 격자 기준 | 50,000 | `--max-table-cells` / `max_table_cells` |
| `limit_list_depth` | 목록 중첩 단계 수 | 32 | `--max-list-depth` / `max_list_depth` |
| `limit_image_bytes` | 이미지·로고 하나의 바이트 수. Base64 / `data:`는 디코딩한 크기 | 20 MiB (20,971,520) | `--max-image-bytes` / `max_image_bytes` |
| `limit_image_pixels` | 이미지·로고 하나의 가로 × 세로 픽셀 수 | 50,000,000 | `--max-image-pixels` / `max_image_pixels` |

이미지 한도는 운영자가 직접 지정한 기관 파일의 로고를 포함해 모든 출처에 적용합니다. 한도를 넘는 Base64는 인코딩된 길이로 먼저 판단해 디코딩 전에 거부합니다. 옵션은 양의 정수만 받으며 무제한 값은 없습니다. API에서는 `InputLimits` 인스턴스를 파서·렌더러·레지스트리·`IBReportConverter`의 `limits` 인자 또는 `RenderOptions(limits=...)`로 넘깁니다. 예:

```python
from input_safety import InputLimits
from md_parser import parse_markdown_file

model = parse_markdown_file(
    "job/draft.md",
    asset_roots=("shared-assets",),
    limits=InputLimits(max_table_cells=60_000),
)
```

이 규칙은 `md-to-word`, `ib-report`, 파일·글자·스트림 API, 직접 구성 모델의 렌더, 레지스트리, 일괄 변환에 적용됩니다. 글자 입력은 UTF-8 바이트 수로 측정합니다. `--format` 등 변환 전 정리 단계를 쓰면 정리 전 원본과 정리 후 파싱할 Markdown의 크기를 각각 검사합니다. 일괄 변환에서는 한도를 넘은 파일만 실패하고 나머지는 계속 처리하되, 전체 종료 상태는 실패를 나타냅니다. 별도 명령인 `md-format`과 `docx-audit`의 입력은 이 크기 한도 대상이 아닙니다.

거부 코드는 `asset_outside_root`와 위 다섯 `limit_*`를 합친 정확히 여섯 개입니다. 코드가 있는 개별 경고·문제 문장은 `[코드] `로 시작하며, 엄격 거부에서 여러 문장을 합쳐도 각 접두를 유지합니다. 기존에 코드가 없던 메시지는 그대로입니다. API 예외는 `ValueError`의 하위인 `ConversionRejection`을 상속합니다. `StrictValidationError.codes`는 포함된 코드 목록이며, `InputLimitError`는 `code`, `limit`, `measured`를 제공합니다. 감싼 예외에서도 `__cause__`를 따라 원래 거부를 읽을 수 있습니다. 플러그인 실행기의 `result.json`에는 문제마다 해당 `code`와 접두를 포함한 문장 전체를 `message`로 갖는 진단을 추가합니다. 기존 `input_validation`, `structural_issue`, 예외 종류 이름, `invalid_arguments` 항목은 유지하며, 기존 엄격 거부 항목의 코드도 계속 `ValueError`입니다.

플러그인 실행기와 휴대용 변환기는 항상 엄격 모드이며 **허용 폴더를 추가하거나 한도를 올리는 옵션이 없습니다.** 필요한 이미지, 기관 파일과 로고를 초안이 있는 작업 폴더(하위 폴더 포함)로 복사하고 상대 경로로 가리키세요. 실제 문서 작업 폴더는 저장소·플러그인 밖에 둡니다. `asset_outside_root`가 나오면 필요한 파일을 복사하고 참조 경로를 수정합니다. 한도 거부는 표시된 한도·측정값에 맞춰 입력·표·이미지를 줄이거나 나누되 원자료의 사실을 유지합니다. 휴대용 창은 코드별 한국어 원인·해결 방법을 보여 주며, 실행기는 `result.json`에 진단을 기록하고 시도마다 새 출력 폴더를 요구합니다. 실행 명령 자체의 런처·초안·출력 폴더 절대 경로는 계속 사용할 수 있습니다.

## 문법·검증의 범위

제목, 문단, 강조, 표, 목록, 인용/콜아웃, 파일·Base64 이미지, 코드블록, 기존 수식·도식 출력을 지원합니다. 글머리표는 `-`, `*`, `+`이고 제목은 `#`부터 `######`까지입니다(5·6단계는 4단계 모양). 역슬래시는 ASCII 문장부호를 모두 이스케이프하며(`\#`, `\%`, `\*`, `\<`), 코드·링크 주소·달러 기호 사이(`$5\%$`)의 역슬래시는 그대로 둡니다. 강조 기호는 CommonMark 규칙대로 글자에 붙어 있어야 하므로 `2 * 3 * 4`나 공백이 뒤따르는 주석 표시(`* 주요 매출처`)는 글자 그대로 남습니다. 문서 변환기가 만든 HTML `<table>`도 Word 표로 만듭니다. `colspan`/`rowspan`은 셀 병합, 블록 태그와 `<br>`는 셀 안 줄바꿈이 되고, `<b>`/`<i>`/`<code>`/`<sup>`/`<sub>`, 색 지정 `<span>`, `<a href>`는 서식으로 옮기며, `<img>`는 셀 이미지, 중첩 표는 해당 셀 안으로 펼칩니다. HTML 글자는 Markdown으로 다시 해석하지 않고 그대로 두며, `{{key}}` 조건 변수는 치환합니다. 첫 행의 행 병합이 덮는 행은 별도 머리행으로 유지합니다(`header_rows`). 닫히지 않은 표와 셀 밖 글자는 보고하며 strict 모드는 거부합니다. 본문과 표 셀 안의 이미지(`![설명](경로)`)는 글자 사이에 넣고 셀 폭에 맞춥니다. `%20`처럼 퍼센트 인코딩된 이미지 경로도 찾으며, `<img>` 한 줄은 이미지로 처리합니다. 이미지 설명(`![설명](경로)`, `<img alt="...">`)은 Word의 대체 텍스트가 됩니다. 설명 없이 들어간 이미지는 파일 이름(Base64·`data:` 이미지는 `embedded image`)과 Word 이름(`Picture 3`)을 적은 경고를 남기며, strict 모드에서도 저장하고 문서는 바뀌지 않습니다. `docx-audit`는 본문에서 대체 텍스트가 없는 그림을 모두 경고하므로, 엔진이 그린 수식·도식·로고(설명을 적을 곳이 없음)도 함께 나옵니다. 외부 HTTP(S)/메일 링크와 숫자 각주 `[^1]` / `[^1]: 설명`는 Word 네이티브 요소로 만듭니다. 로컬 파일 링크도 주소를 `<...>`로 감싸거나, `C:/`·`./`·`../`로 시작하거나, 문서 확장자(`.md`, `.docx`, `.xlsx`, `.pdf`, `.hwp(x)`, 이미지 등)로 끝나면 하이퍼링크로 만듭니다. 상대 경로는 Markdown 폴더 기준이며, 문서 확장자 뒤의 `#`은 문서 내 위치(fragment)입니다. 절대 경로(`C:/...`, `/...`)는 파일 이름이므로 그 안의 `#`, `%`는 글자 그대로입니다. CLI, `IBReportConverter`, 레지스트리가 DOCX를 저장할 때 상대 링크는 DOCX 폴더에서 같은 파일을 가리키도록 고치고, DOCX 폴더 안의 절대 경로는 상대 링크로 바꿔 폴더를 공유하면 동료 PC에서도 열리고 내 PC 경로가 드러나지 않게 합니다. 그 밖의 절대 경로는 `file:///` 링크로 두고 로그로 경고합니다. 직접 저장하는 문서는 Markdown 폴더 기준 상대 링크를 그대로 가집니다. 인라인 코드는 백틱 없이 코드 글꼴(`CODE_FONT`, 기본 Consolas)로 보여 주며, 내용은 글자 그대로(강조·줄바꿈·각주·조건 치환 없음) 두고 표 숫자 서식에서도 제외합니다. 모든 프로파일은 한글과 영문·숫자 사이의 Word 자동 간격을 꺼서 `SPC는`, `제2종`, `300억원`처럼 붙여 씁니다. 각주는 숫자 식별자·한 줄 정의를 사용하십시오.

일반 문서의 References는 본문으로 보존합니다. 문단의 자연스러운 줄바꿈은 공백으로 합치며, 강제 줄바꿈은 `<br>` 또는 줄 끝 역슬래시를 사용합니다.

`<br>`, `<br/>`, `<BR />`는 본문·강조·제목·목록·표 셀 중간에서도 실제 Word 줄바꿈으로 출력합니다. 이스케이프한 `\<br>`와 코드의 리터럴은 보존하며 링크 주소는 바꾸지 않습니다. 줄 끝 공백 2칸 방식은 파서의 명시적 opt-in 설정에서만 지원합니다.

`--strict`는 알려진 입력 손실·각주 오류·요소 렌더 실패·미해결 이미지 표시를 검사하고 실패하면 저장하지 않습니다. 기본 모드에서는 경고와 함께 부분 문서를 생성할 수 있습니다. `docx-audit`는 구조 검사이지 Word→MD나 시각 검사가 아닙니다.

검은 사각형이 보이면 실제 글머리표인지, Word의 **인쇄되지 않는 문단 페이지 제어 표시**인지 먼저 구별하십시오. `docx-audit`의 `pagination_marked_paragraphs`는 표 안의 문단과 스타일 상속까지 고려한 keep-lines/keep-next/page-break-before 문단 개수이며, 글머리표 개수가 아닙니다. `warnings`는 Normal의 과도한 페이지 묶음을 경고하고 `issues`는 구조 오류를 보고합니다. 필요한 제목 제어까지 일괄 제거하거나 사용자의 Word 표시 설정을 임의 변경하지 않습니다. `visual_review: not_performed`는 페이지를 눈으로 확인하지 않았다는 뜻입니다.

아래 사항은 별도 확인 또는 후속 개발 대상입니다.

- 완전한 CommonMark/GFM·복잡한 중첩 문법·임의 HTML 지원
- 회사 DOCX 원본 양식 자동 복제, HWP/HWPX 출력
- 네이티브 편집형 차트·재무모델 계산·법적 적합성 판단
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
uv run ruff check .
uv run ruff format --check .
uv run mypy .
uv build
```

CI는 위 명령과 함께 산출물을 지키는 검사 세 가지를 돌립니다.

- **Word 스키마 게이트.** 별도 CI 작업이 모든 샘플을 strict 모드로 변환하고 버전과 해시를 고정한 OOXML-Validator로 검사합니다. `scripts/word_schema_exceptions.json`에 없는 오류가 하나라도 있으면 실패합니다. 로컬 실행 방법과 예외 규칙은 [scripts/word_schema_validation.md](scripts/word_schema_validation.md)에 있습니다. 통과는 구조 검사 결과이며 시각 검토가 아닙니다.
- **스냅숏.** `tests/test_sample_snapshots.py`가 샘플마다 파싱한 모델과 DOCX 안의 모든 XML 파트를 저장된 스냅숏과 비교합니다. 스냅숏이 없으면 실패합니다. 출력을 의도적으로 바꿨다면 `uv run pytest tests/test_sample_snapshots.py --snapshot-update`로 갱신하고, 차이를 검토한 뒤 바뀐 샘플을 Word에서 확인하십시오.
- **속성 테스트.** `tests/test_markdown_properties.py`가 고정된 예시로 Markdown을 생성해 보이는 글자와 표 칸이 보존되는지, 코드와 이스케이프가 글자 그대로 나오는지, strict 변환이 문서를 저장하거나 파일을 남기지 않고 문서화된 거부로 끝나는지 검사합니다. strict `xfail`로 남은 한 건은 코드 블록 안 트리 표기의 값 강조로, 의도한 서식입니다. `tests/test_config_properties.py`는 같은 저장·거부 검사를 생성한 기관·테마 파일 내용과 별칭을 쓴 YAML에 시간 제한을 두고 적용합니다.

Python 3.12 이상을 지원하며 CI는 Ubuntu와 Windows의 Python 3.12 및 3.13에서 검사합니다. 모델은 `document_model.py`로 분리했으며 기존 `md_parser` 모델 import는 유지합니다. `ib_renderer`와 `md_parser`는 내부 파일(`ib_render_*.py`, `md_parse_*.py`)로 나뉘었으며, 가져오는 곳은 종전대로 `ib_renderer`와 `md_parser`입니다. CLI·API·레지스트리는 하나의 조립 경로를 사용합니다. API에서는 파싱 단계부터 원하는 프로파일을 전달하십시오. 파싱된 모델은 적용한 프로파일을 기록하며, 렌더링 시 다른 프로파일을 요청하면 strict 여부와 무관하게 `ValueError`를 발생시킵니다. 원하는 프로파일로 Markdown을 다시 파싱해야 합니다. 파싱 이력이 없는 직접 구성 모델은 렌더 프로파일을 선택할 수 있습니다.

역변환 모듈과 `roundtrip-audit` 명령은 제거되었습니다. 제거 전 코드는 `d819bbb` 및 로컬 `codex/archive-word-to-md-d819bbb` 브랜치에서 복구할 수 있습니다. 기본 IB 경로는 유지하되 빈칸·숫자 출력·머리말/쪽번호·면책 설정 오류를 고쳤으므로 해당 결과는 의도적으로 달라집니다.

과거 계획서의 양방향 변환 계획보다 이 문서와 2026-09-14 구현계획이 우선합니다.

## 라이선스

[MIT License](LICENSE) — Copyright (c) 2026 Hank. 의존성·글꼴·이미지·샘플 등 제3자 자료의 라이선스는 별도이며 공개 릴리스 전에 확인합니다.
