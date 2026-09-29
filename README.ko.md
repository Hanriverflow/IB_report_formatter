# IB Report Formatter 2.0

IB 보고서와 회사 업무문서를 위한 **Markdown → 편집 가능한 Word 전용 엔진**입니다.

[구현계획](docs/implementation-plan-20260914.md) · [검증 결과](docs/verification-20260914.md) · [예제 7종](samples/profiles) · [English / API 상세](README.md)

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

`title`은 필수이며 `subtitle` 기본값은 `Term Sheet`입니다. 선택 항목인 `date`는 입력한 표기대로 표시하고 `version`은 바닥글에 씁니다. `prepared_by`와 `disclaimer`는 frontmatter나 house 중 한 곳에 공백이 아닌 문구로 있어야 합니다. House YAML은 `prepared_by`, `disclaimer`, `confidential_label`, `confirmation`(`intro`, `items` 목록, `signature`)만 받습니다. 같은 키를 frontmatter에 지정하면 빈 값도 house보다 우선합니다. 기관 문구에는 `( *주요내용* )` 같은 인라인 강조를 쓸 수 있습니다.

실제 딜 문서와 실제 기관 house 파일은 **저장소 밖**에 보관하십시오. Frontmatter의 `house` 상대 경로는 원본 Markdown 폴더 기준입니다. `--house /absolute/path/to/house.yaml`로 파일을 재지정할 수 있으며 CLI의 상대 경로는 실행 폴더 기준입니다. 원본 경로가 없는 문자열·스트림 입력은 절대 house 경로가 필요합니다. 저장소에 포함된 house는 새로 작성한 가상 문구입니다.

장 번호는 `## 1. 본건 개요`처럼 직접 쓰고 소제목은 `###`를 사용합니다. `tables`는 본문 표 순서에 맞추며 설정이 없는 표도 `{}`로 자리를 유지합니다. 위 코드는 표 두 개의 작성 예시이며, 전체 예제를 복사할 때는 예제 파일의 표 사양 순서를 사용하십시오.

- 셀 내용 전체가 `^^`이면 위로, `<<`이면 왼쪽으로 병합합니다. 병합 영역은 직사각형이어야 하며 머리행과 본문 사이를 넘을 수 없습니다. 빈칸은 병합하지 않습니다. 기호를 그대로 쓰려면 `\^^`, `\<<`로 이스케이프합니다. 잘못된 병합 묶음은 원문과 경고를 남기며 strict에서 거부합니다.
- 텀싯은 병합이 기본 활성화됩니다. `spans: false`로 끌 수 있고, 다른 프로파일은 `spans: true`로 명시해야 활성화됩니다.
- `label_columns`는 선행 라벨 열 수를 재지정합니다(0 이상, 전체 열 수 미만의 정수). 생략하면 병합 후 첫 머리 셀의 폭을 사용하되 전체 열 수−1을 넘지 않습니다. `| 구 분 | << | 내 용 |`은 2개, `| 구 분 | 내 용 |`은 1개입니다.
- 라벨 1~2개 뒤에 내용 열 하나가 오는 조건표는 첫 라벨 폭 33.5mm, 음영·굵은 글씨를 씁니다. 둘째 라벨은 30mm, 흰 바탕·보통 굵기·회색 글씨입니다. 그 밖의 격자 표는 내용에 따라 폭을 정하고 라벨에 음영과 보통 굵기를 적용합니다. `label_columns: 0`이면 라벨 음영이 없습니다.
- `columns`로 날짜·코드·숫자 역할을 정하고 `unit`으로 표 위 단위를 표시합니다. 선택 문자열 `note`는 표 아래 오른쪽에 8pt 회색으로 표시하며 모든 프로파일에서 지원합니다. 금융 의미를 추론하거나 상환스케줄을 계산하지는 않습니다.

셀 안 줄바꿈은 `<br>`를 사용합니다. 텀싯 본문 문단의 `<br>`도 줄마다 별도 Word 문단으로 만듭니다. 줄 앞 `•`, `-`, `·`는 각각 0/3/6mm에서 시작하고 3mm 내어쓰기, `①`–`⑳`는 0mm 시작·4.5mm 내어쓰기, `※`는 0mm 시작·4mm 내어쓰기와 8pt 글씨를 적용합니다. 이어지는 줄은 기호 뒤 본문에 맞춰 정렬됩니다. 본문의 Markdown `-` 목록은 기존 목록 처리를 유지합니다.

고객확인란을 넣을 위치에 비어 있고 닫힌 펜스를 둡니다.

````markdown
```confirmation
```
````

House/frontmatter 문구로 안내·체크 항목 행과 음영·가운데 정렬 서명 행을 만들고 한 페이지에 모읍니다. 문구 누락, 내용이 있는 펜스, 닫히지 않은 펜스는 원문을 진단 코드 패널로 보존하며 strict에서 거부합니다. 다른 프로파일에서는 일반 코드 블록입니다.

모든 페이지 머리글에 `confidential_label`(기본 `Strictly Confidential`)을 표시합니다. 빈 라벨, `layout.confidential: false`, `--no-confidential`로 끌 수 있습니다. 바닥글 왼쪽은 `version`이 있을 때 `subtitle version`, 오른쪽은 `PAGE / NUMPAGES`입니다. 첫머리 면책은 `--no-disclaimer`와 무관하게 필수입니다. 기존 **`termsheet` 프리셋**은 표지·목차·끝 면책만 전환하며, 텀싯 프로파일 선택·기관 문구 로딩·조건표 서식은 적용하지 않습니다.

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

제목, 문단, 강조, 표, 목록, 인용/콜아웃, 파일·Base64 이미지, 코드블록, 기존 수식·도식 출력을 지원합니다. 강조 기호는 CommonMark 규칙대로 글자에 붙어 있어야 하므로 `2 * 3 * 4`나 공백이 뒤따르는 주석 표시(`* 주요 매출처`)는 글자 그대로 남습니다. 문서 변환기가 만든 HTML `<table>`도 Word 표로 만듭니다. `colspan`/`rowspan`은 셀 병합, `<br>`/`<p>`/`<li>`는 셀 안 줄바꿈이 되고, `<b>`/`<i>`/`<a href>`는 의미를 유지하며, `<img>`는 셀 이미지, 중첩 표는 해당 셀 안으로 펼칩니다. 머리행이 여러 줄이면 열마다 제목을 줄바꿈으로 이어 한 줄로 합칩니다. 닫히지 않은 표와 셀 밖 글자는 보고하며 strict 모드는 거부합니다. 본문과 표 셀 안의 이미지(`![설명](경로)`)는 글자 사이에 넣고 셀 폭에 맞춥니다. `%20`처럼 퍼센트 인코딩된 이미지 경로도 찾으며, `<img>` 한 줄은 이미지로 처리합니다. 외부 HTTP(S)/메일 링크와 숫자 각주 `[^1]` / `[^1]: 설명`는 Word 네이티브 요소로 만듭니다. 로컬 파일 링크도 주소를 `<...>`로 감싸거나, `C:/`·`./`·`../`로 시작하거나, 문서 확장자(`.md`, `.docx`, `.xlsx`, `.pdf`, `.hwp(x)`, 이미지 등)로 끝나면 하이퍼링크로 만듭니다. 상대 경로는 Markdown 폴더 기준이며, 문서 확장자 뒤의 `#`은 문서 내 위치(fragment)입니다. 절대 경로(`C:/...`, `/...`)는 파일 이름이므로 그 안의 `#`, `%`는 글자 그대로입니다. CLI, `IBReportConverter`, 레지스트리가 DOCX를 저장할 때 상대 링크는 DOCX 폴더에서 같은 파일을 가리키도록 고치고, DOCX 폴더 안의 절대 경로는 상대 링크로 바꿔 폴더를 공유하면 동료 PC에서도 열리고 내 PC 경로가 드러나지 않게 합니다. 그 밖의 절대 경로는 `file:///` 링크로 두고 로그로 경고합니다. 직접 저장하는 문서는 Markdown 폴더 기준 상대 링크를 그대로 가집니다. 인라인 코드는 백틱 없이 코드 글꼴(`CODE_FONT`, 기본 Consolas)로 보여 주며, 내용은 글자 그대로(강조·줄바꿈·각주·조건 치환 없음) 두고 표 숫자 서식에서도 제외합니다. 일반 프로파일(`plain`, `office-letter`, `business-report`, `meeting-minutes`)과 `term-sheet`는 한글과 영문·숫자 사이의 Word 자동 간격을 꺼서 `SPC는`, `제2종`, `300억원`처럼 붙여 씁니다. 각주는 숫자 식별자·한 줄 정의를 사용하십시오.

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
uv build
```

Python 3.12 이상을 지원하며 CI는 Ubuntu와 Windows의 Python 3.12 및 3.13에서 검사합니다. 모델은 `document_model.py`로 분리했으며 기존 `md_parser` 모델 import는 유지합니다. CLI·API·레지스트리는 하나의 조립 경로를 사용합니다. API에서는 파싱 단계부터 원하는 프로파일을 전달하십시오. 파싱된 모델은 적용한 프로파일을 기록하며, 렌더링 시 다른 프로파일을 요청하면 strict 여부와 무관하게 `ValueError`를 발생시킵니다. 원하는 프로파일로 Markdown을 다시 파싱해야 합니다. 파싱 이력이 없는 직접 구성 모델은 렌더 프로파일을 선택할 수 있습니다.

역변환 모듈과 `roundtrip-audit` 명령은 제거되었습니다. 제거 전 코드는 `d819bbb` 및 로컬 `codex/archive-word-to-md-d819bbb` 브랜치에서 복구할 수 있습니다. 기본 IB 경로는 유지하되 빈칸·숫자 출력·머리말/쪽번호·면책 설정 오류를 고쳤으므로 해당 결과는 의도적으로 달라집니다.

과거 계획서의 양방향 변환 계획보다 이 문서와 2026-09-14 구현계획이 우선합니다.

## 라이선스

[MIT License](LICENSE) — Copyright (c) 2026 Hank. 의존성·글꼴·이미지·샘플 등 제3자 자료의 라이선스는 별도이며 공개 릴리스 전에 확인합니다.
