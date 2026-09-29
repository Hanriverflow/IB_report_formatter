# 텀싯 구현 계획

> 작성일: 2026-09-29. 기준 설계: [term-sheet-design-20260929.md](term-sheet-design-20260929.md). 이 문서는 실행 계획이며 검증 기록이 아니다. Claude의 초안과 Codex(`gpt-6-astra`, 읽기 전용)의 독립 초안을 대조해 확정했다. 설계와 다른 결정은 §2에 모았고, 구현이 끝나면 설계서에 반영한다.

## 1. 실행 체계

| 흐름 | 담당 | 브랜치 | 범위 |
|---|---|---|---|
| 기반 | Claude(설계·통합) | `feat/term-sheet` | 설계·계획 문서, 공통 뼈대 커밋 |
| A | Codex `gpt-6-astra` | `feat/term-sheet-profile` (저장소 기본 폴더) | 1단계: 텀싯 프로필, 표 병합 (A1–A3) |
| B | Claude 서브에이전트 | `feat/term-variables` (별도 worktree) | 2단계: 조건 변수, 콘텐츠 컨트롤, 검사 (B1–B3) |
| 통합 | Claude | `feat/term-sheet` | 병합, 샘플 변수화, 전체 검증, 교차 리뷰 반영 |

- A와 B는 모두 `feat/term-sheet`에서 갈라져 동시에 진행한다. 두 흐름이 함께 쓰는 필드·옵션·모듈은 공통 뼈대 커밋에 이미 들어 있다. 따라서 각 흐름은 서로 다른 코드 영역만 고친다.
- **Codex 추론 강도 상한은 `max`다(`ultra` 금지).** 표 병합과 렌더링처럼 어려운 작업은 `max`, 샘플과 문서는 `high`를 쓴다.
- **교차 리뷰:** A의 결과는 Claude가, B의 결과는 Codex가 적대적 관점으로 리뷰한다. 통합 뒤에는 Codex가 `main` 대비 전체 차이를 한 번 더 리뷰한다.
- push와 PR은 소유자가 확인한 뒤에만 한다.

### 공통 뼈대(이미 커밋됨)

- **모델:**
  - `ElementType.CONFIRMATION`
  - `TextRun.term_key`
  - `TableCell.merge`(`"up"`/`"left"`)
  - `Table.spans`(None이면 프로필 기본값), `Table.label_columns`
  - `Heading.runs`, `Blockquote.runs`
  - `DocumentMetadata.display_runs`(조건 치환 때만 채움)
- **옵션:**
  - `RenderOptions.house`/`term_tags`, `ResolvedOptions.house`/`term_tags`, `layout.term_tags`
  - 등록기(registry) 전달, CLI `--house`/`--no-term-tags`
- **frontmatter:** `house`, `terms`, `confirmation`은 문자열로 바뀌지 않고 원래 구조를 유지한다.
- **모듈과 설정:**
  - 빈 `term_sheet.py`, `term_variables.py`(`TERM_TAG_PREFIX`, `TERM_PROPERTY_PREFIX`)를 패키징 목록에 등록했다.
  - `.gitignore`에 `tmp/`, `private/`를 추가했다.

### 공통 규칙(두 흐름 모두)

- `tmp/`, `output/`, `private/`는 열지 않는다. 실제 거래 문서가 있다. 샘플, 테스트, 문서에는 가상 이름만 쓴다(§3).
- 테스트는 가상환경 파이썬으로 실행하고 basetemp는 매번 새 경로로 준다.
  `.venv/Scripts/python.exe -B -m pytest tests/ -q -p no:cacheprovider --basetemp=<새 임시 경로>`
- 커밋 전에 `.venv/Scripts/python.exe -m ruff check .`을 통과해야 한다. `mypy .`는 기존 `output/` 오류 1건 외에 새 오류가 없어야 한다.
- AGENTS.md를 따른다.
  - 단일 조립 경로를 유지한다.
  - 스타일은 불변이고 요청 범위로 둔다.
  - 빈 칸, 이스케이프, 숫자 표기를 보존한다.
  - 출력 변경은 실제 파싱·렌더링 회귀 테스트로 고정한다.
- 테스트를 먼저 쓰고 실패하는 것을 확인한다. 합성 Markdown을 실제 파서와 `IBDocumentRenderer.render`에 넣고, `BytesIO`로 저장·재오픈한 뒤 XML을 확인한다.
- 커밋 메시지는 영어로 쓴다. Claude 커밋은 끝에 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`를, Codex 커밋은 `Co-Authored-By: Codex (gpt-6-astra) <noreply@openai.com>`를 붙인다.

### 확인된 python-docx 1.2.0 동작 (Claude와 Codex가 각각 확인)

- **병합:**
  - `table.cell(r0, c0).merge(table.cell(r1, c1))`는 `w:vMerge`와 `w:gridSpan`을 만든다.
  - 병합 뒤에는 가려진 위치의 `table.cell(r, c)`와 `row.cells[c]`가 모두 기준 칸의 `_tc`를 돌려준다. 좌표별로 도는 기존 스타일 루프가 같은 칸을 여러 번 처리하게 된다.
  - **내용이 있는 칸을 병합하면 문단이 이어 붙는다.** 빈 표에서 병합한 뒤 기준 칸만 한 번 채운다.
- **콘텐츠 컨트롤:**
  - 문단 안 run 수준 `w:sdt`의 글자는 `Paragraph.text`와 `Paragraph.runs`에 보이지 않는다. XML(`.//w:t`)로 읽어야 한다.
  - 컨트롤을 넣은 뒤에 `Paragraph.text`에 값을 대입하면 서식이 사라진다.
- **줄바꿈:** `paragraph.add_run("A\nB")`는 같은 문단 안에 `w:br`을 만든다.
- **문서 속성:** `core_properties.title`은 255자를 넘으면 ValueError를 낸다. 사용자 지정 속성 공개 API는 없으므로 저장소의 `GeneratorSignatureWriter` XML 장치를 재사용한다.
- **파서 이스케이프:** `TextParser`는 `\^^`를 `^^`로 바꾸지만 `\<<`는 그대로 둔다. 병합 기호 이스케이프는 원문(`cell.content`)에서 따로 처리해야 한다.

## 2. 설계 보완 결정

설계서와 다르거나 설계서가 정하지 않은 부분이다. 구현은 이 결정을 따른다.

1. **표 병합은 공통 기능이다.** 병합 출력은 모든 프로필의 `TableRenderer`에서 처리한다. 텀싯은 기본으로 켜지고, 다른 프로필은 `spans: true`일 때만 켜진다. 표 사양의 `spans: false`는 텀싯 기본값도 끈다. 병합이 있는 표는 좌표 루프가 아니라 기준 칸 목록을 한 번씩 처리한다.
2. **잘못된 병합은 묶음 단위로 되돌린다.** 직사각형이 아니거나 경계를 넘는 병합 묶음 하나만 글자로 되돌리고 경고한다. 옆의 올바른 병합은 유지한다.
3. **내어쓰기 좌표.** 기호별 `(시작, 내어쓰기)`를 Word 값으로 옮기면 `left_indent = 시작 + 내어쓰기`, `first_line_indent = −내어쓰기`다.

   | 기호 | 시작 | 내어쓰기 | 비고 |
   |---|---|---|---|
   | `•` | 0 | 3mm | |
   | `-` | 3mm | 3mm | |
   | `·` | 6mm | 3mm | |
   | `①`–`⑳` | 0 | 4.5mm | |
   | `※` | 0 | 4mm | 8pt |
4. **1열 표와 라벨 열.** 기본 라벨 열 수는 `min(머리행 첫 칸 폭, 열 수 − 1)`이다. 명시한 `label_columns`는 0 이상, 열 수 미만이어야 한다.
5. **여러 줄 조건 값은 거부한다.** 줄바꿈이 든 값은 인라인 컨트롤 하나로 추적할 수 없고, 줄별로 태그하면 거짓 불일치가 생기기 때문이다.
6. **가로 표.** 라벨 열 폭은 고정하고, 나머지는 실제 구역의 본문 폭에서 준다. 180mm는 세로 기본값일 뿐이다.
7. **문서 속성 제목과 주제가 255자를 넘으면 명확한 오류로 거부한다.** 조용히 자르지 않는다.
8. **본문 문단 줄 나눔.** 텀싯 프로필에서는 셀과 마찬가지로 본문 문단의 `<br>` 줄도 각각 별도 Word 문단으로 만든다. 그래야 줄마다 내어쓰기가 적용된다.
9. **고객확인란**
   - `ConfirmationBlock(source)` 페이로드로 원래 펜스 원문을 보관한다.
   - 닫힌 빈 펜스만 고객확인란이 된다.
   - 문구가 없거나 펜스에 내용이 있거나 펜스가 닫히지 않았으면, 원문을 코드 상자로 보존하고 진단을 남긴다. strict는 저장 전에 거부한다.
   - 고객확인란은 **1행 1열 표 한 칸** 안에 모든 문단을 넣는다. 그래야 행 분할 금지 하나로 한 쪽에 모인다.
10. **기밀 표시 끄기.** 텀싯 프로필의 기본 `confidential`은 참이다. `--no-confidential`, `layout.confidential: false`, 빈 `confidential_label` 중 하나라도 있으면 머리글을 비운다. 첫머리 면책 문구는 기존 끝 면책 스위치(`--no-disclaimer`)와 무관하게 항상 넣는다.
11. **모듈 의존 방향.** `term_sheet.py`는 모델, 스타일, YAML, python-docx만 import한다. `ib_renderer`는 `term_sheet`를 import할 수 있다. run 렌더링 함수는 렌더러가 인자로 넘긴다.
12. **조건 값은 파싱할 때 run으로 확정한다.** 파서가 run을 만드는 문단, 목록, 표 셀에 더해 장 제목(`Heading.runs`), 인용 상자(`Blockquote.runs`), 제목·부제·날짜(`metadata.display_runs`)도 파싱할 때 채운다. 렌더러는 run이 있으면 그것을 쓰고 다시 해석하지 않는다. 문자열 필드(`Heading.text`, `metadata.title`)는 치환된 글자로 바꿔 두어, 제목 중복 제거·목차·문서 속성이 그대로 동작하게 한다. 표 캡션·단위·출처는 태그 없이 글자로만 치환한다.
13. **`terms:` 유무.** `terms:` 키가 없으면 기존 동작과 같다. 빈 매핑 `{}`는 조건 기능을 켠 것으로 보고, 모든 참조를 "정의되지 않은 키"로 검사한다.
14. **태그는 마지막에 단다.** `TextRenderer.render_runs`는 조건 run의 `w:r`을 렌더링 범위 수집기(ContextVar)에 등록만 한다. 모든 요소 렌더링과 기존 후처리(음수 빨강, 위험 색, 기준 칸 강조, 각주 등)가 끝난 뒤 렌더러가 한꺼번에 `w:sdt`로 감싸고 스냅숏을 쓴다. 그래서 `paragraph.runs`에 의존하는 기존 후처리를 바꾸지 않아도 된다. 링크 안의 조건 run은 등록하지 않는다.
15. **숫자 서식 제외.** 조건 값 run은 표 숫자 서식(천 단위 구분 등)에서 제외한다.
16. **검사 결과 JSON.** 조건 태그나 스냅숏이 없는 문서의 `docx-audit` JSON은 기존과 정확히 같아야 한다. 명시적 직렬화 함수로 비어 있는 `terms`를 생략한다.
17. **python-docx 하한.** `pyproject.toml`은 `python-docx>=0.8.11`을 허용하지만 이번 검증은 1.2.0에서만 했다. 하한 상향은 이번 범위 밖이며 PR 설명에 권고로 적는다.

## 3. 가상 샘플 규칙

- **이름:**

  | 역할 | 가상 이름 |
  |---|---|
  | 회사 | 가나다머티리얼즈㈜ |
  | SPC | 가나다제일차(유) |
  | 작성 기관 | 라마바은행 자본시장부 |
  | 신탁 | 라마바은행 신탁부 |
  | 매출처 | 마바사전자㈜, 아자차케미칼㈜ |
- 계좌번호는 `000-0000-0000-000`만 쓴다. 금액, 금리, 일정은 새로 지어낸 값이다.
- house 샘플 문구는 새로 쓴다. 실제 기관 문구를 옮기지 않는다.

## 4. 흐름 A: 텀싯 프로필과 표 병합 (Codex)

### A1. 프로필, house, 표 사양, 병합 해석 (추론 강도 `max`)

- **프로필:** `PROFILES`에 `"term-sheet": DocumentProfile("term-sheet", confidential=True)`를 추가한다(비IB, A4, 표지·목차·끝 면책 없음).
  - 제목이 없거나 비어 있으면 일반 제목 추론보다 먼저 거부한다.
  - 부제가 비면 `Term Sheet`를 쓴다.
  - 제목과 부제가 255자를 넘으면 거부한다.
- **스타일:**
  - `render_styles.IBStyle`에 다음 필드를 추가한다: `TS_LABEL_BG_HEX="F2F5FC"`, `TS_BORDER_HEX="9AA5C4"`, `TS_MUTED_HEX="555555"`, `TS_CONFIDENTIAL_HEX="888888"`, `TS_LABEL_WIDTH=Inches(33.5/25.4)`, `TS_SUBLABEL_WIDTH=Inches(30/25.4)`, `TS_TITLE_SIZE=Pt(20)`, `TS_SUBTITLE_SIZE=Pt(16)`, `TS_META_SIZE=Pt(10)`, `TS_NOTE_SIZE=Pt(8)`, `TS_DISCLAIMER_SIZE=Pt(7)`, `TS_HEADER_FOOTER_SIZE=Pt(7.5)`.
  - 행 분할 기준(12줄)은 테마 정수 제한(0–9)에 걸리므로 `term_sheet.py`의 모듈 상수로 둔다.
  - `load_style`에서 `profile.name == "term-sheet"`일 때 다음 값을 적용한다.

    | 항목 | 값 |
    |---|---|
    | 강조색 | `NAVY`/`NAVY_HEX` `1A2270` |
    | 제목·본문 글꼴 | 맑은 고딕 |
    | 본문·표 머리·표 본문 | 9pt |
    | `SMALL_SIZE` | 8pt |
    | H1 | 14pt |
    | H2 | 13pt, 앞 15pt, 뒤 7pt |
    | H3 | 10pt, 앞 12pt, 뒤 5pt |
    | 줄 간격 | 1.05 |
    | 문단 뒤 | 3pt |
    | 여백 | 위 18mm, 아래 16mm, 좌우 15mm |
    | `TABLE_HEADER_BG` | `DCE3F5` |
    | 끄는 항목 | 줄무늬, 양쪽 정렬, 기존 제목 테두리 |

  - 표 머리 글자색은 렌더링할 때 `STYLE.NAVY`를 쓴다. 그래야 테마 `primary_color`가 머리 글자색에도 적용된다.
- **house:**
  - `term_sheet.py`에 다음을 둔다.
    - `ConfirmationText`(intro, items는 tuple, signature)
    - `TermSheetTexts`(prepared_by, disclaimer, confidential_label, confirmation), 모두 frozen
    - `load_house(path)`
    - `resolve_term_sheet_texts(metadata, house_path)`
  - 병합 규칙: frontmatter 키가 **있으면**(빈 문자열 포함) house보다 우선한다.
  - 필수와 기본값: `prepared_by`와 `disclaimer`는 공백이 아닌 문자열이어야 한다. `confidential_label`의 기본값은 `Strictly Confidential`이다.
  - `confirmation` 형식: `intro`(str), `items`(비어 있지 않은 str 목록), `signature`(str)만 받는다. 셋 중 하나는 있어야 한다.
  - 거부 대상: 알 수 없는 키, 잘못된 형식, 없는 파일, YAML 오류.
- **house 경로:**
  - `parse_markdown_file`은 frontmatter의 상대 `house` 경로를 Markdown 파일 폴더 기준 절대 경로로 바꾼다. 기존 테마 경로 처리를 참고한다.
  - CLI `--house`는 현재 폴더 기준이며 frontmatter보다 우선한다. 우선하는 경로가 있으면 frontmatter 쪽 파일은 열지 않는다.
  - 원본 경로가 없는 입력(문자열·스트림)의 상대 `house` 경로는 오류로 거부한다.
  - house 해석은 렌더러 시작 단계(요소 오류 복구 밖)에서 한다.
- **표 사양:** `apply_table_specs`의 허용 키에 두 가지를 더한다. `spans`는 bool만 받는다. `label_columns`는 int만 받고(bool 거부) 0 이상, 열 수 미만이어야 한다. 사양이 없는 표도 프로필 기본값을 받고, 순서형 `{}` 자리표시자는 유지한다.
- **병합 해석:**
  - `md_parser`에 `TableSpanResolver`를 두고, `MarkdownParser.parse`에서 `apply_table_specs` 뒤, 표 경고를 모으기 전에 호출한다.
  - 기호는 **원문** `cell.content.strip()`으로 판정한다.
  - 이스케이프 `\^^`와 `\<<`는 글자 그대로 `^^`, `<<`로 `content`와 `runs`를 고친다(병합 기능이 켜진 표에서만).
  - 위·왼쪽 참조를 따라 묶음을 만들고, 기준 칸 하나와 완전한 직사각형인지, 머리행 또는 본문 안에만 있는지 확인한다(§2-1, §2-2).
  - 병합 해석 뒤에 라벨 열 수를 정한다(§2-4).
- **고객확인 펜스:** `document_model`에 `ConfirmationBlock(source: str)`와 `ElementContent` 항목을 추가한다. 텀싯 프로필의 `confirmation` 펜스를 §2-9대로 처리한다. 다른 프로필은 기존대로 코드 블록이다.
- **기존 매개변수 테스트:** `PROFILES` 전체를 도는 기존 테스트(`test_document_profiles.py`, `test_chart_port.py`, `test_preset_port.py`)는 텀싯일 때만 `prepared_by`·`disclaimer`를 frontmatter에 넣는 식으로 고친다. 테스트 의도는 바꾸지 않는다.
- **테스트** (`tests/test_term_sheet_profile.py`, `tests/test_table_spans.py`):
  - 프로필 계약
  - house 스키마와 우선순위, house 진입점(CLI, 파일, 문자열, 등록기)
  - `spans`·`label_columns` 사양
  - 병합: 가로, 세로, 2차원, L자, 경계 넘김, 이스케이프, 빈 칸, 꺼짐, 올바른 묶음 옆의 잘못된 묶음
  - 고객확인 펜스

### A2. 렌더링 (추론 강도 `max`)

- **공통 병합 출력(`TableRenderer`):**
  - 병합이 있는 표는 빈 표를 만들고, 열 폭(`tblGrid`, 칸마다 `tcW`)을 먼저 적용한 뒤, 검증된 직사각형을 병합하고, 기준 칸만 한 번 채우고 칠한다.
  - 병합이 없는 기존 표의 출력은 바이트 단위로 같아야 한다.
- **텀싯 표 모양(텀싯 프로필에서만):**
  - **폭:** 라벨 열이 1이나 2이고 열 수가 라벨 열 수 + 1인 표는 고정 폭이다. 세로 기본 격자는 `[33.5, 146.5]`mm 또는 `[33.5, 30, 116.5]`mm이고, 나머지는 실제 구역 폭에서 준다. 그 밖의 표는 기존 폭 추정을 쓴다. 고정 레이아웃에는 `w:tblLayout fixed`를 쓴다.
  - **음영:** 머리행은 `TABLE_HEADER_BG`에 굵게, 강조색, 가운데 정렬이며 `w:tblHeader`를 넣는다. 라벨 칸은 `TS_LABEL_BG_HEX`에 굵게, 검정, 가운데 정렬이다. 내용 칸은 기존 열 종류, 정렬 표시, 숫자 역할을 따른다.
  - **테두리와 여백:** 6방향 모두 `single sz=4 color=TS_BORDER_HEX`다. 표 수준 `w:tblCellMar`는 위아래 70, 좌우 100 twips다. 모든 칸은 세로 가운데 정렬이고, 칸 문단은 뒤 1.5pt, 줄 간격 1.05다.
  - **줄별 문단:** run 목록을 `\n` 기준으로 줄별로 나누는 도우미를 `term_sheet.py`에 둔다. 서식, 링크, 각주, 빈 줄을 보존한다. 줄마다 별도 문단을 만들고, 내용 칸의 줄 앞 기호에 §2-3 내어쓰기를 준다. 글자는 바꾸지 않는다.
  - **캡션과 단위:** 캡션은 표 위 왼쪽, 굵게 9pt 강조색이다. 단위와 기준일은 표 위 오른쪽, 8pt `TS_MUTED_HEX`로 `(단위 : 억원, 기준일 : …)`이다. 둘 다 다음 요소와 같은 쪽에 둔다. 출처는 표 아래 8pt다.
  - **행 분할:** 텀싯 표에만 적용한다. 행마다 예상 줄 수를 계산해 12줄 이하이면 `w:cantSplit`을 넣는다.
    - 예상 줄 수는 칸별 합의 최댓값이다. 칸별 합은 줄마다 `올림(글자 폭 합 ÷ 칸 안쪽 폭)`을 더한 값이다. 글자 폭은 한글 1.0em, 그 밖 0.55em, 칸 안쪽 폭은 병합 폭에서 여백과 들여쓰기를 뺀 값이다.
    - 여러 행을 덮는 칸은 제외하고, 머리행은 항상 분할 금지다.
  - **표 뒤 간격:** 높이 4pt(정확히)의 빈 문단을 둔다.
- **첫머리와 문단:**
  - `render_term_sheet_opening(doc, metadata, texts, render_runs)`는 제목, 부제, 날짜, 작성 기관, 면책 순서로 그린다.
  - 제목은 비개요 `Title` 스타일을 쓰고, 문서 속성 제목과 주제를 설정한다.
  - `metadata.display_runs`에 run이 있으면 그것을, 없으면 `TextParser.parse_runs(text)`를 `render_runs`로 그린다.
  - 면책 문단은 위아래 0.5pt `TS_BORDER_HEX` 선, 양쪽 정렬, 7pt `TS_MUTED_HEX`다.
  - `True`를 돌려 같은 H1을 건너뛰게 한다.
  - 텀싯 문단은 줄마다 별도 문단(§2-8)이며 기호 내어쓰기를 준다.
  - 텀싯 제목(`##`, `###`)은 강조색이고, `Heading.runs`가 있으면 그것을 쓴다.
- **스타일(`setup_term_sheet_styles`):**
  - `docDefaults`의 `pPr`에 `w:kinsoku w:val="1"`과 `w:wordWrap w:val="0"`을, `rPr`에 `w:lang w:eastAsia="ko-KR"`을 넣는다. 텀싯 문서에만 적용한다.
  - Heading 2는 강조색과 아래 1pt 강조색 선(`w:pBdr/w:bottom sz=8`), Heading 3은 강조색이다.
- **고객확인란(`render_confirmation`):** 1행 1열 표이고 테두리는 0.5pt `TS_BORDER_HEX`다. intro(8pt), items(9pt, `□` 내어쓰기), signature(9pt, 오른쪽 정렬) 순서다. 행 분할 금지이고, 직전 문단은 다음 요소와 같은 쪽에 둔다.
- **머리글·바닥글:** 모든 구역이 생긴 뒤 구역마다 한 번 설정한다. 가로·세로 구역이 섞여도 각 구역의 본문 폭으로 탭을 계산한다.
  - 머리글: 오른쪽 정렬, 기울임 7.5pt `TS_CONFIDENTIAL_HEX` 기밀 표시(§2-10)
  - 바닥글: 왼쪽 `"{부제} {version}"`(version 없으면 빈칸), 오른쪽 탭에 `PAGE / NUMPAGES` 필드, 7.5pt 회색
- **연결:** `IBDocumentRenderer.render`는 텀싯이면 다음을 한다.
  - 문구를 먼저 확정한다.
  - 스타일 생성 뒤 `setup_term_sheet_styles`를 부른다.
  - 첫머리는 `render_office_opening` 대신 `render_term_sheet_opening`을 부른다.
  - 머리글·바닥글은 텀싯 전용 함수를 부른다.
  - `_render_element`는 텀싯일 때 문단, 제목, `CONFIRMATION`을 텀싯 경로로 보낸다.
- **테스트** (`tests/test_term_sheet_tables.py`, `tests/test_term_sheet_profile.py`):
  - 병합 XML(`gridSpan`, `vMerge`, 합산 폭, 기준 칸 글자 한 번, 기호 없음, 중복 `w:shd` 없음)
  - 첫 세로선 공통 위치, 색, 여백, 머리행 반복, 단위 형식
  - 줄별 문단과 `w:ind`, `※` 8pt, 짧은 행·긴 행 `cantSplit`
  - 첫머리 순서, 비개요, 제목 한 번, 문서 속성
  - 가로·세로 섞인 구역의 머리글·바닥글
  - 고객확인란, 텀싯 문서에만 `wordWrap`
  - 병합 없는 기존 표의 출력 불변
  - strict 렌더와 `docx_audit` issues 0

### A3. 샘플과 문서 (추론 강도 `high`)

- **샘플:** `samples/profiles/term-sheet.md`와 `samples/profiles/term-sheet-house.yaml`을 만든다. §3의 가상 이름을 쓰고, 1단계에서는 **값을 글자 그대로** 적는다. 구성은 다음과 같다.
  - 1. 본건 개요(2열)
  - 2. 구조도(①~⑥ 설명)
  - 3. 주요 금융조건(3열, 병합)과 여신조건 격자 표
  - 4. 신탁·대출약정 주요조건(긴 행 포함)
  - 5. 신용보강 및 담보
  - 별첨 상환스케줄(열 역할·단위)
  - `confirmation` 펜스
  - 참고 표
- **문서:**
  - README.md와 README.ko.md: 프로필 표에 `term-sheet`를 추가하고 작성 규칙 요약을 넣는다. `termsheet` 프리셋과 다르다는 점도 적는다.
  - CHANGELOG: 항목을 추가한다.
  - AGENTS.md: 프로필 목록(7종)과 모듈 표를 갱신한다.
- **검증:**
  - 샘플 strict 생성과 `docx-audit` issues 0
  - `md-to-word samples/profiles --batch`
  - 기존 샘플 테스트 통과

## 5. 흐름 B: 조건 변수, 콘텐츠 컨트롤, 검사 (Claude 서브에이전트)

핵심 테스트는 `plain`·`business-report` 프로필로 작성해 1단계와 독립적으로 둔다.

### B1. 조건 변수 해석

- **`term_variables.py`:**
  - `validate_terms(value) -> Dict[str, str]`: 매핑만 받는다. 키는 `^[a-z][a-z0-9_]{0,39}$`이고, 값은 문자열만 받는다. None, 숫자, 목록, 줄바꿈은 거부하며 숫자에는 따옴표 안내를 붙인다.
  - 참조는 `{{키}}`이고 안쪽 공백을 허용한다. 앞이 `\`면 참조가 아니다.
- **토큰화(`TextParser`):**
  - 코드 구간과 이스케이프를 보호한 뒤, 인라인 해석 전에 유효하고 정의된 참조를 사설 영역 토큰으로 바꾼다. 접두사 충돌은 기존 `CODE` 방식을 따른다.
  - 인라인 해석 뒤 토큰이 든 run을 `[앞, 값 run(term_key, 굵게·기울임·색·링크 상속), 뒤]`로 나눈다. 값은 다시 해석하지 않는다.
  - 값 속 `**`, `$`, `<br>`, `|`, `\`, 연속 공백, `( x )`는 그대로 남아야 한다.
  - 활성 조건이 없으면 기존 동작과 같다.
- **파서 연결(`MarkdownParser.parse`):**
  - frontmatter `terms`를 검증한다(오류는 ValueError).
  - 문단, 목록, 표 셀은 run에, 장 제목·인용 상자·제목·부제·날짜는 §2-12의 run 필드에 치환 결과를 채운다.
  - 표 셀 원문(`content`)은 병합 기호 판정용으로 그대로 둔다. 따라서 값이 `^^`여도 병합되지 않는다.
  - 파싱 뒤 정의되지 않은 키와 잘못된 키 이름을 하나씩 모델 경고로 넣는다. 한 번도 쓰이지 않은 키는 `logger.info`로만 알린다.
- **테스트(`tests/test_term_variables.py`):**
  - 스키마
  - 토크나이저(서식 상속, 값 속 특수문자, `\{{`, 링크 주소 보존)
  - 적용 위치: 문단·목록·셀, 장 제목·인용 상자·제목
  - 코드 제외, 정의되지 않은 키 경고와 strict 거부
  - `terms:` 없는 문서의 `{{x}}` 보존, `{}`의 검사 활성화

### B2. 태그와 스냅숏

- **렌더러의 run 사용:** `HeadingRenderer`와 `CalloutRenderer`는 run이 있으면 그것을 쓴다. 텀싯이 아닌 프로필의 표지·공문 첫머리는 치환된 문자열을 쓴다(태그 없음).
- **등록과 감싸기:**
  - `TextRenderer.render_runs`는 조건 run(링크 제외, 태그 켜짐)의 `w:r`과 키, 값을 렌더링 범위 수집기에 등록한다.
  - 렌더러는 모든 요소 렌더링과 후처리 뒤, 구조 검사 전에 등록된 run을 `w:sdt`로 감싼다(§2-14).
  - `w:sdtPr` 순서는 `w:alias`, `w:tag(ibrep:term:키)`, `w:id`(문서 안 고유 양의 정수), `w:text`다. 잠금과 바인딩은 두지 않는다.
- **스냅숏:** 태그를 단 키만 사용자 지정 속성 `ibrep.term.<키>`로 기록한다. `GeneratorSignatureWriter`의 속성 upsert를 일반화하되 기존 속성과 PID를 보존하고, XML을 이스케이프한다.
- **숫자 서식:** 조건 run은 숫자 서식에서 제외한다(§2-15).
- **태그 끄기:** `term_tags`가 거짓이면 컨트롤과 스냅숏 없이 글자만 넣는다. CLI, YAML, API 우선순위를 테스트한다.
- **테스트(`tests/test_term_controls.py`):**
  - 태그·alias·`w:text`
  - 컨트롤 안 run 서식, 음수·위험·기준 칸 스타일 유지
  - 링크 안 조건은 치환되지만 태그 없음
  - 스냅숏은 태그 단 키만, 생성기 속성 보존, 저장·재오픈 값 일치
  - 태그 끄기
  - 렌더러 재사용 시 수집기 초기화

### B3. `docx-audit` 조건 검사

- **현재 글자 추출:** 본문 파트의 `w:sdt` 중 태그가 `ibrep:term:`으로 시작하는 것을 문서 순서로 모은다. `w:t`, `w:tab`, `w:br`을 포함하고, 삽입(`w:ins`)·이동 도착 글자는 포함하고, 삭제(`w:del`)·이동 출발 조상 아래 글자는 제외한다.
- **스냅숏:** 사용자 지정 속성 파트에서 `ibrep.term.` 속성을 읽는다.
- **활성 조건:** 태그 **또는** 스냅숏이 있으면 `terms`를 채운다.
  - `mismatched`: 키별 서로 다른 값 목록
  - `changed`: `{generated, current}`
  - `missing`: 스냅숏에는 있는데 컨트롤이 없는 키
  - `indicative`: 현재 값에 `[`…`]`가 있는 키
- **경고와 출력:** `mismatched`와 `missing`은 사람이 읽을 경고로 `warnings`에 넣는다. `issues`와 종료 코드는 바꾸지 않는다. 명시적 직렬화 함수로 조건이 없으면 JSON에서 `terms`를 생략한다(§2-16). `tests/test_render_hardening.py`의 정확 비교 테스트가 계속 통과해야 한다.
- **테스트(`tests/test_term_audit.py`):**
  - 생성한 문서의 XML을 고쳐 Word 편집을 흉내 낸다: 한 곳만 수정, 모두 같게 수정, 마지막 컨트롤 삭제
  - 변경 추적 삽입·삭제, 탭·줄바꿈, 빈 값, 대괄호 값
  - 조건 없는 문서의 JSON 불변, CLI 종료 코드 0

## 6. 통합 (Claude)

1. 리뷰를 반영한 뒤 `feat/term-sheet`에 A, B 순서로 병합하고 충돌을 해소한다.
2. 텀싯 샘플에 `terms:`를 넣고 본문을 `{{키}}`로 바꾼다. 제목·부제·날짜 태그가 첫머리에서 붙는지 확인한다.
3. 전체 검증을 돌린다.
   - pytest 전체, ruff, mypy
   - `uv build`와 두 배포물 내용 확인, 설치한 wheel로 저장소 밖에서 변환 확인
   - 샘플 strict 생성과 `docx-audit`
   - Word 편집 흉내 후 조건 검사
4. **Word 페이지 검사:**
   - `scripts/word_visual_qa.ps1`로 새 빈 폴더에 페이지 이미지를 만든다(처음은 쪽수 측정용).
   - 쪽수를 확정해 다른 빈 폴더에서 다시 만들고 모든 쪽을 본다.
   - 로컬 참고 양식과 구성·정렬을 비교한다. 참고 양식과 이미지는 커밋하지 않는다.
5. Codex(`max`)가 `main` 대비 전체 차이를 적대적 관점으로 리뷰하고, 확인된 지적을 고친다.
6. 검증 기록 `docs/verification-term-sheet-20260929.md`를 쓰고 설계서에 §2 결정을 반영한다. 커밋 전 실제 거래 관련 이름을 로컬에서 검색한다.
7. 소유자 확인 후 push와 PR을 한다.
