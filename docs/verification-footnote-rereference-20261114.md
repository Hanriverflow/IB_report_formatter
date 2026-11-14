# 검증 기록: 같은 각주를 여러 번 참조한 문서 (2026-11-14)

## 1. 결함

같은 각주를 본문 두 곳 이상에서 참조하면(`본문[^1] … 다시[^1]`) 변환은 strict에서도 종료 코드 0으로 끝나고 스키마 검증도 통과하지만, Word가 문서를 열지 못했다. `docs/improvement-roadmap-20260929.md`의 "뒤로 미루는 것"에 적혀 있던 항목이다.

- 원인: `ib_render_footnotes.py`가 참조마다 같은 ID의 `<w:footnoteReference w:id="1"/>`를 다시 썼다. Word는 한 각주 ID가 본문에 두 번 나오는 문서를 손상된 파일로 본다.
- 재현(시작점 `c97841c`, Word 16.0 build 16.0.20527, Word COM `Documents.Open`):

| 문서 | strict 종료 코드 | `footnoteReference` 개수 | Word |
|---|---:|---:|---|
| 최소 문서(문단 안 두 번) | 0 | 2 | "파일이 손상되었습니다." |
| 제목·문단·표 칸·목록·콜아웃에 한 번씩, `business-report` | 0 | 5 | "파일이 손상되었습니다." |
| 같은 문서, `ib-report` | 0 | 5 | "파일이 손상되었습니다." |
| 각주 두 개를 한 번씩 참조(대조군) | 0 | 2 | 열림 |

## 2. 기준 자료: Word 자신의 방식

Word 16.0.20527 COM으로 새 문서에 각주 하나를 넣고 `Selection.InsertCrossReference(wdRefTypeFootnote=3, wdFootnoteNumberFormatted=16, 1, True)`와 `wdFootnoteNumber=5`로 같은 각주를 두 번 더 가리킨 뒤 Word가 저장한 XML(문서는 저장소 밖에 두었다):

```xml
<w:bookmarkStart w:id="0" w:name="_Ref245569440"/>
<w:r><w:rPr><w:rStyle w:val="ab"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
<w:bookmarkEnd w:id="0"/>
…
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> NOTEREF _Ref245569440 \f \h </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:rStyle w:val="ab"/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

`ab`는 한국어 Word에서 "각주 참조" 스타일의 ID다. Word는 각주 표시를 하나만 두고, 첫 참조 둘레의 숨은 `_Ref` 북마크를 `NOTEREF` 필드로 가리킨다. `\f`는 결과에 각주 참조 서식을 입히고 `\h`는 하이퍼링크다.

## 3. 고친 방식

- `FootnoteRenderer._link_repeated_references`: 모든 참조가 네이티브 각주가 된 뒤(명시 `[^n]`, IB 프로필의 `^n^` 모두) 본문을 문서 순서로 한 번 읽는다. 다시 참조되는 각주의 첫 참조 런을 숨은 북마크 `_Ref_fn_<표시 번호>`로 감싸고, 두 번째부터의 참조 런을 `NOTEREF _Ref_fn_<n> \f \h` 필드(시작·명령·구분·결과·끝 다섯 런)로 바꾼다.
- 결과값은 Word가 보여 줄 번호다. Word는 각주를 참조 순서대로 번호 매기므로, 문서에서 처음 참조된 순서를 센다(`[^2]`가 먼저 나오면 그 각주가 1번).
- 다섯 런 모두 원래 참조 런의 `w:rPr`(각주 참조 스타일, 위 첨자, 글꼴, 크기)를 복사해 갖는다. 필드를 갱신하기 전과 후의 모양이 같다.
- 북마크 ID는 렌더러의 할당기(`IBDocumentRenderer._next_bookmark_id`)에서 받는다. 북마크 이름은 표시 번호를 쓰므로 Word의 40자 제한 안에 든다. E5의 `_Ref_tbl_…`·`_Ref_fig_…`와 겹치지 않는다.
- 한 번만 참조되는 각주는 건드리지 않는다. 그런 문서의 XML은 바이트 단위로 이전과 같다.
- 손으로 만든 속성 자식은 없다(런 속성은 이미 순서대로 만들어진 것을 복사한다). 북마크와 필드 런은 문단 내용이며 `ooxml_order` 대상이 아니다.
- `docx-audit`: `NOTEREF`도 `REF`·`PAGEREF`처럼 북마크가 없으면 issue(`Cross-reference field NOTEREF … has no bookmark named …`)로 보고한다. 결과값이 북마크 안 각주의 표시 번호와 다르면 warning이다. 이 비교는 본문과 설정에 `w:footnotePr`가 없는 기본 번호 매기기일 때만 하고, `customMarkFollows` 참조는 번호를 세지 않으며, `\p` 스위치가 있으면 하지 않는다.

## 4. Word 확인 (Word 16.0 build 16.0.20527)

고친 엔진의 산출물을 Word COM으로 열어 다음 순서로 확인했다. 스크립트는 저장소 밖(`%LOCALAPPDATA%\Temp\ibrf-fn\tools\word-check.ps1`)에 있고, 자기가 띄운 Word만 닫는다.

1. 열기 → 2. 모든 필드 갱신(F9) → 3. 인쇄 미리 보기 보기로 바꿨다가 되돌림 → 4. 첫 참조 한 글자 앞에 새 각주를 넣고 F9 → 5. 다른 이름으로 저장, 닫고 다시 열기.

| 문서 | 열기 | 저장된 결과 | F9 뒤 | 미리 보기 뒤 | 앞에 각주 삽입 + F9 | 저장 후 다시 열기 |
|---|---|---|---|---|---|---|
| 최소 문서 | 열림 | 1 | 1 | 1 | 2 | 2 |
| 다섯 위치, `business-report` | 열림 | 1 ×4 | 1 ×4 | 1 ×4 | 2 ×4 | 2 ×4 |
| 다섯 위치, `ib-report` | 열림 | 1 ×4 | 1 ×4 | 1 ×4 | 2 ×4 | 2 ×4 |
| 순서 섞임(`[^2]` 먼저) | 열림 | 2, 1, 1 | 2, 1, 1 | 2, 1, 1 | 3, 2, 2 | 3, 2, 2 |
| `ib-report` `^1^` 두 번 + References | 열림 | 1 | 1 | 1 | 2 | 2 |
| QA 샘플 `samples/qa/footnote-rereference.md` | 열림 | 1,2,1,3,1,2,4,1 | 같음 | 같음 | 2,3,2,4,2,3,5,2 | 같음 |

모든 `NOTEREF` 결과의 `Font.Superscript`는 -1(참)이었다. 오류 문자열("Error!", "오류!")은 어느 단계에서도 나오지 않았다. 시험 스크립트의 첫 판은 새 각주를 첫 참조와 같은 위치, 곧 `_Ref_fn_1` 북마크 안에 넣어 북마크가 새 각주를 가리키게 됐다(결과 1 유지). 이것은 시험 방법의 결함이라 삽입 위치를 첫 참조 한 글자 앞으로 옮겨 다시 실행했고, 위 표는 그 결과다.

### 쪽 이미지

`scripts/word_visual_qa.ps1`로 새 빈 폴더 `%LOCALAPPDATA%\Temp\ibrf-fn\visual-qa1`에 만들었다. 9쪽을 모두 보았다.

| 문서 | SHA-256 앞 16자 | 쪽 수 | 확인 |
|---|---|---:|---|
| `contexts-business-report.docx` | `479B909B57DB58C0` | 1 | 제목·문단·표 칸·목록·콜아웃의 위 첨자 1, 각주 1개 |
| `contexts-ib-report.docx` | `544EE0E75DB063AB` | 4 | 표지, 목차(제목 `개요`에 번호 없음), 본문 위 첨자 1 다섯 곳, 면책 조항 |
| `footnote-rereference.docx` | `0E5F272D3FB7836B` | 2 | 위 첨자 번호가 표 사양의 `negative: paren` 칸(`(300)¹`)을 포함해 Word 번호와 같음, 각주 4개 |

manifest의 `visualReview`는 스크립트가 쓰는 `pending` 그대로다. 위 표가 이 기록의 시각 검토다.

## 5. 회귀 테스트

`tests/test_footnote_rereference.py`(37개, 모두 실제 파싱·렌더링 경로):

- 두 번·세 번 참조, 북마크가 첫 참조 런만 감싸는지, 필드 런 서식
- 제목·표 칸·글머리 목록·번호 목록·콜아웃·인용 안의 두 번째 참조와, 그 자리에 첫 참조가 오는 경우(12개)
- 각주 여러 개가 섞인 순서의 번호와 북마크 순서
- `ib-report`(목차 포함)와 일곱 프로필 전부
- 기본 모드가 strict와 같은 필드를 쓰는지
- IB의 `^1^` 표시
- 한 번씩만 참조하면 북마크·필드가 없는지
- E5 표 참조(`@tbl-…`)와 E6 `negative: paren` 표 칸이 함께 있는 문서
- 감사: 북마크 없는 `NOTEREF` issue, 오래된 번호 warning, `w:footnotePr`가 있으면 판단하지 않음, `customMarkFollows`

각 테스트는 저장 후 다시 연 문서에 대해 `inspect_document_issues == []`, warning 없음, 각주 ID 중복 없음, 북마크 ID 중복 없음을 확인한다.

### 기존 테스트 두 개의 입력

기존 테스트 4건(함수 2개)이 고친 뒤 실패했다. 둘 다 한 문서에서 `[^1]`을 여러 번 참조하는 입력으로, Word가 열지 못하는 산출물(두 번째 `w:footnoteReference`)을 고정하고 있었다.

| 테스트 | 실패 내용 | 조치 |
|---|---|---|
| `test_render_hardening.py::test_numeric_runs_keep_hyperlinks_spacing_and_explicit_value_scale` | 칸 글자에 NOTEREF 결과 `1`이 붙음(`'12.50%1' != '12.50%'`) | 칸마다 다른 각주 `[^1]`~`[^4]`를 쓰도록 입력만 바꿈. 단언은 그대로 |
| `test_table_rules.py::test_review_five_footnote_then_line_break[paren/minus/triangle]` | 두 번째 칸의 런 구성이 필드 런으로 바뀜 | 두 번째 칸이 `[^2]`를 쓰도록 입력만 바꿈. 단언은 그대로 |

두 테스트의 목적(숫자 런의 링크·간격·단위, 쪼개진 음수와 줄바꿈)은 각주 반복과 관계가 없다. 원래 입력은 `test_footnote_rereference.py`의 `test_numeric_cells_referencing_one_note_four_times`와 `test_negative_style_cells_referencing_one_note_twice`(세 스타일)로 옮겨, 고친 산출물(칸 글자 `12.50%1`, NOTEREF 3개, 링크·굵게 유지, 음수 스타일과 쪼개짐 경고 유지)을 고정했다.

## 6. 게이트

결과는 `stage-1-done.md`(작업 보고 폴더)와 같다. 아래 표는 커밋 전 마지막 실행이다.

| 명령 | 종료 코드 | 결과 |
|---|---:|---|
| `uv run --no-sync ruff check .` | 0 | |
| `uv run --no-sync ruff format --check .` | 0 | |
| `uv run --no-sync mypy .` | 0 | |
| `uv run --no-sync pytest tests/ -q -p no:cacheprovider --basetemp <새 폴더> -W ignore` | 0 | 3492 passed, 5 skipped, 1 xfailed (404 s); 스냅숏 389 passed |
| `uv run --no-sync python -m scripts.validate_word_schema validate …\win-x64.zip` | 0 | 문서 18개, 오류 0, 예외 0 |

스냅숏: 기존 스냅숏 변경 0건. 새 샘플의 스냅숏 22개(모델 1, XML 부분 20, 부분 목록 1)를 `--snapshot-update`로 만들고 내용을 확인했다.

## 7. 범위 밖으로 남긴 것

- 정의 없는 각주 참조(`[^9]`에 `[^9]:`가 없음): strict는 거부(종료 코드 1)하지만, 기본 모드는 정의 없는 `w:footnoteReference`를 쓰고 Word가 열지 못한다("파일이 손상되었습니다", 한 번만 참조한 경우에도 같다). 이 수정과 별개인 기존 동작이며 별도 작업 후보다.
- 한국어 Word가 아닌 Word, Word for Mac, LibreOffice에서는 열어 보지 않았다.
