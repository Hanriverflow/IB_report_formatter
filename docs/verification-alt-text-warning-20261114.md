# 대체 텍스트 경고 검증 (2026-11-14)

이 문서는 로드맵 E9(접근성 진단) 가운데 대체 텍스트 경고 하나의 조사 결과와 검증 기록이다. 명도 대비 진단은 하지 않았다. 경고는 접근성 검사를 대신하지 않으며, 한 검증과 하지 않은 검증을 나눠 적는다.

## 요약

**목표.** 작성자가 설명(대체 텍스트)을 빼먹은 그림마다 경고를 한 번 남긴다. 경고는 strict 모드에서도 저장을 막지 않고, 산출 문서(DOCX)는 바뀌지 않는다. 소유자가 2026-11-14에 "대체 텍스트 경고만 넣는다"를 승인했다.

**바꾼 것.**

| 파일 | 무엇을 |
|---|---|
| `ib_render_media.py` | `ImageRenderer.warn_missing_description`: 들어간 작성자 그림의 설명이 비었거나 공백뿐이면 경고 한 줄을 `IBDocumentRenderer.warnings`와 로그에 남긴다. 블록 그림(`_insert_image`)이 부른다 |
| `ib_render_text.py` | 문단·목록·표 칸 안의 인라인 그림(`_render_inline_image`)도 같은 함수를 부른다 |
| `docx_audit.py` | `inspect_document`가 본문의 그림 가운데 `docPr/@descr`가 없거나 공백뿐인 것마다 warning을 더한다. issue는 더하지 않는다 |
| `converters.py` | DOCX 출력 변환기가 호출자가 준 `render_warnings` 목록에 렌더러 경고를 더한다 |
| `scripts/agent_render.py` | `result.json`에 `render_warnings`(렌더러 경고 전부)를 싣는다. `ok`는 바꾸지 않는다 |
| `tests/test_alt_text_warning.py` | 새 테스트 61개 |
| 문서 | `README.md`, `README.ko.md`, `CHANGELOG.md`, `AGENTS.md`, 로드맵 "뒤로 미루는 것"의 E9 항목, 이 문서 |

**결과.** 표 1의 작성자 그림 11가지 모두에서 경고가 그림마다 한 번 나오고, strict 모드에서 저장된다. 설명이 있는 그림 8가지와 저장소 샘플 17건(16개 파일과 차트를 켠 변형 1건)에서는 경고가 0건이다. 경고를 끈 렌더와 켠 렌더의 XML 파트가 같고, 기존 스냅숏 367개는 바뀌지 않았다.

## 환경

Windows 11(10.0.26300), Python 3.12.12, pytest 9.0.2, python-docx 1.2.0, lxml 6.0.2. 시작 커밋은 `main`의 `c97841c`이다. Word로 문서를 열어 보는 검증은 하지 않았다(산출 문서가 바뀌지 않으므로 쪽 검사도 하지 않았다).

## 1. 지금 엔진이 `descr`에 쓰는 것 (조사)

시험 문서를 실제 파서와 렌더러로 변환해 `wp:docPr`의 `name`·`descr`을 읽었다. python-docx는 그림을 넣을 때 `docPr`에 `id`와 `name`(`Picture N`)만 쓰고, 엔진의 `_apply_alt_text`가 설명이 있을 때만 `descr`와 `title`을 더한다.

**표 1. 그림 경우별 `descr`과 경고**

| 경우 | 입력 예 | 변경 전 `descr` | 설명을 줄 수 있는 사람 | 변환 때 경고 | `docx-audit` 경고 |
|---|---|---|---|---|---|
| 블록 그림, 설명 없음 | `![](pic.png)` | 없음 | 작성자 | 1건 | 1건 |
| 블록 그림, 공백 설명 | `![ ](pic.png)` | `" "` | 작성자 | 1건 | 1건 |
| HTML `<img>` 한 줄, `alt` 없음 | `<img src="pic.png">` | 없음 | 작성자 | 1건 | 1건 |
| HTML `<img>` 한 줄, `alt=""` | `<img src="pic.png" alt="">` | 없음 | 작성자 | 1건 | 1건 |
| Base64 블록 그림 | `![](data:image/png;base64,…)` | 없음 | 작성자 | 1건(`embedded image`) | 1건 |
| 문단 안 인라인 그림 | `앞 ![](pic.png) 뒤` | 없음 | 작성자 | 1건 | 1건 |
| 문단 안 `data:` 그림 | `앞 ![](data:…) 뒤` | 없음 | 작성자 | 1건(`embedded image`) | 1건 |
| 목록 안 인라인 그림 | `- 항목 ![](pic.png)` | 없음 | 작성자 | 1건 | 1건 |
| 파이프 표 칸 안 그림 | `\| ![](pic.png) \|` | 없음 | 작성자 | 1건 | 1건 |
| HTML 표 칸 `<img>` | `<td><img src="pic.png"></td>` | 없음 | 작성자 | 1건 | 1건 |
| HTML 표 칸 `data:` 그림 | `<td><img src="data:…"></td>` | 없음 | 작성자 | 1건(`embedded image`) | 1건 |
| 설명이 있는 그림(위 자리 모두) | `![가상 그림](pic.png)` 등 | 설명 | 작성자 | 0건 | 0건 |
| 번호 붙는 그림(E5) | `![가상 그림](pic.png){#fig-a}` | 설명(필수) | 작성자 | 0건 | 0건 |
| 블록 수식 | `$$x^2$$` | 없음 | 없음(엔진이 그림) | 0건 | 1건 |
| 인라인 수식 | `식 $x^2$ 끝` | 없음 | 없음(엔진이 그림) | 0건 | 1건 |
| 도식 | ```` ```diagram:flow ```` | 없음(제목은 캡션 문단) | 없음(엔진이 그림) | 0건 | 1건 |
| 차트(`--charts`) | ```` ```chart ```` 제목 없음 | `Chart 1`(제목이 있으면 제목) | 작성자(제목) | 0건 | 0건 |
| 텀싯 로고 | `style.logo: pic.png` | 없음 | 없음(설명 키가 없음) | 0건 | 1건 |
| 각주 안 그림 | `[^1]: 각주 ![](pic.png)` | 그림이 들어가지 않음 | — | 0건 | 0건 |
| 파일이 없는 그림 | `![](missing.png)` | 그림이 들어가지 않음(자리표시 글자, 기존 오류) | — | 0건 | 0건 |

차트 `descr`의 `Chart N`은 파서가 붙이는 이름이라 비지 않는다. 손으로 만든 모델에서 `label=""`이고 제목도 없으면 비지만, 파서를 거친 문서에서는 생기지 않아 따로 다루지 않았다.

## 2. 판단한 것과 근거

**변환 때는 작성자가 설명을 줄 수 있는 그림만 경고한다.** 수식·도식·로고에는 Markdown이나 house 파일에 설명을 적을 자리가 없다. 경고해도 작성자가 할 수 있는 일이 없고, 그 그림에 설명을 넣으려면 산출 문서가 바뀌므로 이번 범위(산출 문서 불변)를 넘는다.

**`docx-audit`는 본문의 그림을 모두 본다.** 저장된 파일에는 그림을 누가 만들었는지 적혀 있지 않고, 표시를 더하면 산출 문서가 바뀐다. `docx-audit`는 Word에서 고친 파일도 읽으며, Word에서는 어느 그림에든 대체 텍스트를 넣을 수 있다. 그래서 수식·도식·로고도 audit warning으로 나온다(표 1). 다음은 세지 않는다: Word가 장식용으로 표시한 그림(`adec:decorative val="1"`), 삭제 추적 안의 그림(`w:del`·`w:moveFrom`), `mc:Fallback` 안의 사본, 그림이 아닌 도형. 머리글·바닥글은 읽지 않는다(다른 audit 검사와 같이 본문만). 엔진은 머리글에 그림을 넣지 않는다(DRAFT 워터마크는 VML 도형이라 `docPr`이 없다).

**경고 통로.** strict 거부로 이어지는 `model.warnings`·`IBDocumentRenderer.errors`는 쓰지 않았다. 변환 경고는 E6이 만든 `IBDocumentRenderer.warnings`(매 `render` 호출마다 초기화)와 로그 `[WARNING]`에, audit 경고는 `DocumentAudit.warnings`에 들어간다. 렌더러가 부르는 `inspect_document_issues`는 issue만 보므로 audit warning은 strict에 영향이 없다. 거부 코드(6개)는 더하지 않았고, 경고 문장에 `[code] ` 접두어를 붙이지 않았다. 로드맵 본문 E9 절의 완료 기준은 "`docx_audit` JSON에 경고 코드 추가"였으나, audit JSON의 warnings는 문자열 목록이고 거부 코드는 6개로 고정돼 있어(AGENTS.md) 코드 대신 문장으로 넣었다.

**같은 그림을 두 곳에서 알린다.** 작성자 그림은 변환 경고와 audit 경고가 하나씩 나온다. 두 문구 모두 Word 이름(`Picture N`, Word의 선택 창에 보이는 이름)을 적어 같은 그림임을 알 수 있다. 변환 경고는 Markdown에서 고칠 곳(파일 이름, 적는 법)을, audit 경고는 저장된 파일의 상태를 알린다. 한 통로 안에서는 그림마다 한 번이다.

**`scripts/agent_render.py`.** 변경 전에도 audit warning은 `result.json`의 `audit.warnings`에 실렸으므로 대체 텍스트 audit 경고는 따로 손대지 않아도 실린다. 렌더러 경고(E6의 나뉜 음수 금액 경고 포함)는 실리지 않던 한계가 있었다. 레지스트리를 거쳐 렌더하므로 DOCX 출력 변환기에 `render_warnings` 목록 인자를 더하고, `result.json`에 `render_warnings`를 더했다(`converters.py` 코드 3줄, `agent_render.py` 코드 2줄). `diagnostics`는 실패 진단만 담는 기존 뜻을 지키려고 쓰지 않았고, `ok`는 바뀌지 않는다.

**경고 문구.**

- 변환: `Image without a description (alt text): 'pic.png', Word name 'Picture 1'. Write one as ![description](...) or <img alt="...">.` Base64·`data:` 그림은 `'pic.png'` 자리에 `embedded image`가 온다. 경로는 마지막 이름만 쓰고, 그림 데이터는 넣지 않는다(테스트로 확인).
- audit: `Picture without a description (alt text): Word name 'Picture 1'`.

## 3. 테스트 (`tests/test_alt_text_warning.py`, 61개)

모두 실제 파서(`parse_markdown_file`)와 `IBDocumentRenderer.render`를 거친다.

| 테스트 | 확인하는 것 |
|---|---|
| `test_an_author_image_without_a_description_warns_once` (11경우 × strict·기본) | 경고 문구 전체가 정확히 1건, `errors`와 구조 issue 0건, audit 경고 1건이 같은 Word 이름 |
| `test_an_image_with_a_description_gives_no_warning` (8경우) | 변환 경고 0건, audit 경고 0건 |
| `test_a_chart_is_described_by_its_title_or_label` | 차트 `descr`이 `Chart 1`, 경고 0건 |
| `test_an_engine_made_picture_warns_only_in_the_audit` (블록·인라인 수식, 도식) | 변환 경고·오류 0건, audit 경고 1건 |
| `test_a_term_sheet_logo_warns_only_in_the_audit` | 같음 |
| `test_each_undescribed_picture_warns_once_and_the_warning_holds_no_path_or_data` | 설명 없는 그림 4개(같은 파일 두 번 포함)에 경고 4건, Word 이름 순서 일치, 임시 폴더 경로와 Base64 데이터가 문구에 없음 |
| `test_a_missing_image_file_is_an_error_not_an_alt_text_warning` | 들어가지 않은 그림은 경고 없이 기존 오류 2건 |
| `test_warnings_start_afresh_on_every_render` | 같은 렌더러로 두 번 렌더해도 1건 |
| `test_the_warning_changes_nothing_in_the_saved_document` | 경고 함수를 끈 변환과 켠 변환의 XML 파트 전부가 같다(Base64·도식 그림의 임시 파일 이름 `tmp….png`은 원래 실행마다 달라 같은 글자로 바꿔 비교) |
| `test_strict_cli_conversion_saves_and_logs_the_warning` | `IBReportConverter` strict 변환이 파일을 저장하고 `WARNING` 로그 1건 |
| `test_the_audit_skips_decorative_and_deleted_pictures` | 장식용 표시·삭제 추적 그림은 audit 경고 없음 |
| `test_an_anchored_picture_is_audited_too` | `wp:anchor` 그림도 audit 경고 |
| `test_agent_render_lists_render_warnings_and_stays_ok` | `ok` 참, `diagnostics` 빈 목록, `render_warnings`에 대체 텍스트 경고와 E6 나뉜 금액 경고, `audit.warnings`에 1건 |
| `test_agent_render_has_no_render_warnings_for_a_described_image` | `render_warnings` 빈 목록 |
| `test_repository_samples_give_no_alt_text_warning` (17건) | 저장소 샘플 전체(차트 켠 변형 포함)에서 변환·audit 경고 0건 |

## 4. 게이트

| 명령 | 종료 코드 | 결과 |
|---|---|---|
| `uv run --no-sync ruff check .` | 0 | `All checks passed!` |
| `uv run --no-sync ruff format --check .` | 0 | 106개 파일 형식 맞음 |
| `uv run --no-sync mypy .` | 0 | 소스 43개, 오류 0건 |
| `uv run --no-sync pytest tests/ -q -p no:cacheprovider --basetemp <새 임시 폴더> -W ignore` | 0 | 3510 통과, 5 건너뜀, 1 xfail(기존 strict xfail), 553초. 스냅숏 367개 통과 |
| `git status --short tests/__snapshots__` | 0 | 0줄(스냅숏 변경 0건) |
| `md-to-word doc.md out.docx --strict` (설명 없는 그림 1개, 있는 그림 1개) | 0 | `[WARNING] Image without a description (alt text): 'pic.png', Word name 'Picture 1'. …` 1줄, 파일 저장 |
| `docx-audit out.docx` (위 파일) | 0 | `issues` 빈 목록, `warnings`에 `Picture without a description (alt text): Word name 'Picture 1'` 1건 |

기존 테스트의 단언은 바꾸지 않았고, 건너뛰기·lint 예외를 더하지 않았다. 경고를 더한 뒤에도 Word 스키마 검증 테스트(`tests/test_word_schema_validation.py`)가 통과했다. 산출 문서가 바뀌지 않으므로 Word 쪽 이미지 검사는 하지 않았다.

## 5. 하지 못한 것

- 명도 대비 진단(E9의 다른 절반)은 하지 않았다.
- 수식·도식·로고에 설명을 넣는 일은 하지 않았다. 산출 문서가 바뀌므로 소유자 판단이 필요하다(수식은 식 글자, 도식은 제목, 로고는 house 키를 후보로 볼 수 있다).
- 차트의 `Chart N`처럼 내용을 말하지 않는 설명, 파일 이름을 그대로 적은 설명은 경고하지 않는다. 설명의 질은 판단하지 않는다.
- `docx-audit`는 머리글·바닥글·각주의 그림을 읽지 않는다.
- Word의 접근성 검사기와 결과를 대조하지 않았다. 장식용 표시의 XML은 Office 확장 이름(`http://schemas.microsoft.com/office/drawing/2017/decorative`)으로 넣었고, 이 기기의 Word로 장식용 표시를 해 저장한 문서와 대조하지 않았다.
- `.agents/skills/ib-document/SKILL.md`와 `skills/ib-document/SKILL.md`에 `render_warnings`를 읽으라는 문장을 더하지 않았다(대체 텍스트 audit 경고는 기존 `audit` 안에 실린다).
- 독립 검토는 하지 않았다.
