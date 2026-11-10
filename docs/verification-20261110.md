# 저장소 정비와 안전·검증 기반 검증 (2026-11-10 착수)

이 문서는 저장소 정비(실명 자료 제거, 코드 서식·타입 표기 정리)와 안전·검증 기반(이미지·로고 경로 제한, 입력 한도, Word 스키마 검증 게이트, 스냅숏·속성 테스트) 작업의 검증 기록이다. 한 검증과 하지 않은 검증을 구분해 적으며, 구조 검사 통과는 시각 검토 통과를 뜻하지 않는다.

## 환경

Windows 11(10.0.26300), Python 3.12.12, uv 0.11.28, ruff 0.15.0, mypy 1.14.1, pytest 9.0.2, python-docx 1.2.0, lxml 6.0.2, Microsoft Word 16.0(빌드 16.0.20527). 검증은 2026-11-10과 2026-11-11에 했다. 기준선은 2026-11-10 `main`(`24bab05`)에서 `uv run pytest tests/` **1316 passed, 4 skipped**.

## 실명 자료 제거

- 실명 회사 자료 2개와 이를 읽던 프로토타입 스크립트 1개를 지웠다. 어떤 테스트도 이 파일을 쓰지 않았다.
- 테스트 코드 2개의 회사명·문서 제목을 가상 이름으로, `next_step.md`의 회사명과 지운 파일 경로를 "실제 사례 A/B"로 바꿨다.
- 추적 중인 파일 전체에서 두 회사명을 검색한 결과 **0건**(`git grep`, 종료 코드 1).
- 전체 테스트 **1316 passed, 4 skipped**(기준선과 같음).
- 과거 커밋 기록은 다시 쓰지 않았다. 지운 파일은 이전 커밋에 남아 있다.

## 코드 서식·타입 표기 정리

| 단계 | 내용 | 결과 |
|---|---|---|
| 서식 | `ruff format .`를 전체에 적용(54개 파일 변경) | `ruff format --check .` 통과 |
| 타입 표기 | 린트 무시 목록에서 `B905`, `UP006`, `UP007`, `UP035`, `UP045` 제거, 드러난 위반 799건 수정 | `ruff check .` 위반 0건, 무시 목록은 `E501`, `B008`만 남음 |
| `zip` 호출 10곳 | 주변 코드가 같은 길이를 보장하는 5곳만 `strict=True`, 나머지 5곳은 `strict=False` 명시 | 동작 변화 없음 |
| 타입 검사 | 제외 목록에 `temp/`, `tmp/`, `private/` 추가 | 같은 이름의 모듈이 든 `temp/` 폴더를 만든 상태에서 `mypy .` 통과 |
| CI | 서식 검사 단계 추가, CI 명령 목록을 고정한 테스트 갱신 | 로컬 테스트 통과. CI 결과는 아래 "CI" 절 |

`strict=True`로 둔 곳: `waterfall_positions`가 함께 채우는 두 목록, 바로 윗줄에서 한 목록으로부터 만든 목록, 테스트 안의 길이가 고정된 목록 3곳.

**동작 불변 확인.** 정리 전(`1e57c2b`)과 후에 `samples/profiles`, `samples/qa`, `samples/distribution`의 모든 Markdown 샘플 12개를 엄격 모드로 변환했다(차트 샘플은 차트를 켠 설정으로 한 번 더, 문서 13개). 각 DOCX 안의 모든 XML 파트 **253개**를 비교한 결과 서식 단계 후와 타입 표기 단계 후 모두 정리 전과 **동일**했다. 비교에서 제외한 값은 실행마다 달라지는 세 가지뿐이다: 수정 식별자(`w:rsid*`), 문서 속성의 생성·수정 시각, 의미 북마크 이름 끝의 임의 문자열(`_ibrep_…_xxxx`). 같은 코드로 두 번 변환한 결과가 서로 같음도 먼저 확인했다.

전체 테스트 **1316 passed, 4 skipped**, `uv build` 성공.

`.git-blame-ignore-revs`에는 서식 커밋의 작업 브랜치 해시를 적었다. 병합 방식에 따라 해시가 달라지면 병합 뒤에 고쳐야 한다.

## 이미지·로고 경로 제한

자동 테스트는 `tests/test_asset_policy.py`에 있고 최종 상태에서 모두 통과했다. 아래 표의 "명령줄 확인"은 2026-11-11에 가상 입력으로 `md_to_word.py`를 직접 실행한 결과다.

| 확인한 것 | 테스트 | 명령줄 확인 |
|---|---|---|
| `../`, 절대 경로, UNC, 링크로 폴더를 벗어나는 네 가지가 엄격 모드에서 `asset_outside_root`로 거부되고 산출 파일이 없다 | `test_escape_forms_rejected_before_save`(4가지), `test_junction_escape` | 종료 코드 1, 코드 표시 있음, DOCX 없음(정션 포함 4가지) |
| 일반 모드에서는 자리표시와 코드가 붙은 경고가 남는다 | `test_all_body_image_forms_leave_placeholder_and_coded_warning` | 종료 코드 0, 경고에 코드 있음, DOCX 있음 |
| 폴더 밖 경로의 결과가 파일 존재 여부에 따라 달라지지 않는다 | `test_outside_existence_never_changes_code` | 없는 파일도 같은 코드로 거부 |
| 폴더 안으로 해석되는 절대 경로는 허용 | `test_absolute_inside_and_default_folder_cannot_be_removed` | 종료 코드 0, DOCX 있음 |
| `--asset-root`와 API 인자로 추가한 폴더의 이미지는 허용 | `test_cli_repeatable_asset_root`, `test_api_roots_are_additive` | 종료 코드 0, DOCX 있음 |
| 문서가 지정한 기관 파일의 폴더 밖 로고는 거부, `--house`와 API로 준 기관 파일의 로고는 허용 | `test_house_logo_trust_depends_on_who_named_house`, `test_frontmatter_logo_cannot_borrow_operator_house_trust` | 테스트로만 확인 |
| 글자·스트림 입력과 직접 만든 모델은 허용 폴더를 넘기지 않으면 경로 이미지가 거부되고 넘기면 허용 | `test_no_source_requires_roots`, `test_text_and_stream_accept_parse_roots`, `test_relative_paths_do_not_grant_cwd_access` | 테스트로만 확인 |
| 파싱과 렌더를 나눠 부른 결과가 한 번에 부른 결과와 같다 | `test_registry_two_steps_match_cli_converter` | 테스트로만 확인 |
| 문서 내용으로는 허용 폴더를 넓힐 수 없다 | `test_document_settings_cannot_grant_roots` | 테스트로만 확인 |
| 기존 경고 문장은 그대로다 | `test_existing_missing_image_warning_unchanged`, `test_existing_missing_logo_warning_unchanged` | — |

Windows에서 디렉터리 심볼릭 링크를 만들 권한이 없으면 "링크" 경우는 건너뛴다(생략 5건 중 하나). 같은 경우를 Windows에서는 정션으로(`test_junction_escape`), Ubuntu CI에서는 심볼릭 링크로 검사한다.

## 입력 한도

`tests/test_input_limits.py`가 다섯 한도를 검사하고 최종 상태에서 모두 통과했다.

| 확인한 것 | 테스트 | 명령줄 확인 |
|---|---|---|
| 다섯 한도 각각의 초과가 엄격·일반 모드 모두에서 해당 코드로 중단되고 산출 파일이 없으며, 한도를 올리면 통과 | `test_each_limit_stops_without_output_and_can_be_raised`, `test_cli_adjusts_each_limit` | 표 칸 수: 두 모드 모두 종료 코드 1·DOCX 없음, 한도를 올리면 변환됨. 입력 크기: 종료 코드 1·DOCX 없음 |
| 0, 음수, 소수, `unlimited`는 인자 오류 | `test_cli_invalid_limits_are_argument_errors` | 종료 코드 2 |
| 일괄 변환에서 나머지 파일은 계속 | `test_batch_continues_after_limit_failure` | 테스트로만 확인 |
| 크기는 읽거나 풀기 전에 판정 | `test_file_size_checked_before_read`, `test_stream_stops_at_first_excess_byte`, `test_base64_excess_is_not_decoded`, `test_huge_pixel_dimensions_are_read_without_decoding`, `test_image_bytes_checked_before_opening` | 테스트로만 확인 |
| 이미지 한도는 출처와 무관하게 모든 이미지와 로고에 적용 | `test_image_limits_cover_all_origins` | 테스트로만 확인 |
| 직접 만든 모델도 표 칸 수·목록 깊이 한도를 받는다 | `test_hand_built_models_cannot_bypass_limits` | 테스트로만 확인 |
| 정리 옵션을 쓰면 원본과 정리된 글 모두 판정 | `test_cleanup_checks_original_then_derived` | 테스트로만 확인 |
| 문서 내용으로는 한도를 올릴 수 없다 | `test_frontmatter_cannot_raise_limits` | 테스트로만 확인 |

## 거부 코드와 예외 종류

- 코드가 붙은 문제는 이어 붙인 거부 문장 안에서도 자기 접두를 유지한다: `test_render_rejection_keeps_prefix_on_each_issue`, `test_strict_exception_preserves_uncoded_issues_and_wrapper_cause`.
- `result.json`: 폴더 밖 이미지마다 진단 하나(`test_outside_images_each_have_a_coded_diagnostic`), 한도 초과(`test_input_limit_adds_coded_diagnostic_without_runner_option`), 여러 코드가 섞인 거부(`test_multiple_codes_preserve_individual_sentences`), 감싸진 한도 예외(`test_wrapped_render_limit_preserves_outer_exception`), 코드 없는 엄격 거부는 종전대로 `ValueError`로 보고(`test_uncoded_strict_rejection_retains_value_error_and_message`).
- 실행기와 휴대용 변환기는 허용 폴더·한도 옵션을 받지 않는다: `test_runner_rejects_safety_override_arguments`, `test_portable_rejects_safety_override_arguments`.
- 휴대용 변환기의 한국어 안내 여섯 가지: `test_korean_rejection_causes_and_remedies`, `test_portable_outside_image_rejection_has_korean_guidance`.
- 플러그인 안내문을 따른 가상 예제(작업 폴더에 복사한 기관 파일과 로고)가 실행기에서 변환된다: `test_copied_house_and_logo_render_inside_job_folder`.

## 감사 도구의 XML 읽기

`tests/test_docx_audit_xml.py`: 내부·외부 엔티티 선언이 든 DOCX를 감사해도 내용이 펼쳐지지 않고(`test_audit_does_not_expand_entities`), 정상 문서의 감사 결과와 JSON은 종전과 같다(`test_normal_audit_result_and_json_are_unchanged`).

## Word 스키마 검증 게이트

- **도구.** OOXML-Validator v2.1.6 배포 실행 파일. 내려받은 파일의 SHA-256을 `scripts/word_schema_validator_assets.json`과 대조한다. 해시가 다르면 `verify-asset`이 종료 코드 1로 끝난다(`tests/test_word_schema_validation.py`, 로컬 실행으로도 확인).
- **대상.** 샘플 12개를 엄격 모드로 변환한 문서 13개(차트 샘플은 차트를 켠 설정으로 한 번 더).
- **첫 보고(수정 전).** 문서 13개 모두에서 같은 오류 1건: `/word/settings.xml`의 `w:updateFields`가 스키마가 정한 자리보다 뒤에 있었다. 엔진이 이 요소를 설정 파트 맨 끝에 붙였기 때문이다.
- **수정.** 스키마 순서 삽입 도우미로 제자리에 넣는다. 요소와 속성은 그대로이고 자리만 바뀐다. 바뀐 스냅숏은 문서 13개의 `settings.xml` 13개뿐이며, 각 파일의 차이는 `w:updateFields` 한 줄의 위치다.
- **수정 후.** 로컬(Windows)에서 검증기 실행 결과 문서 13개 모두 오류 0건, 종료 코드 0. Ubuntu CI 결과는 아래 "CI" 절에 적는다.
- **예외 목록.** 비어 있다. Word 재현 기록이 필요한 항목이 없다.
- **게이트.** CI 작업에서 `continue-on-error`를 뺐다. push와 pull request에서 실행되고, 예외 목록에 없는 오류가 있으면 실패한다. 보고서는 실패해도 산출물로 올라간다.

### Word 쪽 확인 (설정 파트 수정 전후)

설정 파트 수정은 XML만 바꾼다. 그래서 수정 전후의 Word 쪽 이미지가 같은지 확인했다.

- **환경.** Windows 11(10.0.26300), Microsoft Word 16.0(빌드 16.0.20527), 2026-11-11.
- **절차.** 수정 전(`6591875`)과 수정 후(`5e8f286`)의 엔진으로 문서 13개를 각각 엄격 모드로 변환했다. 변환한 문서마다 `scripts/word_visual_qa.ps1`을 새 출력 폴더에 실행해 필드를 갱신하고 모든 쪽을 PNG로 만들었다. 매니페스트에는 원본·갱신 DOCX·PNG의 SHA-256, 쪽수, Word 버전이 기록됐다.
- **결과.** 문서 13개, 32쪽. 쪽수는 모든 문서에서 수정 전과 같고, 32쪽 모두 수정 전 이미지와 화소 단위로 같다.
- **눈으로 본 범위.** 수정 후 32쪽을 여섯 장의 모아보기로 줄여 전부 보았다. 빈 쪽, 깨진 표, 잘린 내용, 빠진 머리글·바닥글은 없었다. 줄여서 본 것이므로 글자 단위 교정이나 내용 검토는 아니다. 매니페스트의 `visualReview`는 `pending`으로 두었다.
- 수정 전과 후 사이에는 제목과 코드 블록의 수정(아래 "속성 테스트" 절)도 들어 있다. 따라서 이 비교는 그 수정이 샘플의 쪽 모양을 바꾸지 않았다는 확인도 겸한다.
- 산출물(DOCX, PNG, 매니페스트)은 저장소 밖 임시 폴더에 두었고 저장소에 넣지 않았다.

| 문서 | 쪽수(전 → 후) | 수정 후 DOCX SHA-256 | 쪽 이미지 비교 |
|---|---|---|---|
| `charts-on--qa--charts.docx` | 3 → 3 | `f4a3ad91d2b9293e5cd137bd5c293c50eeb8029cf1ad55fb30afc24552022f30` | 모든 쪽 동일 |
| `default--distribution--minimal-term-sheet.docx` | 2 → 2 | `64b1c581c245524b42d9027f36b09a804c5f1ae32372ea4a124a21442ae1ec42` | 모든 쪽 동일 |
| `default--profiles--business-report.docx` | 1 → 1 | `d1857e0c7be7433ee4405c5cc8e35be383bb485b86ea17b4647fecf0c0b9e21c` | 모든 쪽 동일 |
| `default--profiles--ib-memo.docx` | 1 → 1 | `c76648090fbffde8371b7af8ecae7d7caa362ae9103cb292116cc53650323f8a` | 모든 쪽 동일 |
| `default--profiles--ib-report.docx` | 4 → 4 | `ec33f65a90dfa8b341ca7f868515d564853e619f897b74b387c17bf29df4a2a2` | 모든 쪽 동일 |
| `default--profiles--meeting-minutes.docx` | 1 → 1 | `6a136efc7faec65ddbcf018d83d63061e4df7f5a7f9adf3689e83036ece973fb` | 모든 쪽 동일 |
| `default--profiles--office-letter.docx` | 1 → 1 | `abb6998b29125de791e1220ce721f8ed56b0018f2cdae4ecbda92a4c231a7238` | 모든 쪽 동일 |
| `default--profiles--plain.docx` | 1 → 1 | `19c189ca680ae6150b61bb63c2126b62235c4336b162e69fa4cb3b9548993881` | 모든 쪽 동일 |
| `default--profiles--term-sheet.docx` | 8 → 8 | `5f16d9a36cc73273d6a5fac1640b5d921fa45d3be65118c1c9cd21d8b027cc4b` | 모든 쪽 동일 |
| `default--qa--charts.docx` | 3 → 3 | `55e3e61a9a42a7ae8d38c84693a1e6bef51602dd3a305934039713b81237887d` | 모든 쪽 동일 |
| `default--qa--landscape.docx` | 3 → 3 | `fbc8e0adca02f915a4a25c459e1db9ce249048535b36afd1f9c7148221b701b6` | 모든 쪽 동일 |
| `default--qa--long-table.docx` | 2 → 2 | `f7e800d613f91ef6b2c97e81c4b5552b26de1bb416e4a9edb097db4140b682fc` | 모든 쪽 동일 |
| `default--qa--office-letter-appendix.docx` | 2 → 2 | `dfe2eee0b074512e4589931a53911934a494c781ec6bc51ee38896aa7cb1621d` | 모든 쪽 동일 |

## 스냅숏 테스트

- 대상은 스키마 검증과 같은 문서 13개다. 문서마다 파싱한 모델의 JSON 1개와 DOCX 안의 모든 XML 파트(`.xml`, `.rels`)를 저장한다. 스냅숏 파일은 279개다.
- 정규화하는 값은 실행마다 달라지는 것뿐이다: 수정 식별자(`w:rsid*`), 문서 속성의 생성·수정 시각, 의미 북마크 이름 끝의 임의 문자열. 모델 JSON에서 절대 경로는 샘플 폴더 기준 상대 경로로, 이진 데이터는 길이와 해시로 적는다.
- 스냅숏이 없으면 테스트가 실패한다. 갱신은 `--snapshot-update`로만 한다.
- 같은 스냅숏 파일이 CI의 네 조합에서 통과했다(아래 "CI" 절). 줄 끝 변환으로 스냅숏이 달라지지 않도록 `.gitattributes`에서 스냅숏 폴더를 LF로 고정했다.

## 속성 테스트

`tests/test_markdown_properties.py`. Hypothesis 설정은 `derandomize=True`, `deadline=None`, `database=None`, `max_examples=24`여서 실행마다 같은 예시를 쓴다. 모듈 실행 시간은 Windows에서 약 1분이다.

- **생성 범위.** 문단(줄 병합, 공백·탭, 괄호 포함), 제목(1~6단계), 순서·비순서 목록(`-`, `*`, `+`, `1.`), 파이프 표(머리행 포함 1~5행, 1~4열, 빈 칸 포함), 강조, 코드 구간, 코드 블록(백틱·물결), 역슬래시 이스케이프(ASCII 문장부호 전체), 한글·영문·숫자·문장부호. 엄격 모드 이분성에는 닫히지 않은 구문, 깨진 앞머리, 100자 이내 임의 유니코드 문자열(제어 문자 포함)을 더했다. 짝 없는 서로게이트는 뺐다: 공개 API가 입력을 UTF-8로 바꾸는 단계에서 받지 않기 때문이다.
- **범위에서 뺀 것.** README가 지원 범위 밖이라고 적은 두 가지뿐이다: 문서화된 줄바꿈 외의 원시 HTML, 복잡한 중첩 Markdown(README "Supported Markdown and limitations"의 "Raw HTML, complex nested Markdown and arbitrary Word templates are not guaranteed."). 임의 문자열에는 이 제외를 적용하지 않았다.
- **검사 경로.** 글자·코드·칸의 글자는 `plain` 프로필, 칸 수와 빈 칸은 `plain`과 `ib-report`, 이분성은 두 프로필에서 글자 입력·스트림 입력·명령줄 변환기 세 경로 모두.

**찾아서 고친 결함**(샘플 산출물은 바뀌지 않음, 스냅숏 279개 불변):

| 결함 | 수정 |
|---|---|
| 제목 안의 코드 구간과 이스케이프한 별표에서 별표가 사라졌다(`` # `**가상**` ``, `# \*\*가상\*\*`) | 제목의 굵게 표시 제거가 코드 구간과 이스케이프를 건너뛰고, 이스케이프는 한 번만 푼다. 그 밖의 제목 27가지 형태 × 프로필 4종을 수정 전후로 비교해 XML이 같음을 확인했다 |
| 코드 블록의 공백만 있는 줄이 공백 한 칸으로 바뀌었다 | 그 줄의 공백·탭을 그대로 둔다(빈 줄은 종전대로 공백 한 칸) |
| 깨진 YAML 앞머리, 0 이하 각주 번호, 빈 입력 파일이 문서화된 거부 종류가 아닌 일반 `ValueError`로 끝났다 | 같은 문장으로 `StrictValidationError`를 던진다(`ValueError`의 하위 종류) |
| XML에 넣을 수 없는 제어 문자가 든 입력이 렌더 단계에서 다른 종류의 예외로 끝났다 | 파싱 단계에서 입력 경고로 보고하고, 엄격 모드가 저장 전에 거부한다 |

**고치지 않고 남긴 것**(출력 형식이나 설계를 정해야 해서 소유자 결정이 필요하다. 테스트에는 `xfail(strict=True)`로 남아 있어, 고치면 표시를 떼야 통과한다):

| 항목 | 지금 동작 | 건수 |
|---|---|---|
| `+` 글머리표 | 목록이 아니라 문단으로 나온다(`+ 0` → `+ 0`) | 1 |
| 5·6단계 제목 | 제목이 아니라 문단으로 나온다(`##### 0` → `##### 0`) | 3 |
| 위 두 가지가 함께 든 경우 | 위와 같다 | 1 |
| 엔진이 풀지 않는 이스케이프 14자(`! # % & + , / : ; < = > ? @`) | 역슬래시가 그대로 보인다(`\!` → `\!`). 엔진이 푸는 문자는 `` \ $ ` ^ ~ . * " ' ( ) [ ] { } | _ - ``다. 범위를 넓히면 기존 입력의 보이는 글자가 바뀌고, 표 병합 표시를 다루는 기존 테스트 1개의 기대값(`\<<`)도 바뀐다 | 15 |
| 코드 블록 끝의 빈 줄과 줄 끝 공백 | 사라진다 | 6 |
| 코드 블록 빈 줄의 공백 한 칸 | 원문에 없는 공백 한 칸이 들어간다(줄 높이를 유지하려는 기존 처리) | 2 |
| 코드 블록 안 트리 표기의 값 강조 | `1%` 같은 값이 굵게 나온다(의도한 기존 서식) | 1 |

합계 29건이다.

## CI

작업 브랜치 `dryforge/safety-verification-base`를 GitHub에 올릴 때마다 Actions가 실행됐다. 품질 검사는 Ubuntu·Windows × Python 3.12·3.13 네 조합이고, 린트·서식·타입 검사·전체 테스트(스냅숏, 속성 테스트 포함)·빌드·wheel 내용 확인을 한다.

| 커밋 | 여기까지 들어간 것 | 품질 검사 네 조합 | Word 스키마 검증(Ubuntu) |
|---|---|---|---|
| `6591875` | 스냅숏 테스트 | 모두 성공 | 보고 전용 단계: 문서 13개에서 오류 13건 보고 |
| `f0fd69d` | 속성 테스트 | 모두 성공 | 보고 전용 단계: 같은 오류 13건 보고 |
| `5e8f286` | 설정 파트 수정, 게이트 전환 | 모두 성공 | 게이트: 오류 0건으로 성공 |

마지막 줄의 실행 기록은 `https://github.com/Hanriverflow/IB_report_formatter/actions/runs/42848499047`이다. 이 문서가 들어간 문서 마감 커밋의 CI 결과는 이 문서에 적을 수 없으므로 GitHub의 해당 커밋 표시로 확인한다.

## 최종 상태의 필수 검증

2026-11-11, Windows 11(10.0.26300), Python 3.12.12, uv 0.11.28. 문서 마감 변경까지 들어간 작업 폴더에서 실행했고, 그 뒤로는 이 문서의 결과 문장만 고쳤다.

| 명령 | 결과 |
|---|---|
| `uv run ruff check .` | 종료 코드 0 |
| `uv run ruff format --check .` | 종료 코드 0(파일 83개) |
| `uv run mypy .` | 종료 코드 0(파일 31개) |
| `uv run pytest tests/`(OS 임시 폴더 아래 새 `--basetemp`) | **1683 passed, 5 skipped, 29 xfailed**. 스냅숏 279개 통과 |
| `uv build` | sdist와 wheel 생성. wheel 최상위 모듈 21개가 `pyproject.toml` 목록과 같다(`input_safety.py` 포함) |
| Word 스키마 검증(로컬 Windows, 검증기 v2.1.6) | 문서 13개 오류 0건, 종료 코드 0 |
| 명령줄 확인 13가지(위 경로 제한·입력 한도 표) | 모두 기대와 같음 |

기준선(1316 passed, 4 skipped)보다 통과 367개, 생략 1개가 늘었고 예상 실패 29개가 생겼다. 생략 5건은 모두 이 Windows 계정이 할 수 없는 일이다: 심볼릭 링크 만들기 4건(그중 1건이 이번에 추가한 경로 제한의 "링크" 경우), POSIX 권한 비트 1건. 예상 실패 29건은 "속성 테스트" 절의 표에 있다.

## 하지 않은 검증

- **Word 시각 검토는 위 "Word 쪽 확인"에 적은 범위뿐이다.** 그 확인은 설정 파트 수정 전후의 쪽 이미지가 같은지를 본 것이고, 각 쪽의 내용이 업무상 맞는지를 사람이 검토한 것이 아니다. 매니페스트의 `visualReview`는 `pending`으로 남아 있다.
- 경로 제한·입력 한도는 문서의 겉모양을 바꾸지 않으므로(샘플 XML 동일) 별도의 Word 쪽 확인을 하지 않았다.
- macOS에서는 어떤 것도 실행하지 않았다. Linux는 CI(Ubuntu)의 자동 검사만 했고 Word 확인은 없다.
- 휴대용 변환기 창의 한국어 안내는 문장을 만드는 함수와 변환 함수를 테스트했다. 실제 창을 띄워 눈으로 보지 않았다.
- 플러그인은 실행기 스크립트를 테스트했다. Claude Code·Codex에 설치해 대화로 실행하는 확인은 하지 않았다.
- 검증기의 대상 Office 버전은 Microsoft 365 하나다. 다른 버전 기준으로는 돌리지 않았다.
- `md-format`과 `docx-audit`에 주는 파일의 크기 한도, 퍼징, LibreOffice 렌더 비교는 이번 범위가 아니다.
- 배포 묶음(휴대용 변환기, 플러그인, 소스 묶음)을 새로 만들어 게시하지 않았다. 묶음을 만드는 스크립트의 테스트만 통과했다.

## 남은 확인 사항

- `.git-blame-ignore-revs`의 해시는 작업 브랜치의 서식 커밋이다. `main`에 합칠 때 커밋이 다시 만들어지는 방식(스쿼시, 리베이스)을 쓰면 새 해시로 고쳐야 한다.
- 이미지 크기를 머리 부분만 읽어 판정하는 데 python-docx의 내부 함수(`_ImageHeaderFactory`)를 쓴다. python-docx를 올릴 때 `tests/test_input_limits.py`가 통과하는지 확인한다.
- 스냅숏 폴더는 약 14MB다(문서 13개의 모든 XML 파트).
- 속성 테스트가 남긴 29건은 위 표의 여섯 항목이다.
