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
| 로고 경로도 본문 이미지처럼 적힌 형태와 퍼센트 디코딩한 형태를 각각 해석해 판정한다 | `test_logo_candidates_resolve_independently_through_links`, `test_logos_already_try_written_then_decoded_filenames` | 테스트로만 확인 |

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
| 플러그인 실행기는 입력 크기 한도를 해시 계산보다 먼저 적용하고, 읽는 양도 한도로 묶는다 | `test_markdown_hash_is_bounded_before_parsing` | 테스트로만 확인 |

## 거부 코드와 예외 종류

- 코드가 붙은 문제는 이어 붙인 거부 문장 안에서도 자기 접두를 유지한다: `test_render_rejection_keeps_prefix_on_each_issue`, `test_strict_exception_preserves_uncoded_issues_and_wrapper_cause`.
- `result.json`: 폴더 밖 이미지마다 진단 하나(`test_outside_images_each_have_a_coded_diagnostic`), 한도 초과(`test_input_limit_adds_coded_diagnostic_without_runner_option`), 여러 코드가 섞인 거부(`test_multiple_codes_preserve_individual_sentences`), 감싸진 한도 예외(`test_wrapped_render_limit_preserves_outer_exception`), 코드 없는 엄격 거부는 종전대로 `ValueError`로 보고(`test_uncoded_strict_rejection_retains_value_error_and_message`).
- 실행기와 휴대용 변환기는 허용 폴더·한도 옵션을 받지 않는다: `test_runner_rejects_safety_override_arguments`, `test_portable_rejects_safety_override_arguments`.
- 휴대용 변환기의 한국어 안내 여섯 가지: `test_korean_rejection_causes_and_remedies`, `test_portable_outside_image_rejection_has_korean_guidance`.
- 플러그인 안내문을 따른 가상 예제(작업 폴더에 복사한 기관 파일과 로고)가 실행기에서 변환된다: `test_copied_house_and_logo_render_inside_job_folder`.
- 문서 자신의 설정을 거부하는 오류(앞머리 값, 표 사양, 문서가 지정한 기관·테마 파일의 내용, 조건 변수와 검사 선언)는 종전과 같은 문장의 `StrictValidationError`로 나오고 원래 오류가 원인으로 이어진다. 운영자가 준 인자의 오류와 "다른 프로필로 파싱한 모델" 오류는 종전 종류 그대로다: `tests/test_document_rejections.py`.

## 감사 도구의 XML 읽기

`tests/test_docx_audit_xml.py`: 감사 도구는 python-docx로 열기 전에 내용 종류 목록(`[Content_Types].xml`)을 먼저 읽고, 거기에 XML 종류(`application/xml`, `text/xml`, `+xml`로 끝나는 종류)로 선언된 파트와 모든 연결 정보 파트(`.rels`)를(확장자는 python-docx와 같은 규칙으로 읽어, 확장자가 없는 이름도 대상이다) 엔티티 해석·DTD 로딩·네트워크 접근을 끈 방식으로 먼저 확인한다. DTD나 엔티티 선언이 든 파트가 하나라도 있으면 그 파트 이름을 문제로 보고하고 문서를 열지 않으므로 내용이 펼쳐지지 않는다(`test_package_preflight_rejects_dtd_before_python_docx`: 본문·번호·스타일·연결 정보·내용 종류 목록·참조되지 않은 파트, 내부·외부 선언, 함수 호출과 명령줄 양쪽; `test_audit_does_not_expand_entities`). 정상 문서의 감사 결과와 JSON은 종전과 같다(`test_normal_audit_result_and_json_are_unchanged`).

## Word 스키마 검증 게이트

- **도구.** OOXML-Validator v2.1.6 배포 압축 파일. 검증 스크립트의 `validate`는 내려받은 압축 파일을 받아 `scripts/word_schema_validator_assets.json`의 SHA-256과 대조한 뒤에만 임시 폴더에 풀어 실행한다. 해시가 다르거나 확인하지 않은 파일은 실행되지 않는다(`test_validate_never_executes_unverified_or_mismatched_asset`, `test_validate_runs_the_verified_copy_even_if_download_changes`). `verify-asset`만 따로 돌려도 해시가 다르면 종료 코드 1이다.
- **대상.** 샘플 12개를 엄격 모드로 변환한 문서 13개(차트 샘플은 차트를 켠 설정으로 한 번 더).
- **첫 보고(수정 전).** 문서 13개 모두에서 같은 오류 1건: `/word/settings.xml`의 `w:updateFields`가 스키마가 정한 자리보다 뒤에 있었다. 엔진이 이 요소를 설정 파트 맨 끝에 붙였기 때문이다.
- **수정.** 스키마 순서 삽입 도우미로 제자리에 넣는다. 요소와 속성은 그대로이고 자리만 바뀐다. 바뀐 스냅숏은 문서 13개의 `settings.xml` 13개뿐이며, 각 파일의 차이는 `w:updateFields` 한 줄의 위치다.
- **수정 후.** 로컬(Windows)에서 검증기 실행 결과 문서 13개 모두 오류 0건, 종료 코드 0. Ubuntu CI 결과는 아래 "CI" 절에 적는다.
- **예외 목록.** 비어 있다. Word 재현 기록이 필요한 항목이 없다. 예외는 오류 식별자·파트·위치에 더해 오류 종류와 설명 문장까지 근거 기록과 같을 때만 그 오류를 덮는다(`test_exception_does_not_cover_different_error_at_same_location`).
- **게이트.** CI 작업에서 `continue-on-error`를 뺐다. push와 pull request에서 실행되고, 예외 목록에 없는 오류가 있으면 실패한다. 보고서는 실패해도 산출물로 올라간다.

### Word 쪽 확인 (설정 파트 수정 전후)

설정 파트 수정은 XML만 바꾼다. 그래서 수정 전후의 Word 쪽 이미지가 같은지 확인했다.

- **환경.** Windows 11(10.0.26300), Microsoft Word 16.0(빌드 16.0.20527), 2026-11-11.
- **절차.** 수정 전(`6591875`)과 수정 후(`5e8f286`)의 엔진으로 문서 13개를 각각 엄격 모드로 변환했다. 변환한 문서마다 `scripts/word_visual_qa.ps1`을 새 출력 폴더에 실행해 필드를 갱신하고 모든 쪽을 PNG로 만들었다. 매니페스트에는 원본·갱신 DOCX·PNG의 SHA-256, 쪽수, Word 버전이 기록됐다.
- **결과.** 문서 13개, 32쪽. 쪽수는 모든 문서에서 수정 전과 같고, 32쪽 모두 수정 전 이미지와 화소 단위로 같다.
- **쪽 검토.** 2026-11-11에 수정 후 32쪽을 원본 크기 이미지로 한 쪽씩 보았다(차트를 켠 변형 3쪽은 기본 차트 샘플 3쪽과 파일이 바이트 단위로 같아 그쪽 검토로 갈음했다). 검토는 이 작업을 수행한 AI 에이전트가 했고, 소유자의 검토가 아니다. 검토 대상 매니페스트(`manifest.json`)의 SHA-256은 `e456ca47d39079ad98da488ba8acee4b7e000a3202b2ffc122f2021c3d0ffcb3`이다. 본 것: 머리글·바닥글과 쪽 번호, 표지·목차·고지문, 표의 테두리·머리행·병합·쪽 넘김 시 머리행 반복, 숫자 정렬과 음수 표기, 번호 목록과 붙임, 각주, 가로 구역 전환, 차트 이미지, 확인란 상자. 결과: 13개 문서 모두에서 빈 쪽, 깨진 표, 겹치거나 잘린 글자, 빠진 머리글·바닥글, 어긋난 쪽 번호가 없었다. 긴 표는 코드 001~065가 빠짐없이 이어지고 둘째 쪽에 머리행이 반복된다. 텀싯 8쪽의 병합 칸과 상환 일정 합계 행은 온전하다.
- 매니페스트의 `visualReview` 필드는 스크립트가 `pending`으로 쓰는 값이어서 그대로이고, 검토 기록은 이 항목이다. 이 검토는 쪽 모양에 대한 것이며, 샘플 문구나 수치의 내용이 맞는지에 대한 검토가 아니다.
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
- 정규화하는 값은 실행마다 달라지는 네 가지뿐이다: 수정 식별자(`w:rsid*`와 설정 파트의 같은 값), 문서 속성의 생성·수정 시각, 의미 북마크 이름 끝의 임의 문자열, 그리고 고지문 끝의 저작권 문장에 엔진이 넣는 실행 연도(그 문장 형태에 한해 연도만 바꾸며, 본문의 다른 날짜와 연도는 그대로 둔다). 모델 JSON에서 절대 경로는 샘플 폴더 기준 상대 경로로, 이진 데이터는 길이와 해시로 적는다.
- 스냅숏이 없으면 테스트가 실패한다. 갱신은 `--snapshot-update`로만 한다.
- 같은 스냅숏 파일이 CI의 네 조합에서 통과했다(아래 "CI" 절). 줄 끝 변환으로 스냅숏이 달라지지 않도록 `.gitattributes`에서 스냅숏 폴더를 LF로 고정했다.

## 속성 테스트

`tests/test_markdown_properties.py`. Hypothesis 설정은 `derandomize=True`, `deadline=None`, `database=None`, `max_examples=24`여서 실행마다 같은 예시를 쓴다. 전체 테스트 실행 시간은 Windows에서 5~6분이고 그중 절반가량이 이 모듈이다.

- **생성 범위.** 문단(줄 병합, 공백·탭, 괄호 포함), 제목(1~6단계), 순서·비순서 목록(`-`, `*`, `+`, `1.`), 파이프 표(머리행 포함 1~5행, 1~4열, 빈 칸 포함), 강조, 코드 구간, 코드 블록(백틱·물결), 역슬래시 이스케이프(ASCII 문장부호 전체), 한글·영문·숫자·문장부호. 엄격 모드 이분성에는 닫히지 않은 구문, 깨진 앞머리, 100자 이내 임의 유니코드 문자열(제어 문자 포함)을 더했다. 짝 없는 서로게이트는 뺐다: 공개 API가 입력을 UTF-8로 바꾸는 단계에서 받지 않기 때문이다.
- **범위에서 뺀 것.** README가 지원 범위 밖이라고 적은 두 가지뿐이다: 문서화된 줄바꿈 외의 원시 HTML, 복잡한 중첩 Markdown(README "Supported Markdown and limitations"의 "Raw HTML, complex nested Markdown and arbitrary Word templates are not guaranteed."). 임의 문자열에는 이 제외를 적용하지 않았다.
- **검사 경로.** 글자·코드·칸의 글자는 `plain` 프로필, 칸 수와 빈 칸은 `plain`과 `ib-report`, 이분성은 두 프로필에서 글자 입력·스트림 입력·명령줄 변환기 세 경로 모두.
- **최종 검토 뒤에 넓힌 것.** 엄격 모드 이분성에 YAML로는 맞지만 값의 종류나 내용이 틀린 앞머리(프로필, 배치, 스타일, 표 사양, 조건 변수, 검사, 공문 설정 등)와 YAML 이스케이프로 들어오는 제어 문자를 더했다. 코드 블록은 지금 동작하는 범위(줄마다 글자 보존, 서식 없음)를 생성 입력으로 실제 검사하고, 알려진 빈틈만 따로 예상 실패로 둔다.

**찾아서 고친 결함**(샘플 산출물은 바뀌지 않음, 스냅숏 279개 불변):

| 결함 | 수정 |
|---|---|
| 제목 안의 코드 구간과 이스케이프한 별표에서 별표가 사라졌다(`` # `**가상**` ``, `# \*\*가상\*\*`) | 제목의 굵게 표시 제거가 코드 구간과 이스케이프를 건너뛰고, 이스케이프는 한 번만 푼다. 그 밖의 제목 27가지 형태 × 프로필 4종을 수정 전후로 비교해 XML이 같음을 확인했다 |
| 코드 블록의 공백만 있는 줄이 공백 한 칸으로 바뀌었다 | 그 줄의 공백·탭을 그대로 둔다(빈 줄은 종전대로 공백 한 칸) |
| 깨진 YAML 앞머리, 0 이하 각주 번호, 빈 입력 파일이 문서화된 거부 종류가 아닌 일반 `ValueError`로 끝났다 | 같은 문장으로 `StrictValidationError`를 던진다(`ValueError`의 하위 종류) |
| XML에 넣을 수 없는 제어 문자가 든 입력이 렌더 단계에서 다른 종류의 예외로 끝났다 | 파싱 단계에서 입력 경고로 보고하고, 엄격 모드가 저장 전에 거부한다 |
| 앞머리 설정 값이 틀린 문서(예: `layout: {cover: wrong}`)가 문서화된 거부 종류가 아닌 일반 `ValueError`로 끝났다 | 문서 설정을 검사하는 경계마다 같은 문장의 `StrictValidationError`로 바꿔 던진다. 운영자 인자의 오류는 종전 종류 그대로다 |
| YAML 이스케이프로 들어온 제어 문자(예: 제목의 NUL)가 원문 검사에 걸리지 않고 렌더 단계에서 다른 종류의 예외로 끝났다 | 해석된 메타데이터와 문서가 지정한 기관·테마 파일의 문자열도 같은 진단으로 검사한다 |

**고치지 않고 남긴 것**(출력 형식이나 설계를 정해야 해서 소유자 결정이 필요하다. 테스트에는 `xfail(strict=True)`로 남아 있어, 고치면 표시를 떼야 통과한다):

| 항목 | 지금 동작 | 건수 |
|---|---|---|
| `+` 글머리표 | 목록이 아니라 문단으로 나온다(`+ 0` → `+ 0`) | 1 |
| 5·6단계 제목 | 제목이 아니라 문단으로 나온다(`##### 0` → `##### 0`) | 3 |
| 위 두 가지가 함께 든 경우 | 위와 같다 | 1 |
| 엔진이 풀지 않는 이스케이프 14자(`! # % & + , / : ; < = > ? @`) | 역슬래시가 그대로 보인다(`\!` → `\!`). 엔진이 푸는 문자는 `` \ $ ` ^ ~ . * " ' ( ) [ ] { } | _ - ``다. 범위를 넓히면 기존 입력의 보이는 글자가 바뀌고, 표 병합 표시를 다루는 기존 테스트 1개의 기대값(`\<<`)도 바뀐다 | 15 |
| 코드 블록 끝의 빈 줄과 줄 끝 공백 | 사라진다 | 7 |
| 코드 블록 빈 줄의 공백 한 칸 | 원문에 없는 공백 한 칸이 들어간다(줄 높이를 유지하려는 기존 처리) | 4 |
| 코드 블록 안 트리 표기의 값 강조 | `1%` 같은 값이 굵게 나온다(의도한 기존 서식) | 1 |

합계 32건이다.

## CI

작업 브랜치 `dryforge/safety-verification-base`를 GitHub에 올릴 때마다 Actions가 실행됐다. 품질 검사는 Ubuntu·Windows × Python 3.12·3.13 네 조합이고, 린트·서식·타입 검사·전체 테스트(스냅숏, 속성 테스트 포함)·빌드·wheel 내용 확인을 한다.

| 커밋 | 여기까지 들어간 것 | 품질 검사 네 조합 | Word 스키마 검증(Ubuntu) |
|---|---|---|---|
| `6591875` | 스냅숏 테스트 | 모두 성공 | 보고 전용 단계: 문서 13개에서 오류 13건 보고 |
| `f0fd69d` | 속성 테스트 | 모두 성공 | 보고 전용 단계: 같은 오류 13건 보고 |
| `5e8f286` | 설정 파트 수정, 게이트 전환 | 모두 성공 | 게이트: 오류 0건으로 성공 |
| `75efb5c` | 첫 검토의 지적 반영과 그 기록 | 모두 성공 | 게이트: 오류 0건으로 성공 |
| `3bd6419` | 두 번째 검토의 지적 반영과 그 기록 | 모두 성공 | 게이트: 오류 0건으로 성공 |
| `34dade8` | 확인 검토의 지적 반영과 그 기록 | 모두 성공 | 게이트: 오류 0건으로 성공 |

마지막 줄의 실행 기록은 `https://github.com/Hanriverflow/IB_report_formatter/actions/runs/42909444757`이다. 마지막 확인 검토의 권고를 반영한 커밋과 이 문서가 들어간 문서 마감 커밋의 CI 결과는 이 문서에 적을 수 없으므로 GitHub의 해당 커밋 표시로 확인한다.

## 최종 상태의 필수 검증

2026-11-11, Windows 11(10.0.26300), Python 3.12.12, uv 0.11.28. 마지막 확인 검토의 권고까지 반영한 뒤의 상태에서 실행했고, 그 뒤로는 문서만 고쳤다.

| 명령 | 결과 |
|---|---|
| `uv run ruff check .` | 종료 코드 0 |
| `uv run ruff format --check .` | 종료 코드 0(파일 84개) |
| `uv run mypy .` | 종료 코드 0(파일 31개) |
| `uv run pytest tests/`(OS 임시 폴더 아래 새 `--basetemp`) | **2318 passed, 5 skipped, 32 xfailed**. 스냅숏 279개 통과(갱신 없음). 두 번째 검토 뒤에는 180자 길이의 `--basetemp`로도 한 번 돌려 통과를 확인했다 |
| `uv build` | sdist와 wheel 생성. wheel 최상위 모듈 21개가 `pyproject.toml` 목록과 같다(`input_safety.py` 포함) |
| Word 스키마 검증(로컬 Windows, 검증기 v2.1.6, 압축 파일 해시 확인 포함) | 문서 13개 오류 0건, 종료 코드 0 |
| 명령줄 확인 13가지(위 경로 제한·입력 한도 표) | 모두 기대와 같음 |

기준선(1316 passed, 4 skipped)보다 통과 1002개, 생략 1개가 늘었고 예상 실패 32개가 생겼다. 생략 5건은 모두 이 Windows 계정이 할 수 없는 일이다: 심볼릭 링크 만들기 4건(그중 1건이 이번에 추가한 경로 제한의 "링크" 경우), POSIX 권한 비트 1건. 예상 실패 32건은 "속성 테스트" 절의 표에 있다.

## 최종 검토의 지적과 처리

2026-11-11에 전체 변경(시작 커밋 `24bab05`부터)을 요구사항에 대조하는 독립 검토를 한 번 돌렸다. 지적 8건과 권고 2건이 나왔고 모두 처리했다. 처리 뒤에도 샘플 산출물은 바뀌지 않았다(스냅숏 279개 갱신 없이 통과).

| 지적 | 처리 |
|---|---|
| 감사 도구가 python-docx로 먼저 여는 파트(본문, 번호, 스타일 등)는 강화된 XML 읽기를 거치지 않았다 | 열기 전에 패키지의 XML 파트를 확인하도록 고침(파트를 고르는 기준은 두 번째 검토 뒤에 다시 고쳤다) |
| 플러그인 실행기가 입력 한도를 적용하기 전에 원본 전체를 읽어 해시를 계산했다 | 한도 확인을 해시보다 앞에 두고 읽는 양을 한도로 묶음 |
| 앞머리 설정 값이 틀린 문서와 YAML 이스케이프 제어 문자가 일반 예외로 끝났다 | 문서 설정 검사 경계에서 전용 거부 예외로 바꾸고, 속성 테스트의 생성 범위를 넓힘 |
| 로고 경로가 미리 해석되어 퍼센트 디코딩 후보 판정이 본문 이미지와 달랐다 | 로고도 적힌 형태 그대로 공용 판정 함수에 넘김 |
| 로컬 검증 흐름에서 해시 확인 없이 검증기를 실행할 수 있었다 | `validate`가 압축 파일을 받아 같은 실행 안에서 확인한 뒤에만 실행 |
| 예외 항목이 같은 위치의 다른 오류까지 덮을 수 있었다 | 오류 종류와 설명 문장까지 일치해야 덮도록 고침 |
| 코드 블록 속성이 통째로 예상 실패여서 생성 입력이 실제로 검사되지 않았다 | 동작하는 범위는 실제 검사로, 알려진 빈틈만 예상 실패로 분리 |
| Word 쪽 검토가 축소본 확인에 그쳤다 | 32쪽을 원본 크기로 검토하고 매니페스트 해시와 함께 기록("Word 쪽 확인" 절) |
| (권고) 스냅숏 도우미가 경로처럼 보이는 모든 글자를 경로로 바꿨다 | 경로 값인 항목만 바꾸도록 고침 |
| (권고) 저작권 연도 정규화가 기록에서 빠졌다 | "스냅숏 테스트" 절에 적음 |

검토가 직접 보지 못했다고 밝힌 것: 저장소 밖에 있는 Word 쪽 이미지와 매니페스트, 그리고 GitHub의 실제 CI 실행. 이 둘은 위 "Word 쪽 확인"과 "CI" 절의 기록이 근거다.

## 두 번째 검토의 지적과 처리

2026-11-11에 위 수정 커밋(`c6d5e84`)만을 대상으로 독립 검토를 한 번 더 돌렸다. 첫 검토의 지적 가운데 여섯 건은 닫혔다고 확인했고, 두 건(감사 도구, 문서 설정 거부)에서 남은 빈틈과 수정이 새로 만든 문제를 합쳐 다섯 건을 지적했다. 모두 처리했다. 처리 뒤에도 샘플 산출물은 바뀌지 않았다(스냅숏 279개 갱신 없이 통과).

| 지적 | 처리 |
|---|---|
| 감사 도구가 XML 파트를 파일 이름으로 골랐다. 본문 파트의 이름을 `document.bin`이나 `document.XML`로 바꾸고 내용 종류 목록과 연결 정보를 맞추면 검사를 피했고, 반대로 이름만 `.xml`인 이진 첨부는 잘못 거부됐다 | 내용 종류 목록(`[Content_Types].xml`)을 먼저 안전하게 읽고, 거기에 XML 종류로 선언된 파트와 모든 연결 정보 파트를 검사한다. 이름과 확장자는 python-docx와 같이 대소문자를 가리지 않고 맞춘다. 내용 종류 목록이 없거나 읽을 수 없으면 문서를 열지 않고 문제로 보고한다 |
| 앞머리의 틀린 날짜(예: `date: 2026-99-01`)가 YAML 문법 오류가 아니라 값 변환 오류여서 일반 `ValueError`로 끝났다 | 선언된 앞머리에서는 `Invalid YAML frontmatter: `로 시작하는 `StrictValidationError`로 거부한다. 선언되지 않은 블록은 문법 오류 때와 같은 기존 처리(단순한 `키: 값` 블록은 글자 값으로, 그 밖은 본문으로)를 따른다 |
| 머리글·표지·제목 영역에 쓰이는 값의 종류가 틀리면(예: `sender: {organization: 12}`) 일반 `TypeError`로 끝났고, 직접 만든 모델의 각주에 든 NUL은 XML 쓰기 단계의 일반 `ValueError`로 끝났다 | 렌더러가 글자로 쓰는 메타데이터 값을 렌더 전에 검사해 `StrictValidationError`로 거부한다(미리 검사하는 이 방식은 확인 검토 뒤에 바꿨다. 아래 절 참조). 지금 변환되는 숫자·날짜 값은 그대로 변환된다(앞머리 56가지 형태를 엄격·비엄격 모드에서 수정 전후로 비교해 확인). 엄격 모드는 모델 전체의 문자열에서 XML에 넣을 수 없는 문자를 저장 전에 거부한다 |
| 로고 경로를 미리 해석하지 않게 바꾸면서 Windows의 드라이브 상대 경로(`Z:logo.png`)가 절대 경로가 되지 못해 실패했고, 운영자가 지정한 기관 파일의 경로 오류가 문서 거부로 표시됐다 | 링크를 따라가거나 `..`를 접지 않는 방식으로 절대 경로를 만든다. 운영자가 지정한 기관 파일의 경로 오류는 일반 오류로, 앞머리 로고의 오류는 문서 거부로 구분한다 |
| 테마 테스트 다섯 개가 임시 폴더 경로 길이에 따라 실패했다(앞머리에 선언 키가 없어 120자를 넘는 값이 있으면 본문으로 처리됨) | 테스트 입력에 `profile:`을 넣고 테마 설정이 실제로 읽혔는지 확인한다. 180자 경로와 짧은 경로 양쪽에서 전체 테스트가 통과한다 |
| (권고) 예상 실패 표의 건수 한 줄이 틀렸다 | 7건으로 고침(합계 32건은 그대로) |
| (권고) 변경 기록의 "원래 오류가 원인으로 이어진다"가 모든 거부에 해당하는 것처럼 읽혔다 | 원인이 이어지는 경우와 직접 던지는 경우를 나눠 적음 |

**처리 결과를 직접 확인하다가 찾아 고친 것.** 위 셋째 지적의 수정이 엄격하지 않은 모드까지 모델 전체를 검사해, 본문 문단·목록·인용·코드·표 칸에 제어 문자가 든 문서를 거부했다. 수정 전에는 그 요소만 오류 표시로 바뀌고 문서가 저장됐다. 엄격하지 않은 모드는 종전대로 요소 단위로 넘어가게 되돌렸고, 요소 단위 보호 밖에서 쓰이는 글자(각주, 보고서 프로필의 제목)가 XML 쓰기에서 실패하면 그때 같은 진단 문장의 `StrictValidationError`로 바꿔 던진다(종전에는 일반 `ValueError`). 이때의 판별 방식은 확인 검토 뒤에 바꿨다(아래 절 참조). `tests/test_document_rejections.py`의 `test_non_strict_*` 테스트가 이 동작을 고정한다.

**엄격 모드 이분성 속성이 지금 생성하는 입력.** 범위가 정해진 Markdown 블록과 인라인 구문, 닫히지 않은 구문, 임의 유니코드 문자열, 깊이가 정해진 임의 YAML 값, 일부 중첩 설정, 값 변환에 실패하는 YAML 스칼라(틀린 날짜·시각 등), 렌더러가 글자로 쓰는 메타데이터 키의 임의 값(`plain`과 `ib-report`, 표지 유무), 직접 만든 모델의 문단·제목·글자 조각·표 칸·목록·각주·메타데이터 문자열.

**속성이 생성하지 않아 확인되지 않은 입력.** 문서가 지정한 기관·테마 파일의 내용, 이미지 파일의 바이트와 폴더 구조, 임의의 YAML 태그·별칭, 중첩 설정의 모든 조합, 임의로 틀린 모양의 모델 객체. 쓰이지 않는 메타데이터 키와 임의의 파이썬 모델 객체는 전면적인 형식 검사를 받지 않는다. 이 가운데 일부는 개별 테스트로만 확인했다. 따라서 "엄격 변환은 저장하거나 문서화된 거부로 끝난다"는 위 생성 범위에 대해서만 확인된 것이다.

두 번째 검토가 직접 보지 못했다고 밝힌 것: Ubuntu와 Python 3.13에서의 실행, GitHub의 실제 CI, 실제 검증기 압축 파일, 저장소 밖의 Word 쪽 이미지. 앞의 셋은 "CI" 절과 "최종 상태의 필수 검증" 절의 기록이 근거다.

## 확인 검토의 지적과 처리

2026-11-11에 두 번째 검토의 수정 커밋(`611a0de`)을 대상으로 확인 검토를 돌렸다. 다섯 건 가운데 로고 경로와 테마 테스트 두 건은 닫혔다고 확인했고, 나머지 세 건에서 빈틈 네 가지를 지적했다. 모두 처리했다. 처리 뒤에도 샘플 산출물은 바뀌지 않았다(스냅숏 279개 갱신 없이 통과).

| 지적 | 처리 |
|---|---|
| 감사 도구가 파트의 확장자를 python-docx와 다르게 읽었다. 확장자가 없는 이름(`word/document`)이나 점으로 시작하는 이름(`word/.document`)에 빈 확장자 기본값으로 XML 종류를 선언하면 검사를 피했다 | 확장자를 python-docx와 같은 규칙으로 읽는다(`test_extensionless_xml_main_part_is_checked`: 고치기 전의 python-docx가 그 파트를 읽어 엔티티를 펼치는 것을 테스트 안에서 먼저 보인다) |
| YAML 값 변환 오류 가운데 `ValueError`가 아닌 것들이 일반 예외로 끝났다(`!!timestamp x` → `AttributeError`, `!!bool x` → `KeyError`, `!!int ""` → `IndexError`) | YAML을 읽는 곳을 `input_safety.load_yaml` 하나로 모았다. 이 종류의 오류와 지나치게 깊은 중첩을 `ValueError`로 바꿔, 앞머리·기관 파일·테마 파일·차트 블록의 기존 처리가 그대로 받는다 |
| 메타데이터를 미리 검사하는 방식이 지금까지 변환되던 문서를 거부했다. 예: `sender: {organization: [가, 나]}`(머리글에 이어 붙여 나오던 값), 쓰이지 않는 부제목의 `None` | 미리 검사하는 함수를 없앴다. 표지·제목 영역·맺음·고지문·머리글을 만드는 각 호출에서 값의 종류가 틀려 나는 `TypeError`·`AttributeError`만 그 자리에서 `StrictValidationError`로 바꾼다. 쓰기 전에는 아무것도 검사하지 않으므로 지금 변환되는 값은 그대로 변환된다 |
| 엄격하지 않은 모드의 예외 처리가 lxml의 오류 문장에 의존했다. 짝 없는 서로게이트는 놓쳤고, 운영자 인자의 오류 문장에 같은 글자가 들어 있으면 문서 거부로 바뀌었다 | 문장을 보지 않는다. 운영자 인자 검사가 모두 끝나고 문서 조립이 시작된 뒤에 `ValueError`가 빠져나왔고, 모델에 XML에 넣을 수 없는 문자가 실제로 있을 때만 그 문자를 이름 붙인 `StrictValidationError`로 바꾼다. 원래 오류가 원인으로 이어진다 |

**메타데이터 비교.** 프로필 7종 × 표지 유무 × 엄격·비엄격 × 메타데이터 항목 22개(제목·부제목·회사 등 6개, `date`·`sender`·`attendees` 등 11개, `sender`의 하위 5개) × 값 10가지(`None`, 정수, 실수, 참·거짓, 문자열 목록, 정수 목록, 사전, 날짜, 빈 문자열, 문자열)와 빈 제목에 다른 종류의 부제목을 둔 경우까지 6,468가지를 직접 만든 모델로 렌더해 첫 검토 직후 커밋(`75efb5c`)과 비교했다. `75efb5c`에서 저장되던 3,138가지는 모두 그대로 저장된다. 일반 예외로 끝나던 1,406가지(`AttributeError` 1,258, `TypeError` 120, `ValueError` 28)는 모두 `StrictValidationError`가 됐고, 일반 예외로 끝나는 경우는 남지 않았다. 이 비교는 저장 여부와 예외 종류만 본 것이고, 저장된 문서의 내용 비교는 스냅숏 테스트가 맡는다.

**직접 확인하다가 찾아 고친 것.**

- 차트 블록의 YAML에 변환할 수 없는 값(틀린 날짜 등)이 있으면 파서가 일반 `ValueError`로 끝났다. YAML 문법 오류와 같이 그 차트의 오류로 기록한다(`test_chart_fence_with_unconstructible_yaml_is_a_chart_error`).
- 제목 영역이 직접 검사하는 참석자 목록 오류(`attendees must be ...`)가 보고서 프로필에서 일반 `ValueError`로 끝났다. 같은 문장의 `StrictValidationError`로 바꿨다.

**남는 한계.** 엄격하지 않은 모드에서, 모델에 XML에 넣을 수 없는 문자가 있는 문서를 조립하다가 그 문자와 무관한 `ValueError`가 빠져나오면 그 문자를 이름 붙인 거부로 보고된다(원래 오류는 원인으로 남는다). 표지·제목 영역·머리글을 만드는 코드 자체의 결함으로 `TypeError`·`AttributeError`가 나도 메타데이터 종류 오류로 보고된다. 두 경우 모두 원인 사슬에서 원래 오류를 볼 수 있다.

확인 검토가 직접 보지 못했다고 밝힌 것: 전체 테스트 실행(대상 테스트 403개와 스냅숏 테스트만 돌림), 실제 Word 쪽 렌더, GitHub의 실제 CI.

## 마지막 확인 검토

2026-11-11에 확인 검토의 수정 커밋(`d26ad0b`)을 대상으로 한 번 더 검토를 돌렸다. 결과는 승인이다: 지적 네 건이 모두 닫혔고 막아야 할 문제는 없다고 했다. 검토가 따로 해 본 것: 감사 도구에 대한 파일 이름·내용 종류 조합 28가지, YAML 읽기 결과가 `safe_load`와 같은지(C 확장 유무 양쪽), 메타데이터 784가지(이전에 저장되던 문서는 모두 XML까지 같게 저장됨), 렌더러를 성공·실패 뒤에 다시 써도 상태가 남지 않는지. 권고 두 건은 반영했다.

| 권고 | 처리 |
|---|---|
| 이번 변경 이전부터 있던 일반 예외 두 가지: 4,300자리를 넘는 각주 번호(`[^999…]`)와 앞머리의 같은 길이 정수(`title: 0xfff…`). 파이썬이 그 길이의 정수를 글자로, 글자를 정수로 바꾸기를 거부한다 | 각주·참조 번호는 `Number is too long to read (...)`로, YAML 정수는 `YAML integer is too long to use`로 거부한다(`test_overlong_*`). 파이썬이 바꿀 수 있는 긴 번호는 그대로다 |
| 기록이 "YAML을 읽는 곳을 하나로 모았다"고 했는데 도식 블록은 `yaml.safe_load`를 그대로 썼다(그 블록의 기존 처리가 오류를 모두 받아 코드 블록으로 남기므로 문제는 아니었다) | 도식 블록도 `load_yaml`로 읽게 했다. 변환 경로에 남은 `yaml.safe_load` 호출은 없다. 별도 도구인 변환본 정리기(`converted_md_cleaner.py`)는 대상이 아니어서 그대로다 |

이 마지막 반영분은 독립 검토를 다시 받지 않았다. 근거는 추가한 테스트와 전체 테스트, CI다.

마지막 확인 검토가 직접 보지 못했다고 밝힌 것: 전체 테스트 실행과 6,468가지 비교의 재실행(대상 테스트 818개와 스냅숏만 돌림), GitHub의 실제 CI, Word 쪽 모양.

## 하지 않은 검증

- **Word 시각 검토는 위 "Word 쪽 확인"에 적은 범위뿐이다.** 수정 전후 쪽 이미지의 일치와 AI 에이전트의 쪽 모양 검토까지 했다. 소유자가 직접 보는 검토, 각 쪽의 내용이 업무상 맞는지에 대한 검토는 하지 않았다.
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
- 속성 테스트가 남긴 예상 실패 32건은 위 표의 일곱 줄이다.
