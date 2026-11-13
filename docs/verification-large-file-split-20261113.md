# 큰 파일 분할 1단계(옮기기) 검증 (2026-11-13 착수)

이 문서는 `ib_renderer.py`와 `md_parser.py`를 기능별 내부 파일 12개로 나누는 작업의 검증 기록이다. 이 작업은 코드를 옮기기만 하며, 만들어지는 Word 문서·명령줄 동작·오류와 로그 문장은 바뀌지 않아야 한다. 아래에는 옮기기 전의 기준 값과, 옮긴 뒤 "달라진 것이 없다"를 판정할 장치를 적는다. 한 검증과 하지 않은 검증을 구분해 적으며, 구조 검사 통과는 시각 검토 통과를 뜻하지 않는다.

## 환경

Windows 11(10.0.26300), Python 3.12.12, uv 0.11.28, ruff 0.15.0, mypy 1.14.1, pytest 9.0.2, python-docx 1.2.0, lxml 6.0.2, syrupy 6.1.1, hypothesis 6.170.15, Microsoft Word 16.0(빌드 16.0.20527). 시작 커밋은 `main`의 `5975cff`(PR #23 병합 커밋, `5975cffa7f267963e7e1b694ba612397def74fc4`)이고, 작업 브랜치는 이 커밋에서 갈라졌다.

## 1. 기준 값 (시작 커밋, 옮기기 전)

| 검증 | 명령 | 결과 | 종료 코드 |
|---|---|---|---|
| 린트 | `uv run ruff check .` | 위반 0건 | 0 |
| 서식 | `uv run ruff format --check .` | 86개 파일 모두 서식 일치 | 0 |
| 타입 | `uv run mypy .` | 31개 소스 파일, 문제 0건 | 0 |
| 전체 테스트 | `uv run pytest tests/ --basetemp <OS 임시 폴더 아래 새 절대 경로>` | **2575 passed, 5 skipped, 1 xfailed** | 0 |
| 스냅숏 | `uv run pytest tests/test_sample_snapshots.py` | **279 snapshots passed**(테스트 26개), 갱신 없음 | 0 |
| Word 스키마 검증 | `scripts/word_schema_validation.md`의 절차(`validate`, Windows용 압축 파일) | 문서 **13개**, 오류 **0건**(`ok=true`, `uncovered_error_count=0`) | 0 |
| 배포 묶음 | `uv build`(저장소 밖 새 폴더로) | wheel의 최상위 모듈 **21개**, wheel `only-include` 목록의 `.py` 21개와 같음 | 0 |

생략 5개와 예상 실패 1개의 내역(시작 시점 그대로 유지돼야 한다):

- 생략: 이 계정에서 심볼릭 링크를 만들 수 없어 생략되는 테스트 4개(`test_asset_policy.py`, `test_codex_bundle.py`, `test_memo_rendering.py`, `test_plugin_bundle.py`), Windows에서 POSIX 권한 비트 검사를 생략하는 테스트 1개(`test_cli_hardening.py`).
- 예상 실패: `test_markdown_properties.py::test_fenced_code_has_no_semantic_emphasis`(코드 블록 안 트리 값 강조, `AGENTS.md`에 적힌 엄격 `xfail`).

Word 스키마 검증 도구는 `scripts/word_schema_validator_assets.json`에 고정된 주소에서 내려받은 OOXML-Validator v2.1.6 `win-x64.zip`이다. 내려받은 파일의 SHA-256(`0f22c0da…cf1acb`)이 고정 값과 같았고, 스크립트가 실행 중에 다시 확인했다.

### 변환 시간과 가져오기 시간

같은 기기에서 측정마다 새 파이썬 프로세스를 띄워 세 번씩 재고 중앙값을 적는다. 변환 시간은 `scripts.sample_corpus.render_corpus`로 샘플 문서 13개를 엄격 모드로 새 임시 폴더에 변환하는 데 걸린 시간이다(인터프리터 시작 시간 제외). 가져오기 시간은 아무것도 가져오지 않은 새 프로세스에서 `import <모듈>` 한 줄에 걸린 시간이다. 측정 스크립트는 저장소 밖에 둔 일회용 도구다.

| 측정 | 세 번의 값(초) | 중앙값(초) |
|---|---|---|
| 샘플 13개 변환 | 2.2677, 2.2912, 2.3130 | **2.2912** |
| `import ib_renderer` | 0.4462, 0.4138, 0.4337 | **0.4337** |
| `import md_parser` | 0.3199, 0.3153, 0.3825 | **0.3199** |

옮긴 뒤 같은 방법으로 다시 재서, 20%를 넘게 느려졌으면 원인을 확인해 보고한다.

## 2. 대표 문서의 Word 쪽 기준

- **문서.** `samples/profiles/ib-report.md`. 표지, 목차, 머리말·쪽 번호, 제목, 본문, 금융표, 출처, 면책 고지를 모두 거치므로 옮겨 가는 렌더러 대부분과 남는 렌더러(표지·고지문)를 함께 지난다.
- **절차.** 시작 커밋의 코드로 `uv run md-to-word samples/profiles/ib-report.md <저장소 밖 새 폴더>/ib-report.docx --strict`로 변환하고, `scripts/word_visual_qa.ps1`을 새 빈 출력 폴더에 실행해 필드를 갱신하고 모든 쪽을 PNG로 만들었다. DOCX, 쪽 이미지, 매니페스트는 저장소 밖 OS 임시 폴더 아래 새 폴더에 보관했다.
- **결과.** 4쪽(1쪽 표지, 2쪽 목차, 3쪽 본문과 금융표, 4쪽 면책 고지). 매니페스트: 원본 DOCX SHA-256 `9822b2af…41fe2a`(실행 전후 같음), 쪽 수 4, Word 16.0 빌드 16.0.20527.
- **검토.** 네 쪽을 모두 열어 보았다. 깨진 배치, 빈 쪽, 잘린 표는 없다. 매니페스트의 `visualReview`는 스크립트가 쓰는 `pending` 그대로이며, 이 항목이 검토 기록이다. 이 검토는 쪽 모양에 대한 것이고 샘플 문구나 수치의 내용이 맞는지에 대한 검토가 아니다.
- **옮긴 뒤.** 같은 문서를 같은 절차로 다시 만들어 쪽 이미지를 이 기준과 비교한다.

| 쪽 | PNG SHA-256 |
|---|---|
| 1 | `74B1FDB963E63FF910212F94E2788E732E6850D5EF144E5F5262FA755F39DEB2` |
| 2 | `CAB8697DD937998B7C272F8D8FCA61300B3BF6E930C44AF612F9D36DD7DCB771` |
| 3 | `5D71517FF07046ABF2CC296D64964B294A3B2C83165BA9E407918F3A8E4A53E4` |
| 4 | `38C9C33D5EC01DCA03DA88145DF221F7B5537C367401DC3C580F9AD0E856AA0D` |

## 3. 새 보호 테스트 `tests/test_module_layout.py` (1~3번 검사)

옮기기 전에 넣고, 옮기기 전 코드에서 통과함을 확인했다(테스트 7개). 4~8번 검사는 옮기면서 채운다.

1. **목록의 이름.** `ib_renderer`가 정의하는 24개와 이 파일을 거쳐 쓰이는 가져온 이름 6개(`STYLE`, `IBStyle`, `uuid4`, `Document`, `time`, `platform`), `md_parser`가 정의하는 29개와 문서 모델에서 다시 내보내는 이름 26개를 글자로 적어 고정했다. 모두 원래 모듈에서 가져올 수 있어야 한다. 가져온 이름 6개와 문서 모델 이름 26개는 원래 출처의 객체와 같은 객체여야 한다.
2. **저장소가 두 모듈을 거쳐 쓰는 이름.** 최상위 파이썬 파일과 `tests/`, `scripts/`, `tools/` 아래 파이썬 파일(87개)을 구문 분석해 네 가지 형태를 모두 찾는다: `from ib_renderer import X`, `ib_renderer.X`(그 파일이 `import ib_renderer`로 묶은 이름만), 글자로 적은 `"ib_renderer.X"`(파일 이름 `….py`는 제외), 모듈과 이름을 따로 넘기는 호출(`setattr(ib_renderer, "X", …)`, `monkeypatch.setattr(ib_renderer, "X", …)` 등). `md_parser`도 같다. 찾은 이름이 모두 그 모듈에 있어야 한다. 읽지 않는 폴더: `.venv`, `.git`, `reports`, `temp`, `tmp`, `private`, 테스트 스냅숏 폴더, 저장소 루트의 읽을 수 없는 임시 폴더(`Users*AppDataLocalTemp*`). 검색이 실제로 동작하는지 확인하려고, 네 형태 각각에서 지금 있는 참조가 찾아지는지도 검사한다.
   - 시작 시점의 결과: 참조 254건(가져오기 246, 글자 4, 호출 2, 속성 2). `ib_renderer`를 거쳐 쓰는 이름 18개, `md_parser`를 거쳐 쓰는 이름 24개이며, 모두 요구사항 6절의 목록 안에 있다. 목록 밖의 새 이름은 없었다.
3. **바꿔치기되는 이름.** 위 검색에서 테스트가 원래 모듈에서 바꿔치기하는 모듈 수준 이름은 `uuid4`, `Document`, `time` 세 개다(`platform`은 모듈 객체의 함수를 바꿔치기하므로 해당하지 않는다). 이 이름을 읽는 최상위 정의는 원래 파일에 있어야 하고, 같은 원래 파일에서 갈라져 나온 파일(`ib_render_*.py`, `md_parse_*.py`)에는 없어야 한다. 시작 시점: `uuid4`·`Document`는 `IBDocumentRenderer`가, `time`은 `CoverRenderer`·`DisclaimerRenderer`가 읽는다. 요구사항 4절에서 남기기로 한 정의와 같다.

**검사가 실제로 실패하는지 확인.** 저장소 밖에서 두 가지를 일부러 깨 보았다. `ib_renderer.time`을 지우면 2번 검사가 실패했다. `CoverRenderer`를 갈라져 나온 파일로 옮긴 사본을 만들면 3번 검사가 그 파일의 `CoverRenderer`를 지목하며 실패했다.

**새 테스트를 더한 뒤.** 아래 "4절"의 결과를 본다.

## 4. 새 테스트를 더한 뒤 (T1 커밋)

| 검증 | 결과 | 종료 코드 |
|---|---|---|
| `uv run ruff check .` | 위반 0건 | 0 |
| `uv run ruff format --check .` | 87개 파일 서식 일치 | 0 |
| `uv run mypy .` | 문제 0건 | 0 |
| `uv run pytest tests/` | **2582 passed, 5 skipped, 1 xfailed** = 기준 2575 + 새 테스트 7. 생략·예상 실패 수는 같음 | 0 |

## 5. 본문 동일성 비교 도구

저장소 밖에 둔 일회용 도구다. 시작 커밋의 `ib_renderer.py`, `md_parser.py`를 `git show`로 읽어 최상위 정의(클래스, 함수, 값 대입)를 이름별로 모으고, 작업 사본에서 요구사항 4절이 정한 파일(그 파일이 아직 없으면 원래 파일)에서 같은 이름의 정의를 찾아 두 가지를 비교한다.

- **구문 트리.** 위치 정보를 뺀 구문 트리가 같아야 한다. 5절 규칙 5로 함수 안에 더한 가져오기 문이 있으면 그 줄을 뺀 나머지를 비교하고, 더한 줄을 따로 출력한다.
- **원문 줄.** 정의가 차지하는 줄(장식자부터 마지막 줄까지)이 글자 그대로 같아야 한다. 구문 트리에 없는 주석과 빈 줄의 변화도 여기서 잡힌다. 정의 사이의 구획 주석은 5절 규칙 4에 따라 새 파일에서 다시 배치될 수 있으므로, 새 구성 어디에도 없는 줄을 정보로 출력한다.

그 밖에 새 파일의 `logger`가 `logging.getLogger("ib_renderer")` 또는 `logging.getLogger("md_parser")` 형태인지, 새 파일에 정해지지 않은 최상위 문장이 없는지, 옮긴 이름이 원래 파일에 남아 있지 않은지 확인한다. 찾은 정의가 53개가 아니거나 하나라도 다르면 실패다.

- **시작 시점의 결과.** 정의 **53개**(`ib_renderer.py` 24, `md_parser.py` 29)를 모두 찾았고, 53개 모두 같음. 정의 사이 구획 주석은 각각 27줄, 24줄이며 모두 그대로 있음. 종료 코드 0.
- **도구 자체 확인.** 시작 커밋의 두 파일로 만든 임시 저장소에서 확인했다: 글꼴 클래스 두 개를 새 파일로 옮기면 통과, 함수 안에 가져오기 문을 더하면 "더한 가져오기 문"으로 표시하고 통과, 메서드 이름 한 글자를 바꾸면 구문 트리 차이로 실패, 주석 한 줄을 더하면 원문 줄 차이로 실패, `logger`를 `__name__`으로 두면 실패, 한 클래스만 옮기고 같은 파일로 갈 다른 클래스를 남기면 실패, `--layout split`에서 새 파일이 없으면 실패.

## 6. 하지 않은 것

- macOS·Linux에서의 측정. 이 기록의 시간 값은 위 Windows 기기에서만 쟀다. CI의 Ubuntu 결과는 PR 단계에서 확인한다.
- 배포 묶음(플러그인·Codex 묶음, 휴대용 변환기)의 실제 재생성과 게시. 범위 밖이다.
