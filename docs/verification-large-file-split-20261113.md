# 큰 파일 분할 1단계(옮기기) 검증 (2026-11-13 착수)

이 문서는 `ib_renderer.py`와 `md_parser.py`를 기능별 내부 파일 12개로 나누는 작업의 검증 기록이다. 이 작업은 코드를 옮기기만 하며, 만들어지는 Word 문서·명령줄 동작·오류와 로그 문장은 바뀌지 않아야 한다. 아래에는 옮기기 전의 기준 값과, 옮긴 뒤 "달라진 것이 없다"를 판정할 장치를 적는다. 한 검증과 하지 않은 검증을 구분해 적으며, 구조 검사 통과는 시각 검토 통과를 뜻하지 않는다.

## 요약: 무엇을 옮겼고, 왜 달라진 것이 없다고 말할 수 있는가

**무엇을 옮겼나.** 두 큰 파일에 들어 있던 최상위 정의(클래스·함수·값) 53개 가운데 43개를 새 파일 12개로 옮겼다. 렌더러 쪽은 `ib_renderer.py`에서 19개를 `ib_render_*.py` 7개로, 파서 쪽은 `md_parser.py`에서 24개를 `md_parse_*.py` 5개로 옮겼다. 나머지 10개(각 파일의 중심 클래스와 몇 가지)는 제자리에 남았다. `ib_renderer.py`는 4,758줄에서 1,365줄로, `md_parser.py`는 3,979줄에서 1,561줄로 줄었다. 이 엔진을 쓰는 쪽은 종전처럼 `ib_renderer`와 `md_parser`에서 가져온다. 두 파일이 옮긴 이름을 모두 다시 내보내기 때문이다.

**달라진 것이 없다고 말할 수 있는 근거.** 서로 독립된 여섯 가지 확인이 모두 같은 결론이다.

1. **옮긴 코드가 글자 그대로다.** 옮긴 정의마다 시작 커밋의 원문과 새 파일의 원문을 줄 단위로, 그리고 파이썬 구문 구조로 비교했다. 53개 모두 같고, 다른 것은 0개다(5절, 8절).
2. **만들어지는 Word 문서가 같다.** 샘플 13개 문서의 파싱 결과와 DOCX 안의 모든 XML 파트를 고정한 스냅숏 279개가, 고치지 않은 채로 모든 커밋에서 통과했다.
3. **Word에서 보이는 쪽이 같다.** 대표 문서 하나를 옮기기 전후에 만들어 비교했다. DOCX 안의 파트 20개가 바이트 단위로 같고, Word로 그린 4쪽의 이미지도 바이트 단위로 같다(8절).
4. **기존 테스트가 고치지 않은 채로 모두 통과한다.** 시작 시점에 통과하던 2,575개가 그대로 통과하고, 생략 5개와 예상 실패 1개도 그대로다. 테스트 수가 늘어난 것은 이번에 더한 보호 테스트 61개뿐이다.
5. **바깥에서 가져오는 이름이 그대로다.** 저장소의 모든 파이썬 파일이 두 파일을 거쳐 쓰는 이름을 찾아, 그 이름이 모두 두 파일에 있고 새 파일의 것과 같은 객체임을 테스트로 고정했다(3절, 6절).
6. **배포와 형식 검사가 그대로다.** Word 스키마 검증(13개 문서 오류 0건), 린트·서식·타입 검사, 배포 묶음 테스트가 시작 시점과 같은 결과다. 변환·가져오기 시간도 2% 안에서 같다.

**받아들인 차이.** 옮긴 클래스와 함수가 속한 모듈 이름(`__module__`)과, 그것이 드러나는 곳(예외 추적 기록의 파일 이름)은 새 파일 이름으로 바뀐다. 저장소에 이 값에 기대는 코드는 없다. 옮긴 줄의 `git blame` 기록은 옮긴 커밋을 가리키며, 이전 기록은 `git blame -C -C`로 볼 수 있다.

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

## 6. 렌더러 옮기기 (T2)

`ib_renderer.py`의 정의 19개를 요구사항 4절 표의 새 파일 7개로, 표의 순서대로 한 파일씩 옮겼다. 새 파일 하나에 커밋 하나다. `IBDocumentRenderer`, `_metadata_text`, `CoverRenderer`, `DisclaimerRenderer`, `logger`는 `ib_renderer.py`에 남았다. `ib_renderer.py`는 4,758줄에서 1,365줄이 됐다.

### 옮긴 방법

파일마다 같은 절차를 따랐다.

1. 옮길 정의의 줄(장식자부터 마지막 줄까지)을 글자 그대로 잘라 새 파일에 원래 순서대로 붙였다.
2. 새 파일 맨 위에는 모듈 설명 한 줄과, 옮긴 코드가 실제로 쓰는 이름의 가져오기 문만 두었다. 가져오기 문은 `ib_renderer.py`에 있던 것과 같은 모듈, 같은 별칭이다. 앞서 옮긴 새 파일의 이름은 그 파일에서 직접 가져왔다. `ib_renderer`에서 가져온 새 파일은 없다.
3. 로그를 쓰는 새 파일 넷(`ib_render_fonts`, `ib_render_media`, `ib_render_text`, `ib_render_tables`)에는 `logger = logging.getLogger("ib_renderer")`를 두었다. 나머지 셋(`ib_render_document_style`, `ib_render_footnotes`, `ib_render_blocks`)은 로그를 쓰지 않는다.
4. `ib_renderer.py`에 `from 새_파일 import 이름 as 이름`을 더하고, 쓰지 않게 된 가져오기 문을 지웠다. `platform`은 `import platform as platform`으로 남겼다.
5. `ruff`로 가져오기 문만 정렬하고 안 쓰는 것을 지운 뒤 `ruff format`을 돌렸다. 바뀐 줄은 가져오기 문과 정의 사이의 빈 줄뿐이며, 옮긴 본문이 다시 정렬된 곳은 없다(본문 동일성 비교로 확인).
6. `pyproject.toml`의 목록 세 곳과 `tests/test_module_layout.py`의 옮긴 파일 목록에 그 파일을 더했다.

**구획 주석.** 한 정의만 이끄는 구획 주석(`DOCUMENT SIGNATURE`, `FONT STYLER`, `DOCUMENT STYLER`, `TABLE STYLER`, `SAVING`, `LATEX RENDERER (NEW)`)은 그 정의와 함께 옮겼다. 남는 코드를 이끄는 구획 주석(`STYLE CONFIGURATION`, `ELEMENT RENDERERS`, `DOCUMENT RENDERER (ORCHESTRATOR)`)은 `ib_renderer.py`에 그대로 두었다. 시작 커밋의 구획 주석 27줄은 모두 새 구성 안에 있다.

**새 파일 사이의 가져오기.** 각 파일의 모듈 수준 `import` 문 기준이다. 글꼴 파일(`ib_render_fonts`)은 다른 새 파일을 가져오지 않는다. 문서 스타일·각주·그림 파일은 글꼴 파일을 가져온다. 글자 파일(`ib_render_text`)은 글꼴·각주·그림 파일을 가져온다. 블록 파일과 표 파일은 글꼴·문서 스타일·글자 파일을 가져온다. 서로를 가져오는 쌍은 없다.

### 요구사항 4절 표와 달라진 것

없다. 묶음 변경, 5절 규칙 5로 더한 줄(함수 안 가져오기 문), 표와 달리 남긴 정의가 모두 없다. 기존 테스트가 원래 파일에서 바꿔치기하는 모듈 수준 이름(`uuid4`, `Document`, `time`)을 읽는 정의는 4절대로 `ib_renderer.py`에 남았고, 그 밖에 새로 드러난 경우는 없다(새 테스트 3번 검사).

### 보호 테스트 4~8번

`tests/test_module_layout.py`에 요구사항 7절의 4~8번을 더했다. 옮긴 파일 목록은 파일을 하나 옮길 때마다 늘렸다.

4. 옮긴 정의가 정해진 새 파일에 있고, 원래 파일에는 없으며, `ib_renderer`에서 가져온 것과 같은 객체다.
5. wheel 목록과 sdist 목록의 `.py` 항목이 서로 같고 최상위 파이썬 파일과 같다. 새 모듈이 린트의 자사 모듈 목록에 있다. 디스크의 `ib_render_*.py`, `md_parse_*.py`가 옮긴 파일 목록과 정확히 같다(12개 여부는 T3 끝에서 고정).
6. 새 파일은 모듈 수준에서 원래 모듈을 가져오지 않는다. `term_sheet.py`는 `ib_renderer`를 가져오지 않는다.
7. 새 파일과 원래 두 모듈을 하나씩 새 파이썬 프로세스에서 가져와도 오류가 없다.
8. 새 파일의 `logger`는 원래 모듈의 `logger`와 같은 객체다.

### 커밋별 결과

모든 커밋에서 `ruff check .`, `ruff format --check .`, `mypy .`, `pytest tests/`의 종료 코드가 0이고, 스냅숏 279개가 갱신 없이 통과했다. 기존 테스트 파일은 고치지 않았다. 통과 수 = 기준 2575 + `tests/test_module_layout.py`의 테스트 수.

| 커밋 | 새 파일 | 옮긴 정의 | 통과 | 생략 | 예상 실패 |
|---|---|---|---|---|---|
| `4e20751` | `ib_render_fonts.py` | `FontStyler`, `FontPolicy` | 2591 (2575 + 16) | 5 | 1 |
| `b3a1d15` | `ib_render_document_style.py` | `GeneratorSignatureWriter`, `DocumentStyler`, `TableStyler` | 2595 (+ 20) | 5 | 1 |
| `af36696` | `ib_render_footnotes.py` | `FootnoteRenderer`, `NativeFootnotesPart`, `rebase_local_links` | 2599 (+ 24) | 5 | 1 |
| `b4b5494` | `ib_render_media.py` | `ImageRenderer`, `ChartRenderer`, `LaTeXRenderer` | 2603 (+ 28) | 5 | 1 |
| `c4666c8` | `ib_render_text.py` | `TextRenderer` | 2607 (+ 32) | 5 | 1 |
| `f1f813e` | `ib_render_blocks.py` | `TOCRenderer`, `HeadingRenderer`, `ParagraphRenderer`, `ListRenderer`, `CalloutRenderer` | 2611 (+ 36) | 5 | 1 |
| `9fd2046` | `ib_render_tables.py` | `TableRenderer`, `_BLACK` | 2615 (+ 40) | 5 | 1 |

첫 커밋에서 `ruff format --check`가 새 테스트 파일 한 곳(긴 목록 내포 한 줄)을 다시 정렬하라고 해서, 커밋 전에 테스트 파일만 정렬했다. 옮긴 본문과는 관계없다.

### T2 끝의 확인

- **본문 동일성.** 시작 커밋과 비교한 정의 **53개**, 다른 것 **0개**. 옮긴 19개는 정해진 새 파일에, 남은 5개는 `ib_renderer.py`에 있다(파서 29개는 아직 `md_parser.py`). 더한 가져오기 문(규칙 5) 0개. 새 파일 넷의 `logger`가 규칙 2의 형태. 종료 코드 0.
- **배포 묶음.** `uv build` 종료 코드 0. wheel의 최상위 모듈 **28개**, wheel `only-include` 목록과 같음. sdist에도 새 파일 7개가 들어 있음.
- **아직 하지 않은 것.** Word 스키마 검증, 변환·가져오기 시간, 대표 문서의 Word 쪽 비교는 T4의 최종 검증에서 한다. 스냅숏 279개(DOCX 안의 모든 XML 파트)가 커밋마다 같았으므로 산출 문서는 이 시점까지 같다.

## 7. 파서 옮기기 (T3)

`md_parser.py`의 정의 24개를 요구사항 4절 표의 새 파일 5개로, 표의 순서대로 한 파일씩 옮겼다. 새 파일 하나에 커밋 하나다. `MarkdownParser`, `_read_with_encoding`, `_decode_bytes`, `parse_markdown_file`, `logger`와 문서 모델 재수출(`from document_model import 이름 as 이름`)은 `md_parser.py`에 그대로 남았다. `md_parser.py`는 3,979줄에서 1,561줄이 됐다.

### 옮긴 방법

T2와 같은 절차다. 새 파일은 문서 모델 이름을 `document_model`에서 직접 가져오고, `md_parser`를 가져오지 않는다. 로그를 쓰는 새 파일 넷(`md_parse_text`, `md_parse_frontmatter`, `md_parse_tables`, `md_parse_blocks`)에는 `logger = logging.getLogger("md_parser")`를 두었다. `md_parse_html_tables`는 로그를 쓰지 않는다. 로그 이름 `md_parser`에 기대는 기존 테스트(`tests/test_term_variables.py`)는 고치지 않고 통과한다.

**구획 주석.** 옮기는 정의만 이끄는 구획 주석(`LOCAL LINK TARGETS`, `HTML TABLES (converter output)`, `LaTeX PARSER (NEW v3)`, `BASE64 IMAGE PARSER (NEW v3)`)은 함께 옮겼다. `ENUMS`, `MARKDOWN PARSER (MAIN)`, `ENCODING UTILITIES (ENHANCED v3)`, `CONVENIENCE FUNCTION`은 `md_parser.py`에 남았다. 시작 커밋의 구획 주석 24줄은 모두 새 구성 안에 있다.

**정의 사이의 빈 줄.** 원래 파일에서 빈 줄만을 사이에 두고 이어져 있던 정의는 새 파일에서도 그 빈 줄을 그대로 둔다. 그래서 `_DRIVE_PATH_RE`, `_URI_DRIVE_PATH_RE`, `_URI_SCHEME_RE` 세 줄은 원래처럼 붙어 있고, `LOCAL LINK TARGETS` 구획 주석부터 `TextParser` 앞까지의 줄은 시작 커밋과 글자 그대로 같다. (처음 옮기는 도구는 정의 사이에 늘 빈 줄 두 개를 넣었다. 커밋 전에 이를 발견해 그 단계를 되돌리고 도구를 고쳐 다시 옮겼다. T2에서 함께 옮긴 이웃 정의들은 원래도 빈 줄 두 개로 떨어져 있었으므로 T2의 결과는 이 수정과 관계없다.)

**새 파일 사이의 가져오기.** 각 파일의 모듈 수준 `import` 문 기준이다. 글자 파일(`md_parse_text`)과 앞머리 파일(`md_parse_frontmatter`)은 다른 새 파일을 가져오지 않는다. 표·HTML 표·블록 파일은 글자 파일을 가져온다. 서로를 가져오는 쌍은 없다. 모듈을 읽는 시점의 참조는 `LaTeXParser`(클래스 본문에서 `TextParser`)와 `MarkdownParser`(클래스 본문에서 `FenceScanner`) 두 곳이며, 둘 다 가져오기 순서와 맞는다.

### 요구사항 4절 표와 달라진 것

없다. 묶음 변경, 5절 규칙 5로 더한 줄, 표와 달리 남긴 정의가 모두 없다. 기존 테스트가 `md_parser`에서 바꿔치기하는 모듈 수준 이름은 없다(새 테스트 3번 검사).

### 보호 테스트

옮긴 파일 목록에 파서 파일 5개를 차례로 더했다. 마지막 커밋에서 끝난 구성을 고정하는 검사를 하나 더했다: 새 파일은 정확히 12개이고, 옮긴 정의는 `ib_renderer`에서 19개, `md_parser`에서 24개이며, 옮긴 정의와 남는 정의를 합하면 요구사항 6절의 정의 목록(24개, 29개)과 같고, 원래 두 파일에는 남는 정의만 있다.

### 커밋별 결과

모든 커밋에서 `ruff check .`, `ruff format --check .`, `mypy .`, `pytest tests/`의 종료 코드가 0이고, 스냅숏 279개가 갱신 없이 통과했다. 기존 테스트 파일은 고치지 않았다.

| 커밋 | 새 파일 | 옮긴 정의 | 통과 | 생략 | 예상 실패 |
|---|---|---|---|---|---|
| `37a4fe7` | `md_parse_text.py` | `TextParser`, `file_uri`, `rebase_link_target`, `_rebase_link_target`, `_parse_term_runs`, `_substitute_text`, `_document_int`, `_DRIVE_PATH_RE`, `_URI_DRIVE_PATH_RE`, `_URI_SCHEME_RE` | 2619 (2575 + 44) | 5 | 1 |
| `438b44e` | `md_parse_frontmatter.py` | `FrontmatterParser`, `_infer_metadata_from_elements`, `_extract_leading_metadata_pair`, `_infer_company_from_title` | 2623 (+ 48) | 5 | 1 |
| `d4803a1` | `md_parse_tables.py` | `TableParser`, `TableSpanResolver` | 2627 (+ 52) | 5 | 1 |
| `620c2c7` | `md_parse_html_tables.py` | `_HtmlCell`, `_html_image`, `_HtmlTableReader`, `HtmlTableParser` | 2631 (+ 56) | 5 | 1 |
| `411302e` | `md_parse_blocks.py` | `FenceScanner`, `FootnoteParser`, `LaTeXParser`, `Base64ImageParser` | 2636 (+ 61) | 5 | 1 |

### T3 끝의 확인

- **본문 동일성.** 모든 새 파일이 있어야 하는 방식(`--layout split`)으로 시작 커밋과 비교한 정의 **53개**, 다른 것 **0개**. 옮긴 43개(렌더러 19, 파서 24)는 정해진 새 파일에, 남은 10개(각 5개)는 원래 파일에 있다. 더한 가져오기 문(규칙 5) 0개. 새 파일 여덟의 `logger`가 규칙 2의 형태. 구획 주석 51줄(27 + 24) 모두 새 구성 안에 있음. 종료 코드 0.
- **배포 묶음.** `uv build` 종료 코드 0. wheel의 최상위 모듈 **33개**, wheel `only-include` 목록과 같음. sdist에도 새 파일 12개가 들어 있음.
- **새 파일 수.** 최상위 `ib_render_*.py`, `md_parse_*.py`가 정확히 12개(요구사항 7절 5번 검사 통과).
- **아직 하지 않은 것.** Word 스키마 검증, 변환·가져오기 시간, 대표 문서의 Word 쪽 비교는 T4의 최종 검증에서 한다.

## 8. 최종 검증 (T4)

옮기기를 모두 마친 코드(`d8cc90a`)에서 요구사항 9절의 검증을 처음부터 다시 돌렸다. 같은 기기, 같은 명령, 같은 방법이다. 그 뒤의 커밋은 문서만 바꿨다(엔진 코드·테스트·설정 변경 없음).

| 검증 | 시작 시점(`5975cff`) | 최종(`d8cc90a`) | 판정 |
|---|---|---|---|
| `uv run ruff check .` | 종료 코드 0 | 종료 코드 0 | 같음 |
| `uv run ruff format --check .` | 0 (86개 파일) | 0 (99개 파일: 새 엔진 파일 12개와 새 테스트 1개) | 같음 |
| `uv run mypy .` | 0 (31개 소스 파일) | 0 (43개 소스 파일: 새 엔진 파일 12개) | 같음 |
| 전체 테스트 통과 | 2,575 | 2,636 = 2,575 + 새 테스트 61 | 기준 충족 |
| 생략 / 예상 실패 | 5 / 1 | 5 / 1 | 같음 |
| 스냅숏 | 279개 통과 | 279개 통과(갱신 없음) | 같음 |
| 배포 묶음 테스트 4종 | 통과 | 통과(`test_codex_bundle`, `test_plugin_bundle`, `test_plugin_runtime`, `test_portable_launcher`) | 같음 |
| 본문 동일성 | 정의 53개 확인 | 비교 53개, 다른 것 0개, 더한 줄 0개 | 같음 |
| Word 스키마 검증 | 문서 13개, 오류 0건 | 문서 13개, 오류 0건 | 같음 |
| `uv build` wheel 최상위 모듈 | 21개(목록과 같음) | 33개(목록과 같음) | 기준 충족 |
| 샘플 13개 변환 시간(중앙값) | 2.2912초 | 2.2964초 | +0.2% |
| `import ib_renderer`(중앙값) | 0.4337초 | 0.4282초 | −1.3% |
| `import md_parser`(중앙값) | 0.3199초 | 0.3138초 | −1.9% |
| 대표 문서 Word 쪽 이미지 | 4쪽 | 4쪽, 쪽마다 이미지 해시 같음 | 같음 |

**시간.** 세 번씩 잰 값(초): 변환 2.3031, 2.2829, 2.2964 / `import ib_renderer` 0.4347, 0.4282, 0.4254 / `import md_parser` 0.3138, 0.3162, 0.3070. 가장 큰 변화가 +0.2%로, 20% 기준보다 훨씬 작다. 파일을 나누면 가져올 파일 수가 늘지만 측정에서 드러나는 차이는 없다.

**Word 쪽 확인.** 2절의 기준과 같은 입력(`samples/profiles/ib-report.md`, SHA-256 `ef264b86…cbb54`)을 같은 명령으로 저장소 밖 새 빈 폴더에 다시 변환했다.

- DOCX 파일 해시는 기준 `9822b2af…41fe2a`, 최종 `4747cc93…c9ba`로 다르다. 그러나 파일 안의 파트 20개는 목록과 내용이 모두 바이트 단위로 같다. 파일 해시가 다른 것은 압축 파일이 각 파트에 저장 시각을 적기 때문이다.
- `scripts/word_visual_qa.ps1`을 새 빈 출력 폴더에 예상 쪽 수 4로 실행했다. 실제 4쪽으로 예상과 같고(`pageCountMatches: true`), 원본 DOCX는 실행 전후 같았다. Word 16.0(빌드 16.0.20527)으로 기준과 같다.
- 쪽 이미지 4장의 SHA-256이 2절의 기준과 쪽마다 같다. 네 쪽을 모두 열어 보았다: 표지, 목차, 본문과 금융표, 면책 고지가 기준과 같은 모습이다. 매니페스트의 `visualReview`는 스크립트가 쓰는 `pending` 그대로이며, 이 항목이 검토 기록이다.

**문서 맞추기(요구사항 8절).** `AGENTS.md`(모듈 표, 새 요소 추가 안내, 내부 파일 안내 한 문장), `README.md`·`README.ko.md`(구조를 설명하는 기존 문장 옆에 "내부 파일로 나뉘었고 가져오는 곳은 같다" 한 문장. 가져오기 예시는 그대로), `docs/format-preprocessing-guide.md`(모듈 표 두 줄), `CHANGELOG.md`(Unreleased에 한 줄), `docs/improvement-roadmap-20260929.md`(현황: 분할 1단계 완료, 다음 차례는 한국어 조판·줄바꿈)를 고쳤다. 날짜가 붙은 과거 계획·검증 문서는 고치지 않았다.

**8절 범위 밖에서 발견했지만 고치지 않은 것.**

- `git_checklist.md`의 "필수 실행 파일 포함 여부" 목록은 `ib_renderer.py`, `md_parser.py`를 들지만 새 파일 12개는 없다. 이 목록은 이번 작업 전에도 `document_model.py`, `input_safety.py` 등 다른 엔진 파일을 빠뜨리고 있었다.
- 저장소 최상위 `plan.md`의 "Deferred" 항목(`ib_renderer.py`, `md_parser.py`, `md_formatter.py`, `md_to_word.py` 분할)은 앞의 두 파일 1단계가 끝난 지금 일부만 맞는다.

## 9. 독립 검토

- **누가.** Codex `gpt-6.1-sol`, effort `xhigh`, 읽기 전용(Codex CLI 0.166.0, 종료 코드 0). 이 작업을 하지 않은 별도 세션이 맡겼고, 검토에는 결론이 아니라 요구사항 전문과 변경 내용을 주었다.
- **대상.** 시작 커밋 `5975cff`부터 `77fd93a`까지의 커밋 16개 전체.
- **판정.** 승인. 막는 지적은 없고 참고 지적이 2건이다. 검토를 맡긴 세션이 두 건 모두 파일을 열어 사실임을 확인했다.
- **검토가 직접 확인한 것.** 정의 53개 전부를 구문 트리와 원문 줄로 비교해 차이가 없었고(중간 커밋에서도 같음), 43개 정의가 정해진 새 파일 12개에 있고 원래 파일마다 5개가 남았으며, 다시 내보내기와 `logger`가 같은 객체이고, 새 파일과 원래 두 모듈을 각각 새 프로세스에서 가져올 수 있고, 엔진 33개의 모듈 수준 가져오기에 순환이 없으며, wheel·sdist 목록이 엔진 파일과 같고, 새 `noqa`·`type: ignore`·`__all__`·생략 표시가 없으며, 정해진 문서만 바뀌었다는 것.

**참고 지적과 처리.**

1. **바꿔치기 검사가 직접 대입을 잡지 못한다**(`tests/test_module_layout.py`의 3번 검사). 테스트가 `md_parser.X = …`처럼 모듈 속성에 직접 대입해 이름을 바꿔치기하면, 새 테스트는 이를 "읽기"로만 기록하므로 3번 검사의 대상에서 빠진다. **처리: 고치지 않았다.** 요구사항 7절 2번이 찾을 형태를 네 가지로 정했고 직접 대입은 그 안에 없으며, 지금 저장소에는 그런 대입이 없다. 범위를 넓히지 않는다는 소유자 방침에 따라 기록하고 소유자에게 보고한다.
2. **가져오기 방향을 적은 문장의 화살표가 앞뒤로 달랐다**(이 문서 6절, 7절의 "새 파일 사이의 가져오기"). **처리: 고쳤다.** 화살표 대신 "글자 파일은 글꼴·각주·그림 파일을 가져온다"처럼 방향을 말로 적고, 각 새 파일의 실제 모듈 수준 `import` 문과 대조했다. 파서 쪽 문장도 같은 방식으로 맞췄다. 문서만 바뀌었으므로 코드 검증은 다시 돌리지 않았다.

**검토가 다시 돌리지 않은 것.** 요청에 따라 검토는 테스트, 린트·서식·타입 검사, `uv build`, Word 스키마 검증, 문서 변환을 실행하지 않았다. 저장소 밖의 DOCX·쪽 이미지·매니페스트를 직접 대조하지 않았고, 시간 측정도 하지 않았다. 검토 환경에 PyInstaller가 없어 휴대용 변환기 빌드가 함수 안 가져오기까지 따라가는지는 확인하지 못했다(모듈 수준의 가져오기 경로가 새 모듈까지 이어지는 것은 확인했다). CI 결과도 보지 않았다. 이 값들은 이 기록의 8절과 10절(CI)에 근거한다.

## 10. CI

- **대상.** 작업 브랜치 `claude/large-file-split-1-1ea9dc`의 커밋 `0860991`(PR #24). 이 커밋의 엔진 코드·테스트·설정은 8절의 최종 검증을 한 `d8cc90a`와 같다(그 사이 커밋은 문서만 바꿨다).
- **실행이 두 개인 이유.** CI 설정(`.github/workflows/ci.yml`)이 `push`와 `pull_request` 두 사건에서 모두 돈다. 실행 `43441589419`는 브랜치를 올린 `push`, 실행 `43441749241`은 PR의 `pull_request`이며, 둘 다 같은 커밋 `0860991`을 검사했다.
- **결과.** 두 실행 모두 검사 5개가 모두 성공했다(합계 10개 성공, 실패 0).

| 검사 | `push` 실행 `43441589419` | `pull_request` 실행 `43441749241` |
|---|---|---|
| `quality (ubuntu-latest, 3.12)` | 성공 (5분 22초) | 성공 (6분 47초) |
| `quality (ubuntu-latest, 3.13)` | 성공 (6분 50초) | 성공 (6분 32초) |
| `quality (windows-latest, 3.12)` | 성공 (7분 24초) | 성공 (9분 36초) |
| `quality (windows-latest, 3.13)` | 성공 (9분 34초) | 성공 (9분 40초) |
| `Word schema gate` | 성공 (28초) | 성공 (23초) |

품질 검사는 린트, 서식, 타입 검사, 전체 테스트, 배포 묶음 빌드, wheel 내용이 `only-include` 목록과 같은지, 최상위 파이썬 파일의 Python 3.12 문법을 확인한다. Word 스키마 검증은 Ubuntu에서 샘플 문서 13개를 엄격 모드로 변환해 고정한 검증기로 검사한다.

이 절은 `0860991`의 결과다. 이 절을 더한 뒤의 문서 커밋에 대한 CI는 PR #24의 검사 표시로 확인한다.

## 11. 하지 않은 것

- **이 기록을 더한 문서 커밋의 CI.** 위 10절은 `0860991`의 결과다. 그 뒤의 문서 커밋에 대한 CI 결과는 이 기록에 적지 않고 PR의 검사 표시로 확인한다.
- **macOS·Linux에서의 측정.** 이 기록의 시간 값과 Word 쪽 확인은 위 Windows 기기에서만 했다. Linux는 CI로 확인한다. macOS는 확인하지 않는다.
- **배포 묶음의 실제 재생성과 게시.** 플러그인·Codex 묶음과 휴대용 변환기를 다시 만들어 게시하지 않았다. 범위 밖이다. 묶음 스크립트가 읽는 wheel 목록과 묶음 테스트 4종은 위에서 확인했다.
- **대표 문서 외 샘플의 Word 쪽 확인.** XML이 모두 같으므로(스냅숏) 대표 문서 하나만 Word로 확인했다(요구사항 9절).
