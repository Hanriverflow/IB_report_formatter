# IB Report Formatter 플러그인 사용 안내

Claude Code 또는 Codex에 이 플러그인을 설치하면 **현재 작업 폴더의 자료를 주고 Word 문서 작성을 요청**할 수 있습니다. AI가 근거 정리와 Markdown 초안을 작성하고, 내장 Python 엔진이 엄격 검사 후 DOCX를 만듭니다. 별도 모델 API 키나 변환용 EXE는 필요 없습니다. 사용하는 Claude Code/Codex의 계정과 이용 조건이 적용됩니다.

이 안내의 대상은 **Claude Code와 Codex**입니다. Claude 웹 또는 네이티브 Cowork에서 설치·실행이 검증됐다는 의미는 아닙니다. 일반 소스 폴더를 Codex 프로젝트로 여는 방법은 [GitHub의 기존 Codex 안내](https://github.com/Hanriverflow/IB_report_formatter/blob/4dda4e523fab8047289f29038b7f9abd583daed9/docs/codex-start.html)를 참고하세요.

플러그인에는 엔진 사용법을 담은 `README.md`와 `README.ko.md`가 포함됩니다. README에 연결된 개발 계획·과거 검증 기록 등 전체 저장소 문서는 이 최소 배포본에 포함되지 않을 수 있으므로 해당 자료는 [GitHub 저장소](https://github.com/Hanriverflow/IB_report_formatter)에서 확인하세요.

## 1. ZIP 전체를 압축 해제

전달받은 플러그인 ZIP을 예를 들어 `C:\도구\ib-report-formatter`에 압축 해제합니다. 다음 파일들이 함께 있는 폴더를 지정해야 합니다.

```text
ib-report-formatter/
  plugin.json
  .claude-plugin/
  skills/
    setup/SKILL.md
    ib-document/SKILL.md
  scripts/plugin_runtime.py
  pyproject.toml
  uv.lock
  samples/
  docs/
```

숨김 폴더를 포함해 전체 구조를 유지하세요. 이 폴더는 플러그인 설치 원본입니다. 실제 거래조건표, 기관 문구, 생성 문서는 별도의 작업 폴더에 보관합니다. GitHub 계정이나 소스 수정은 필요 없습니다.

## 2. 사용하는 앱에 설치

아래 경로는 실제로 압축 해제한 폴더로 바꾸세요. 설치 명령은 해당 호스트의 CLI가 실행되는 터미널에서 사용합니다. 설치 후 새 대화에서 플러그인을 사용하세요. 아래는 로컬 배포본 설치 방법이며, GitHub 기본 브랜치에 플러그인이 이미 게시되어 있다는 뜻이 아닙니다.

### Claude Code

```powershell
claude plugin marketplace add "C:/도구/ib-report-formatter"
claude plugin install ib-report-formatter@ib-report-formatter-local
```

Claude Code에서 다음 명령으로 초기 환경을 준비합니다.

```text
/ib-report-formatter:setup
```

설치 전에 한 세션에서만 시험하려면 `claude --plugin-dir "C:/도구/ib-report-formatter"`로 시작할 수 있습니다. 이 방식은 영구 설치가 아닙니다.

### Codex

```powershell
codex plugin marketplace add "C:/도구/ib-report-formatter"
codex plugin add ib-report-formatter@ib-report-formatter-local
```

Codex 앱에서 문서를 작성할 **자신의 작업 폴더**를 프로젝트로 열고 새 대화에서 요청합니다.

```text
IB Report Formatter 플러그인의 setup 스킬로 실행 환경을 준비해줘.
설치된 플러그인 위치에서 스킬을 읽고, 실제 작업 파일은 이 작업 폴더에 보관해줘.
```

호스트 버전에 따라 플러그인 목록과 호출 UI가 다를 수 있습니다. 스킬이 발견되지 않으면 설치된 플러그인의 `skills/setup/SKILL.md`와 `skills/ib-document/SKILL.md`를 실제 절대 경로로 지정해 읽도록 요청하세요. 원본 ZIP의 다른 복사본과 설치된 플러그인을 혼동하지 않도록 합니다. `plugin` 명령 자체가 없다면 해당 호스트에서 플러그인을 지원하는 버전인지 먼저 확인해야 합니다.

## 3. 최초 실행 환경

플러그인은 `uv`와 Python 3.12 실행 환경을 사용합니다. Python은 필요하면 uv가 준비하며, 엔진 의존성은 동봉된 잠금 파일에 따라 설치합니다. uv가 없다면 [uv 공식 설치 안내](https://docs.astral.sh/uv/getting-started/installation/)를 따르거나 사용 중인 에이전트에 환경 준비를 요청하세요. 런처 자체가 uv를 시스템에 설치하지는 않습니다.

초기 환경은 플러그인 설치 폴더 밖의 캐시에 생성합니다. 원본 플러그인 파일이나 사용자 프로젝트의 의존성 파일을 수정하지 않습니다. 최초 Python·라이브러리 다운로드와 호스트 모델 사용에는 인터넷 연결이 필요할 수 있습니다. ZIP 자체가 실행 환경을 포함하거나 AI 작성까지 오프라인으로 제공하는 것은 아닙니다.

에이전트가 실행하는 초기 설정의 형태는 다음과 같습니다. 일반 사용자가 직접 실행할 필요는 없습니다.

```powershell
uv run --no-project --python 3.12 python "C:/실제/플러그인/scripts/plugin_runtime.py" setup
```

성공하면 JSON의 `status`가 `ready`이고 실제 캐시·실행 환경 경로가 표시됩니다. 다른 위치를 지정하려면 `setup --cache-dir "C:/문서도구캐시"`를 사용할 수 있습니다. 이후 `render`에도 동일한 `--cache-dir`를 전달하거나 `IB_REPORT_FORMATTER_CACHE_DIR`로 지정합니다. 이 경로는 플러그인 밖이어야 합니다.

## 4. 가상 자료로 첫 문서 만들기

Claude Code에서는 아래 요청 앞에 `/ib-report-formatter:ib-document`를 넣으세요. Codex에서는 “IB Report Formatter 플러그인의 ib-document 스킬을 사용해줘”라고 요청하면 됩니다.

```text
IB Report Formatter 플러그인의 ib-document 스킬을 사용해줘.
플러그인 안 samples/harness/request.txt와 brief.md를 읽고
가상 ABCP 텀시트를 작성해줘. 기관 문구는 같은 폴더의 house.yaml을 사용해.
expected-terms.json을 확정값 비교 기준으로 사용하고 기준값을 임의 수정하지 마.
미확정 주선수수료와 ABCP 조건은 확인 필요로 표시해줘.
현재 작업 폴더에 새 결과 폴더를 만들고, 근거 목록·MD·Word·검사 결과를 남겨줘.
원자료 대조, 구조 검사, 실제 페이지 검토 여부를 각각 알려줘.
```

샘플은 새로 작성한 가상 조건과 가상 기관 문구입니다. 대출금리와 ABCP 발행금리처럼 서로 다른 항목을 구분해야 하며, 샘플 값을 실거래의 기본값으로 쓰면 안 됩니다. 모든 페이지를 실제로 열어 확인하지 않았다면 시각 검토는 미수행입니다.

## 5. 실제 텀시트 요청

실제 파일 경로로 바꿔 요청하세요. 지원되는 다른 문서 프로필은 `ib-report`, `ib-memo`, `plain`, `office-letter`, `business-report`, `meeting-minutes`입니다.

```text
IB Report Formatter 플러그인의 ib-document 스킬을 사용해줘.
C:/문서작업/거래A/거래조건.md를 근거로 term-sheet 프로필의 텀시트를 작성해줘.
기관 공통 문구는 C:/문서작업/거래A/house.yaml을 사용해줘.
거래개요, 주요 금융조건, 상환조건, 확인 필요 사항으로 구성해줘.
당사자·금액·금리·일정은 원자료의 표기와 확정 여부를 유지해줘.
자료가 없거나 충돌하는 항목은 추정하지 말고 질문 목록에 남겨줘.
C:/문서작업/거래A 아래 새 작업 폴더에 결과를 저장해줘.
```

에이전트는 근거 목록에서 값을 정리한 뒤 `known-terms.json`과 MD의 `terms:`를 작성합니다. `{{amount}}` 같은 변수로 반복되는 값을 사용하고, 독립적으로 정리한 비교 기준과 대조합니다. 기관 파일과 이미지 경로는 원본 기준을 보존합니다. PDF·Word·HWP 자료를 읽는 도구는 호스트 환경에 따라 달라지며 이 플러그인 자체가 역변환기를 제공하지는 않습니다.

텀시트에는 제목과 작성기관·공통 고지가 필요합니다. 필수 문구를 제공하지 않았다면 승인된 것처럼 임의 작성하지 않습니다. 미확정 금리나 상환표 모순 때문에 수치 검사를 통과할 수 없다면 초안과 구체적 확인사항을 전달합니다.

## 6. 결과와 수정

| 결과물 | 의미 |
|---|---|
| DOCX, 수정 가능한 MD | 생성된 문서와 재생성할 원본. 엄격 변환 실패 시 DOCX가 없을 수 있음 |
| `source-ledger.md` | 주요 사실의 값·출처·상태. 단순 MD 변환은 생략 가능 |
| `known-terms.json` | 원자료에서 먼저 정리한 일부 조건의 정확한 비교 기준 |
| `questions.md` | 미확정·누락·충돌 조건이 있을 때 작성 |
| 결과 JSON | 실행 단계, 진단, 감사, 실제 경로와 해시 |

각 변환 시도는 **아직 존재하지 않는 새 출력 폴더**를 사용합니다. 빈 폴더라도 재사용하지 않습니다. 작성 오류는 진단을 근거로 최대 두 차례 수정하며, 실패한 검사 삭제·엄격 모드 해제·원자료 수치 변경으로 통과시키지 않습니다.

내용 수정은 새 근거와 함께 요청하세요. 원자료 대조는 문서 내용의 근거를, 엄격 변환과 구조 감사는 엔진 규칙을, 시각 검토는 실제 페이지를 확인합니다. 이 세 가지 상태는 별개이며, 구조 검사 통과는 금융·법률적 타당성 또는 페이지 배치의 인증이 아닙니다.

## 설치 참고

- [Codex 공식 플러그인 안내](https://developers.openai.com/plugins/build/plugins)
- [Claude Code 공식 플러그인 안내](https://code.claude.com/docs/en/plugins)
- [Claude Code 공식 플러그인 마켓플레이스 안내](https://code.claude.com/docs/en/plugin-marketplaces)

이 문서는 사용 방법입니다. 설치·런타임 시험 결과와 실제 모델을 사용한 문서 작성 시험 결과는 별도의 검증 범위로 기록해야 합니다. 로컬 런처가 성공했다는 이유만으로 모든 호스트 UI나 서비스 환경의 실행을 검증했다고 표현하지 않습니다.
