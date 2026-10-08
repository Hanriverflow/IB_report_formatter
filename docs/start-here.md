# 배포본 선택과 시작하기

IB Report Formatter는 **같은 Markdown → Word 엔진**을 세 가지 방식으로 제공합니다. 먼저 실행 방식을 고르고, 문서 종류는 프로필로 선택하세요.

## 1. 문서 변환기: Windows 실행형

**이미 작성한 Markdown을 Word로 바꾸려는 분**에게 적합합니다. AI가 문서를 작성하는 앱은 아닙니다.

- 필요 환경: Windows x64. Python이 포함되어 별도 설치가 필요 없고 AI 계정도 쓰지 않습니다.
- 파일명: `ib-report-formatter-windows-x64-2.0.0.zip` (현재 엔진 버전 기준).
- ZIP **전체**를 새 폴더에 압축 해제합니다. EXE만 따로 옮기지 마세요.
- `시작하기.html`을 브라우저로 열고 `예제`의 가상 MD/YAML을 확인합니다.
- `문서변환기.exe`에서 MD 파일, 출력 폴더, 문서 유형을 고르고 변환합니다.
- 원본 내용과 실제 Word 페이지를 검토합니다. 변환 성공은 금융·법률적 승인이나 페이지 검토 완료가 아닙니다.

[상세 HTML 안내 원본](https://github.com/Hanriverflow/IB_report_formatter/blob/main/docs/distribution/시작하기.html)은 GitHub에서는 소스로 보일 수 있습니다. 배포 ZIP에 들어 있는 파일을 브라우저로 여세요.

## 2. AI 문서 작성 도우미: 플러그인

**기존 업무 폴더의 자료로 AI에게 작성·수정부터 Word 변환까지 맡기려는 분**에게 적합합니다. 일반 AI 작성 사용자의 기본 선택입니다.

- 지원 대상: Claude Code / Codex의 플러그인 지원 버전과 본인 계정.
- 파일명: `ib-report-formatter-plugin-0.1.0.zip` (현재 플러그인 버전 기준).
- uv가 필요하며 Python 3.12와 잠긴 의존성은 최초 설정 시 준비합니다. ZIP에 런타임이나 EXE는 포함되지 않습니다.
- ZIP을 도구용 폴더에 전체 압축 해제하고 [플러그인 설치 안내](plugin-guide.md)에 따라 호스트에 설치합니다.
- 실제 자료와 결과는 **플러그인 밖의 자신의 작업 폴더**에 보관합니다. 호스트에서 그 작업 폴더를 열어 `setup`, `ib-document` 스킬을 사용합니다.
- 별도 프로젝트 API 키는 필요 없지만 호스트의 계정·이용 조건이 적용됩니다. 초기 환경 다운로드와 AI 사용에는 인터넷이 필요할 수 있습니다.

`.plugin` 파일은 동일한 ZIP의 다른 확장자입니다. 별도 제품도, Claude 웹·Cowork 호환 인증도 아닙니다. [기록된 검증 범위와 미검증 항목](https://github.com/Hanriverflow/IB_report_formatter/blob/main/docs/verification-plugin-20261008.md)을 확인하세요.

<a id="codex-project-kit"></a>
## 3. Codex 프로젝트 키트: 고급 사용자용 소스형

**프로젝트 전체를 열어 예제를 배우거나 작업 흐름·코드를 수정하려는 분**에게 적합합니다. 플러그인보다 변환 기능이 많은 상위 제품이 아닙니다.

- 필요 환경: Codex 계정·앱, uv, Python 3.12 이상. 런타임은 포함되지 않습니다.
- 파일명: `ib-report-formatter-codex-source-2.0.0.zip` (현재 엔진 버전 기준).
- ZIP을 전체 압축 해제하고 `AGENTS.md`, `pyproject.toml`, `.agents`, `samples`가 함께 있는 **소스 폴더 자체**를 Codex 프로젝트로 엽니다.
- `$ib-document`로 가상 예제 작성을 요청합니다. 인식되지 않으면 `.agents/skills/ib-document/SKILL.md`를 읽도록 요청합니다.
- 실제 거래자료·기관 설정·결과물은 저장소 밖에 보관합니다.
- 소스 묶음에는 플러그인·빌드 관련 파일도 있지만, 기본 시작 방법은 소스 프로젝트 방식입니다. 설치형 플러그인처럼 쓰려면 2번 안내를 따르세요.

[상세 Codex HTML 안내](codex-start.html)는 압축 해제 후 브라우저로 열 수 있습니다.

## 문서 종류는 프로필로 선택

| 목적 | 프로필 |
|---|---|
| 공문·업무보고·회의록·일반 문서 | `office-letter`, `business-report`, `meeting-minutes`, `plain` |
| IB 보고서·메모 | `ib-report`, `ib-memo` |
| 금융거래 텀시트·기관별 스타일 | `term-sheet` + 기관별 YAML 설정 |

예제는 가상 자료입니다. 기관 문구·로고 설정을 바꿀 때 엔진 브랜치를 나눌 필요는 없습니다. HWP/Word를 직접 입력하는 역변환기는 제공하지 않습니다. 외부에서 변환한 MD의 정리 기능과는 구분하세요.

## 다운로드·버전·검증 상태

- [GitHub Releases](https://github.com/Hanriverflow/IB_report_formatter/releases)에 게시된 자산만 공개 배포본으로 취급합니다. 자산이 없다면 아직 공개 다운로드가 게시되지 않은 상태입니다.
- GitHub의 **Code → Download ZIP**은 저장소 소스이며 Windows 실행 프로그램 ZIP이 아닙니다.
- 엔진 **2.0.0 (Beta)**와 플러그인 **0.1.0**은 별도 버전입니다. 현재 플러그인에는 엔진 2.0.0이 포함됩니다. 실제 묶음의 manifest와 검증 기록을 함께 확인하세요.
- 테스트·구조 검사, 원자료 일치 확인, Word의 실제 페이지 검토는 각각 다른 단계입니다. 한 단계의 성공으로 나머지를 보장하지 않습니다.
- 개발 브랜치는 개발 이력입니다. 최종 사용자에게는 검증된 Release 자산을 전달하고, 빌드는 [관리자용 안내](https://github.com/Hanriverflow/IB_report_formatter/blob/main/docs/distribution-build.md)를 따릅니다.
