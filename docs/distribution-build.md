# 두 배포본 만들기

같은 Markdown → Word 엔진을 두 방식으로 전달합니다.

| 배포본 | 받는 사람이 하는 일 | 필요한 환경 |
|---|---|---|
| Windows portable ZIP | 작성한 MD를 실행 프로그램에서 변환 | Windows x64, Python 설치 불필요 |
| Codex source ZIP | 자료와 요청을 주고 Codex에서 MD 작성·변환 | Codex 계정·앱, uv, Python 3.12 이상 |

사용 안내는 [실행 프로그램 시작하기](distribution/시작하기.html)와 [Codex 시작하기](codex-start.html)를 참고합니다. Codex 방식에는 별도 프로젝트 API 키가 필요하지 않습니다. 실행 프로그램 자체는 LLM을 호출하지 않습니다.

`docs/distribution/시작하기.html`은 Windows 배포본에 복사할 안내 원본입니다. 예제 링크는 빌드 후 ZIP 루트에 배치된 `시작하기.html`에서 작동합니다. 저장소·Codex 소스 묶음에서는 `samples/distribution`과 `samples/profiles`의 원본을 참고하십시오.

아래 명령은 저장소 루트에서 실행하는 **빌드 방법**이며, 특정 바이너리의 새 빌드·검증 또는 GitHub Release 게시를 확인한 기록이 아닙니다. 생성물은 Git에서 제외된 `dist/`에 두고 빌드 스크립트·설명서·가상 예제만 커밋합니다.

## Windows 실행 프로그램

Windows x64에서 uv와 Python 3.12를 사용합니다. 별도 빌드 환경에 잠긴 런타임 의존성과 PyInstaller를 설치합니다. 출력 폴더는 매번 새 경로를 지정합니다.

```powershell
uv venv --python 3.12 dist/portable-build-env
uv export --locked --extra full --no-dev --no-emit-project --output-file dist/portable-requirements.txt
uv pip install --python dist/portable-build-env/Scripts/python.exe -r dist/portable-requirements.txt "pyinstaller==6.22.0"
dist/portable-build-env/Scripts/python.exe scripts/build_portable.py --output-dir dist/portable-release-001
```

빌더는 실행 프로그램과 지원 파일, 텀시트 작성 안내, 가상 MD/YAML 및 변환 예제, 라이선스, 파일 해시가 담긴 `manifest.json`을 ZIP으로 묶습니다. 의존성 설치와 누락된 Tcl/Tk 라이선스 취득에는 인터넷이 필요할 수 있습니다.

`--app-dir`는 이미 빌드한 실행 프로그램 폴더를 재사용하는 옵션입니다. **현재 소스와 같은 코드로 빌드되었음을 확인한 경우에만 사용하십시오.** 오래된 실행 프로그램을 재사용하면 새 설명서·예제와 실제 변환 동작이 서로 달라질 수 있습니다. 새 배포에는 위 명령처럼 이 옵션을 생략하는 편이 명확합니다.

ZIP을 새 폴더에 전체 압축 해제하고 GUI 실행, 예제 변환, 오류 표시를 확인합니다. 수신 환경을 대표하는 Python 미설치 Windows PC에서도 확인하고, Word 페이지 배치는 별도로 검토합니다. manifest의 `visual_review: pending`은 페이지 검토 완료를 뜻하지 않습니다.

## Codex 소스 묶음

```powershell
uv sync --locked
uv run python scripts/build_codex_bundle.py --output-dir dist
```

빌더는 명시된 공개 파일만 포함하고 ZIP, SHA-256 파일, 묶음 내부의 `bundle-manifest.json`을 생성합니다. 소스·작성 스킬·검증 도구·시작 안내·가상 자료가 포함되며 Python 런타임, 개인 자료, 실제 거래 문서와 생성 결과는 포함되지 않습니다. 파일 목록은 `scripts/build_codex_bundle.py`에서 관리합니다.

받는 사람은 전체 압축 해제 후 `docs/codex-start.html`을 읽고 폴더를 Codex 프로젝트로 엽니다. 새 폴더에서 의존성 설치와 가상 요청 예제를 실행해 묶음만으로 작성·변환이 되는지 확인합니다. 원자료 일치와 페이지 검토는 구조 검사 결과와 별도로 기록합니다.

## 전달

검증한 ZIP과 해시를 공유폴더 또는 GitHub Release 자산으로 전달할 수 있습니다. Git 저장소에는 생성된 ZIP·EXE를 직접 추가하지 않습니다. 배포할 때 사용한 소스 커밋, 환경과 검증 범위를 함께 남기십시오.
