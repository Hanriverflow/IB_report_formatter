# GitHub 배포 전 체크리스트

IB Report Formatter를 GitHub에 올리기 전에 아래 항목을 순서대로 확인하세요.

## 1) 민감정보/불필요 산출물 제외

- [ ] 사내 민감정보가 포함된 원본 `.md` 파일 제외
- [ ] 보고서 결과물(`*.docx`) 제외
- [ ] 로컬 환경/캐시 파일 제외 (`.venv/`, `__pycache__/`, `.pytest_cache/`, `.mypy_cache/`, `.ruff_cache/`)
- [ ] 로컬 도구 상태 폴더 제외 (`.claude/`, `.sisyphus/`)
- [ ] `.gitignore`가 루트에 존재하는지 확인

빠른 확인:

```bash
git status --short
```

## 2) 필수 실행 파일 포함 여부

아래 파일/폴더는 저장소에 포함되어야 합니다.

- [ ] `md_to_word.py`
- [ ] `md_parser.py`
- [ ] `md_formatter.py`
- [ ] `ib_renderer.py`
- [ ] `pyproject.toml`
- [ ] `uv.lock`
- [ ] `tests/`
- [ ] `README.md`
- [ ] `README.ko.md`

선택 포함:

- [ ] `AGENTS.md` (협업 가이드)
- [ ] `docs/` (내부/민감 노트 제거 후)

## 3) 로컬 검증

- [ ] 테스트 통과

```bash
uv run pytest -q
```

- [ ] 기본 변환 동작 확인

```bash
uv run md_to_word.py --list
```

## 4) README 확인

- [ ] `README.md`의 설치/실행 가이드 최신 상태
- [ ] `README.ko.md`의 설치/실행 가이드 최신 상태
- [ ] GitHub 업로드 포함/제외 기준 반영 여부 확인

## 5) 커밋 전 최종 점검

- [ ] 커밋 대상 파일 재확인

```bash
git status
```

- [ ] 스테이징 파일 목록 재확인

```bash
git diff --cached --name-only
```

## 6) 최초 업로드 절차 (미초기화 저장소)

```bash
git init
git add .
git commit -m "Initial release: IB report formatter"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

## 7) 업데이트 배포 절차 (기존 저장소)

```bash
git add .
git commit -m "Update: <change summary>"
git push
```

## 8) 다른 PC에서 설치 확인 (최종)

클린 환경에서 아래만으로 실행 가능해야 합니다.

```bash
git clone <your-repo-url> IB_report_formatter
cd IB_report_formatter
uv sync
uv run md_to_word.py --list
```

문제가 없다면 배포 준비 완료입니다.
