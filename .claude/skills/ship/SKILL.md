---
name: ship
description: Use when 릴리즈·버전 범프·배포 준비를 요청받았을 때 ("릴리즈 하자", "버전 올려", "1.0.4 준비", bump version, release). 버전 문자열이 들어가는 파일을 하나라도 고치기 전에 반드시 이 절차를 따른다.
---

# Ship — 릴리즈 절차

## 개요

이 저장소의 릴리즈는 **버전 일관성 3점 세트(pyproject.toml == CHANGELOG.md == uv.lock)를 한 커밋에** 담는 것이다.
과거 릴리즈에서 uv.lock 누락(`chore: sync uv.lock for v1.0.1`)과 메타데이터 뒷수습
(`chore: update test count and fix pyproject readme field`)으로 fix-up 커밋이 두 번 발생했다.
이 절차는 그 두 사고의 재발 방지책이다.

## 절차

### 0. 사전 조건

```bash
git branch --show-current        # main이어야 함
git status --short               # 릴리즈와 무관한 변경이 없어야 함 (기존 untracked는 예외)
git pull
```

이후 `/preflight` 전체 실행 → 전부 PASS일 때만 진행. FAIL 상태에서 릴리즈 금지.

### 1. 버전 결정

```bash
LAST=$(git log --format='%H' --grep='bump release metadata' -1)
git log --oneline ${LAST}..HEAD
```

위 커밋 목록으로 semver 판정: `fix:`만 있으면 **patch**, `feat:`가 하나라도 있으면 **minor**,
하위호환이 깨지는 변경(기본 출력 변화, CLI 인터페이스 제거)이 있으면 **멈추고 사용자에게 확인** — 이 도구의 사용자는 major 범프 여부를 직접 정한다.

### 2. 편집 세트 — 아래 전부를 하나의 커밋에

| 파일 | 할 일 |
|------|------|
| `pyproject.toml` | `version = "X.Y.Z"` 갱신 |
| `CHANGELOG.md` | 최상단에 `## [X.Y.Z] - YYYY-MM-DD` 섹션. 1단계 커밋 목록을 Added/Changed/Fixed/Documentation으로 분류해 사용자 관점 문장으로 요약 (기존 1.0.2 항목의 문체를 모방) |
| `uv.lock` | pyproject 수정 **후** `uv lock` 실행 (자동 갱신됨) |
| `README.md` + `README.ko.md` | 기능 목록이 바뀌었을 때만, 두 파일 함께 |
| `AGENTS.md` | 테스트 개수 표기를 실측값으로 (`uv run pytest -q \| tail -1`) |

날짜는 실제 오늘 날짜(YYYY-MM-DD). 커밋 이력에서 복사한 과거 날짜를 쓰지 않는다.

### 3. 일관성 검증 (통과 전 커밋 금지)

```bash
grep -m1 '^version' pyproject.toml
grep -m1 '^## \[' CHANGELOG.md
grep -A2 'name = "ib-report-formatter"' uv.lock | grep version
```

세 출력의 버전이 모두 `X.Y.Z`로 일치해야 한다. 하나라도 다르면 **지금** 고친다.
fix-up 커밋으로 나중에 수습하는 것은 이 스킬의 실패다.

### 4. 커밋

```bash
git add pyproject.toml CHANGELOG.md uv.lock          # 바꾼 문서가 더 있으면 명시적으로 추가
git diff --cached --name-only                        # 민감 파일 없음 확인 (M9)
```

커밋 메시지 (전례 d819bbb):

```
chore: bump release metadata to X.Y.Z

- CHANGELOG: (한 줄 요약)
- (동반 문서 갱신이 있으면 나열)

Co-Authored-By: Claude <모델명> <noreply@anthropic.com>
```

### 5. 태그·push

- **태그를 만들지 않는다.** 이 저장소는 git tag를 쓰지 않는 것이 관례다 (`git tag -l`이 비어 있음).
- **push는 사용자 승인 후에만** (CLAUDE.md §6-8). 커밋까지 완료한 뒤 push 여부를 물어라.

## 합리화 차단표

| 변명 | 현실 |
|------|------|
| "uv.lock은 커밋 후에 따로" | v1.0.1에서 그렇게 해서 fix-up 커밋이 생겼다. `uv lock`은 3초다. |
| "CHANGELOG는 커밋 메시지 복붙으로" | CHANGELOG는 사용자 관점 문장이다. 기존 1.0.2 항목과 문체를 비교하라. |
| "preflight는 방금 했으니 생략" | 버전 파일을 편집한 뒤의 상태는 "방금"이 아니다. 3단계 검증이 그 보완이다. |
| "태그도 만들어두면 좋지 않나" | 관례에 없는 산출물을 임의로 추가하지 않는다. 필요하면 사용자가 시킨다. |

## 레드 플래그

- 버전 문자열을 두 파일만 고치고 커밋하려 한다 → 3점 세트 미완성.
- CHANGELOG 날짜를 추정으로 쓰고 있다 → 오늘 날짜를 확인하라.
- push까지 한 번에 하려 한다 → 승인 게이트 위반.
