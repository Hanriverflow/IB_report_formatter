---
name: ib-report-docx
description: md 보고서를 저장소의 변환 엔진으로 IB 스타일 Word로 만들거나 docx로 변환할 때 사용한다. "IB 리포트 만들어", "마크다운을 Word로", "보고서를 docx로 변환" 같은 요청에서 작성 규약, 플래그 선택, 실행, QA를 안내한다.
---

# IB Report DOCX

## 트리거와 적용 범위

이 스킬은 Markdown 원고를 이 저장소의 기존 변환 엔진으로 IB 스타일 `.docx`로 만들 때 사용한다.

- "md 보고서를 IB 스타일 Word로 만들어"
- "이 마크다운을 docx로 변환해"
- "IB 리포트 만들어"
- "붙여넣은 리서치 결과를 Word 보고서로 정리해"

이 스킬은 원고 구조를 엔진이 이해하는 Markdown으로 맞추고, CLI 플래그를 고르고, 결과를 점검하는 얇은 어댑터다. 색상, 폰트, 여백, 표 스타일 등 하우스 스타일의 단일 소스는 기존 렌더링 엔진이다.

## Authoring 규약

변환 전에 원본 Markdown을 아래 규약에 맞추되, 이미 구조가 올바르면 불필요하게 다시 쓰지 않는다.

### YAML frontmatter

직접 지원되는 최상위 키만 사용한다.

| 키 | 용도 |
|---|---|
| `title` | 보고서 제목 |
| `subtitle` | 부제 |
| `company` | 기관 또는 회사 |
| `ticker` | 티커 |
| `sector` | 업종 |
| `analyst` | 작성자 |

```markdown
---
title: 한빛산업 기업분석
subtitle: 실적 회복과 재무 안정성
company: 한빛산업
ticker: "012345"
sector: 산업재
analyst: 기업금융팀
---
```

frontmatter가 없으면 첫 `#` 헤딩에서 제목을 추론할 수 있고, 문서 상단의 볼드 키-값 문단도 메타데이터로 추론된다.

```markdown
**기준일:** 2026-08-18
**분석 대상 기간:** 최근 3개년
**분석 기준:** 연결 재무제표
**작성자:** 기업금융팀
**기관:** 한빛산업
**업종:** 산업재
**수신:** 투자심의위원회
```

지원 라벨은 `기준일`, `작성일`, `as of`, `report date`, `date`,
`분석 대상 기간`, `analysis period`, `분석 기준`, `analysis basis`,
`prepared by`, `작성자`, `analyst`, `institution`, `company`, `기관`,
`sector`, `업종`, `prepared for`, `recipient`, `수신`이다.
형식은 반드시 `**라벨:** 값`으로 쓴다.

### 헤딩

헤딩은 `#`부터 `####`까지만 사용해 문서 계층을 명확히 한다.

```markdown
# 기업분석 보고서
## 1. Executive Summary
### 1.1 핵심 실적
#### 주요 가정
```

`#`은 문서 제목, `##`는 주요 섹션, `###`와 `####`는 하위 논점에 사용한다.

### 표

표 타입은 헤더 행의 키워드로 자동 판별된다. 우선순위는 `UPSIDE_DOWNSIDE` → `BEP_SENSITIVITY` → `RISK_MATRIX` → `FINANCIAL` → `GENERIC`이며, 여러 타입의 키워드가 섞이면 앞선 타입이 이긴다.

| 판별 타입 | 헤더 키워드 |
|---|---|
| `UPSIDE_DOWNSIDE` | upside, downside, 상승, 하락, 요인 |
| `BEP_SENSITIVITY` | bep, sensitivity, cmr, contribution margin, fixed cost, 손익분기, 민감도, 고정비, 변동비 |
| `RISK_MATRIX` | risk, impact, probability, likelihood, 리스크, 위험, 영향, 확률 |
| `FINANCIAL` | revenue, income, ebitda, profit, margin, expense, 매출, 수익, 이익, 손익, 순이익, 영업, 비용 또는 연도·실적 지표 |
| `GENERIC` | 위 키워드가 없는 표 |

연도·실적 지표에는 `2024`, `2025`, `2026`, `yoy`, `cagr`, `a)`, `b)`, `e)`, `년도`, `연도`, `실적`이 포함된다.

일반 표 예시:

```markdown
| 구분 | 설명 |
|---|---|
| 사업 | 산업용 부품 제조 |
```

재무 표 예시:

```markdown
| Revenue | 2025 | 2026E | YoY |
|---|---:|---:|---:|
| 매출액 | 120,000 | 135,000 | 12.5% |
```

리스크 표 예시:

```markdown
| Risk | Impact | Probability | 대응 |
|---|---|---|---|
| 원재료 상승 | High | Medium | 장기 계약 확대 |
```

### 콜아웃

콜아웃은 blockquote 첫 줄에 인식 가능한 라벨을 둔다. 허용 라벨은 `시사점`, `참고`, `주의`, `결론`, `요약`, `핵심`, `KEY INSIGHT`, `NOTE`, `WARNING`이며 영문은 대소문자를 구분하지 않는다.

```markdown
> [시사점] 마진 회복은 판가 인상보다 원가 안정화의 영향이 크다.
>
> 추가 확인이 필요한 가정은 민감도 표에 반영한다.
```

라벨 목록 밖의 표현을 새 콜아웃 이름으로 만들지 않는다.

### LaTeX

인라인 수식은 `$...$`, 블록 수식은 `$$...$$`로 쓴다.

```markdown
기업가치는 $EV = EBITDA \times Multiple$로 계산한다.

$$
WACC = \frac{E}{D+E}R_e + \frac{D}{D+E}R_d(1-T)
$$
```

### 이미지

로컬 파일은 Markdown 파일 기준으로 해석 가능한 경로를 사용한다.
Base64 이미지도 같은 이미지 문법으로 넣는다.

```markdown
![매출 추이](assets/revenue_chart.png)
![회사 로고](data:image/png;base64,iVBORw0KGgoAAA...)
```

### 각주

인라인 마커는 `.1` 또는 `^1^`을 사용한다.
`.1`은 공백, 줄끝 또는 `, ; : -` 앞에 놓는다.

```markdown
회사는 신규 공장 투자를 발표했다.1
신용등급은 안정적이다^2^.

## 참고문헌
1. 회사 공시자료, 설비투자 계획
2. 신용평가사 정기평가 보고서
```

각주 본문은 문서 끝의 `references`, `sources`, `citations`, `works cited`, `참고문헌`, `출처` 중 하나가 포함된 섹션 아래에 `1. 인용문` 형식으로 둔다.

### 구분선

파서가 인식하는 본문 구분선은 `---`와 `## ---` 두 형태다.

```markdown
---

## ---
```

`--separator-mode auto`에서 `---`는 가로줄, `## ---`는 페이지 나눔이다.
frontmatter의 시작·끝 `---`는 문서 최상단 YAML 경계이므로 본문 구분선과 혼동하지 않는다.

### 차트 (opt-in)

`--charts` 플래그를 켰을 때만 `chart` 언어 펜스 안 YAML이 PNG 차트로 렌더링된다. 플래그가 없으면 코드 패널로만 표시된다. `chart_type`은 `bar`·`line`·`waterfall`(waterfall은 series 정확히 1개) 중 하나이고, `labels`는 문자열 또는 숫자(따옴표 없는 연도도 자동으로 문자열 변환), `series`는 `name`+`values`(labels와 길이 일치) 목록이다. 스펙이 잘못되면 렌더러가 조용히 실패하지 않고 코드 패널로 폴백한다.

````markdown
```chart
chart_type: bar
labels: [2024, 2025, 2026]
series:
  - name: 매출
    values: [120000, 135000, 150000]
```
````

## 플래그 결정 트리

아래 순서로 필요한 플래그만 고른다.

| 입력 또는 산출물 상태 | 판정과 조치 |
|---|---|
| 헤딩·문단·표가 이미 분리됨 | `--format` 없이 변환 |
| 한 줄 붙여넣기, 소수의 긴 줄, 구조 붕괴 | `uv run md_formatter.py --check "입력.md"` 실행 후 필요하면 `--format` |
| DeepResearch 전용 마커가 감지됨 | `--deepresearch-cleaner auto --cite-mode footnote` 추가 |
| 표지 불필요 | `--no-cover` 추가 |
| 짧은 문서라 목차 불필요 | `--no-toc` 추가 |
| 면책문구가 요구되지 않음 | `--no-disclaimer` 추가 |
| 두 구분선 의미를 함께 보존 | `--separator-mode auto` 유지 |
| 모든 구분선을 가로줄로 통일 | `--separator-mode rule` 사용 |
| 모든 구분선을 페이지 나눔으로 통일 | `--separator-mode page-break` 사용 |
| 문서 유형이 정해져 있음(IB 리포트·텀시트·법률메모·강의노트) | `--preset ib-report|termsheet|legal-memo|lecture-note`(또는 커스텀 YAML 경로) |
| 하우스 스타일과 다른 색상·폰트 프로필 필요 | `--theme <name>`(`themes/<name>.yaml`) 또는 커스텀 YAML 경로 |
| `chart` 펜스를 이미지로 렌더링해야 함 | `--charts` 추가(미지정 시 코드 패널로 폴백) |

DeepResearch 전용 마커는 `cite...`, `entity...`, `image_group...` 같은 전용 블록을 뜻한다.
일반 Markdown에는 클리너 플래그를 습관적으로 붙이지 않는다.
표지, 목차, 면책문구는 기본적으로 유지하고 사용자 요청이나 산출물 목적이 분명할 때만 생략한다.
필드별 우선순위는 명시적 CLI 플래그 > `--preset` > 내장 기본값이며, 프리셋은 `--no-cover`/`--no-toc`/`--no-disclaimer`로 이미 끈 섹션을 다시 켜지 못한다.

## 실행 절차

1. 입력 파일이 저장소 안에 있고 `.md` 확장자인지 확인한다.
2. Authoring 규약을 점검하고 필요한 경우 원본 Markdown만 수정한다.
3. 플래그 결정 트리에 따라 옵션을 확정한다.
4. 입력과 출력 경로를 모두 명시해 변환한다.

기본 명령은 항상 다음 형태로 실행한다.

```bash
uv run md_to_word.py "입력.md" "출력.docx"
```

플래그가 필요하면 출력 경로 뒤에 붙인다.

```bash
uv run md_to_word.py "원시 보고서.md" "IB 보고서.docx" --format --deepresearch-cleaner auto --cite-mode footnote --separator-mode auto
```

한국어, 공백, 괄호가 포함된 파일명과 경로는 반드시 따옴표로 감싼다.
출력 파일이 Word에서 열려 잠긴 경우 엔진이 타임스탬프 접미사를 붙인 새 이름으로 저장하는 것은 정상 동작이다.
이때 로그의 실제 `Saved` 경로를 후속 QA 대상으로 사용한다.

## QA 루프

1. 변환 명령 직후 exit code가 성공인지 확인한다. PowerShell에서는 `$LASTEXITCODE -eq 0`을 확인한다.
2. 로그에 기록된 실제 산출 경로에 `.docx`가 존재하는지 `Test-Path -LiteralPath "출력.docx"`로 확인한다.
3. 기존 산출물을 읽기 전용으로 열어 표 개수와 헤딩 존재 여부를 스팟체크한다.

`uv run python` REPL에서 먼저 `import sys`를 실행한 뒤, 다음 세 줄을 붙여넣는다.

```python
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document; doc = Document("출력.docx")
print(f"tables={len(doc.tables)}; has_heading={any((p.style.name or '').startswith('Heading') for p in doc.paragraphs)}")
```

이 코드는 생성이 아니라 기존 산출물 검사용이다.
예상한 표가 없거나 헤딩이 인식되지 않으면 렌더러를 건드리지 말고 원본 Markdown 구조를 수정해 다시 변환한다.
같은 검증 실패에 대한 수정은 최대 두 번까지만 시도한다.
두 번 실패하면 추가 수정을 중단하고 저장소 `CLAUDE.md` §6의 검증 실패 프로토콜에 따라 보고한다.
보고에는 실패 출력 전문, `git diff` 요약, 원인 가설을 포함하며 통과할 때까지 조용히 반복하지 않는다.

## 금지 사항

- `python-docx`로 새 `.docx`를 직접 생성하지 않는다. 반드시 기존 변환 엔진을 사용한다.
- `IBStyle` 또는 `ib_renderer.py`를 수정하라고 지시하지 않는다.
- 하우스 스타일을 이 스킬에 복제하거나 색상·폰트·여백을 하드코딩하지 않는다.
- Word에서 Markdown을 추출하는 요청은 이 스킬 범위 밖이다. README의 범위 외 안내에 따라 `kordoc` 같은 전용 외부 도구를 안내한다.
- `reports/`의 원문, 고객 정보, 내부 분석 자료를 외부 서비스나 공개 위치로 전송하지 않는다.
- 변환 실패를 숨기거나, QA를 생략하거나, 성공하지 않은 산출물을 완료로 보고하지 않는다.
