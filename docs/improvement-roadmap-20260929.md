# Python 하한 결정이 엔진의 다음 단계를 연다

> 작성 기준일: 2026-09-29. 대상: IB_report_formatter 소유자(코드 유지보수자 겸 DCM 실무자). 이 문서는 리서치 노트 4건(`tooling_landscape`, `korean_document_requirements`, `ib_document_conventions`, `qa_security_llm_integration`)을 종합한 개선 로드맵이며 구현 계획이나 검증 기록이 아니다. 규모는 1인 기준 추정치(S = 1–3일, M = 1–2주, L = 3주 이상, 테스트·페이지 검사 포함)이고 출처가 아닌 작성자 판단이다. Word→MD 역변환은 소유자가 폐기했으므로 어떤 형태로도 제안하지 않는다.

다음 단계의 핵심은 기능을 더 붙이는 일이 아니다. **엔진이 만든 DOCX를 Word 안에서 계속 편집하고 검증할 수 있는 문서로 만드는 일**이 핵심이다. 가장 큰 변수는 소유자 결정 하나다. Python 3.8 하한을 유지하면 markdown-it-py는 2023-06의 3.0.0, python-docx는 1.1.2(댓글 API 없음)에 고정된다. 3.10으로 올리면 최신 AST 파서, Word 댓글, docxcompose 2.x, python-hwpx가 한 번에 열린다 ([PyPI JSON API](https://pypi.org/pypi/markdown-it-py/json); [python-docx HISTORY](https://raw.githubusercontent.com/python-openxml/python-docx/master/HISTORY.rst)). 3.8은 2024년 10월에 수명이 끝났다 ([Python devguide](https://devguide.python.org/versions/)). 결정과 무관하게 지금 3.8에서 착수할 수 있는 일도 많다. 세 가지가 가장 효과가 크다. 첫째는 입력 안전 한도와 이미지 경로 제한이다. 둘째는 Microsoft Open XML SDK 검증기를 CI 게이트로 두는 것이다. 셋째는 DocumentModel 스냅숏 테스트로, 이후 파서 교체의 안전망이 된다. 그다음은 Word 네이티브 편집성이다. SEQ 캡션, REF 상호참조, 한국어 어절 줄바꿈, IB 표 합계 규칙, 공문 "끝/이하 빈칸" 규칙이 여기에 속한다. HWPX 출력은 2026-05-18 정부 온나라 시스템의 개방형 문서 의무화로 수요 근거가 생겼다. 그러나 README가 명시한 제품 경계와 "단일 조립 경로" 원칙을 바꾸는 일이므로 소유자 결정 없이 착수하지 않는다. 컴플라이언스 표시(심사필, 위험고지)도 마찬가지다. 엔진은 항목이 있는지만 점검할 수 있고 법적 충분성은 판단하지 않는다는 범위를 먼저 정해야 한다.

## 소유자 결정 기록 (2026-09-29)

- **D1 승인:** 지원 Python 하한을 **3.12**로 올린다(본문 권고안 3.10보다 높게 소유자가 결정). CI는 3.12·3.13에서 검증한다. 이 결정으로 E12(AST 파서 전환)와 E13(Word 댓글)의 선행 조건이 풀린다. 하한 변경 자체는 의존성 업그레이드·코드 현대화와 분리된 PR로 진행한다.
- **D2 결정: DOCX를 먼저 충분히 고도화한다.** HWPX 출력(E14)은 보류하고 README의 제품 경계 문구를 유지한다. 우선순위는 P0~P2의 DOCX 품질·편집성 항목이다. HWPX는 DOCX 고도화가 끝난 뒤 수요와 형식 독립 계층 설계를 다시 검토할 때 재논의한다.
- **표지 없는 문서의 제목:** 표지를 렌더링하지 않을 때(`--no-cover`, termsheet·legal-memo 프리셋 등) 문서 제목을 본문 맨 위에 둔다. 이 로드맵과 별도로 구현한다.
- D3~D5는 미결이다.

## 로드맵 한눈에 보기: 소유자 결정 5건, 엔지니어링 17건

현재 상태는 다음과 같다. 프로필 6종, YAML 표 사양, 네이티브 번호·각주·TOC 필드, strict 모드, 번호 참조까지 보는 `docx_audit`, Windows Word COM 페이지 이미지 QA, 약 450개 테스트, Ubuntu·Windows CI가 이미 있다. PR #7(입력 손실·저장 안전·성능 보강 및 CI), PR #8(MIT LICENSE), 테마·프리셋·차트 이식 브랜치가 진행 중이다. 아래 표는 이를 전제로 한다. 저장소 `main`을 확인한 결과 **표 머리행 반복(`w:tblHeader`), 이미지 대체텍스트(`docPr/@descr`), `w:updateFields`, `yaml.safe_load`, 의미 북마크, 엔진 서명용 사용자 지정 속성, 공문 "끝." 배치, 표 단위 가로 구역**은 이미 구현돼 있다. 따라서 이 로드맵은 이들을 다시 제안하지 않고 그다음 단계만 다룬다.

| ID | 방향 | 우선순위 | 선행 조건 / 막는 결정 | 규모 | 검증 가능한 완료 기준(요약) |
|---|---|---|---|---|---|
| D1 | Python 하한 3.8 → 3.12(권고안 3.10에서 상향) | **승인(2026-09-29)** | 소유자 | — | AGENTS.md·pyproject·CI 매트릭스 동시 갱신 |
| D2 | HWPX 출력을 제품 경계에 포함할지 | **보류: DOCX 우선(2026-09-29)** | 소유자, D1 | — | README 범위 문구 유지 |
| D3 | 컴플라이언스 표시 범위(존재 점검 vs 미지원) | 결정 | 소유자 | — | 지원 항목과 비보증 문구 문서화 |
| D4 | 음수 표시 기본값(빨강 → 괄호) 등 출력 기본값 변경 | 결정 | 소유자 | — | 프로필별 기본값 표 확정 |
| D5 | 배포 채널(PyPI, Agent Skill, MCP) | 결정 | 소유자, PR #7·#8 병합 | — | 패키지명·공개 API 범위 확정 |
| E1 | 입력 안전 한도와 asset-root 이미지 경로 제한 | P0 | 없음(3.8 가능) | S–M | 경로 탈출·한도 초과 fixture를 strict가 코드화된 진단으로 거부 |
| E2 | Open XML SDK 검증기 CI 게이트 | P0 | 없음 | S | 모든 sample 산출물 0 errors, 검증기 버전 고정 |
| E3 | DocumentModel·정규화 XML 스냅숏과 Hypothesis 속성 테스트 | P0 | 없음 | M | 전 sample 스냅숏 존재, 속성 4종 이상 통과 |
| E4 | 한국어 조판 플래그(어절 줄바꿈 등) | P1 | 없음 | S | styles.xml 속성 확인 + 대표 페이지 검사 |
| E5 | SEQ 캡션·REF 상호참조·표/그림 목록 | P1 | 없음 | M | F9 후 번호 불변, 미해결 참조 strict 거부 |
| E6 | IB 표 의미 확장(합계선, NM, 파생 %) | P1 | D4(기본값) | M | 명시 좌표만 적용, 일반 프로필 비추론 테스트 |
| E7 | 공문(office-letter) 정밀도 | P1 | ㉮ 번호 형식 검증 | M | 표로 끝나는 본문 규칙·붙임 형식 회귀 테스트 |
| E8 | 문서 속성·DRAFT 표시 | P1 | 워터마크 마크업 검증 | S–M | Word 디자인 > 워터마크로 제거 가능 |
| E9 | 접근성 진단 | P1 | 없음 | S | 대체텍스트 누락·대비 미달이 warning으로 보고 |
| E10 | DCM 차트 유형(PNG) | P1 | 차트 브랜치 병합 | M | 만기 프로필·트랜치 스택·이중축 샘플 페이지 검사 |
| E11 | 스타일 계약과 reference-doc | P2 | 제품 경계 확인 | M | 참조 문서 교체만으로 서식 변경, 코드 무변경 |
| E12 | AST 파서(markdown-it-py) 전환 | P2 | **D1** | L | 레거시 대비 DocumentModel diff 0 또는 승인 목록 |
| E13 | Word 댓글 출력 | P2 | **D1**(python-docx 1.2) | S–M | 댓글이 Word 검토 창에 표시 |
| E14 | HWPX 출력 변환기 | 보류 | **D1, D2(보류)** | L | 한컴오피스 열림·렌더 검사 통과 |
| E15 | 편집형 네이티브 Word 차트 | P3 | E10, PR #392 호환성 검증 | L | Word "데이터 편집" 시 임베디드 통합문서 일치 |
| E16 | Excel 범위 연동과 숫자 대사(reconciliation) | P3 | 의존성 추가 결정 | L | 원천 불일치 숫자를 strict가 보고 |
| E17 | Agent Skill → MCP 서버, 패키징 | P3 | D5 | M | JSON 진단 스키마 버전 고정, Trusted Publishing |

## 소유자 결정 다섯 건이 로드맵의 절반을 잠그고 있다

이 절의 항목은 엔지니어링으로 해결할 수 없다. 제품 경계, 호환성 약속, 법적 책임 범위를 바꾸는 일이기 때문이다. 각 결정은 AGENTS.md나 README 문구 변경으로 기록되어야 한다.

### D1. Python 3.10 하한은 파서·댓글·HWPX를 한 번에 연다

모든 선택지를 동시에 묶는 제약은 Python 3.8 호환 규칙이다. PyPI 릴리스 이력상 3.8 지원의 마지막 버전은 다음과 같다 ([PyPI JSON API](https://pypi.org/pypi/markdown-it-py/json)). **markdown-it-py 3.0.0(2023-06)**, mdit-py-plugins 0.4.2(2024-09), marko 2.2.0, **python-docx 1.1.2(2024-05)**. 이후 버전의 요구 조건은 이렇다. markdown-it-py 4.0.0부터 `>=3.10`이다. 4.x에는 참조 파서의 이차 복잡도 수정과 "gfm-like2" 프리셋이 들어 있다 ([markdown-it-py releases](https://github.com/executablebooks/markdown-it-py/releases)). python-docx 1.2.0은 "Add support for comments"와 "Drop support for Python 3.8"을 함께 발표했다 ([python-docx HISTORY](https://raw.githubusercontent.com/python-openxml/python-docx/master/HISTORY.rst)). docxcompose 2.2.0은 `>=3.10`이다 ([PyPI docxcompose](https://pypi.org/pypi/docxcompose/json)). python-hwpx도 `>=3.10`이다 ([PyPI python-hwpx](https://pypi.org/project/python-hwpx/)). 3.8에 남는 비용은 누적된다. 규칙 기반 파서는 이미 펜스, 이스케이프, 목록, 참조 링크에서 반복 수정이 필요했다. 반면 3.8 호환 AST 파서 중 활발히 배포되는 것은 mistune 3.3뿐이다. mistune은 스스로 "sane CommonMark"이며 토큰 우선순위와 이스케이프 같은 복잡한 경우를 처리하지 못한다고 밝힌다 ([Mistune docs](https://mistune.lepture.com/en/v3/)). 이미 수동으로 고쳐 온 종류의 경계 사례가 다시 생긴다는 뜻이다.

**권고: 3.10으로 올린다.** 대안은 3.8 코어를 유지하고 3.10 전용 기능을 optional extra와 런타임 버전 게이트로 분리하는 것이다. 이 방식은 CI 매트릭스와 테스트 분기가 두 배가 된다. 1인 유지보수 프로젝트에는 부담이 크다. 결정 후 해야 할 일은 네 가지다. AGENTS.md의 "Python 3.8+ compatible"과 `Path.with_stem()` 금지 조항을 개정하고, `pyproject.toml`의 `requires-python`을 올리고, CI 매트릭스를 3.10–3.13으로 바꾸고, 3.8 사용자에게 마지막 호환 릴리스를 명시한다. 완료 기준은 CI에서 3.8 job이 제거되고 3.10 최저 버전 job이 녹색인 것이다.

### D2. HWPX는 수요 근거가 생겼지만 "단일 조립 경로" 원칙과 충돌한다

수요 근거는 분명해졌다. **2026-05-18부터 중앙부처와 지자체는 온나라 문서시스템에서 개방형(HWPX 기반) 문서를 써야 하고, 개방형 파일만 첨부할 수 있다.** 이는 2026-05-12 국무회의 의결 사항이다 ([ZDNet Korea](https://zdnet.co.kr/view/?no=20260512173412); [경향신문](https://www.khan.co.kr/article/202605121345001)). 같은 날 시행된 「행정업무의 운영 및 혁신에 관한 규정」 제5조②은 행정기관 문서에 "개방형 문서 형식"과 문서요지·키워드 포함을 요구한다 ([law.go.kr 행정업무규정](https://www.law.go.kr/법령/행정업무의운영및혁신에관한규정)). DCM 업무에는 공공기관 발행사, 지자체, 공사채 거래상대방이 있다. 이들에게 가는 문서라면 HWPX 출력의 가치가 실재한다. 구현 수단도 생겼다. python-hwpx 6.6.0(2026-09-27)은 순수 Python으로 HWPX를 새로 만들 수 있다. Apache-2.0이지만 Alpha 상태이고 Python ≥3.10을 요구한다 ([PyPI python-hwpx](https://pypi.org/project/python-hwpx/)). kordoc(MIT, TypeScript)은 이미 Markdown에서 공문 프리셋으로 HWPX를 생성하므로 선행 사례이자 비교 대상이다 ([GitHub kordoc](https://github.com/chrisryugj/kordoc)).

막는 요소는 기술보다 설계에 있다. README.ko.md는 HWP/HWPX 출력을 후속 대상으로 분류해 두었다. AGENTS.md는 "CLI/API/registry must share `IBDocumentRenderer.render`; do not reintroduce a second document assembler"를 요구한다. HWPX는 OWPML이라는 다른 XML 어휘로 써야 하므로 렌더러가 하나 더 필요할 수밖에 없다. 따라서 HWPX를 받아들이려면 원칙을 재정의해야 한다. 프로필 의미(번호 체계, 표 의미, 메타데이터, 끝/붙임 배치)를 형식 독립 계층에서 한 번만 결정하고, DOCX와 HWPX 작성기는 그 결과만 직렬화하도록 바꾸는 것이다. 이는 약 3,400줄인 `ib_renderer.py`를 분할하는 작업을 전제로 한다. 결정 전에 할 수 있는 저비용 중간 단계도 있다. 한컴오피스에서 잘 열리는 DOCX를 보장하고 "DOCX → 한글에서 HWPX로 저장" 경로를 문서화하는 것이다. 노트에서 이 중간 단계의 품질은 검증되지 않았다. HWPX 리더를 추가해 Word→MD 방향을 우회적으로 되살리는 일은 제품 경계상 금지다.

### D3. 컴플라이언스 표시는 "존재 점검"까지만 약속해야 한다

금융투자협회 「금융투자회사의 영업 및 업무에 관한 규정」에는 템플릿으로 만들 수 있는 구체적 요소가 있다. 제2-47조는 "OO회사 준법감시인 심사필 제 호 (20 . . ~ 20. . .)" 형식의 심사필 표시를 정한다 ([KOFIA 제2-47조](https://law.kofia.or.kr/service/law/detailArticlePrint.do?seq=136&historySeq=1474&contentSeq=238655)). 제2-37조는 원금손실 가능성과 예금자보호 여부 등 위험고지를 **9포인트 이상**으로, 바탕색과 구별되게 표시하도록 한다 ([KOFIA 규정 historySeq 1374](https://law.kofia.or.kr/service/law/lawFullScreenContent.do?seq=136&historySeq=1374)). 조사분석자료에는 분석사 이름, 재산적 이해관계, 2년간 투자등급·목표가격 변동과 그래프, 준법감시인 사전 승인이 요구된다 ([KOFIA 영업 및 업무에 관한 규정](https://law.kofia.or.kr/service/law/lawFullScreenContent.do?seq=136&historySeq=296)). 금융소비자보호법 제22조는 광고에 설명서 확인 권유와 과거 실적의 비보장 문구를 요구한다 ([국가법령정보센터 금소법](https://www.law.go.kr/LSW/lsInfoP.do?lsId=013704&ancYnChk=0)).

소유자가 정할 범위는 둘로 나뉜다. 하나는 엔진이 `compliance:` frontmatter 블록을 받아 고정 위치에 렌더링하고 strict 모드에서 **존재 여부, 유효기간 경과, 글자 크기 9pt 미만만 점검**하는 것이다. 다른 하나는 아예 지원하지 않는 것이다. 권고는 전자다. 다만 두 가지 불확실성이 있다. 첫째, 노트가 참조한 협회 규정 판(historySeq)은 최신본이 아닐 수 있다. 둘째, **전문투자자 대상 ABCP 설명자료 같은 DCM 기관용 자료가 투자광고로서 심사필 대상인지에 대한 권위 있는 근거를 찾지 못했다.** 따라서 이 기능은 "법적 충분성을 판단하지 않는다"는 비보증 문구와 함께 출시해야 한다. 이는 AGENTS.md의 "Diagnostics do not establish visual or financial correctness" 원칙과도 일치한다. 샘플은 가상 회사명과 가상 심사필 번호로만 만들고, 실제 딜 문서는 fixture로 쓰지 않는다.

### D4·D5. 출력 기본값과 배포 채널은 사용자 계약을 바꾼다

현재 IB 표의 "음수 빨강"은 회사 선호에 가깝다. 널리 따르는 관행은 괄호 표시이고, 빨강은 흑백 인쇄에 불리하다 ([BIWS Excel Formatting Best Practices](https://palikhov.wordpress.com/wp-content/uploads/2019/11/biws-excel-formatting-best-practices.pdf); [Financial Edge](https://www.fe.training/free-resources/financial-modeling/financial-model-formatting/)). 한국 정부·예산 문서에서는 `△`가 감소를 뜻한다 ([NOON](https://www.noononda.com/news/999)). 기본값을 괄호로 바꾸면 기존 사용자의 출력이 달라지므로 소유자 결정 사항이다. 표시 방식을 표 단위 옵션(`-`, `( )`, `△`)으로 제공하는 일은 엔지니어링이다(E6). 이때 `△`는 증감 화살표 ▲▼와 섞이지 않아야 하고, 값의 크기를 바꾸지 않아야 한다. D5는 E17의 전제다. PyPI 공개 여부, 패키지명, 그리고 CLI 플래그와 JSON 진단 스키마를 SemVer 공개 API로 선언할지를 정해야 한다.

## P0: 3.8에서 지금 바로 착수할 검증·안전 기반

이 세 항목은 어떤 결정과도 무관하다. 동시에 이후 모든 변경(파서 교체, HWPX, 네이티브 차트)의 회귀 위험을 낮춘다. PR #7과 차트 브랜치를 병합한 직후가 착수 시점이다.

### E1. LLM이 쓴 Markdown이 로컬 파일을 문서에 실어 내보내는 경로를 막는다

현재 `md_parser.py`는 상대 이미지 경로를 MD 파일 폴더 기준으로 해석한다. 절대 경로는 그대로 받아들인다. 기준 폴더 밖으로 나가는 경로를 막는 장치는 `main`에서 확인되지 않았다. PR #7이 이를 다뤘는지는 병합 전에 확인해야 한다. DOCX는 기계 밖으로 나가는 산출물이다. 따라서 프롬프트 인젝션으로 생성된 `![](C:/Users/.../x.png)` 같은 참조는 **로컬 파일 유출 경로**가 된다(노트의 추론). OWASP는 원격 URL을 가져올 때 호스트 허용 목록, 리다이렉트 금지, 사설·메타데이터 IP 차단을 권고한다 ([OWASP SSRF Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/Server_Side_Request_Forgery_Prevention_Cheat_Sheet.html)). 자원 고갈도 실제 사례가 있다. 56KB짜리 .docx 하나가 153초의 CPU를 소비한 사례가 보고됐다 ([PersonalClaw #2747](https://github.com/PersonalClaw/PersonalClaw/issues/2747)). python-docx는 `resolve_entities=False`로 XXE는 막지만, 엔티티 확장이나 블록 수 폭증 같은 DoS 한도는 호출자 책임이다 ([DEV: XXE in document-processing APIs](https://dev.to/roxdavirox/xxe-in-document-processing-apis-the-attack-surface-nobody-hardens-ih5)).

구현 범위는 다음과 같다. `--asset-root`(기본값은 MD 파일 폴더)를 두고, `Path.resolve()` 후 그 밖을 가리키는 절대 경로, UNC 경로, 심볼릭 링크는 기본적으로 거부한다. 원격 이미지는 계속 비활성으로 둔다. 입력 바이트, 표 행×열, 목록 중첩 깊이, 이미지 픽셀·바이트, base64 크기에 한도를 둔다. 테마나 템플릿처럼 엔진이 직접 파싱하는 XML은 `no_network=True, load_dtd=False`로 강화한 lxml 파서를 거치게 한다 ([CodeQL py/xxe](https://codeql.github.com/codeql-query-help/python/py-xxe/)). **완료 기준:** 경로 탈출(`../`, 절대 경로, UNC, symlink) fixture 4종과 한도 초과 fixture가 strict 모드에서 고유 진단 코드로 거부되고, 기존 samples 전체는 계속 통과한다.

### E2. Microsoft 검증기는 "잘 열리는" 문서의 절반 가까이를 불합격시켰다

`docx_audit`는 이 프로젝트 고유의 구조 검사다. 그러나 OOXML 스키마 적합성은 보지 않는다. 가장 강한 근거는 dolanmiu/docx 사례다. 이 프로젝트는 2026-09 CI에 Open XML SDK 검증기를 붙였다. 그러자 **데모 119개 중 54개가 정상적으로 열리는데도 스키마 검증에 실패**했다. 수정은 numbering, styles, 요소 순서, 속성값 전반에 걸쳤다 ([dolanmiu/docx PR #3543](https://github.com/dolanmiu/docx/pull/3543)). 이 엔진에서 손으로 조립하는 XML(번호 정의, 각주, 필드 코드, 사용자 지정 속성, OMML)이 바로 같은 위험 지점이다. 도구로는 .NET CLI인 OOXML-Validator가 있다. JSON으로 출력하고 `Office2019`/`Microsoft365` 대상을 고를 수 있다 ([OOXML-Validator README](https://github.com/mikeebowen/OOXML-Validator/blob/main/README.md)). npm 래퍼 `@xarsh/ooxml-validator`도 있다. Docxtor는 검증기 버전을 고정하고 종료 코드 0/1/2로 결과를 구분했다 ([Docxtor PR #198](https://github.com/mikolaj92/Docxtor/pull/198)). **완료 기준:** 별도 CI job에서 버전을 고정한 검증기가 `samples/profiles`와 `samples/qa`의 모든 산출물에 대해 0 errors를 내고, 오류의 part URI와 XPath가 JSON 아티팩트로 올라간다. 제품 의존성에는 .NET이나 npm을 추가하지 않는다. 사전 확인할 사항은 OOXMLValidatorCLI가 사전 빌드 바이너리나 `dotnet tool` 패키지를 제공하는지다(노트 Gap).

### E3. 파서를 바꾸기 전에 "무엇이 바뀌었는지"를 측정할 장치가 먼저다

DOCX는 타임스탬프와 rsid가 섞인 zip이다. 그래서 바이트 단위 골든 파일은 깨지기 쉽다. 스냅숏 대상은 **파싱된 DocumentModel의 JSON**과 **정규화한 `document.xml`/`numbering.xml`/`styles.xml`**이어야 한다. syrupy는 스냅숏이 없을 때도 실패하고 가변 ID용 matcher를 제공한다 ([syrupy](https://github.com/syrupy-project/syrupy)). pytest-regressions도 대안이다 ([pytest-regressions](https://pytest-regressions.readthedocs.io/en/latest/overview.html)). Hypothesis 속성 테스트는 AGENTS.md 규칙과 일대일로 대응시킬 수 있다. 가시 텍스트 보존, 코드와 이스케이프 불변, 빈 셀을 포함한 셀 수 보존, strict 모드의 "성공 또는 거부" 이분성, `md_formatter` 멱등성, 셀 숫자 표기의 바이트 보존이 그 대상이다. markdown-it-py도 CommonMark 스펙 테스트와 OSS-Fuzz를 함께 운영한다 ([markdown-it-py AGENTS.md](https://github.com/executablebooks/markdown-it-py/blob/master/AGENTS.md)). atheris 퍼징은 Linux 야간 job에 적합하다 ([OSS-Fuzz Python guide](https://google.github.io/oss-fuzz/getting-started/new-project-guide/python-lang/)). **완료 기준:** 모든 sample에 DocumentModel 스냅숏이 있고, 위 속성 중 최소 4종(텍스트 보존, 빈 셀 보존, 이스케이프 불변, strict 이분성)이 Windows와 Ubuntu CI 모두에서 통과한다.

LibreOffice 렌더 비교는 이 단계의 선택 항목이다. Malgun Gothic과 메트릭이 호환되는 무료 글꼴은 찾지 못했다. LibreOffice 버전이 오르면 줄바꿈과 표 분할도 달라진다 ([DEV: replacing headless LibreOffice](https://dev.to/nixan/we-replaced-headless-libreoffice-with-a-single-rust-binary-for-docx-pdf-7po)). 그러므로 Linux 렌더는 **Word PNG가 아니라 고정된 Docker 이미지로 만든 Linux 기준선과만 비교하는 비차단 드리프트 탐지기**로 둔다. 합격 증거는 계속 Word COM 페이지 검사만 인정한다 ([next-gen-editor #89](https://github.com/IbraheemAlz/next-gen-editor/issues/89)).

## P1: Word 네이티브 편집성과 한국어 조판이 체감 품질을 가른다

뱅커가 결과물을 받은 뒤 Word에서 번호를 고치고, 표를 옮기고, 버전을 바꾸는 순간에 엔진의 가치가 드러난다. 이 단계의 항목은 모두 3.8에서 가능하다. 모두 출력이 바뀌므로 AGENTS.md에 따라 실제 파싱·렌더 회귀 테스트와 대표 페이지 검사가 필수다.

### E4. `w:wordWrap=0` 한 줄이 한국어 문서의 가장 눈에 띄는 결함을 없앤다

> **2026-09-29 정정:** Word 실측 결과 방향이 반대다. `w:wordWrap w:val="0"`은 한글을 음절 단위로 끊고, 요소가 없거나 `1`이면 어절 단위로 끊는다. 텀싯 프로필은 `wordWrap=1`과 `autoSpaceDE`/`autoSpaceDN=0`을 쓴다([텀싯 계획 §2-18](term-sheet-plan-20260929.md)). 아래 본문은 작성 당시 기록으로 남긴다.

Word는 기본적으로 한국어를 음절 단위로 끊는다. 한 어절이 두 줄로 쪼개지는 현상이 여기서 나온다. `w:wordWrap`은 Word의 "한글 단어 잘림 허용" 옵션에 해당하며, Word는 이 속성을 **한국어 텍스트에만** 적용한다 ([MS-OI29500 §17.3.1.45](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5c96914d-820a-4f49-a239-40024839685e)). 그래서 영문 문단에는 부작용이 없다. `w:kinsoku`는 행두·행말 금칙을 한·중·일 텍스트에 적용한다 ([MS-OI29500 kinsoku](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/84486937-c3e4-4d37-8672-68892080944d)). `w:autoSpaceDE`/`autoSpaceDN`은 한글과 영문·숫자 사이 간격을 자동으로 조정한다. 재무 표의 숫자 간격을 고르게 하려고 한국 템플릿이 이를 끄는 경우가 있으므로 프로필별 선택 사항으로 둔다 ([datypic w:pPr](http://www.datypic.com/sc/ooxml/e-w_pPr-6.html)). 현재 엔진은 `w:eastAsia` 글꼴은 설정하지만 이 줄바꿈 속성들은 쓰지 않는다. 추가할 것은 네 가지다. 한국어 프로필의 본문, 목록, 표 셀 스타일에 `wordWrap=0`과 `kinsoku`를 명시하고, `w:lang w:eastAsia="ko-KR"`을 지정하고, 숫자와 영문은 `ascii`/`hAnsi` 글꼴로 분리한다. **완료 기준:** 한국어 프로필 산출물의 styles.xml에서 해당 속성이 확인되고, 대표 샘플의 Word COM 페이지 이미지에서 어절 중간 분리가 0건이며, 영문 IB 샘플의 쪽수가 변하지 않는다. 한국어 Normal.dotm의 kinsoku와 autoSpace 기본값은 미확인이다. 그러므로 명시적으로 켜고 끄는 방식이 안전하다.

### E5. SEQ 캡션과 REF 상호참조는 Pandoc도 DOCX에서 하지 않는 일이다

Pandoc은 `native_numbering`으로 DOCX 표와 그림에 카운터 필드를 붙인다. 그러나 새로 고칠 수 있는 상호참조 필드(`xrefs_name`/`xrefs_number`)는 **odt 출력에서만** 지원한다 ([Pandoc manual](https://pandoc.org/MANUAL.html)). python-docx 엔진이 캡션 SEQ 필드와 `_Ref` 북마크, 그리고 REF/PAGEREF 필드를 함께 내보내면 이 영역에서 Pandoc을 앞선다. Word 문법은 `{ SEQ 표 \* ARABIC }`이고, 장 번호를 붙인 캡션에는 `\s` 스위치를 쓴다 ([Microsoft Support: SEQ](https://support.microsoft.com/en-US/Word/field-codes-seq-sequence-field)). SEQ 번호를 교차참조 대화상자에 나타나게 하려면 같은 이름의 캡션 레이블이 있어야 한다 ([Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/4762967/cross-referencing-a-seq-field)). 작성 문법은 새로 만들지 않는다. Quarto의 ID 접두사 관례(`{#tbl-x}`, `@tbl-x`, `#fig-`, `#sec-`)를 그대로 쓴다 ([Quarto Cross References](https://quarto.org/docs/authoring/cross-references.html)). 작성자와 LLM이 이미 아는 문법이라 학습 비용이 없다. 엔진에는 이미 의미 북마크와 `w:updateFields`가 있으므로 확장 비용이 낮다. 필드에는 미리 계산한 캐시 결과를 넣어, F9를 누르기 전에도 올바르게 읽히게 한다. 표 목록과 그림 목록은 `{ TOC \h \z \c "표" }`로 만든다. **완료 기준:** 샘플 문서에서 Word F9 후에도 번호와 참조 텍스트가 변하지 않고, 표 목록이 생성되며, 존재하지 않는 ID를 참조하면 `docx_audit`가 issue로 보고하고 strict 모드가 저장을 거부한다.

### E6. IB 표는 "명시한 좌표에만" 합계선·NM·파생 % 규칙을 적용한다

널리 따르는 IB 표 관행은 다음과 같다. 합계 위에는 단선, 아래에는 이중선을 긋고 합계는 굵게 한다 ([Financial Edge](https://www.fe.training/free-resources/financial-modeling/financial-model-formatting/)). 계산된 %는 이탤릭으로 하되 가정 %는 이탤릭으로 하지 않는다. 배수는 `0.0x`로 쓰고, 음수이거나 100x를 넘는 배수는 "NM"으로 표시한다. 통화 기호는 첫 행과 합계 행에만 둔다 ([BIWS PDF](https://palikhov.wordpress.com/wp-content/uploads/2019/11/biws-excel-formatting-best-practices.pdf)). 비교기업 분석에서 NM 값은 평균과 중앙값 계산에서 제외한다 ([Macabacus](https://macabacus.com/valuation/comparable-companies)). BIWS 스스로 색상 코딩 외에는 회사마다 표준이 다르다고 밝힌다. 따라서 모두 **관행**이며 추론으로 적용하지 않는다. YAML 표 사양에 넣을 항목은 다음과 같다. `total_rows`/`subtotal_rows`(1부터 시작하는 본문 행 좌표), `percent` 열의 `italic: derived|none`, 문자 그대로 보존할 sentinel 토큰(`NM`, `n.m.`, `NA`, `n.a.`, `N/A`, `–`)과 자동 범례, `FY24A`/`FY25E` 추정 열 음영, 단위 캡션 `(단위: 백만원)`. 숫자 정렬은 Word의 소수점 탭(`w:tab w:val="decimal"`)을 쓴다. 이 기능은 일반 프로필 4종에는 적용하지 않는다. 금융 표 의미를 추론하지 않는다는 AGENTS.md 규칙 때문이다. **완료 기준:** 명시한 행에만 테두리와 굵게가 적용되고, 사양 없는 표와 일반 프로필 표의 출력이 바이트 단위로 이전과 같으며, sentinel 셀이 원문 그대로 오른쪽 정렬된다.

### E7. 공문 프로필은 "표로 끝나는 본문"과 8단계 번호에서 아직 규정과 어긋난다

「행정업무의 운영 및 혁신에 관한 규정 시행규칙」 제4조⑤는 본문이 표로 끝나는 경우를 따로 정한다. 표가 마지막 칸까지 채워졌으면 표 아래 왼쪽 기본선에서 한 글자 띄우고 "끝"을 쓴다. 표가 중간에 끝나면 "끝"을 쓰지 않고 마지막 기재 칸 다음 칸에 **"이하 빈칸"**을 쓴다 ([law.go.kr 시행규칙](https://www.law.go.kr/법령/행정업무의운영및혁신에관한규정시행규칙)). 현재 `office_layout.py`는 마지막 문단이나 붙임 뒤에 "  끝."을 붙이는 경로만 가지고 있고, "이하 빈칸" 처리는 저장소에서 찾을 수 없다. 편람·지침 층위의 형식도 있다. 붙임 뒤에는 콜론 없이 두 칸을 띄우고, 붙임이 하나면 번호를 매기지 않는다. 하위 항목은 2타씩 들여 쓰고, 항목 기호 뒤는 1타를 띄우며, 둘째 줄부터는 내어쓰기로 맞춘다 ([경기도교육청 공문서 작성법](https://www.goe.go.kr/resource/old/BBSMSTR_000000000028/BBS_202410150153084250.pdf)). 8단계 항목 기호 `1. 가. 1) 가) (1) (가) ① ㉮`는 시행규칙 제2조①의 법정 순서다. 선택형 lint도 추가할 수 있다. 대상은 `2026.09.29` → `2026. 9. 29.`, `오후 3시 20분` → `15:20`, `~` → `∼`, `345천원` 경고, 금액 한글 병기 도우미다. 이 lint는 코드, 이스케이프, 링크 URL을 절대 바꾸지 않아야 한다. 제5조② 대응도 비용이 작다. YAML `summary:`와 `keywords:`를 DOCX 핵심 속성에 쓰고, 필요하면 요지 상자를 보이게 한다. **완료 기준:** 가득 찬 표, 중간에 끝나는 표, 단일 붙임, 복수 붙임 네 경우의 회귀 테스트와 페이지 검사를 통과하고, 번호 정의가 8단계를 모두 네이티브로 렌더한다. 착수 전 확인할 것이 있다. **㉮(원문자 한글)에 해당하는 OOXML `ST_NumberFormat` 값이 있는지 확인되지 않았다.** 없으면 8단계는 자동 증가하지 않는 리터럴 `lvlText`로 처리해야 하며, 이 제약을 문서화해야 한다.

### E8·E9. 문서 속성, DRAFT 표시, 접근성은 저비용 편집성 항목이다

frontmatter에서 핵심 속성(title, subject, keywords, contentStatus)과 사용자 지정 속성(DealName, Version, Classification, AsOfDate)을 채울 수 있다. 머리글과 바닥글에서 `{ DOCPROPERTY Version }`으로 참조하면, 뱅커는 파일 속성에서 버전을 한 번 바꾸고 F9만 누르면 된다. 사용자 지정 속성 파트는 이미 엔진 서명용으로 존재하므로 확장만 하면 된다. `status: draft`는 DRAFT 워터마크와 머리글 스탬프를 켜고, `final`은 둘 다 끈다. Word는 워터마크를 모든 머리글의 VML 도형으로 저장하고 `PowerPlusWaterMarkObject…` ID로 인식한다고 한다. 같은 마크업을 쓰면 사용자가 Word 기본 기능으로 워터마크를 제거할 수 있다 ([dolanmiu/docx PR #3221](https://github.com/dolanmiu/docx/pull/3221)). 다만 **이 스키마는 Microsoft 1차 문서가 아닌 오픈소스 구현에서 확인한 것이다. DOCPROPERTY 지원 페이지도 404여서 검증하지 못했다.** 따라서 Word에서 워터마크를 넣어 저장한 문서를 역으로 분석한 결과를 fixture로 삼아야 한다. **완료 기준:** Word의 디자인 > 워터마크 > 제거로 삭제되고, 속성 변경 후 F9를 누르면 머리글이 갱신된다.

접근성은 한국에서 웹·앱에 대한 구속 기준(KWCAG 2.2 = KS X OT0003:2022)은 있다. 그러나 **DOCX·HWPX 문서에 대한 구속 기준은 찾지 못했다** ([KWCAG 2.2](https://a11ykr.github.io/kwcag22/)). 대체텍스트와 머리행 반복은 이미 구현돼 있다. 남은 것은 두 가지 진단이다. 하나는 대체텍스트가 비었을 때의 warning이다. 다른 하나는 흰 글씨·남색 배경 요약 상자의 명도 대비가 4.5:1에 미달할 때의 warning이다. 둘 다 구조 issue가 아니라 warning으로 분리한다. **완료 기준:** `docx_audit` JSON에 두 경고 코드가 추가되고 기존 issue 수는 변하지 않는다.

### E10. 차트 이식 브랜치 다음에는 DCM 차트 세 종이 온다

진행 중인 matplotlib 막대, 선, 워터폴 차트는 M&A 피치북 쪽 표준에 가깝다. 풋볼필드와 워터폴은 피치북의 대표 차트다 ([Wall Street Prep](https://www.wallstreetprep.com/knowledge/football-field-valuation-real-example-excel-template/)). DCM 실무에는 만기가 몰리는 "maturity tower"를 관리하고 12–18개월 전에 차환을 준비하는 서사가 핵심이다 ([ibinterviewquestions DCM guide](https://ibinterviewquestions.com/guides/debt-capital-markets)). 여기서 필요한 순서는 다음과 같다. 첫째, 상품별(CP, ABCP, ABSTB, 채권, 대출) 누적 **만기 프로필**. 둘째, 신용보강 층을 보여 주는 100% **트랜치 스택**. 셋째, 금리(%)와 스프레드(bp)의 **이중축 선 차트**. 풋볼필드는 그 뒤다. 민감도 히트맵은 이미지보다 셀 음영을 쓴 네이티브 표가 낫다. 편집이 가능하기 때문이다. ABCP 지급 우선순위 워터폴은 데이터 차트가 아니라 도식이므로 `diagram_renderer.py` 영역이다. **다만 DCM 차트의 배치와 색상을 정한 공개 권위 출처는 찾지 못했다.** 사양은 소유자의 실무 판단으로 확정해야 한다. **완료 기준:** 세 차트 유형이 `--charts` opt-in으로 렌더되고, 출처 줄과 단위 표기가 붙으며, 가상 데이터 샘플의 페이지 검사를 통과한다.

## P2: 3.10 전환 뒤 파서를 교체하고, 스타일 계약을 공개한다

### E11. 회사 템플릿은 "양식 복제"가 아니라 "스타일 참조"로 받는다

README는 회사 DOCX 원본 양식의 자동 복제를 범위 밖으로 두었다. Pandoc의 reference-doc 모델은 이와 다르다. 본문 내용은 무시하고 **스타일시트와 문서 속성(여백, 용지, 머리글·바닥글)만 가져온다** ([Pandoc manual](https://pandoc.org/MANUAL.html)). 그래서 경계와 충돌하지 않는다. 필요한 작업은 세 가지다. 엔진이 쓰는 스타일 이름의 계약(`IB Body`, `IB Heading 1–4`, `IB Table`, `IB Caption`, 콜아웃 스타일)을 공개한다. 참조 문서에 같은 이름이 있으면 덮어쓰지 않는다. `custom-style` 속성을 가진 div와 span으로 임의의 명명 스타일에 연결하는 탈출구를 둔다. 표 테두리나 색처럼 엔진이 직접 계산하는 서식을 가능한 한 명명 스타일로 옮기면, 템플릿 담당자가 코드 없이 서식을 바꿀 수 있다. 템플릿 셸에 본문을 끼워 넣는 기능은 docxcompose 2.x(≥3.10)나 docxtpl의 Subdoc이 맡는다 ([docxtpl docs](https://docxtpl.readthedocs.io/en/latest/)). `.dotx`를 여는 방식은 검증되지 않았다(콘텐츠 형식 변환이 필요할 것으로 추정). **완료 기준:** 스타일만 바꾼 참조 문서 두 벌로 같은 MD를 렌더했을 때 코드 변경 없이 글꼴과 색이 바뀌고, 계약에 없는 스타일은 Normal을 상속해 생성된다.

### E12. markdown-it-py 전환은 DocumentModel 경계 뒤에서 차등 테스트로 진행한다

mdit-py-plugins는 이 엔진에 필요한 구문 대부분을 이미 제공한다. front_matter, footnote, deflist, tasklists, container, admon, attrs(`{#id .class}`), dollarmath가 그것이다 ([mdit-py-plugins docs](https://mdit-py-plugins.readthedocs.io/en/latest/)). 4.2.0의 `make_fence_rule()`은 펜스 규칙을 포크하지 않고 사용자 정의 펜스를 등록하게 해 준다 ([markdown-it-py releases](https://github.com/executablebooks/markdown-it-py/releases)). Python 선례로는 MyST-Parser가 mistletoe에서 markdown-it-py로 옮긴 사례가 있다(PR #123). 이유는 세 가지였다. 전역 상태 때문에 스레드 안전하지 않았고, 방문자 패턴이 없었고, 확장하려면 블록 토큰 대부분을 서브클래싱해야 했다 ([MyST-Parser PR #123](https://github.com/executablebooks/MyST-Parser/pull/123/files/f20333d99e8f60b4e7a7151fef827f0aeb554318); [mistletoe-ebp docs](https://mistletoe-ebp.readthedocs.io/en/latest/)).

전환은 네 단계로 한다. 첫째, `MarkdownParser` → `DocumentModel` 경계 뒤에 새 파서를 넣고 렌더러는 건드리지 않는다. 둘째, `parser="legacy"|"mdit"` 플래그를 두고, E3의 스냅숏 코퍼스 전체에서 두 파서의 DocumentModel을 비교해 모든 차이를 "버그 수정"과 "회귀"로 분류한다. 셋째, 한국어 항목 기호(가., ①, (1))와 콜아웃 레이블(요약, 시사점, 주의, 참고)은 프로필이 제어하는 사후 변환이나 블록 규칙으로 옮긴다. 이때 일반 프로필이 짧은 번호 문장을 제목으로 승격하지 않는 규칙을 지킨다. 넷째, 줄 병합, 괄호 공백 정리, 선택형 두 칸 줄바꿈 같은 기존 문단 정규화는 파서 해킹이 아니라 렌더러 쪽 텍스트 정책으로 다시 구현한다. 블록 토큰의 `map`(원본 줄 범위)을 쓰면 strict 모드 입력 손실 진단에 줄 번호를 붙일 수 있다. 이 속성은 markdown-it 표준이지만 이번 조사에서 재검증하지는 않았다. CommonMark 스펙 예제는 HTML 동치 오라클이 아니라 **충돌·무손실 코퍼스**로만 쓴다. 엔진은 CommonMark 준수를 주장하지 않기 때문이다. **완료 기준:** 코퍼스 전체에서 두 파서의 DocumentModel 차이가 0이거나 문서화된 승인 목록에만 남고, 기본값을 `mdit`로 바꾼 뒤 한 릴리스 동안 legacy를 유지하다 제거한다.

### E13. Word 댓글은 검토 흐름을 문서 안으로 가져온다

python-docx 1.2.0부터 댓글 API가 생겼다 ([python-docx HISTORY](https://raw.githubusercontent.com/python-openxml/python-docx/master/HISTORY.rst)). Markdown의 검토 메모(예: `[본문]{.comment author="..."}` 같은 span)를 Word 댓글로 내보내면, 초안 검토 의견이 별도 메일이 아니라 문서 안의 검토 창에 남는다. 작성 문법은 E12의 attrs 플러그인에 맞춰 정한다. 가치는 있지만 D1 이후에만 가능하고 필수 기능은 아니므로 P2 후순위다. **완료 기준:** 댓글이 Word 검토 창에 작성자와 함께 표시되고, 댓글 없는 문서의 출력은 변하지 않는다.

## P3: HWPX·편집형 차트·Excel 연동·에이전트 인터페이스는 조건부 확장이다

**E14 HWPX 변환기**는 D1과 D2가 모두 승인되어야 착수한다. 설계 조건은 D2에서 말한 대로다. 프로필 의미를 형식 독립 계층에서 한 번만 결정하고, `HwpxOutputConverter`를 레지스트리에 `DocxOutputConverter`와 나란히 등록한다. 위험은 세 가지다. python-hwpx는 Alpha라 API가 바뀔 수 있다. python-hwpx가 공문 8단계 번호 정의와 "쪽 X / Y" 필드를 얼마나 지원하는지 검증되지 않았다. 레이아웃 승인은 실제 한컴오피스에서만 가능하다(LibreOffice와 H2Orestart는 충실도가 부족하다) ([GitHub python-hwpx](https://github.com/airmang/python-hwpx); [H2Orestart](https://github.com/ebandal/H2Orestart)). **완료 기준:** 프로필별 가상 샘플의 HWPX가 한컴오피스에서 복구 경고 없이 열리고, 번호, 머리글·바닥글, 표가 DOCX 출력과 같은 의미로 렌더되며, 페이지 검사 기록이 해시와 함께 남는다.

**E15 편집형 네이티브 차트**는 python-docx에 차트 API가 없는 상태가 2015년부터 이어진다는 데서 출발한다 ([python-docx #179](https://github.com/python-openxml/python-docx/issues/179)). 가장 가벼운 경로는 python-pptx 1.0.2(MIT, ≥3.8)의 `ChartXmlWriter`/`WorkbookWriter`로 `c:chartSpace`와 임베디드 xlsx를 생성하고 docx 패키지에 차트 파트를 연결하는 것이다 ([python-pptx chart data](https://python-pptx.readthedocs.io/en/stable/dev/analysis/cht-chart-data.html)). Word의 "데이터 새로 고침"은 임베디드 통합문서를 다시 읽는다. 두 데이터를 하나의 `ChartData`에서 함께 생성해야 어긋나지 않는다 ([Botched Deployments](https://botched-deployments.com/posts/python-docx-charts)). 이 경로를 보여 준 미병합 PR #392가 python-docx 1.x와 python-pptx 1.0.2에서도 동작하는지는 검증되지 않았다 ([python-docx PR #392](https://github.com/python-openxml/python-docx/pull/392)). 따라서 착수 전에 1–2일짜리 스파이크로 확인한다. matplotlib PNG는 strict 안전 경로이자 폴백으로 유지한다. **완료 기준:** Word에서 차트의 "데이터 편집"을 열면 계열 값이 원천과 일치하고, E2 검증기를 통과하며, Word COM 페이지 검사를 통과한다.

**E16 Excel 연동과 숫자 대사**는 금융 문서에서 LLM의 숫자 환각을 막는 결정론적 장치다. UpSlide와 Macabacus는 셀 하나를 본문 텍스트로 연결하는 기능을 메모 작성에 매우 유용하다고 설명한다. 반면 네이티브 OLE 링크는 행 삽입에 깨지고 통합문서 전체를 복사해 파일을 키운다고 지적한다 ([Macabacus blog](https://macabacus.com/blog/linking-from-excel-to-powerpoint-and-word); [UpSlide](https://support.upslide.net/hc/en-us/articles/360015613080-How-to-link-Excel-data-in-Word)). Python 엔진이 할 수 있는 일은 두 가지다. 하나는 렌더 시점에 xlsx 범위를 읽어 셀의 `number_format`을 적용하는 것이다. 다른 하나는 출처(파일, 시트, 범위, SHA-256, 읽은 시각)를 사용자 지정 속성에 기록하는 것이다. 실시간 링크는 할 수 없다. 제약도 있다. openpyxl은 서식이 적용된 문자열을 만들지 않으므로 서식 문법 해석기가 필요하다. `data_only=True`는 Excel이 마지막으로 저장한 캐시 값만 읽는다. openpyxl은 새 의존성이다. 대사 기능은 이렇게 동작한다. 본문과 표의 모든 숫자를 선언된 원천 집합과 비교하고, 역할(%·bp·배수)을 인식해 일치하지 않는 숫자를 진단으로 낸다. strict 모드에서 원천 매니페스트가 선언되어 있으면 실패시킨다. **이 설계는 노트의 추론이며 외부 근거는 없다.** **완료 기준:** 가상 모델 xlsx에서 한 셀을 바꾸면 해당 숫자가 불일치로 보고되고, 캐시 값이 없는 통합문서는 strict가 거부한다.

**E17 에이전트 인터페이스와 패키징**은 순서가 중요하다. 먼저 Agent Skill을 만든다. SKILL.md는 metadata 약 100토큰만 상시 로드되고 스크립트 코드는 컨텍스트에 들어가지 않는다. 그래서 `md-to-word --strict`와 `docx-audit`의 JSON 출력을 되돌려 주는 스킬만으로도 작성 → 검증 → 수정 루프가 성립한다 ([Claude Docs: Agent Skills](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview)). MCP 서버는 그다음이다. 기존 docx MCP 서버들은 `add_paragraph` 수준의 저수준 래퍼다 ([Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)). 이 방식은 단일 렌더 경로라는 이 프로젝트의 가치를 우회한다. 따라서 도구는 `list_profiles`, `validate_markdown`, `render_docx`, `audit_docx` 네 개로 제한한다. 각 도구는 `outputSchema`를 선언하고, strict 거부는 `isError: true`와 구조화된 진단 목록으로 반환한다 ([MCP spec 2025-11-25](https://modelcontextprotocol.io/specification/2025-11-25/server/tools)). 2026-09 기준으로 더 새로운 MCP 스펙 판이 있는지는 확인하지 않았다. 배포는 PyPI Trusted Publishing과 `pypa/gh-action-pypi-publish`를 쓴다. 이 조합은 PEP 740 증명을 기본으로 생성한다 ([gh-action-pypi-publish](https://github.com/pypa/gh-action-pypi-publish)). `uv publish`는 증명을 직접 생성하지 않는다 ([uv docs](https://docs.astral.sh/uv/guides/package/)). **완료 기준:** JSON 진단에 `schema_version`이 있고, TestPyPI 배포에 증명이 첨부되며, 기존 wheel 페이로드 검사가 비공개 샘플과 QA 산출물을 배제함을 계속 확인한다.

## 법정 요건과 관행을 구분해 엔진 동작을 정한다

엔진은 "규정이니 기본으로 강제"와 "관행이니 선택 사항"을 섞으면 안 된다. 민간 회사의 대외공문에 공문서 규정은 구속력이 없다. 반면 공공기관 사용자에게는 같은 규칙이 자체 문서관리 규정을 통해 구속력을 가진다. 금융투자협회 규정은 회원사의 투자광고와 조사분석자료에 적용되지만, DCM 기관용 자료에 대한 적용 여부는 확인되지 않았다.

| 항목 | 근거 등급 | 적용 대상 | 엔진 처리 원칙 |
|---|---|---|---|
| 날짜 `2026. 9. 29.`, 24시간 표기, A4, 아라비아 숫자 | 법정(규정 제7조) | 행정기관 | office-letter 기본값, 민간은 관행으로 안내 |
| 항목 기호 8단계 `1. 가. 1) 가) (1) (가) ① ㉮` | 법정(시행규칙 제2조①) | 행정기관 | 네이티브 번호로 구현, ㉮ 형식 미검증 |
| 두문·본문·결문, 붙임·"끝"·"이하 빈칸" | 법정(시행규칙 제4조) | 행정기관 | E7 규칙 엔진 |
| 금액 한글 병기 `금113,560원(금일십일만…)` | 법정 문언 "적어야 한다", 교육청 지침은 "필요시" | 행정기관 | 선택형 도우미, 문언과 실무 충돌 명시 |
| 개방형 문서 형식, 문서요지·키워드 | 법정(규정 제5조②, 2026-05-19 시행) | 행정기관 | 속성·요지 상자 opt-in |
| 온나라 HWPX 사용 의무 | 정부 방침(2026-05-18, 언론 보도 기반) | 중앙부처·지자체 | D2 결정 근거 |
| 2타 들여쓰기, 여백 30/20/15/15mm, 줄간격 123% | 편람·지침 | 행정기관 관행 | 선택형 "공문 페이지" 프리셋, 2025 편람 미확인 |
| 심사필 표시, 위험고지 9pt 이상 | 협회 규정(제2-47조, 제2-37조) | 금융투자회사 투자광고 | D3 범위 내 존재 점검만 |
| 조사분석자료 이해관계·등급 이력 공시 | 협회 규정(제2-32·33조) | 조사분석자료 | D3, 법적 충분성 비보증 |
| 설명서 확인 권유, 과거 실적 비보장 문구 | 법률(금소법 제22조) | 금융상품 광고 | D3 위험고지 세트에 포함 |
| 음수 괄호·△, 단위 캡션, 합계 이중선, 파생 % 이탤릭, NM | 관행 | 회사별 | 표 사양 옵션, 추론 금지(E6) |
| 대체텍스트, 머리행, 명도 대비 4.5:1 | 웹·앱에만 법정(KWCAG 2.2), 문서는 권고 | 전반 | 진단 warning(E9) |
| Pretendard·Noto CJK 임베딩 | 라이선스(SIL OFL) | 전반 | 글꼴 프리셋 선택 사항, 나눔 글꼴 라이선스 미확인 |

출처: [law.go.kr 행정업무규정](https://www.law.go.kr/법령/행정업무의운영및혁신에관한규정), [law.go.kr 시행규칙](https://www.law.go.kr/법령/행정업무의운영및혁신에관한규정시행규칙), [경기도교육청 공문서 작성법](https://www.goe.go.kr/resource/old/BBSMSTR_000000000028/BBS_202410150153084250.pdf), [2018 행정업무운영 편람](https://active.cbyouth.net/upload/userfile/education/1526004898@2018_%ED%96%89%EC%A0%95%EC%97%85%EB%AC%B4%EC%9A%B4%EC%98%81%20%ED%8E%B8%EB%9E%8C_%EC%B5%9C%EC%A2%85-%EB%B3%B5%EC%82%AC.pdf), [KOFIA 규정](https://law.kofia.or.kr/service/law/lawFullScreenContent.do?seq=136&historySeq=1374), [KWCAG 2.2](https://a11ykr.github.io/kwcag22/), [Pretendard](https://github.com/orioncactus/pretendard).

## 착수 전에 확인할 미검증 항목

아래 항목은 리서치 노트의 Gap이다. 해당 방향을 시작하기 전에 1차 자료나 실제 Word·한컴 동작으로 확인해야 한다. 확인 전에는 설계를 확정하지 않는다.

| 관련 방향 | 미검증 사항 | 확인 방법 |
|---|---|---|
| E7 | OOXML `ST_NumberFormat`에 ㉮㉯㉰(U+326E–327B) 형식이 있는지 | Word에서 목록 서식을 지정해 저장한 뒤 numbering.xml 분석 |
| E7 | 2025 행정업무운영 편람(2026-01-02 게시)의 여백과 123% 유지 여부 | 편람 PDF·HWPX 원문 확인 ([행정안전부](https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardArticle.do?bbsId=BBSMSTR_000000000012&nttId=122878)) |
| E7 | 규정 제2조(적용범위) 원문 | 국가법령정보센터 |
| E8 | DOCPROPERTY 필드 문법, 워터마크 VML 스키마 | Word 생성 문서 역분석 |
| E4 | 한국어 Normal.dotm의 kinsoku와 autoSpace 기본값, 맑은 고딕 fsType | Word 기본 문서 분석 |
| E11 | python-docx의 `.dotx` 로딩 동작 | 스파이크 테스트 |
| E12 | markdown-it-py 토큰 `map` 동작, 4.x 대비 mistune과 marko 성능 | 코퍼스 벤치마크 |
| E14 | python-hwpx의 공문 8단계 번호와 쪽번호 필드 지원, OWPML 구현 라이선스 조건 | 스파이크 테스트, 한컴 문서 확인 |
| E15 | PR #392 방식의 python-docx 1.x 호환성, pandoc-crossref의 DOCX 필드 출력 여부 | 스파이크 테스트 |
| E16 | Word LINK 필드 문법 | Microsoft 1차 문서 |
| D3 | 협회 규정 제2-37조·별표 9의 현행 문언, DCM 기관 자료의 심사필 대상 여부 | 현행 규정 원문과 준법감시 부서 확인 |
| E2 | OOXMLValidatorCLI 배포 형태, Word 복구 대화상자의 COM 탐지 방법 | 저장소 릴리스 확인, 실험 |
| E17 | 2025-11-25 이후 MCP 스펙 개정 여부 | 스펙 사이트 확인 |

## 결론

이번 조사가 바꾼 판단은 두 가지다. 첫째, 파서 교체는 "언젠가 할 리팩터링"이 아니라 Python 하한 결정의 종속 변수다. 3.8을 유지하면 3.8을 지원하는 CommonMark 준수 파서는 2023년판에 멈춘 markdown-it-py 3.0.0뿐이다. 따라서 규칙 기반 파서를 계속 고쳐 쓰는 것이 합리적인 선택으로 남는다. 둘째, HWPX는 2026년 5월을 기점으로 "요청 시 검토"에서 "공공 거래상대방이 기대하는 형식"으로 성격이 바뀌었다. 그래도 결정 순서는 D1 → 형식 독립 계층 분리 → D2다. 순서를 바꾸면 AGENTS.md가 금지한 두 번째 조립 경로가 사실상 생긴다.

실행 순서의 원칙은 "측정 장치를 먼저, 출력 변경은 그다음"이다. E1–E3은 결정 없이 3.8에서 끝낼 수 있다. 이 셋이 있어야 E5–E7의 출력 변경과 E12의 파서 교체가 "무엇이 바뀌었는지"를 증명할 수 있다. 컴플라이언스, Excel 대사, 차트 같은 금융 특화 기능은 모두 "존재와 일치 여부는 점검하되 법적·재무적 적합성은 보증하지 않는다"는 같은 원칙 아래 둔다. 그래야 이 엔진이 뱅커의 검토를 대체하는 도구가 아니라, 검토 가능한 초안을 결정론적으로 만드는 도구로 남는다.
