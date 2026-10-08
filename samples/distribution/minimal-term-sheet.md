---
profile: term-sheet
title: "{{company}} 운전자금대출 주요조건(안)"
subtitle: "Term Sheet"
date: "2026-10-08"
version: "v1.0 / 가상 예시"
house: minimal-house.yaml
terms:
  company: "가나다머티리얼즈㈜"
  amount: "100억원"
  rate: "4.50%"
  first_draw: "2026-10-15"
  maturity: "2028-10-15"
tables:
  - columns: [text, text, text]
    label_columns: 2
  - columns: [text, date, number, number]
    label_columns: 0
    unit: "억원"
    note: "※ 이자·수수료 및 실제 영업일 조정은 포함하지 않은 가상 예시"
    schedule: {repayment: 3, balance: 4, principal: "{{amount}}", total: true}
---

## 1. 주요 금융조건

| 구 분 | << | 내 용 |
|---|---|---|
| 차주 | << | {{company}} |
| 대출 | 약정금액 | {{amount}} |
| ^^ | 적용금리 | 연 {{rate}} (가상 고정금리) |
| ^^ | 실행·만기 | • 실행예정일: {{first_draw}}<br>• 최종만기일: {{maturity}} |
| 자금용도 | << | • 원재료 구매 및 운전자금<br>• 실제 사용범위는 최종 약정에서 확정 |
| 선행조건 | << | • 내부 승인 및 최종 계약 체결<br>• 필요한 서류 제출과 선행조건 충족 |

## 2. 원금 상환일정

아래 원금상환액과 상환 후 잔액은 작성자가 직접 입력한 값입니다. 프로그램은 선언된 검사 규칙에 따라 일치 여부를 확인하며 금액을 자동 계산하여 채우지 않습니다.

| 회차 | 지급예정일 | 원금상환액 | 상환 후 잔액 |
|---|---|---:|---:|
| 1 | 2027-10-15 | 50 | 50 |
| 2 | 2028-10-15 | 50 | 0 |
| 합계 | | 100 | 0 |

※ 모든 기관명, 거래조건, 금액과 일정은 문서 작성 연습용으로 만든 예시입니다.

```confirmation
```
