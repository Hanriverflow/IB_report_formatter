---
profile: ib-report
title: 가상전자 회사채 발행 검토 차트
company: 가상증권
charts: true
layout:
  cover: false
  toc: false
  disclaimer: false
---

# 가상전자 회사채 발행 검토 차트

모든 회사 이름과 수치는 기능 검증용 가상 자료입니다. 실제 기업, 거래, 평가사 자료와 관계없습니다.

## 만기 구조

```chart
chart_type: stacked_bar
title: 2027~2028년에 만기가 몰리며, 2028년 차환 부담이 4,700억원으로 가장 크다
subtitle: "연도별 만기 도래액, 단위: 억원"
labels: ["2026", "2027", "2028", "2029", "2030", 2031 이후]
series:
  - name: 공모 회사채
    values: [1200, 2500, 3100, 900, 1500, 800]
  - name: 사모사채·CP
    values: [600, 400, 700, 300, 0, 0]
  - name: 은행 차입
    values: [300, 500, 900, 200, 400, 0]
totals: true
number_format: ',.0f'
note: 2026년 9월 말 잔액 기준
source: 회사 제시 자료(가상)
```

만기는 2028년에 가장 크다.

## 차입 부담

```chart
chart_type: bar_line
title: 순차입금이 줄면서 순차입금/EBITDA는 3.4배에서 2.1배로 낮아졌다
subtitle: "막대: 순차입금(억원)  선: 순차입금/EBITDA(배)"
labels: ["2022", "2023", "2024", "2025", 2026.9 LTM]
series:
  - name: 순차입금
    values: [4500, 4200, 3900, 3500, 3300]
  - name: 순차입금/EBITDA
    values: [3.4, 3.1, 2.8, 2.4, 2.1]
    axis: right
number_format: ',.0f'
right_number_format: multiple
source: 회사 제시 자료(가상)
```

## 동종사 비교

```chart
chart_type: hbar
title: 가상전자의 회사채 잔액은 AA- 동종 5개사 중 두 번째로 많다
subtitle: "AA- 등급 동종 기업 회사채 잔액(2026년 9월 말), 단위: 억원"
labels: [가상화학, 가상전자, 가상건설, 가상유통, 가상해운]
series:
  - name: 회사채 잔액
    values: [18500, 14200, 9800, 7600, 5100]
highlight: [가상전자]
number_format: ',.0f'
source: 가상 수치
```

## 스프레드 추이

```chart
chart_type: line
title: 가상전자 3년물 스프레드는 2년 사이 18bp 줄어 등급 평균보다 5bp 낮아졌다
subtitle: "국고채 3년 대비 민평 스프레드(월말), 단위: bp"
labels: ["24.10", "24.11", "24.12", "25.01", "25.02", "25.03", "25.04", "25.05", "25.06", "25.07", "25.08", "25.09", "25.10", "25.11", "25.12", "26.01", "26.02", "26.03", "26.04", "26.05", "26.06", "26.07", "26.08", "26.09"]
series:
  - name: AA- 평균
    values: [70, 71, 73, 74, 72, 71, 71, 70, 69, 69, 68, 67, 66, 66, 65, 65, 64, 63, 63, 62, 62, 61, 61, 61]
  - name: 가상전자
    values: [74, 76, 79, 78, 75, 73, 74, 72, 70, 69, 68, 66, 65, 66, 64, 63, 61, 60, 59, 58, 58, 57, 57, 56]
reference:
  - name: 희망 스프레드
    value: 60
source: 가상 수치
```

## 수요예측 결과

```chart
chart_type: bar
title: 수요예측 참여 5,800억원 중 -15bp 이하에 3,700억원이 몰렸다
subtitle: "개별민평 대비 가산금리 구간별 참여금액(3년물, 모집 1,000억원), 단위: 억원"
labels: [-30bp, -25bp, -20bp, -15bp, -10bp, -5bp, 0bp, +5bp, +10bp]
series:
  - name: 참여금액
    values: [300, 600, 1100, 1700, 900, 500, 400, 200, 100]
highlight: [-15bp]
number_format: ',.0f'
note: 강조한 구간이 확정 가산금리(-15bp)
source: 가상 수치
```

확정 가산금리 구간에 참여가 가장 많았다.

## 밸류에이션 범위

```chart
chart_type: range
title: 방법별 주당 가치는 11,800~13,200원 구간에서 겹친다
subtitle: "평가 방법별 주당 가치 범위, 단위: 원"
labels: [52주 최저·최고, 유사회사 PER, 유사거래 EV/EBITDA, "DCF (WACC ±0.5%p)"]
low: [9800, 10500, 11800, 11000]
high: [13200, 14000, 16000, 14500]
reference:
  - name: 현재 주가
    value: 12300
show_overlap: true
number_format: ',.0f'
source: 가상 수치
```

## 현금 증감

```chart
chart_type: waterfall
title: 영업현금으로 차입을 갚고도 기말 현금이 늘었다
subtitle: "단위: 억원"
labels: [기초 현금, 영업현금, 차입 상환, 배당]
series:
  - name: 증감
    values: [1200, 900, -600, -150]
start_total: true
total_label: 기말 현금
number_format: ',.0f'
source: 회사 제시 자료(가상)
```
