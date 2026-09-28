---
profile: business-report
title: 가상기업 차트 검증
charts: true
---

# 가상기업 차트 검증

모든 수치는 기능 검증용 가상 자료입니다. 실제 기업이나 거래와 관계없습니다.

## 부문별 매출 비교

```chart
chart_type: bar
title: 가상기업 부문별 매출
labels: [제조, 서비스, 유통]
series:
  - name: 전기
    values: [120, 80, 60]
  - name: 당기
    values: [150, 95, 75]
y_label: 백만원
source: 기능 검증용 가상 자료
number_format: ',.0f'
```

## ---

## 분기별 지표 추이

입력값 12.5는 12.5%로 표시하며 비율을 환산하지 않습니다.

```chart
type: line
title: 가상기업 분기별 이익률
labels: [일분기, 이분기, 삼분기, 사분기]
series:
  - name: 계획
    values: [12.5, 13.0, 13.5, 14.0]
  - name: 실적
    values: [12.0, 13.5, 12.8, 14.2]
unit: 이익률
number_format: percent
source: 기능 검증용 가상 자료
```

## ---

## 현금 증감 분석

0부터 시작해 100, -30, 20을 누적합니다. 합계 막대는 90입니다.

```chart
chart_type: waterfall
title: 가상기업 현금 증감
labels: [기초, 비용, 유입]
series:
  - name: 증감
    values: [100, -30, 20]
y_label: 백만원
total_label: 기말
source: 기능 검증용 가상 자료
```
