# 📉 VIP 도달·이탈 진단 대시보드

VIP DAU 역신장 사유를 추적하기 위한 Streamlit 대시보드입니다. **채널 수신거부(이탈)** 와 **앱 미보유/삭제**로 도달 가능 모수가 얼마나 줄어드는지를 등급·그룹별로 진단합니다. 파일을 업로드하면 바로 시각화됩니다.

## 배경

- VIP(SP·PT·GD·SV·BK) DAU가 전년비 10%+ 역신장
- 가설: 수신거부 + 앱 미보유/삭제로 **실제 도달 가능 모수**가 줄어 DAU에 악영향
- 발송량을 늘려도 도달 모수가 작으면 DAU 기여에 한계

## 핵심 지표

| 지표 | 정의 |
|---|---|
| **앱 미보유/삭제** | `ACT_PUSH_MEM − TOT_PUSH_MEM` (앱푸시 동의했으나 발송 불가) |
| **앱푸시 도달률** | `TOT_PUSH_MEM / ACT_PUSH_MEM` |
| **수신거부율** | `OUT / TOT` |
| **순증감** | `NEW − OUT` |

## 그룹 & 등급 순서

- **등급 순서**(상위→하위): SP · PT · GD · SV · BK · PP · RD
- **그룹** — VIP: SP·PT·GD·SV·BK / 일반: PP·RD

## 대시보드 구성

1. **핵심 진단** — VIP 도달률·미보유/삭제 규모 자동 인사이트 + KPI
2. **그룹 비교** — VIP vs 일반 카드 + 도달률 일별 추이
3. **앱푸시 도달 진단** — 등급별 발송가능 vs 미보유/삭제 + 도달률
4. **수신거부 분석** — 채널별 추이, 등급별 거부율, 등급×채널 히트맵
5. **그룹 내 등급별 인사이트** — 그룹 안에서 등급 간 비교 + 자동 코멘트
6. **상세 테이블** — 등급별 종합 + CSV 다운로드

## 실행 / 배포

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit Cloud: 레포 연결 → Main file `app.py` → Deploy → 앱에서 export 파일(.xls/.xlsx) 업로드

## 데이터 형식

`STD_DD`, `GRADE_CD`, `ACT_MEM`, `ACT_PUSH_MEM`, `TOT_{PUSH/SMS/EMAIL}_MEM`,
`NEW_{PUSH/SMS/EMAIL}_MEM`, `OUT_{PUSH/SMS/EMAIL}_MEM` 컬럼이 첫 시트에 있어야 합니다. (`GRADE_CD='TOTAL'` 합계행과 `SQL` 시트는 자동 무시)
