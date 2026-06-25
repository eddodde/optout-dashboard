# 📭 등급별·채널별 수신거부 대시보드

회원 등급(GRADE_CD)과 채널(PUSH/SMS/EMAIL)별로 **수신거부(이탈)·신규수신·순증감**을 분석하는 Streamlit 대시보드입니다. 파일을 업로드하면 바로 시각화됩니다.

## 분석 내용

- **핵심 지표**: 총 수신거부, 신규수신, 순증감, 기간 수신거부율
- **채널별 요약**: PUSH / SMS / EMAIL 각각의 거부·신규·순증감·거부율
- **일별 추이**: 채널별 수신거부 추이, 신규 vs 거부 순증감
- **등급 × 채널 매트릭스**: 수신거부 건수 / 거부율 / 순증감 히트맵
- **등급별 분석**: 등급별 신규 vs 거부, 수신거부율 비교
- **상세 테이블**: 등급 × 채널 집계 + CSV 다운로드

## 컬럼 규칙

| 접두어 | 의미 |
|---|---|
| `ACT_` | 유효회원수 |
| `TOT_` | 수신자수 |
| `NEW_` | 신규수신 |
| `OUT_` | 수신거부(이탈) |

지표 정의:
- 수신거부율 = `OUT / TOT`
- 순증감 = `NEW − OUT`
- 도달률 = `TOT / ACT`

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 배포 (Streamlit Cloud)

1. 이 레포를 GitHub에 push
2. [streamlit.io/cloud](https://streamlit.io/cloud) 에서 레포 연결
3. Main file: `app.py` 설정 후 Deploy
4. 앱 화면에서 export 파일(.xls/.xlsx) 업로드

## 데이터 형식

`STD_DD`, `GRADE_CD`, `ACT_MEM`, `ACT_PUSH_MEM`, `TOT_{PUSH/SMS/EMAIL}_MEM`,
`NEW_{PUSH/SMS/EMAIL}_MEM`, `OUT_{PUSH/SMS/EMAIL}_MEM` 컬럼을 포함한 시트가 첫 번째 시트에 있어야 합니다. (`SQL` 시트 등 나머지는 무시)
